using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Vb6.ActiveX.Hosting.Interop;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>
  /// Hosts VB6 ActiveX DLL UI on a dedicated STA thread (KB Q247791 pattern).
  /// <see cref="Open()"/> may be called from STA or MTA; the caller does not need to pump messages.
  /// <see cref="CreateInstance(string)"/> objects are tracked and released on the STA thread
  /// when the host is disposed (or earlier via <see cref="Release"/>). Do not call
  /// <c>Marshal.ReleaseComObject</c> yourself.
  /// </summary>
#if NET8_0_OR_GREATER
  [SupportedOSPlatform("windows")]
#endif
  public sealed class ComHost : IDisposable, IAsyncDisposable
  {
    [ThreadStatic]
    private static ComHost? t_current;

    private readonly ConcurrentQueue<WorkItem> _work = new ConcurrentQueue<WorkItem>();
    private readonly List<object> _instances = new List<object>();
    private readonly AutoResetEvent _wakeup = new AutoResetEvent(false);
    private readonly Thread _staThread;
    private readonly ManualResetEventSlim _started = new ManualResetEventSlim(false);
    private readonly ComHostOptions _options;
    private readonly SynchronizationContext? _eventContext;

    private VB6OleMessageFilterSession? _session;
    private HostWindowTracker? _windows;
    private ExceptionDispatchInfo? _startError;
    private volatile bool _stop;
    private bool _disposed;
    private bool _shutdownStarted;
    private int _shutdownTick;

    private ComHost(ComHostOptions options)
    {
      _options = options;
      _eventContext = options.EventContext ?? SynchronizationContext.Current;
      _staThread = new Thread(ThreadMain)
      {
        Name = string.IsNullOrWhiteSpace(options.ThreadName) ? "Vb6.ActiveX.Hosting" : options.ThreadName,
        IsBackground = true,
      };
      _staThread.SetApartmentState(ApartmentState.STA);
    }

    /// <summary>Starts a dedicated STA thread, registers the OLE message filter, and runs a message pump.</summary>
    public static ComHost Open()
    {
      return Open(options: null, CancellationToken.None);
    }

    /// <summary>Starts a dedicated STA thread using <paramref name="options"/>.</summary>
    public static ComHost Open(ComHostOptions options)
    {
      return Open(options, CancellationToken.None);
    }

    /// <summary>Starts a dedicated STA thread. Cancelling stops startup and tears down the thread.</summary>
    public static ComHost Open(CancellationToken cancellationToken)
    {
      return Open(options: null, cancellationToken);
    }

    /// <summary>Starts a dedicated STA thread using <paramref name="options"/>.</summary>
    public static ComHost Open(ComHostOptions? options, CancellationToken cancellationToken)
    {
      ComHostOptions resolved = options ?? new ComHostOptions();
      ThrowIfUnsupportedProcess(resolved);

      ComHost host = new ComHost(resolved);
      host._staThread.Start(host);
      try
      {
        host._started.Wait(cancellationToken);
      }
      catch (OperationCanceledException)
      {
        host.RequestStop();
        host.JoinStaThread(CancellationToken.None);
        throw;
      }

      if (host._startError != null)
      {
        host._staThread.Join();
        host._startError.Throw();
      }

      host.Log(LogLevel.Debug, "ComHost opened on STA thread '" + host._staThread.Name + "'.");
      return host;
    }

    /// <summary>Raised when a VB6 modal form is shown. Marshaled off the host STA thread.</summary>
    public event EventHandler<HostWindowEventArgs>? ModalWindowOpened;

    /// <summary>Raised when a VB6 modal form is destroyed. Marshaled off the host STA thread.</summary>
    public event EventHandler<HostWindowEventArgs>? ModalWindowClosed;

    /// <summary>Raised when a VB6 modeless form is shown. Marshaled off the host STA thread.</summary>
    public event EventHandler<HostWindowEventArgs>? ModelessWindowOpened;

    /// <summary>Raised when a VB6 modeless form is destroyed. Marshaled off the host STA thread.</summary>
    public event EventHandler<HostWindowEventArgs>? ModelessWindowClosed;

    /// <summary>HRESULT from <c>CoRegisterMessageFilter</c> for the installed filter (diagnostics).</summary>
    public int MessageFilterRegistrationHResult => Session.RegisterHResult;

    /// <summary>
    /// Creates a COM object by ProgID on the host STA thread (late-bound friendly: assign to <c>dynamic</c>).
    /// Method and property calls on the returned object also run on that thread, so
    /// <c>Show vbModal</c> can be closed (same as <see cref="Invoke"/>).
    /// The host owns the RCW until <see cref="Release"/> or <see cref="Dispose()"/>.
    /// Use <see cref="CreateInstance{T}(string)"/> for a raw typed RCW.
    /// </summary>
    public object CreateInstance(string progId)
    {
      return StaComProxy.Wrap(this, Invoke(() => Track(CreateInstanceCore(progId))));
    }

    /// <summary>
    /// Creates a COM object by ProgID and casts it to <typeparamref name="T"/>
    /// (interop class or interface). The raw RCW is returned; call UI methods
    /// through <see cref="Invoke"/> or <see cref="InvokeAsync"/> so modal forms can close.
    /// </summary>
    public T CreateInstance<T>(string progId)
      where T : class
    {
      object instance = Invoke(() => Track(CreateInstanceCore(progId)));
      if (instance is T typed)
      {
        return typed;
      }

      throw new ComHostException(
        "Created COM object could not be cast to " + typeof(T).FullName + ".",
        progId,
        0);
    }

    /// <summary>
    /// Creates a COM object by CLSID on the host STA thread.
    /// Late-bound calls on the returned object run on that thread (see <see cref="CreateInstance(string)"/>).
    /// The host owns the RCW until <see cref="Release"/> or <see cref="Dispose()"/>.
    /// </summary>
    public object CreateInstance(Guid clsid)
    {
      return StaComProxy.Wrap(this, Invoke(() => Track(CreateInstanceCore(clsid))));
    }

    /// <summary>
    /// Releases a <see cref="CreateInstance(string)"/> object on the host STA thread.
    /// Optional: <see cref="Dispose()"/> releases any that are still tracked.
    /// Do not use <paramref name="instance"/> afterward.
    /// </summary>
    public void Release(object instance)
    {
      if (instance == null)
      {
        throw new ArgumentNullException(nameof(instance));
      }

      Invoke(() => UntrackAndRelease(StaComProxy.Unwrap(instance)));
    }

    /// <summary>Tracked VB6 forms (live titles). Safe to call from any thread.</summary>
    public IReadOnlyList<HostWindowInfo> GetWindows()
    {
      return Windows.Snapshot();
    }

    /// <summary>Runs <paramref name="action"/> on the host STA thread. Inline if already on that thread.</summary>
    public void Invoke(Action action, CancellationToken cancellationToken = default)
    {
      if (action == null)
      {
        throw new ArgumentNullException(nameof(action));
      }

      _ = Invoke(() =>
      {
        action();
        return 0;
      }, cancellationToken);
    }

    /// <summary>Runs <paramref name="func"/> on the host STA thread. Inline if already on that thread.</summary>
    public T Invoke<T>(Func<T> func, CancellationToken cancellationToken = default)
    {
      return InvokeAsync(func, cancellationToken).GetAwaiter().GetResult();
    }

    /// <summary>Runs <paramref name="action"/> on the host STA thread.</summary>
    public Task InvokeAsync(Action action, CancellationToken cancellationToken = default)
    {
      if (action == null)
      {
        throw new ArgumentNullException(nameof(action));
      }

      return InvokeAsync(() =>
      {
        action();
        return 0;
      }, cancellationToken);
    }

    /// <summary>Runs <paramref name="func"/> on the host STA thread.</summary>
    public Task<T> InvokeAsync<T>(Func<T> func, CancellationToken cancellationToken = default)
    {
      if (func == null)
      {
        throw new ArgumentNullException(nameof(func));
      }

      ThrowIfDisposed();
      cancellationToken.ThrowIfCancellationRequested();
      if (ReferenceEquals(Thread.CurrentThread, _staThread))
      {
        return Task.FromResult(func());
      }

      TaskCompletionSource<T> tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
      WorkItem item = new WorkItem(() =>
      {
        try
        {
          tcs.TrySetResult(func());
        }
        catch (Exception ex)
        {
          tcs.TrySetException(ex);
        }
      })
      {
        OnAbandoned = () => tcs.TrySetException(new ObjectDisposedException(nameof(ComHost))),
      };

      CancellationTokenRegistration registration = default;
      if (cancellationToken.CanBeCanceled)
      {
        registration = cancellationToken.Register(() =>
        {
          item.Cancelled = true;
          tcs.TrySetCanceled(cancellationToken);
          try
          {
            _wakeup.Set();
          }
          catch (ObjectDisposedException)
          {
          }
        });
      }

      _work.Enqueue(item);
      _wakeup.Set();
      return AwaitAndDisposeRegistration(tcs.Task, registration);
    }

    /// <summary>Posts <c>WM_CLOSE</c> to every tracked VB6 form. Safe to call from a background thread.</summary>
    public void CloseWindows()
    {
      Windows.Close(kind: null);
    }

    /// <summary>Posts <c>WM_CLOSE</c> to tracked modal VB6 forms. Safe to call from a background thread.</summary>
    public void CloseModalWindows()
    {
      Windows.Close(HostWindowKind.Modal);
    }

    /// <summary>Posts <c>WM_CLOSE</c> to tracked modeless VB6 forms. Safe to call from a background thread.</summary>
    public void CloseModelessWindows()
    {
      Windows.Close(HostWindowKind.Modeless);
    }

    /// <summary>Stops the STA thread, closes remaining VB6 forms, and releases tracked RCWs.</summary>
    public void Dispose()
    {
      Dispose(CancellationToken.None);
    }

    /// <summary>Stops the host thread. Cancelling abandons the join if the STA thread does not exit in time.</summary>
    public void Dispose(CancellationToken cancellationToken)
    {
      if (_disposed)
      {
        return;
      }

      _disposed = true;
      RequestStop();
      if (!ReferenceEquals(Thread.CurrentThread, _staThread))
      {
        JoinStaThread(cancellationToken);
      }
    }

    /// <summary>Stops the host off the STA thread so <c>await using</c> does not block the caller.</summary>
    public ValueTask DisposeAsync()
    {
      if (ReferenceEquals(Thread.CurrentThread, _staThread))
      {
        Dispose();
        return default;
      }

      return new ValueTask(Task.Run(() => Dispose(CancellationToken.None)));
    }

    internal static ComHost? Current => t_current;

    internal static void DrainPostedWork()
    {
      t_current?.DrainWork();
    }

    internal static bool TryGetWakeup(out WaitHandle wakeup)
    {
      ComHost? host = t_current;
      if (host == null)
      {
        wakeup = null!;
        return false;
      }

      wakeup = host._wakeup;
      return true;
    }

    internal static void NotifyModalLoopStarted(IMsoComponent? component)
    {
      t_current?.Windows.OnModalLoopStarted(component);
    }

    internal static void NotifyModalLoopPulse()
    {
      t_current?.Windows.EnsureModalEnabled();
    }

    internal void RaiseWindowOpened(HostWindowEventArgs e)
    {
      EventHandler<HostWindowEventArgs>? handler = e.Kind == HostWindowKind.Modal
        ? ModalWindowOpened
        : ModelessWindowOpened;
      Raise(handler, e);
    }

    internal void RaiseWindowClosed(HostWindowEventArgs e)
    {
      EventHandler<HostWindowEventArgs>? handler = e.Kind == HostWindowKind.Modal
        ? ModalWindowClosed
        : ModelessWindowClosed;
      Raise(handler, e);
    }

    internal void Log(LogLevel level, string message)
    {
      ILogger? logger = _options.Logger;
      if (logger != null)
      {
        logger.Log(level, message);
        return;
      }

      Debug.WriteLine("Vb6.ActiveX.Hosting: " + message);
    }

    internal static void ThrowIfUnsupportedProcess(ComHostOptions options)
    {
      if (options.ThrowIf64BitProcess && Environment.Is64BitProcess)
      {
        throw new PlatformNotSupportedException(
          "Vb6.ActiveX.Hosting loads 32-bit VB6 ActiveX DLLs in-process. "
          + "Use an x86 / AnyCPU-prefer-32-bit process, or set ComHostOptions.ThrowIf64BitProcess to false.");
      }
    }

    private VB6OleMessageFilterSession Session => _session ?? throw new InvalidOperationException("ComHost has not started.");

    private HostWindowTracker Windows => _windows ?? throw new InvalidOperationException("ComHost has not started.");

    private static async Task<T> AwaitAndDisposeRegistration<T>(Task<T> task, CancellationTokenRegistration registration)
    {
      try
      {
        return await task.ConfigureAwait(false);
      }
      finally
      {
        registration.Dispose();
      }
    }

    private static void ThreadMain(object? state)
    {
      ComHost host = (ComHost)state!;
      t_current = host;
      try
      {
        host._session = new VB6OleMessageFilterSession();
        host._windows = new HostWindowTracker(host);
      }
      catch (Exception ex)
      {
        host._startError = ExceptionDispatchInfo.Capture(ex);
        host._started.Set();
        host.ReleaseThreadResources();
        return;
      }

      host._started.Set();
      try
      {
        VB6StaMessagePump.RunWhile(
          () =>
          {
            host.DrainWork();
            return !host._stop || host.ContinueShutdownPump();
          },
          host._wakeup);
      }
      finally
      {
        host.ReleaseThreadResources();
      }
    }

    private void DrainWork()
    {
      while (_work.TryDequeue(out WorkItem? item))
      {
        if (item.Cancelled)
        {
          continue;
        }

        item.Action();
      }
    }

    private bool ContinueShutdownPump()
    {
      int timeoutMs = ToTimeoutMilliseconds(_options.ShutdownTimeout);
      if (!_shutdownStarted)
      {
        _shutdownStarted = true;
        _shutdownTick = Environment.TickCount;
        try
        {
          _windows?.Close(kind: null);
        }
        catch (InvalidOperationException)
        {
        }
      }

      if (_windows == null || !_windows.HasTrackedWindows)
      {
        return false;
      }

      if (unchecked(Environment.TickCount - _shutdownTick) >= timeoutMs)
      {
        Log(LogLevel.Warning, "Shutdown timeout reached; destroying remaining VB6 windows.");
        _windows.DestroyTrackedWindows();
        return false;
      }

      return true;
    }

    private void RequestStop()
    {
      _stop = true;
      try
      {
        Windows.Close(kind: null);
      }
      catch (InvalidOperationException)
      {
      }

      try
      {
        _wakeup.Set();
      }
      catch (ObjectDisposedException)
      {
      }
    }

    private void JoinStaThread(CancellationToken cancellationToken)
    {
      if (!_staThread.IsAlive)
      {
        return;
      }

      int timeoutMs = ToTimeoutMilliseconds(_options.ShutdownTimeout) + 3000;
      int started = Environment.TickCount;
      while (_staThread.IsAlive)
      {
        cancellationToken.ThrowIfCancellationRequested();
        if (unchecked(Environment.TickCount - started) >= timeoutMs)
        {
          Log(LogLevel.Warning, "STA thread did not exit within the shutdown timeout.");
          return;
        }

        _staThread.Join(50);
      }
    }

    private void ReleaseThreadResources()
    {
      FailPendingWork();
      ReleaseTrackedInstances();
      ModalWindowOpened = null;
      ModalWindowClosed = null;
      ModelessWindowOpened = null;
      ModelessWindowClosed = null;
      _windows?.Dispose();
      _windows = null;
      _session?.Dispose();
      _session = null;
      MsoComponentManagerStub.Shutdown();
      if (ReferenceEquals(t_current, this))
      {
        t_current = null;
      }

      _wakeup.Dispose();
      _started.Dispose();
    }

    private void FailPendingWork()
    {
      while (_work.TryDequeue(out WorkItem? item))
      {
        item.Cancelled = true;
        item.OnAbandoned?.Invoke();
      }
    }

    private void ThrowIfDisposed()
    {
      if (_disposed || _stop)
      {
        throw new ObjectDisposedException(nameof(ComHost));
      }
    }

    private object Track(object instance)
    {
      lock (_instances)
      {
        _instances.Add(instance);
      }

      return instance;
    }

    private void UntrackAndRelease(object instance)
    {
      lock (_instances)
      {
        _instances.Remove(instance);
      }

      TryFinalRelease(instance);
    }

    private void ReleaseTrackedInstances()
    {
      List<object> copy;
      lock (_instances)
      {
        copy = new List<object>(_instances);
        _instances.Clear();
      }

      foreach (object instance in copy)
      {
        TryFinalRelease(instance);
      }
    }

    private static void TryFinalRelease(object instance)
    {
      if (!Marshal.IsComObject(instance))
      {
        return;
      }

      try
      {
        Marshal.FinalReleaseComObject(instance);
      }
      catch (InvalidComObjectException)
      {
      }
      catch (COMException)
      {
      }
    }

    private static object CreateInstanceCore(string progId)
    {
      if (string.IsNullOrWhiteSpace(progId))
      {
        throw new ArgumentException("ProgID is required.", nameof(progId));
      }

      Type? t = Type.GetTypeFromProgID(progId, throwOnError: false);
      if (t == null)
      {
        throw new ComHostException(
          "ProgID is not registered or is invalid.",
          progId,
          ComHostException.ClassNotRegistered);
      }

      try
      {
        return Activator.CreateInstance(t)
          ?? throw new ComHostException("Activator.CreateInstance returned null.", progId, 0);
      }
      catch (COMException ex)
      {
        throw new ComHostException("Failed to create COM object.", progId, ex.HResult, ex);
      }
    }

    private static object CreateInstanceCore(Guid clsid)
    {
      Type? t = Type.GetTypeFromCLSID(clsid);
      if (t == null)
      {
        throw new ComHostException("CLSID is not registered.", clsid, ComHostException.ClassNotRegistered);
      }

      try
      {
        return Activator.CreateInstance(t)
          ?? throw new ComHostException("Activator.CreateInstance returned null.", clsid, 0);
      }
      catch (COMException ex)
      {
        throw new ComHostException("Failed to create COM object.", clsid, ex.HResult, ex);
      }
    }

    private void Raise(EventHandler<HostWindowEventArgs>? handler, HostWindowEventArgs e)
    {
      if (handler == null)
      {
        return;
      }

      void InvokeHandler()
      {
        try
        {
          handler(this, e);
        }
        catch (Exception ex)
        {
          Log(LogLevel.Error, "Window event handler threw: " + ex);
        }
      }

      if (_eventContext != null)
      {
        _eventContext.Post(_ => InvokeHandler(), null);
        return;
      }

      _ = ThreadPool.QueueUserWorkItem(_ => InvokeHandler());
    }

    private static int ToTimeoutMilliseconds(TimeSpan timeout)
    {
      if (timeout <= TimeSpan.Zero)
      {
        return 0;
      }

      if (timeout.TotalMilliseconds >= int.MaxValue)
      {
        return int.MaxValue;
      }

      return (int)timeout.TotalMilliseconds;
    }

    private sealed class WorkItem
    {
      internal WorkItem(Action action)
      {
        Action = action;
      }

      internal Action Action { get; }

      internal Action? OnAbandoned { get; set; }

      internal volatile bool Cancelled;
    }
  }
}
