using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Vb6.ActiveX.Hosting.Interop;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>
  /// Thread-local CBT hook plus a message-only sink. Form creation is recorded immediately;
  /// modal vs modeless is decided after VB6 either enters <c>FPushMessageLoop</c> or the
  /// posted classify message is dispatched (modeless <c>Show</c> has already returned).
  /// </summary>
  internal sealed class HostWindowTracker : IDisposable
  {
    private const int WH_CBT = 5;
    private const int HCBT_CREATEWND = 3;
    private const int HCBT_DESTROYWND = 4;
    private const int HWND_MESSAGE = -3;
    private const uint WM_APP = 0x8000;
    private const uint WM_CLASSIFY = WM_APP + 0x6C01;
    private const uint WM_CLOSE = 0x0010;
    private const uint WM_DESTROY = 0x0002;
    private const int MsocWindowFrameToplevel = 0;
    private const int MsocWindowComponent = 2;

    private readonly string _sinkClassName = "Vb6.ActiveX.Hosting.WindowSink." + Guid.NewGuid().ToString("N");
    private readonly ComHost _host;
    private readonly CbtProc _cbtProc;
    private readonly WndProc _wndProc;
    private readonly HashSet<IntPtr> _pending = new HashSet<IntPtr>();
    private readonly Dictionary<IntPtr, HostWindowEventArgs> _open = new Dictionary<IntPtr, HostWindowEventArgs>();
    private readonly object _gate = new object();

    private IntPtr _hook;
    private IntPtr _sink;
    private IntPtr _sinkModule;
    private bool _classRegistered;
    private bool _disposed;

    private delegate IntPtr CbtProc(int nCode, IntPtr wParam, IntPtr lParam);

    private delegate IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    internal HostWindowTracker(ComHost host)
    {
      _host = host ?? throw new ArgumentNullException(nameof(host));
      _cbtProc = OnCbt;
      _wndProc = OnSink;
      _sink = CreateSinkWindow();
      _hook = SetWindowsHookEx(WH_CBT, _cbtProc, IntPtr.Zero, GetCurrentThreadId());
      if (_hook == IntPtr.Zero)
      {
        Debug.WriteLine("HostWindowTracker: SetWindowsHookEx failed, err=" + Marshal.GetLastWin32Error());
      }
    }

    internal void OnModalLoopStarted(IMsoComponent? component)
    {
      if (_disposed)
      {
        return;
      }

      IntPtr hwnd = TryGetComponentWindow(component);
      if (hwnd == IntPtr.Zero || !IsVb6Form(hwnd))
      {
        hwnd = GetActiveWindow();
      }

      if (hwnd == IntPtr.Zero || !IsVb6Form(hwnd))
      {
        hwnd = SinglePendingOrZero();
      }

      if (hwnd != IntPtr.Zero)
      {
        PromoteToModal(hwnd);
      }
    }

    internal void Close(HostWindowKind? kind)
    {
      List<IntPtr> handles = new List<IntPtr>();
      lock (_gate)
      {
        foreach (KeyValuePair<IntPtr, HostWindowEventArgs> pair in _open)
        {
          if (kind == null || pair.Value.Kind == kind.Value)
          {
            handles.Add(pair.Key);
          }
        }

        if (kind == null || kind.Value == HostWindowKind.Modeless)
        {
          foreach (IntPtr pending in _pending)
          {
            handles.Add(pending);
          }
        }
      }

      foreach (IntPtr hwnd in handles)
      {
        PostClose(hwnd);
      }
    }

    internal static void PostClose(IntPtr hwnd)
    {
      if (hwnd != IntPtr.Zero && IsWindow(hwnd))
      {
        _ = PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
      }
    }

    internal IReadOnlyList<HostWindowInfo> Snapshot()
    {
      List<HostWindowInfo> list = new List<HostWindowInfo>();
      lock (_gate)
      {
        foreach (KeyValuePair<IntPtr, HostWindowEventArgs> pair in _open)
        {
          if (!IsWindow(pair.Key))
          {
            continue;
          }

          list.Add(new HostWindowInfo(pair.Key, pair.Value.Kind, ReadTitle(pair.Key), PostClose));
        }

        foreach (IntPtr hwnd in _pending)
        {
          if (!IsWindow(hwnd))
          {
            continue;
          }

          list.Add(new HostWindowInfo(hwnd, HostWindowKind.Modeless, ReadTitle(hwnd), PostClose));
        }
      }

      return list;
    }

    internal bool HasTrackedWindows
    {
      get
      {
        lock (_gate)
        {
          return _open.Count > 0 || _pending.Count > 0;
        }
      }
    }

    internal void DestroyTrackedWindows()
    {
      List<IntPtr> handles;
      lock (_gate)
      {
        handles = new List<IntPtr>(_open.Count + _pending.Count);
        handles.AddRange(_open.Keys);
        handles.AddRange(_pending);
        _open.Clear();
        _pending.Clear();
      }

      foreach (IntPtr hwnd in handles)
      {
        if (IsWindow(hwnd))
        {
          _ = DestroyWindow(hwnd);
        }
      }
    }

    public void Dispose()
    {
      if (_disposed)
      {
        return;
      }

      _disposed = true;
      if (_hook != IntPtr.Zero)
      {
        _ = UnhookWindowsHookEx(_hook);
        _hook = IntPtr.Zero;
      }

      if (_sink != IntPtr.Zero)
      {
        _ = DestroyWindow(_sink);
        _sink = IntPtr.Zero;
      }

      if (_classRegistered)
      {
        _ = UnregisterClass(_sinkClassName, _sinkModule);
        _classRegistered = false;
      }

      lock (_gate)
      {
        _open.Clear();
        _pending.Clear();
      }
    }

    private IntPtr OnCbt(int nCode, IntPtr wParam, IntPtr lParam)
    {
      if (nCode >= 0 && !_disposed)
      {
        if (nCode == HCBT_CREATEWND && IsVb6Form(wParam))
        {
          lock (_gate)
          {
            _pending.Add(wParam);
          }

          if (_sink != IntPtr.Zero)
          {
            _ = PostMessage(_sink, WM_CLASSIFY, wParam, IntPtr.Zero);
          }
        }
        else if (nCode == HCBT_DESTROYWND)
        {
          OnDestroyed(wParam);
        }
      }

      return CallNextHookEx(_hook, nCode, wParam, lParam);
    }

    private IntPtr OnSink(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
      if (msg == WM_CLASSIFY)
      {
        ClassifyModeless(wParam);
        return IntPtr.Zero;
      }

      if (msg == WM_DESTROY)
      {
        return DefWindowProc(hWnd, msg, wParam, lParam);
      }

      return DefWindowProc(hWnd, msg, wParam, lParam);
    }

    private void ClassifyModeless(IntPtr hwnd)
    {
      if (_disposed || hwnd == IntPtr.Zero || !IsWindow(hwnd) || !IsVb6Form(hwnd))
      {
        lock (_gate)
        {
          _pending.Remove(hwnd);
        }

        return;
      }

      HostWindowEventArgs? args = null;
      lock (_gate)
      {
        if (!_pending.Remove(hwnd) || _open.ContainsKey(hwnd))
        {
          return;
        }

        args = new HostWindowEventArgs(hwnd, HostWindowKind.Modeless, ReadTitle(hwnd), PostClose);
        _open[hwnd] = args;
      }

      _host.RaiseWindowOpened(args);
    }

    private void PromoteToModal(IntPtr hwnd)
    {
      if (_disposed || hwnd == IntPtr.Zero || !IsWindow(hwnd))
      {
        return;
      }

      HostWindowEventArgs? opened = null;
      lock (_gate)
      {
        _pending.Remove(hwnd);
        if (_open.TryGetValue(hwnd, out HostWindowEventArgs? existing))
        {
          if (existing.Kind == HostWindowKind.Modal)
          {
            return;
          }

          _open.Remove(hwnd);
        }

        opened = new HostWindowEventArgs(hwnd, HostWindowKind.Modal, ReadTitle(hwnd), PostClose);
        _open[hwnd] = opened;
      }

      _host.RaiseWindowOpened(opened);
    }

    private void OnDestroyed(IntPtr hwnd)
    {
      HostWindowEventArgs? args;
      lock (_gate)
      {
        _pending.Remove(hwnd);
        if (!_open.TryGetValue(hwnd, out args))
        {
          return;
        }

        _open.Remove(hwnd);
      }

      _host.RaiseWindowClosed(args);
    }

    private IntPtr SinglePendingOrZero()
    {
      lock (_gate)
      {
        if (_pending.Count != 1)
        {
          return IntPtr.Zero;
        }

        foreach (IntPtr hwnd in _pending)
        {
          return hwnd;
        }
      }

      return IntPtr.Zero;
    }

    private static IntPtr TryGetComponentWindow(IMsoComponent? component)
    {
      if (component == null)
      {
        return IntPtr.Zero;
      }

      try
      {
        IntPtr hwnd = component.HwndGetWindow(MsocWindowFrameToplevel, 0);
        if (hwnd != IntPtr.Zero)
        {
          return hwnd;
        }

        return component.HwndGetWindow(MsocWindowComponent, 0);
      }
      catch (Exception ex)
      {
        Debug.WriteLine("HostWindowTracker.HwndGetWindow: " + ex.Message);
        return IntPtr.Zero;
      }
    }

    private IntPtr CreateSinkWindow()
    {
      _sinkModule = GetModuleHandle(null);
      WndClass wc = new WndClass
      {
        lpfnWndProc = _wndProc,
        hInstance = _sinkModule,
        lpszClassName = _sinkClassName,
      };

      ushort atom = RegisterClass(ref wc);
      if (atom == 0)
      {
        Debug.WriteLine("HostWindowTracker: RegisterClass failed, err=" + Marshal.GetLastWin32Error());
        return IntPtr.Zero;
      }

      _classRegistered = true;
      IntPtr hwnd = CreateWindowEx(0, _sinkClassName, string.Empty, 0, 0, 0, 0, 0, new IntPtr(HWND_MESSAGE), IntPtr.Zero, wc.hInstance, IntPtr.Zero);
      if (hwnd == IntPtr.Zero)
      {
        Debug.WriteLine("HostWindowTracker: CreateWindowEx failed, err=" + Marshal.GetLastWin32Error());
        _ = UnregisterClass(_sinkClassName, _sinkModule);
        _classRegistered = false;
      }

      return hwnd;
    }

    internal static bool IsVb6Form(IntPtr hwnd)
    {
      if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
      {
        return false;
      }

      string className = ReadClassName(hwnd);
      return className.StartsWith("ThunderRT6Form", StringComparison.Ordinal)
        || className.StartsWith("ThunderForm", StringComparison.Ordinal)
        || className.Equals("ThunderRT6MDIForm", StringComparison.Ordinal)
        || className.Equals("ThunderMDIForm", StringComparison.Ordinal);
    }

    private static string ReadClassName(IntPtr hwnd)
    {
      StringBuilder sb = new StringBuilder(256);
      _ = GetClassName(hwnd, sb, sb.Capacity);
      return sb.ToString();
    }

    private static string ReadTitle(IntPtr hwnd)
    {
      int len = GetWindowTextLength(hwnd);
      if (len <= 0)
      {
        return string.Empty;
      }

      StringBuilder sb = new StringBuilder(len + 1);
      _ = GetWindowText(hwnd, sb, sb.Capacity);
      return sb.ToString();
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WndClass
    {
      public uint style;
      public WndProc lpfnWndProc;
      public int cbClsExtra;
      public int cbWndExtra;
      public IntPtr hInstance;
      public IntPtr hIcon;
      public IntPtr hCursor;
      public IntPtr hbrBackground;
      public string? lpszMenuName;
      public string lpszClassName;
    }

    [DllImport("user32.dll", EntryPoint = "SetWindowsHookExW", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, CbtProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegisterClassW", SetLastError = true)]
    private static extern ushort RegisterClass(ref WndClass lpWndClass);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "UnregisterClassW", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterClass(string lpClassName, IntPtr hInstance);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "CreateWindowExW", SetLastError = true)]
    private static extern IntPtr CreateWindowEx(
      int dwExStyle,
      string lpClassName,
      string lpWindowName,
      int dwStyle,
      int x,
      int y,
      int nWidth,
      int nHeight,
      IntPtr hWndParent,
      IntPtr hMenu,
      IntPtr hInstance,
      IntPtr lpParam);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetActiveWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassNameW")]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowTextW")]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowTextLengthW")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);
  }
}
