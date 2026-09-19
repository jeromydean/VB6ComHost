using System;
using System.ComponentModel;
using System.Threading;
using System.Runtime.InteropServices;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>
  /// Minimal STA <c>PeekMessage</c> / <c>DispatchMessage</c> loop. Used internally by <see cref="ComHost"/>.
  /// Prefer <see cref="ComHost"/> unless you are pumping an STA thread yourself.
  /// </summary>
  [EditorBrowsable(EditorBrowsableState.Advanced)]
  public static class VB6StaMessagePump
  {
    private const uint PM_NOREMOVE = 0;
    private const uint PM_REMOVE = 0x0001;
    private const uint WM_QUIT = 0x0012;
    private const uint QS_ALLINPUT = 0x04FF;
    private const uint MWMO_INPUTAVAILABLE = 0x0004;

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
      public int X;
      public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSG
    {
      public IntPtr hwnd;
      public uint message;
      public UIntPtr wParam;
      public IntPtr lParam;
      public uint time;
      public POINT pt;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TranslateMessage(ref MSG lpMsg);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref MSG lpMsg);

    [DllImport("user32.dll")]
    private static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WaitMessage();

    [DllImport("user32.dll")]
    private static extern void PostQuitMessage(int nExitCode);

    [DllImport("user32.dll")]
    private static extern uint MsgWaitForMultipleObjectsEx(
      uint nCount,
      IntPtr pHandles,
      uint dwMilliseconds,
      uint dwWakeMask,
      uint dwFlags);

    [DllImport("user32.dll")]
    private static extern uint MsgWaitForMultipleObjects(
      uint nCount,
      IntPtr[] pHandles,
      [MarshalAs(UnmanagedType.Bool)] bool bWaitAll,
      uint dwMilliseconds,
      uint dwWakeMask);

    internal static bool Peek(out MSG msg, bool remove)
    {
      return PeekMessage(out msg, IntPtr.Zero, 0, 0, remove ? PM_REMOVE : PM_NOREMOVE);
    }

    internal static bool TryGetMessage(out MSG msg)
    {
      return GetMessage(out msg, IntPtr.Zero, 0, 0) != -1;
    }

    internal static bool IsQuit(in MSG msg)
    {
      return msg.message == WM_QUIT;
    }

    internal static void Wait()
    {
      _ = WaitMessage();
    }

    internal static void WaitForInput(uint milliseconds)
    {
      _ = MsgWaitForMultipleObjectsEx(0, IntPtr.Zero, milliseconds, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
    }

    internal static void TranslateAndDispatch(ref MSG msg)
    {
      _ = TranslateMessage(ref msg);
      _ = DispatchMessage(ref msg);
    }

    internal static void RepostQuit(in MSG msg)
    {
      PostQuitMessage((int)msg.wParam.ToUInt32());
    }

    /// <summary>
    /// Runs until <paramref name="cancellationToken"/> is canceled or <c>WM_QUIT</c> is processed.
    /// Must be called from the same STA thread that created the COM object and registered the message filter.
    /// </summary>
    /// <exception cref="InvalidOperationException">Current thread is not STA.</exception>
    public static void Run(CancellationToken cancellationToken = default)
    {
      if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
      {
        throw new InvalidOperationException(
          "VB6StaMessagePump.Run must run on an STA thread. Use [STAThread] on Main or SetApartmentState before COM work.");
      }

      while (true)
      {
        if (cancellationToken.IsCancellationRequested)
        {
          break;
        }

        bool any = false;
        while (PeekMessage(out MSG msg, IntPtr.Zero, 0, 0, PM_REMOVE))
        {
          any = true;
          if (msg.message == WM_QUIT)
          {
            return;
          }

          _ = TranslateMessage(ref msg);
          _ = DispatchMessage(ref msg);
        }

        if (!any)
        {
          SleepWithCancelCheck(cancellationToken, 50);
        }
      }
    }

    private static void SleepWithCancelCheck(CancellationToken cancellationToken, int milliseconds)
    {
      const int stepMs = 10;
      int waited = 0;
      while (waited < milliseconds)
      {
        if (cancellationToken.IsCancellationRequested)
        {
          return;
        }

        int chunk = Math.Min(stepMs, milliseconds - waited);
        Thread.Sleep(chunk);
        waited += chunk;
      }
    }

    /// <summary>
    /// Runs while <paramref name="shouldContinue"/> returns true (evaluated when the queue is empty).
    /// </summary>
    public static void RunWhile(Func<bool> shouldContinue)
    {
      RunWhile(shouldContinue, wakeup: null);
    }

    /// <summary>
    /// Runs while <paramref name="shouldContinue"/> returns true (evaluated when the queue is empty).
    /// When <paramref name="wakeup"/> is signaled, the wait returns so the caller can drain work.
    /// </summary>
    public static void RunWhile(Func<bool> shouldContinue, WaitHandle? wakeup)
    {
      if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
      {
        throw new InvalidOperationException(
          "VB6StaMessagePump.RunWhile must run on an STA thread. Use [STAThread] on Main or SetApartmentState before COM work.");
      }

      if (shouldContinue == null)
      {
        throw new ArgumentNullException(nameof(shouldContinue));
      }

      while (true)
      {
        bool any = false;
        while (PeekMessage(out MSG msg, IntPtr.Zero, 0, 0, PM_REMOVE))
        {
          any = true;
          if (msg.message == WM_QUIT)
          {
            return;
          }

          _ = TranslateMessage(ref msg);
          _ = DispatchMessage(ref msg);
        }

        if (!shouldContinue())
        {
          break;
        }

        if (wakeup != null)
        {
          WaitForInputOrHandle(wakeup, any ? 0u : 50u);
        }
        else
        {
          Thread.Sleep(any ? 0 : 50);
        }
      }
    }

    internal static void WaitForInputOrHandle(WaitHandle handle, uint milliseconds)
    {
      if (handle == null)
      {
        throw new ArgumentNullException(nameof(handle));
      }

      IntPtr[] handles = { handle.SafeWaitHandle.DangerousGetHandle() };
      _ = MsgWaitForMultipleObjects(1, handles, false, milliseconds, QS_ALLINPUT);
    }
  }
}
