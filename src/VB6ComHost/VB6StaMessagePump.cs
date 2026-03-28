using System;
using System.Threading;
using System.Runtime.InteropServices;

namespace VB6ComHost
{
  /// <summary>
  /// Minimal STA <c>PeekMessage</c> / <c>DispatchMessage</c> loop for hosts that are not WinForms or WPF.
  /// VB6 modeless forms still need a real Windows message pump on the same STA thread as the COM object.
  /// </summary>
  public static class VB6StaMessagePump
  {
    private const uint PM_REMOVE = 0x0001;
    private const uint WM_QUIT = 0x0012;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
      public int X;
      public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
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

        Thread.Sleep(any ? 0 : 50);
      }
    }
  }
}
