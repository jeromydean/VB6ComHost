using System;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>Snapshot of a tracked VB6 form. <see cref="Title"/> is read at snapshot time.</summary>
  public sealed class HostWindowInfo
  {
    private readonly Action<IntPtr> _close;

    internal HostWindowInfo(IntPtr handle, HostWindowKind kind, string title, Action<IntPtr> close)
    {
      Handle = handle;
      Kind = kind;
      Title = title ?? string.Empty;
      _close = close ?? throw new ArgumentNullException(nameof(close));
    }

    /// <summary>Win32 HWND of the VB6 form.</summary>
    public IntPtr Handle { get; }

    /// <summary>Modal vs modeless classification at snapshot time.</summary>
    public HostWindowKind Kind { get; }

    /// <summary>Window caption captured when the snapshot was taken.</summary>
    public string Title { get; }

    /// <summary>Posts <c>WM_CLOSE</c>. Safe to call from any thread.</summary>
    public void Close()
    {
      _close(Handle);
    }
  }
}
