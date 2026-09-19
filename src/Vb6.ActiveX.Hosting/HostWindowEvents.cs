using System;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>How a VB6 form is being shown.</summary>
  public enum HostWindowKind
  {
    /// <summary>Non-modal VB6 form (<c>vbModeless</c>).</summary>
    Modeless = 0,

    /// <summary>Modal VB6 form (<c>vbModal</c>).</summary>
    Modal = 1,
  }

  /// <summary>
  /// A VB6 form discovered by <see cref="ComHost"/>. <see cref="Close"/> posts <c>WM_CLOSE</c>
  /// and may be called from any thread.
  /// </summary>
  public sealed class HostWindowEventArgs : EventArgs
  {
    private readonly Action<IntPtr> _close;

    internal HostWindowEventArgs(IntPtr handle, HostWindowKind kind, string title, Action<IntPtr> close)
    {
      Handle = handle;
      Kind = kind;
      Title = title ?? string.Empty;
      _close = close ?? throw new ArgumentNullException(nameof(close));
    }

    /// <summary>Win32 HWND of the VB6 form.</summary>
    public IntPtr Handle { get; }

    /// <summary>Modal vs modeless classification at the time the event was raised.</summary>
    public HostWindowKind Kind { get; }

    /// <summary>Current window caption (may be empty if the title is not set yet).</summary>
    public string Title { get; }

    /// <summary>Asks the form to close (<c>WM_CLOSE</c>). Safe to call from a background thread.</summary>
    public void Close()
    {
      _close(Handle);
    }
  }
}
