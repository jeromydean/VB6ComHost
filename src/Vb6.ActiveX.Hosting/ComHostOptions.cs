using System;
using System.Threading;
using Microsoft.Extensions.Logging;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>Optional settings for <see cref="ComHost.Open(ComHostOptions, CancellationToken)"/>.</summary>
  public sealed class ComHostOptions
  {
    /// <summary>How long dispose waits for VB6 forms to close before destroying remaining windows. Default 2 seconds.</summary>
    public TimeSpan ShutdownTimeout { get; set; } = TimeSpan.FromSeconds(2);

    /// <summary>Name of the dedicated STA thread.</summary>
    public string ThreadName { get; set; } = "Vb6.ActiveX.Hosting";

    /// <summary>
    /// Context used to raise window events. If null, <see cref="SynchronizationContext.Current"/>
    /// is captured at <c>Open</c>; if that is also null, events are posted to the thread pool
    /// so handlers do not run on the host STA thread.
    /// </summary>
    public SynchronizationContext? EventContext { get; set; }

    /// <summary>When true (default), <c>Open</c> throws <see cref="PlatformNotSupportedException"/> in a 64-bit process.</summary>
    public bool ThrowIf64BitProcess { get; set; } = true;

    /// <summary>Optional logger for host diagnostics.</summary>
    public ILogger? Logger { get; set; }
  }
}
