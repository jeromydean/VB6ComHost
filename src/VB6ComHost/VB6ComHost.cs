using System;
using System.Threading;

namespace VB6ComHost
{
  /// <summary>
  /// Registers a pure managed STA OLE message filter suitable for VB6 ActiveX DLL modeless UI (KB Q247791 pattern).
  /// Dispose restores the previous filter. The host must still run a Windows message pump while modeless forms are shown (WinForms, WPF, or <see cref="VB6StaMessagePump"/>).
  /// </summary>
  public sealed class VB6ComHost : IDisposable
  {
    private readonly VB6OleMessageFilterSession _session;

    private VB6ComHost(VB6OleMessageFilterSession session)
    {
      _session = session;
    }

    /// <summary>Creates a host on the current thread. Requires STA.</summary>
    /// <exception cref="InvalidOperationException">Current thread is not STA.</exception>
    public static VB6ComHost Open()
    {
      if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
      {
        throw new InvalidOperationException(
          "VB6ComHost requires an STA thread. Use [STAThread] on Main or Thread.SetApartmentState(ApartmentState.STA) before Open.");
      }

      return new VB6ComHost(new VB6OleMessageFilterSession());
    }

    /// <summary>HRESULT from <c>CoRegisterMessageFilter</c> for the installed filter (diagnostics).</summary>
    public int MessageFilterRegistrationHResult => _session.RegisterHResult;

    /// <summary>Creates a COM object by ProgID (late-bound friendly: assign to <c>dynamic</c>).</summary>
    public object CreateInstance(string progId)
    {
      Type? t = Type.GetTypeFromProgID(progId, throwOnError: false);
      if (t == null)
      {
        throw new ArgumentException("ProgID is not registered or is invalid: " + progId, nameof(progId));
      }

      return Activator.CreateInstance(t) ?? throw new InvalidOperationException("Activator.CreateInstance returned null.");
    }

    /// <summary>Creates a COM object by CLSID.</summary>
    public object CreateInstance(Guid clsid)
    {
      Type t = Type.GetTypeFromCLSID(clsid);
      return Activator.CreateInstance(t) ?? throw new InvalidOperationException("Activator.CreateInstance returned null.");
    }

    public void Dispose()
    {
      _session.Dispose();
    }
  }
}
