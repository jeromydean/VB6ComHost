using System;
using System.Globalization;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>Failure while opening the host or creating/activating a VB6 ActiveX object.</summary>
  public class ComHostException : Exception
  {
    /// <summary>HRESULT <c>REGDB_E_CLASSNOTREG</c> (0x80040154).</summary>
    public const int ClassNotRegistered = unchecked((int)0x80040154);

    /// <summary>HRESULT <c>CO_E_CLASSSTRING</c> (0x800401F3).</summary>
    public const int InvalidClassString = unchecked((int)0x800401F3);

    /// <summary>Creates an exception with no extra activation context.</summary>
    public ComHostException()
    {
    }

    /// <summary>Creates an exception with the specified message.</summary>
    /// <param name="message">Error text shown to the caller.</param>
    public ComHostException(string message)
      : base(message)
    {
    }

    /// <summary>Creates an exception with the specified message and inner exception.</summary>
    /// <param name="message">Error text shown to the caller.</param>
    /// <param name="innerException">The exception that caused this failure.</param>
    public ComHostException(string message, Exception innerException)
      : base(message, innerException)
    {
    }

    /// <summary>Creates an exception for a failed ProgID activation.</summary>
    /// <param name="message">Error text shown to the caller.</param>
    /// <param name="progId">ProgID passed to <see cref="ComHost.CreateInstance(string)"/>.</param>
    /// <param name="hresult">COM HRESULT, or 0 if none.</param>
    /// <param name="innerException">Optional inner exception.</param>
    public ComHostException(string message, string? progId, int hresult, Exception? innerException = null)
      : base(Format(message, progId, clsid: null, hresult), innerException)
    {
      ProgId = progId;
      if (hresult != 0)
      {
        HResult = hresult;
      }
    }

    /// <summary>Creates an exception for a failed CLSID activation.</summary>
    /// <param name="message">Error text shown to the caller.</param>
    /// <param name="clsid">CLSID passed to <see cref="ComHost.CreateInstance(Guid)"/>.</param>
    /// <param name="hresult">COM HRESULT, or 0 if none.</param>
    /// <param name="innerException">Optional inner exception.</param>
    public ComHostException(string message, Guid clsid, int hresult, Exception? innerException = null)
      : base(Format(message, progId: null, clsid, hresult), innerException)
    {
      Clsid = clsid;
      if (hresult != 0)
      {
        HResult = hresult;
      }
    }

    /// <summary>ProgID passed to <see cref="ComHost.CreateInstance(string)"/>, when applicable.</summary>
    public string? ProgId { get; }

    /// <summary>CLSID passed to <see cref="ComHost.CreateInstance(Guid)"/>, when applicable.</summary>
    public Guid? Clsid { get; }

    private static string Format(string message, string? progId, Guid? clsid, int hresult)
    {
      string suffix = string.Empty;
      if (!string.IsNullOrEmpty(progId))
      {
        suffix += " ProgID '" + progId + "'.";
      }

      if (clsid.HasValue)
      {
        suffix += " CLSID " + clsid.Value.ToString("B") + ".";
      }

      if (hresult != 0)
      {
        suffix += " HRESULT 0x" + hresult.ToString("X8", CultureInfo.InvariantCulture) + ".";
      }

      if (hresult == ClassNotRegistered || hresult == InvalidClassString)
      {
        suffix += " Register the 32-bit ActiveX DLL (regsvr32 from SysWOW64) or use registration-free COM.";
      }

      return string.IsNullOrEmpty(suffix) ? message : message + suffix;
    }
  }
}
