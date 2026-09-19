using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Vb6.ActiveX.Hosting;

namespace Vb6.ActiveX.Hosting.Interop
{
  /// <summary>COM <c>IServiceProvider</c>; must be public for <c>Marshal.GetComInterfaceForObject</c>.</summary>
  [EditorBrowsable(EditorBrowsableState.Never)]
  [ComVisible(true)]
  [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IOleServiceProvider
  {
    /// <summary>Returns a service identified by <paramref name="rguidService"/> as <paramref name="riid"/>.</summary>
    [PreserveSig]
    int QueryService(IntPtr rguidService, IntPtr riid, IntPtr ppvObject);
  }

  internal static class OleServiceProviderIid
  {
    internal static readonly Guid Value = new Guid("6D5140C1-7436-11CE-8034-00AA006009FA");
  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  internal sealed class VB6OleServiceProviderShim : IOleServiceProvider
  {
    private const int S_OK = 0;
    private const int E_NOINTERFACE = unchecked((int)0x80004002);
    private const int E_POINTER = unchecked((int)0x80004003);

    private static readonly object Gate = new object();
    private static VB6OleServiceProviderShim? _instance;

    internal static IntPtr AcquirePointer()
    {
      lock (Gate)
      {
        _instance ??= new VB6OleServiceProviderShim();
        return Marshal.GetComInterfaceForObject(_instance, typeof(IOleServiceProvider));
      }
    }

    public int QueryService(IntPtr rguidService, IntPtr riid, IntPtr ppvObject)
    {
      string w = IntPtr.Size == 8 ? "X16" : "X8";
      if (ppvObject == IntPtr.Zero || !VB6OleMemory.IsPointerSlotWritable(ppvObject))
      {
        Debug.WriteLine(
          "VB6OleServiceProviderShim QueryService E_POINTER ppv=0x" + ppvObject.ToString(w)
          + (ppvObject == new IntPtr(-1)
            ? " (caller did not pass a writable void**; see Output)"
            : string.Empty));
        Debug.WriteLine(
          "  raw args (servprov.h: svc, riid, ppv): rguidService=0x" + rguidService.ToString(w)
          + " riid=0x" + riid.ToString(w)
          + " ppv=0x" + ppvObject.ToString(w));

        if (riid != IntPtr.Zero && VB6OleMemory.IsReadable(riid, 16))
        {
          Guid gTry = Marshal.PtrToStructure<Guid>(riid);
          Debug.WriteLine("  decoded riid slot as GUID: " + ComKnownIidNames.Describe(gTry));
        }

        return E_POINTER;
      }

      Marshal.WriteIntPtr(ppvObject, IntPtr.Zero);

      if (rguidService == IntPtr.Zero || riid == IntPtr.Zero
        || !VB6OleMemory.IsReadable(rguidService, 16)
        || !VB6OleMemory.IsReadable(riid, 16))
      {
        Debug.WriteLine("VB6OleServiceProviderShim QueryService E_POINTER (null or unreadable REFGUID/RIID)");
        return E_POINTER;
      }

      Guid gService = Marshal.PtrToStructure<Guid>(rguidService);
      Guid gRiid = Marshal.PtrToStructure<Guid>(riid);
      Debug.WriteLine(
        "VB6OleServiceProviderShim QueryService rguidService="
        + ComKnownIidNames.Describe(gService)
        + " riid="
        + ComKnownIidNames.Describe(gRiid));

      if (MsoComponentManagerIids.IsQueryServicePair(gService, gRiid))
      {
        Marshal.WriteIntPtr(ppvObject, MsoComponentManagerStub.AcquirePointer());
        return S_OK;
      }

      return E_NOINTERFACE;
    }
  }
}
