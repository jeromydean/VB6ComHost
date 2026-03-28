using System;
using System.Runtime.InteropServices;

namespace VB6ComHost
{
  /// <summary>Office <c>IMsoComponentManager</c>; must be public for <c>Marshal.GetComInterfaceForObject</c>.</summary>
  [ComVisible(true)]
  [Guid("000C060B-0000-0000-C000-000000000046")]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IMsoComponentManager
  {
      [PreserveSig]
      int QueryService([In] ref Guid guidService, [In] ref Guid riid, out IntPtr ppvObj);

      [PreserveSig]
      int FDebugMessage(uint dwReserved, uint message, UIntPtr wParam, UIntPtr lParam);

      [PreserveSig]
      int FRegisterComponent(IntPtr piComponent, IntPtr pcrinfo, IntPtr pdwComponentID);

      [PreserveSig]
      int FRevokeComponent(UIntPtr dwComponentID);

      [PreserveSig]
      int FUpdateComponentRegistration(UIntPtr dwComponentID, IntPtr pcrinfo);

      [PreserveSig]
      int FOnComponentActivate(UIntPtr dwComponentID);

      [PreserveSig]
      int FSetTrackingComponent(UIntPtr dwComponentID, [MarshalAs(UnmanagedType.Bool)] bool fTrack);

      void OnComponentEnterState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude,
          uint dwReserved);

      [PreserveSig]
      int FOnComponentExitState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude);

      [PreserveSig]
      int FInState(uint uStateID, IntPtr pvoid);

      [PreserveSig]
      int FContinueIdle();

      [PreserveSig]
      int FPushMessageLoop(UIntPtr dwComponentID, uint uReason, IntPtr pvLoopData);

      [PreserveSig]
      int FCreateSubComponentManager(IntPtr piunkOuter, IntPtr piunkServProv, [In] ref Guid riid, IntPtr ppvObj);

      [PreserveSig]
      int FGetParentComponentManager(IntPtr ppicm);

      [PreserveSig]
      int FGetActiveComponent(uint dwgac, IntPtr ppic, IntPtr pcrinfo, uint dwReserved);
  }

  internal static class MsoComponentManagerIids
  {
      internal static readonly Guid IidIMsoComponentManager = new Guid("000C060B-0000-0000-C000-000000000046");
      internal static readonly Guid SidSMsoComponentManager = new Guid("000C0601-0000-0000-C000-000000000046");

      internal static bool IsQueryServicePair(Guid guidService, Guid riid)
      {
          bool docOrder = guidService.Equals(IidIMsoComponentManager) && riid.Equals(SidSMsoComponentManager);
          bool swapped = guidService.Equals(SidSMsoComponentManager) && riid.Equals(IidIMsoComponentManager);
          return docOrder || swapped;
      }
  }

  /// <summary>
  /// Minimal Office-style component manager stub. Implements only <see cref="IMsoComponentManager"/> on the CCW;
  /// <c>QueryInterface</c> for <c>SID_SMsoComponentManager</c> is not customized (.NET Standard 2.0 cannot use
  /// <c>Marshal.GetComInterfaceForObject(..., CustomQueryInterfaceMode.Ignore)</c> to implement that without recursion).
  /// </summary>
  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  internal sealed class MsoComponentManagerStub : IMsoComponentManager
  {
      private const int E_NOINTERFACE = unchecked((int)0x80004002);

      public int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObj)
      {
          ppvObj = IntPtr.Zero;
          return E_NOINTERFACE;
      }

      public int FDebugMessage(uint dwReserved, uint message, UIntPtr wParam, UIntPtr lParam) => 1;

      public int FRegisterComponent(IntPtr piComponent, IntPtr pcrinfo, IntPtr pdwComponentID)
      {
          if (pdwComponentID != IntPtr.Zero)
          {
              Marshal.WriteIntPtr(pdwComponentID, new IntPtr(1));
          }

          return 1;
      }

      public int FRevokeComponent(UIntPtr dwComponentID) => 1;

      public int FUpdateComponentRegistration(UIntPtr dwComponentID, IntPtr pcrinfo) => 1;

      public int FOnComponentActivate(UIntPtr dwComponentID) => 1;

      public int FSetTrackingComponent(UIntPtr dwComponentID, bool fTrack) => 1;

      public void OnComponentEnterState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude,
          uint dwReserved)
      {
      }

      public int FOnComponentExitState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude) => 1;

      public int FInState(uint uStateID, IntPtr pvoid) => 0;

      public int FContinueIdle() => 1;

      public int FPushMessageLoop(UIntPtr dwComponentID, uint uReason, IntPtr pvLoopData) => 1;

      public int FCreateSubComponentManager(IntPtr piunkOuter, IntPtr piunkServProv, ref Guid riid, IntPtr ppvObj)
      {
          if (ppvObj != IntPtr.Zero)
          {
              Marshal.WriteIntPtr(ppvObj, IntPtr.Zero);
          }

          return 0;
      }

      public int FGetParentComponentManager(IntPtr ppicm)
      {
          if (ppicm != IntPtr.Zero)
          {
              Marshal.WriteIntPtr(ppicm, IntPtr.Zero);
          }

          return 0;
      }

      public int FGetActiveComponent(uint dwgac, IntPtr ppic, IntPtr pcrinfo, uint dwReserved)
      {
          if (ppic != IntPtr.Zero)
          {
              Marshal.WriteIntPtr(ppic, IntPtr.Zero);
          }

          return 0;
      }
  }

}