using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VB6ComHost
{
  /// <summary>COM <c>IMessageFilter</c>; must be public for <c>Marshal.GetComInterfaceForObject</c>.</summary>
  [ComVisible(true)]
  [Guid("00000016-0000-0000-C000-000000000046")]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IMessageFilter
  {
      [PreserveSig]
      uint HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, IntPtr lpInterfaceInfo);

      [PreserveSig]
      uint RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType);

      [PreserveSig]
      uint MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType);
  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  internal sealed class VB6OleMessageFilter : IMessageFilter, ICustomQueryInterface
  {
      private const uint SERVERCALL_ISHANDLED = 0;
      private const uint PENDINGMSG_CANCELCALL = 0;

      CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out IntPtr ppv)
      {
          ppv = IntPtr.Zero;
          Debug.WriteLine("VB6OleMessageFilter QueryInterface " + ComKnownIidNames.Describe(iid));
          if (iid.Equals(OleServiceProviderIid.Value))
          {
              ppv = VB6OleServiceProviderShim.AcquirePointer();
              Debug.WriteLine("VB6OleMessageFilter: IServiceProvider QI -> Handled (VB6OleServiceProviderShim)");
              return CustomQueryInterfaceResult.Handled;
          }

          return CustomQueryInterfaceResult.NotHandled;
      }

      public uint HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, IntPtr lpInterfaceInfo)
      {
          return SERVERCALL_ISHANDLED;
      }

      public uint RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType)
      {
          return uint.MaxValue;
      }

      public uint MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType)
      {
          return PENDINGMSG_CANCELCALL;
      }
  }

  internal sealed class VB6OleMessageFilterSession : IDisposable
  {
      private IntPtr _filterIface;
      private IntPtr _previousFilter;

      public VB6OleMessageFilterSession()
      {
          Debug.WriteLine("VB6ComHost: registering managed IMessageFilter + IServiceProvider shim.");
          VB6OleMessageFilter filter = new VB6OleMessageFilter();
          _filterIface = Marshal.GetComInterfaceForObject(filter, typeof(IMessageFilter));
          RegisterHResult = Ole32Interop.CoRegisterMessageFilter(_filterIface, out _previousFilter);
      }

      public int RegisterHResult { get; }

      public void Dispose()
      {
          if (_filterIface == IntPtr.Zero)
          {
              return;
          }

          _ = Ole32Interop.CoRegisterMessageFilter(_previousFilter, out _);
          Marshal.Release(_filterIface);
          _filterIface = IntPtr.Zero;
      }
  }

}