using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Vb6.ActiveX.Hosting;

namespace Vb6.ActiveX.Hosting.Interop
{
  /// <summary>COM <c>IMessageFilter</c>; must be public for <c>Marshal.GetComInterfaceForObject</c>.</summary>
  [EditorBrowsable(EditorBrowsableState.Never)]
  [ComVisible(true)]
  [Guid("00000016-0000-0000-C000-000000000046")]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IMessageFilter
  {
      /// <summary>Handles an incoming COM call on this thread.</summary>
      [PreserveSig]
      uint HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, IntPtr lpInterfaceInfo);

      /// <summary>Called when a rejected call may be retried.</summary>
      [PreserveSig]
      uint RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType);

      /// <summary>Called while a COM call is pending and Windows messages arrive.</summary>
      [PreserveSig]
      uint MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType);
  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  internal sealed class VB6OleMessageFilter : IMessageFilter, ICustomQueryInterface
  {
      private const uint SERVERCALL_ISHANDLED = 0;
      private const uint PENDINGMSG_WAITDEFPROCESS = 2;

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
          return PENDINGMSG_WAITDEFPROCESS;
      }
  }

  internal sealed class VB6OleMessageFilterSession : IDisposable
  {
      private VB6OleMessageFilter? _filter;
      private IntPtr _filterIface;
      private IntPtr _previousFilter;

      public VB6OleMessageFilterSession()
      {
          Debug.WriteLine("Vb6.ActiveX.Hosting: registering managed IMessageFilter + IServiceProvider shim.");
          _filter = new VB6OleMessageFilter();
          _filterIface = Marshal.GetComInterfaceForObject(_filter, typeof(IMessageFilter));
          RegisterHResult = Ole32Interop.CoRegisterMessageFilter(_filterIface, out _previousFilter);
      }

      public int RegisterHResult { get; }

      public void Dispose()
      {
          if (_filterIface == IntPtr.Zero)
          {
              return;
          }

          _ = Ole32Interop.CoRegisterMessageFilter(_previousFilter, out IntPtr current);
          if (current != IntPtr.Zero)
          {
              Marshal.Release(current);
          }

          if (_previousFilter != IntPtr.Zero)
          {
              Marshal.Release(_previousFilter);
              _previousFilter = IntPtr.Zero;
          }

          Marshal.Release(_filterIface);
          _filterIface = IntPtr.Zero;
          _filter = null;
      }
  }

}