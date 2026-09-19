using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Vb6.ActiveX.Hosting;

namespace Vb6.ActiveX.Hosting.Interop
{
  /// <summary>Office <c>IMsoComponentManager</c>; must be public for <c>Marshal.GetComInterfaceForObject</c>.</summary>
  [EditorBrowsable(EditorBrowsableState.Never)]
  [ComVisible(true)]
  [Guid("000C060B-0000-0000-C000-000000000046")]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IMsoComponentManager
  {
      /// <summary>Returns a service identified by <paramref name="guidService"/> as <paramref name="riid"/>.</summary>
      [PreserveSig]
      int QueryService([In] ref Guid guidService, [In] ref Guid riid, out IntPtr ppvObj);

      /// <summary>Optional debug hook; unused by this host.</summary>
      [PreserveSig]
      int FDebugMessage(uint dwReserved, uint message, UIntPtr wParam, UIntPtr lParam);

      /// <summary>Registers an <see cref="IMsoComponent"/> (VB6 registers before a modal pump).</summary>
      [PreserveSig]
      int FRegisterComponent(IntPtr piComponent, IntPtr pcrinfo, IntPtr pdwComponentID);

      /// <summary>Unregisters a previously registered component.</summary>
      [PreserveSig]
      int FRevokeComponent(UIntPtr dwComponentID);

      /// <summary>Updates registration data for a component.</summary>
      [PreserveSig]
      int FUpdateComponentRegistration(UIntPtr dwComponentID, IntPtr pcrinfo);

      /// <summary>Notifies the manager that a component became active.</summary>
      [PreserveSig]
      int FOnComponentActivate(UIntPtr dwComponentID);

      /// <summary>Marks a component as the tracking (mouse-capture) component.</summary>
      [PreserveSig]
      int FSetTrackingComponent(UIntPtr dwComponentID, [MarshalAs(UnmanagedType.Bool)] bool fTrack);

      /// <summary>Broadcasts that a component entered a modal or busy state.</summary>
      void OnComponentEnterState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude,
          uint dwReserved);

      /// <summary>Broadcasts that a component left a modal or busy state.</summary>
      [PreserveSig]
      int FOnComponentExitState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude);

      /// <summary>Returns whether the manager is currently in <paramref name="uStateID"/>.</summary>
      [PreserveSig]
      int FInState(uint uStateID, IntPtr pvoid);

      /// <summary>Returns whether idle processing should continue.</summary>
      [PreserveSig]
      int FContinueIdle();

      /// <summary>Runs a nested message loop until the registered component says the modal form is gone.</summary>
      [PreserveSig]
      int FPushMessageLoop(UIntPtr dwComponentID, uint uReason, IntPtr pvLoopData);

      /// <summary>Creates a nested component manager. Not used by this host.</summary>
      [PreserveSig]
      int FCreateSubComponentManager(IntPtr piunkOuter, IntPtr piunkServProv, [In] ref Guid riid, IntPtr ppvObj);

      /// <summary>Returns the parent component manager. Not used by this host.</summary>
      [PreserveSig]
      int FGetParentComponentManager(IntPtr ppicm);

      /// <summary>Returns the active or tracking component.</summary>
      [PreserveSig]
      int FGetActiveComponent(uint dwgac, IntPtr ppic, IntPtr pcrinfo, uint dwReserved);
  }

  /// <summary>Office <c>IMsoComponent</c>; must be public for RCW / <c>Marshal.GetComInterfaceForObject</c>.</summary>
  [EditorBrowsable(EditorBrowsableState.Never)]
  [ComImport]
  [Guid("000C0600-0000-0000-C000-000000000046")]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IMsoComponent
  {
      /// <summary>Optional debug hook implemented by the component.</summary>
      [PreserveSig]
      int FDebugMessage(IntPtr dwReserved, uint message, UIntPtr wParam, IntPtr lParam);

      /// <summary>Gives the component a chance to translate a message before dispatch.</summary>
      [PreserveSig]
      int FPreTranslateMessage(IntPtr pMsg);

      /// <summary>Notifies the component that a host state (modal, busy) was entered or left.</summary>
      void OnEnterState(uint uStateID, [MarshalAs(UnmanagedType.Bool)] bool fEnter);

      /// <summary>Notifies the component that the application was activated or deactivated.</summary>
      void OnAppActivate([MarshalAs(UnmanagedType.Bool)] bool fActive, uint dwOtherThreadID);

      /// <summary>Notifies the component that it is no longer the active component.</summary>
      void OnLoseActivation();

      /// <summary>Notifies the component of an activation change among registered components.</summary>
      void OnActivationChange(
          IntPtr pic,
          [MarshalAs(UnmanagedType.Bool)] bool fSameComponent,
          IntPtr pcrinfo,
          [MarshalAs(UnmanagedType.Bool)] bool fHostIsActivating,
          IntPtr pchostinfo,
          uint dwReserved);

      /// <summary>Performs idle work. Return nonzero to be called again.</summary>
      [PreserveSig]
      int FDoIdle(uint grfidlef);

      /// <summary>Returns nonzero while the component wants the nested message loop to continue.</summary>
      [PreserveSig]
      int FContinueMessageLoop(uint uReason, IntPtr pvLoopData, IntPtr pMsgPeeked);

      /// <summary>Asks whether the component can terminate.</summary>
      [PreserveSig]
      int FQueryTerminate([MarshalAs(UnmanagedType.Bool)] bool fPromptUser);

      /// <summary>Tells the component to terminate.</summary>
      void Terminate();

      /// <summary>Returns a window handle requested by the manager.</summary>
      IntPtr HwndGetWindow(uint dwWhich, uint dwReserved);
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
  /// Office-style component manager. VB6 modal <c>Show</c> from an ActiveX DLL calls
  /// <see cref="IMsoComponentManager.FPushMessageLoop"/>; that method must run a nested
  /// STA pump (peek, <c>FContinueMessageLoop</c>, get, dispatch) until the component
  /// says the modal form is gone. Returning immediately makes <c>Show vbModal</c> return
  /// to the host while the form is still up, so another "modal" can open.
  /// </summary>
  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  internal sealed class MsoComponentManagerStub : IMsoComponentManager
  {
      private const int E_NOINTERFACE = unchecked((int)0x80004002);
      private const uint MsoContextAll = 0;
      private const uint MsoContextMine = 1;
      private const uint MsoLoopDoEvents = 2;
      private const uint MsoLoopModalForm = 4;
      private const uint MsoLoopModalAlert = 5;
      private const uint MsoLoopDoEventsModal = 6;
      private const uint MsoLoopMain = 0;
      private const uint MsoIdleAll = 0xFFFFFFFFu;
      private const uint MsoGacActive = 0;
      private const uint MsoGacTracking = 1;
      private const uint MsoGacTrackingOrActive = 2;

      [ThreadStatic]
      private static MsoComponentManagerStub? t_instance;

      private readonly Dictionary<UIntPtr, IMsoComponent> _components = new Dictionary<UIntPtr, IMsoComponent>();
      private readonly int _msgSize = Marshal.SizeOf(typeof(VB6StaMessagePump.MSG));
      private IntPtr _msgScratch;
      private int _nextCookie;
      private IMsoComponent? _active;
      private IMsoComponent? _tracking;
      private uint _currentState;
      private bool _disposed;

      internal static IntPtr AcquirePointer()
      {
          t_instance ??= new MsoComponentManagerStub();
          if (t_instance._disposed)
          {
              t_instance = new MsoComponentManagerStub();
          }

          return Marshal.GetComInterfaceForObject(t_instance, typeof(IMsoComponentManager));
      }

      internal static void Shutdown()
      {
          t_instance?.ReleaseNative();
          t_instance = null;
      }

      public int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObj)
      {
          ppvObj = IntPtr.Zero;
          return E_NOINTERFACE;
      }

      public int FDebugMessage(uint dwReserved, uint message, UIntPtr wParam, UIntPtr lParam) => 1;

      public int FRegisterComponent(IntPtr piComponent, IntPtr pcrinfo, IntPtr pdwComponentID)
      {
          if (_disposed || piComponent == IntPtr.Zero || pdwComponentID == IntPtr.Zero)
          {
              return 0;
          }

          IMsoComponent component;
          try
          {
              component = (IMsoComponent)Marshal.GetObjectForIUnknown(piComponent);
          }
          catch (Exception ex)
          {
              Debug.WriteLine("IMsoComponentManager.FRegisterComponent: " + ex.Message);
              return 0;
          }

          _nextCookie++;
          UIntPtr id = new UIntPtr(unchecked((ulong)(uint)_nextCookie));
          _components[id] = component;
          Marshal.WriteIntPtr(pdwComponentID, new IntPtr(_nextCookie));
          Debug.WriteLine("IMsoComponentManager.FRegisterComponent id=" + _nextCookie);
          return 1;
      }

      public int FRevokeComponent(UIntPtr dwComponentID)
      {
          if (!_components.TryGetValue(dwComponentID, out IMsoComponent? component))
          {
              return 0;
          }

          if (ReferenceEquals(_active, component))
          {
              _active = null;
          }

          if (ReferenceEquals(_tracking, component))
          {
              _tracking = null;
          }

          _components.Remove(dwComponentID);
          ReleaseComponent(component);
          return 1;
      }

      public int FUpdateComponentRegistration(UIntPtr dwComponentID, IntPtr pcrinfo)
      {
          return _components.ContainsKey(dwComponentID) ? 1 : 0;
      }

      public int FOnComponentActivate(UIntPtr dwComponentID)
      {
          if (!_components.TryGetValue(dwComponentID, out IMsoComponent? component))
          {
              return 0;
          }

          _active = component;
          return 1;
      }

      public int FSetTrackingComponent(UIntPtr dwComponentID, bool fTrack)
      {
          if (!_components.TryGetValue(dwComponentID, out IMsoComponent? component))
          {
              return 0;
          }

          bool isTracking = ReferenceEquals(_tracking, component);
          if (isTracking == fTrack)
          {
              return 0;
          }

          _tracking = fTrack ? component : null;
          return 1;
      }

      public void OnComponentEnterState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude,
          uint dwReserved)
      {
          _currentState = uStateID;
          if (uContext == MsoContextAll || uContext == MsoContextMine)
          {
              NotifyEnterState(uStateID, enter: true);
          }
      }

      public int FOnComponentExitState(
          UIntPtr dwComponentID,
          uint uStateID,
          uint uContext,
          uint cpicmExclude,
          IntPtr rgpicmExclude)
      {
          _currentState = 0;
          if (uContext == MsoContextAll || uContext == MsoContextMine)
          {
              NotifyEnterState(uStateID, enter: false);
          }

          return 1;
      }

      public int FInState(uint uStateID, IntPtr pvoid) => _currentState == uStateID ? 1 : 0;

      public int FContinueIdle()
      {
          return VB6StaMessagePump.Peek(out _, remove: false) ? 0 : 1;
      }

      public int FPushMessageLoop(UIntPtr dwComponentID, uint uReason, IntPtr pvLoopData)
      {
          if (!_components.TryGetValue(dwComponentID, out IMsoComponent? requesting))
          {
              Debug.WriteLine("IMsoComponentManager.FPushMessageLoop: unknown component " + dwComponentID);
              return 0;
          }

          Debug.WriteLine("IMsoComponentManager.FPushMessageLoop enter reason=" + DescribeLoopReason(uReason));
          uint previousState = _currentState;
          IMsoComponent? previousActive = _active;
          _active = requesting;
          if (IsModalLoopReason(uReason))
          {
              ComHost.NotifyModalLoopStarted(requesting);
          }

          try
          {
              return RunPushedLoop(requesting, uReason, pvLoopData);
          }
          finally
          {
              _currentState = previousState;
              _active = previousActive;
              Debug.WriteLine("IMsoComponentManager.FPushMessageLoop exit reason=" + DescribeLoopReason(uReason));
          }
      }

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
          IMsoComponent? component = dwgac switch
          {
              MsoGacActive => _active,
              MsoGacTracking => _tracking,
              MsoGacTrackingOrActive => _tracking ?? _active,
              _ => null
          };

          if (component == null)
          {
              return 0;
          }

          if (ppic != IntPtr.Zero)
          {
              Marshal.WriteIntPtr(ppic, Marshal.GetComInterfaceForObject(component, typeof(IMsoComponent)));
          }

          return 1;
      }

      private int RunPushedLoop(IMsoComponent requesting, uint uReason, IntPtr pvLoopData)
      {
          while (true)
          {
              ComHost.DrainPostedWork();
              IMsoComponent component = _tracking ?? _active ?? requesting;

              if (VB6StaMessagePump.Peek(out VB6StaMessagePump.MSG peeked, remove: false))
              {
                  if (!ContinueLoop(component, uReason, pvLoopData, peeked))
                  {
                      return 1;
                  }

                  if (!VB6StaMessagePump.TryGetMessage(out VB6StaMessagePump.MSG msg))
                  {
                      return 0;
                  }

                  if (VB6StaMessagePump.IsQuit(in msg))
                  {
                      if (uReason != MsoLoopMain)
                      {
                          VB6StaMessagePump.RepostQuit(in msg);
                      }

                      return 1;
                  }

                  if (component.FPreTranslateMessage(MsgPtr(in msg)) == 0)
                  {
                      VB6StaMessagePump.TranslateAndDispatch(ref msg);
                  }
              }
              else
              {
                  if (uReason == MsoLoopDoEvents || uReason == MsoLoopDoEventsModal)
                  {
                      return 1;
                  }

                  bool wantIdle = DoIdle();
                  if (!ContinueLoop(component, uReason, pvLoopData, peekedMsg: null))
                  {
                      return 1;
                  }

                  if (wantIdle)
                  {
                      if (ComHost.TryGetWakeup(out WaitHandle wakeup))
                      {
                          VB6StaMessagePump.WaitForInputOrHandle(wakeup, 100);
                      }
                      else
                      {
                          VB6StaMessagePump.WaitForInput(100);
                      }
                  }
                  else if (!VB6StaMessagePump.Peek(out _, remove: false))
                  {
                      if (ComHost.TryGetWakeup(out WaitHandle wakeup))
                      {
                          VB6StaMessagePump.WaitForInputOrHandle(wakeup, 0xFFFFFFFF);
                      }
                      else
                      {
                          VB6StaMessagePump.Wait();
                      }
                  }
              }
          }
      }

      private bool ContinueLoop(IMsoComponent component, uint uReason, IntPtr pvLoopData, VB6StaMessagePump.MSG? peekedMsg)
      {
          IntPtr pMsg = peekedMsg.HasValue ? MsgPtr(peekedMsg.Value) : IntPtr.Zero;
          return component.FContinueMessageLoop(uReason, pvLoopData, pMsg) != 0;
      }

      private bool DoIdle()
      {
          bool continueIdle = false;
          foreach (IMsoComponent idle in SnapshotComponents())
          {
              continueIdle |= idle.FDoIdle(MsoIdleAll) != 0;
          }

          return continueIdle;
      }

      private void NotifyEnterState(uint uStateID, bool enter)
      {
          foreach (IMsoComponent component in SnapshotComponents())
          {
              component.OnEnterState(uStateID, enter);
          }
      }

      private IMsoComponent[] SnapshotComponents()
      {
          if (_components.Count == 0)
          {
              return Array.Empty<IMsoComponent>();
          }

          IMsoComponent[] copy = new IMsoComponent[_components.Count];
          _components.Values.CopyTo(copy, 0);
          return copy;
      }

      private IntPtr MsgPtr(in VB6StaMessagePump.MSG msg)
      {
          if (_msgScratch == IntPtr.Zero)
          {
              _msgScratch = Marshal.AllocHGlobal(_msgSize);
          }

          Marshal.StructureToPtr(msg, _msgScratch, false);
          return _msgScratch;
      }

      private void ReleaseNative()
      {
          if (_disposed)
          {
              return;
          }

          _disposed = true;
          foreach (IMsoComponent component in SnapshotComponents())
          {
              ReleaseComponent(component);
          }

          _components.Clear();
          _active = null;
          _tracking = null;
          if (_msgScratch != IntPtr.Zero)
          {
              Marshal.FreeHGlobal(_msgScratch);
              _msgScratch = IntPtr.Zero;
          }
      }

      private static void ReleaseComponent(IMsoComponent component)
      {
          try
          {
              _ = Marshal.ReleaseComObject(component);
          }
          catch (InvalidComObjectException)
          {
          }
      }

      private static bool IsModalLoopReason(uint uReason)
      {
          return uReason == MsoLoopModalForm || uReason == MsoLoopModalAlert;
      }

      private static string DescribeLoopReason(uint uReason)
      {
          return uReason switch
          {
              1 => "FocusWait",
              MsoLoopDoEvents => "DoEvents",
              3 => "Debug",
              MsoLoopModalForm => "ModalForm",
              MsoLoopModalAlert => "ModalAlert",
              MsoLoopDoEventsModal => "DoEventsModal",
              _ => uReason.ToString()
          };
      }
  }
}
