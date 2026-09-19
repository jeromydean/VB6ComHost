# 🧩 Vb6.ActiveX.Hosting

.NET-side support for **hosting Visual Basic 6 ActiveX DLLs** that show **modeless** (and optionally modal) forms and expect a cooperating Win32 host. The focus is a **pure, managed C#** approach on Windows: register an STA OLE message filter and related COM plumbing so VB6’s runtime (`MSVBVM60`) accepts the host, without requiring a native helper DLL for the scenario this project targets.

---

## ⚠️ VB6 runtime message this project addresses

> **Non-modal forms cannot be displayed in this host application from an ActiveX DLL, ActiveX Control, or Property Page.**

That message appears when VB6 decides the host is **not** a suitable container for **modeless** UI from an **ActiveX DLL** (or related binaries). In code, the same gate is often reflected as **`App.NonModalAllowed`** being **False** when you expected modeless forms to work. The hosting pattern here is aimed at satisfying the checks behind that message—together with a proper **STA** thread and a **Windows message pump**—so modeless VB6 forms can run inside a .NET process.

Related search terms people use alongside that error: *VB6 modeless ActiveX DLL*, *NonModalAllowed*, *CoRegisterMessageFilter*, *IMessageFilter*, *IServiceProvider*, *IMsoComponentManager*, *KB 247791*, *VB6 host application*.

---

## 🎯 What this repository provides

**`Vb6.ActiveX.Hosting`** (class library, **.NET Standard 2.0**, Windows-only at runtime):

1. Registers a COM-visible STA **`IMessageFilter`** via `CoRegisterMessageFilter` (see [KB Q247791](https://support.microsoft.com/help/247791) context).
2. Exposes **`IServiceProvider`** from that filter (via a CCW shim) so `QueryService` can answer the Office-style **`IMsoComponentManager` / `SID_SMsoComponentManager`** pair. The manager keeps VB6’s **`IMsoComponent`** registration and implements **`FPushMessageLoop`** as a nested `PeekMessage` / `GetMessage` pump so **`Show vbModal`** blocks until the form is dismissed.
3. **`ComHost.Open()`** / **`Open(ComHostOptions)`** starts a dedicated **STA** thread and message pump (MTA callers are fine). **`CreateInstance`** / **`CreateInstance<T>`** / **`Invoke`** / **`InvokeAsync`** run on that thread. The host **tracks** each `CreateInstance` RCW and **releases it on dispose** (or earlier via **`Release`**). Window events are posted off the STA thread (`SynchronizationContext` or the thread pool). **`GetWindows()`** snapshots open forms. **`CloseWindows()`** (or **`HostWindowEventArgs.Close()`**) posts **`WM_CLOSE`**.
4. **`ComHostException`** for activation failures (ProgID, HRESULT, regsvr32 hint). A 64-bit process throws **`PlatformNotSupportedException`** unless `ThrowIf64BitProcess` is false.
5. **`VB6StaMessagePump`** and COM interfaces live for interop / advanced hosts; prefer **`ComHost`**.

**`Vb6.ActiveX.Hosting.ConsoleSample`** (SDK-style **.NET Framework 4.8** executable, **x86**): demonstrates an **MTA** console calling **`ComHost`**, **`Microsoft.CSharp`** for `dynamic`, sync **`CreateInstance`** (**N** / **M**), async **`InvokeAsync`** (**I** / **A**), **`await using`** / **`DisposeAsync`**, and UTF-8 console output (emoji look best in **Windows Terminal**).

**`vb6/ActiveXLibrary/`** — sample VB6 ActiveX DLL sources (`WindowLauncher` with `ShowNonModal` / `ShowModal`).

**`lib/com/x86/ActiveXLibrary.dll`** — checked-in 32-bit build for convenience; the console project copies it to output on build. You still need **`regsvr32`** (or registration-free COM) for ProgID activation.

Typical caller code (sync `CreateInstance` + `using` / `Dispose`):

```csharp
var options = new ComHostOptions
{
  EventContext = SynchronizationContext.Current, // WinForms/WPF UI thread; omit for thread-pool events
  ShutdownTimeout = TimeSpan.FromSeconds(2),
};

using (ComHost host = ComHost.Open(options))
{
  host.ModelessWindowOpened += (_, e) => { /* not on the host STA thread */ };
  dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
  launcher.ShowNonModal(1); // late-bound calls run on the host STA thread
  foreach (HostWindowInfo window in host.GetWindows())
  {
    _ = window.Title;
  }
} // Dispose closes remaining forms and releases tracked RCWs on the STA thread
```

Same host, async (`InvokeAsync` + `await using` / `DisposeAsync`). `ShowModal` does not block the caller:

```csharp
await using (ComHost host = ComHost.Open())
{
  await host.InvokeAsync(() =>
  {
    dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
    launcher.ShowModal(1);
  });
}
```

The NuGet listing uses [`src/Vb6.ActiveX.Hosting/README.md`](src/Vb6.ActiveX.Hosting/README.md). Pack from that project (GitHub repo is currently [jeromydean/VB6ComHost](https://github.com/jeromydean/VB6ComHost)):

```bash
dotnet pack src/Vb6.ActiveX.Hosting/Vb6.ActiveX.Hosting.csproj -c Release
```

The package targets **netstandard2.0** and **net8.0-windows**, is **Windows-only**, and must be consumed from a **32-bit (x86)** process when loading classic VB6 DLLs.

---

## 📋 Requirements

| Requirement | Why |
|-------------|-----|
| **Windows** | COM, `ole32`, `user32`, VB6 runtime are Win32-centric. |
| **STA thread** | VB6 COM and the message filter still require STA; **`ComHost.Open()`** creates that thread and pump so callers do not have to. |
| **32-bit (x86) process** when loading 32-bit VB6 DLLs in-process | VB6 ActiveX DLLs are normally **32-bit**; a 64-bit host cannot load them in-process. |
| **.NET** | Library targets **.NET Standard 2.0**; the sample targets **net48** and references **`Microsoft.CSharp`** for `dynamic`. |

---

## 📁 Repository layout

```text
lib/com/x86/              # Prebuilt sample ActiveXLibrary.dll (COM, 32-bit)
vb6/ActiveXLibrary/       # VB6 project sources for that DLL
src/
  Vb6.ActiveX.Hosting.slnx              # Visual Studio solution (.slnx)
  Vb6.ActiveX.Hosting/                  # Class library (netstandard2.0)
  Vb6.ActiveX.Hosting.ConsoleSample/    # SDK-style console host (net48, x86)
  Vb6.ActiveX.Hosting.Tests/            # xUnit (net48, x86)
```

### 🔨 Build

Build from the repo root or `src`:

```bash
dotnet build src/Vb6.ActiveX.Hosting.slnx
```

Open **`src/Vb6.ActiveX.Hosting.slnx`** in Visual Studio (2022 with `.slnx` support, or newer).

---

## 🔒 Modal forms and the nested message pump

When an **ActiveX DLL** calls **`Show vbModal`**, VB6 does not always run its own `GetMessage` loop. If the host advertises **`IMsoComponentManager`** (required for modeless hosting), VB6 registers an **`IMsoComponent`** and calls **`FPushMessageLoop`**. That call must pump messages until **`IMsoComponent.FContinueMessageLoop`** returns false (form closed). A stub that returns immediately makes **`ShowModal`** return to managed code while the form is still visible, so another “modal” can open. `ComHost` runs that nested pump on its dedicated STA thread.

---

## 📄 License

See [LICENSE](LICENSE) (MIT).
