# 🧩 VB6ComHost

.NET-side support for **hosting Visual Basic 6 ActiveX DLLs** that show **modeless** forms and expect a cooperating Win32 host. The focus is a **pure, managed C#** approach on Windows: register an STA OLE message filter and related COM plumbing so VB6’s runtime (`MSVBVM60`) accepts the host, without requiring a native helper DLL for the scenario this project targets.

---

## ⚠️ VB6 runtime message this project addresses

If you are searching for the exact text you see at run time, it is:

> **Non-modal forms cannot be displayed in this host application from an ActiveX DLL, ActiveX Control, or Property Page.**

That message appears when VB6 decides the host is **not** a suitable container for **modeless** UI from an **ActiveX DLL** (or related binaries). In code, the same gate is often reflected as **`App.NonModalAllowed`** being **False** when you expected modeless forms to work. The hosting pattern here is aimed at satisfying the checks behind that message—together with a proper **STA** thread and a **Windows message pump**—so modeless VB6 forms can run inside a .NET process.

Related search terms people use alongside that error: *VB6 modeless ActiveX DLL*, *NonModalAllowed*, *CoRegisterMessageFilter*, *IMessageFilter*, *IServiceProvider*, *IMsoComponentManager*, *KB 247791*, *VB6 host application*.

---

## 🎯 What this code is meant to accomplish

When a VB6 class (compiled as an **ActiveX DLL**) shows a **modeless** form, the runtime checks that the host process looks like a “real” OLE/Office-style container (historically described in context of [KB Q247791](https://support.microsoft.com/help/247791)). If those checks fail, behavior such as `App.NonModalAllowed` or modeless UI can fail even though the same code works inside VB6.exe.

**This repository’s goal** is to provide a small **.NET Standard 2.0** library (**`VB6ComHost`**) that:

1. **Registers a COM-visible STA `IMessageFilter`** via `CoRegisterMessageFilter`, matching what VB6 probes on the hosting thread.
2. **Exposes `IServiceProvider` from that filter** (via a separate CCW shim) so `QueryService` can answer the **Office-style** `IMsoComponentManager` / `SID_SMsoComponentManager` pair with a **minimal managed stub**—enough for typical VB6 modeless checks in tested configurations.
3. **Lets the host create coclasses by ProgID or CLSID** (`Activator` / `Type.GetTypeFromProgID`) for late binding (`dynamic`, `InvokeMember`, or PIAs).
4. Documents that the host must still run a **Windows message pump** on the **same STA thread** (WinForms/WPF, or a classic `PeekMessage` / `DispatchMessage` loop). Modeless VB6 UI will not run correctly without it.

**Companion client** (**`VB6ComClient`**): a sample executable (currently .NET Framework–style) intended to reference the library and demonstrate **STA**, **32-bit (x86) when talking to 32-bit VB6 binaries**, and a message loop while COM objects are alive.

Together, the intent is: **from C#, host a registered VB6 ActiveX DLL coclass, call into it (e.g. show a modeless form), and keep it responsive** using only managed interop plus the OS message APIs.

---

## 📋 Requirements

| Requirement | Why |
|-------------|-----|
| **Windows** | COM, `ole32`, `user32`, VB6 runtime are Win32-centric. |
| **STA thread** | VB6 COM objects and the message filter registration expect single-threaded apartment; use `[STAThread]` or equivalent before opening the host. |
| **32-bit (x86) process** | VB6 ActiveX DLLs are normally **32-bit**; a 64-bit host cannot load them in-process. |
| **.NET** | Library targets **.NET Standard 2.0**; clients can be .NET Framework 4.7.2+ or modern .NET on Windows. |

---

## 📁 Repository layout

```text
src/
  VB6ComHost.slnx          # Visual Studio solution (filterable `.slnx` format)
  VB6ComHost/              # Class library (netstandard2.0) — hosting API lives here
  VB6ComClient/            # Sample / test host executable
```

Open **`src/VB6ComHost.slnx`** in Visual Studio (2022 17.10+ / 17.12+ with `.slnx` support, or newer). You can also build from the command line with MSBuild/dotnet as your toolchain supports for SDK-style and classic projects.

---

## 🚧 Current status

The solution is wired with **`VB6ComHost`** (library) and **`VB6ComClient`** (client) and a **MIT** license. The **implementation** of the message filter, service provider shim, optional STA message pump, and public host API is **to be filled in** (or ported from your prior work) so the bullets under *What this code is meant to accomplish* are satisfied in this tree.

---

## 📄 License

See [LICENSE](LICENSE) (MIT).
