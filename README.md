# 🧩 VB6ComHost

.NET-side support for **hosting Visual Basic 6 ActiveX DLLs** that show **modeless** (and optionally modal) forms and expect a cooperating Win32 host. The focus is a **pure, managed C#** approach on Windows: register an STA OLE message filter and related COM plumbing so VB6’s runtime (`MSVBVM60`) accepts the host, without requiring a native helper DLL for the scenario this project targets.

---

## ⚠️ VB6 runtime message this project addresses

> **Non-modal forms cannot be displayed in this host application from an ActiveX DLL, ActiveX Control, or Property Page.**

That message appears when VB6 decides the host is **not** a suitable container for **modeless** UI from an **ActiveX DLL** (or related binaries). In code, the same gate is often reflected as **`App.NonModalAllowed`** being **False** when you expected modeless forms to work. The hosting pattern here is aimed at satisfying the checks behind that message—together with a proper **STA** thread and a **Windows message pump**—so modeless VB6 forms can run inside a .NET process.

Related search terms people use alongside that error: *VB6 modeless ActiveX DLL*, *NonModalAllowed*, *CoRegisterMessageFilter*, *IMessageFilter*, *IServiceProvider*, *IMsoComponentManager*, *KB 247791*, *VB6 host application*.

---

## 🎯 What this repository provides

**`VB6ComHost`** (class library, **.NET Standard 2.0**, Windows-only at runtime):

1. Registers a COM-visible STA **`IMessageFilter`** via `CoRegisterMessageFilter` (see [KB Q247791](https://support.microsoft.com/help/247791) context).
2. Exposes **`IServiceProvider`** from that filter (via a CCW shim) so `QueryService` can answer the Office-style **`IMsoComponentManager` / `SID_SMsoComponentManager`** pair with a minimal managed stub.
3. **`VB6ComHost.Open()`** / **`CreateInstance(progId)`** / **`Dispose`** for hosting on the current STA thread.
4. **`VB6StaMessagePump`** — a small `PeekMessage` / `DispatchMessage` loop for hosts that are not WinForms or WPF.

**`VB6ComHost.ConsoleSample`** (SDK-style **.NET Framework 4.8** executable, **x86**): demonstrates **`[STAThread]`**, **`Microsoft.CSharp`** for `dynamic`, multiple modeless windows (**N**), modal demo (**M**), and UTF-8 console output (emoji look best in **Windows Terminal**).

**`vb6/ActiveXLibrary/`** — sample VB6 ActiveX DLL sources (`WindowLauncher` with `ShowNonModal` / `ShowModal`).

**`lib/com/x86/ActiveXLibrary.dll`** — checked-in 32-bit build for convenience; the console project copies it to output on build. You still need **`regsvr32`** (or registration-free COM) for ProgID activation.

---

## 📋 Requirements

| Requirement | Why |
|-------------|-----|
| **Windows** | COM, `ole32`, `user32`, VB6 runtime are Win32-centric. |
| **STA thread** | VB6 COM objects and the message filter registration expect single-threaded apartment; use `[STAThread]` or equivalent before opening the host. |
| **32-bit (x86) process** when loading 32-bit VB6 DLLs in-process | VB6 ActiveX DLLs are normally **32-bit**; a 64-bit host cannot load them in-process. |
| **.NET** | Library targets **.NET Standard 2.0**; the sample targets **net48** and references **`Microsoft.CSharp`** for `dynamic`. |

---

## 📁 Repository layout

```text
lib/com/x86/              # Prebuilt sample ActiveXLibrary.dll (COM, 32-bit)
vb6/ActiveXLibrary/       # VB6 project sources for that DLL
src/
  VB6ComHost.slnx         # Visual Studio solution (.slnx)
  VB6ComHost/             # Class library (netstandard2.0)
  VB6ComHost.ConsoleSample/  # SDK-style console host (net48, x86)
```

### 🔨 Build

Build from the repo root or `src`:

```bash
dotnet build src/VB6ComHost.slnx
```

Open **`src/VB6ComHost.slnx`** in Visual Studio (2022 with `.slnx` support, or newer).

---

## 🚧 Known hosting quirks (modal)

When an **ActiveX DLL** shows **`Show vbModal`** and the caller is a **.NET COM host**, some configurations report that **`ShowModal` returns to managed code before the VB6 form is actually dismissed**, which can make it look as if several “modal” windows can be open at once or as if the console prints “returned” while a form is still visible. That behavior comes from the **VB6 runtime + hosting boundary**, not from the message pump being “disconnected.” The sample is primarily aimed at **modeless** hosting; treat **modal** as a demo only unless you validate it for your DLL and host.

---

## 📄 License

See [LICENSE](LICENSE) (MIT).
