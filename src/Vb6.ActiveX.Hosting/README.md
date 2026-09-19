# Vb6.ActiveX.Hosting

Host **VB6 ActiveX DLL** modeless and modal forms inside a .NET process. `ComHost.Open()` starts a dedicated STA thread and message pump (MTA callers are fine). You do not register an OLE message filter or pump messages yourself.

Windows-only. Classic 32-bit VB6 DLLs must be loaded from an **x86** (or AnyCPU prefer-32-bit) process.

## Install

```bash
dotnet add package Vb6.ActiveX.Hosting
```

## Sync

```csharp
using (ComHost host = ComHost.Open())
{
  host.ModelessWindowOpened += (_, e) => { /* not on the host STA thread */ };
  dynamic launcher = host.CreateInstance("YourLib.YourClass");
  launcher.ShowNonModal();
}
```

`CreateInstance` (non-generic) returns a late-bound proxy: `ShowModal` / `ShowNonModal` run on the host STA thread, so a modal form can be closed with the title-bar close box or `CloseWindows()`. The host tracks each RCW and releases it on `Dispose` (or earlier via `Release`). Do not call `Marshal.ReleaseComObject`. `CreateInstance<T>` returns the raw RCW — call UI methods through `Invoke` / `InvokeAsync`.

`Show vbModal` blocks the caller until the form closes. From another thread, `CloseWindows()` or `HostWindowEventArgs.Close()` posts `WM_CLOSE`.

## Async

```csharp
await using (ComHost host = ComHost.Open())
{
  await host.InvokeAsync(() =>
  {
    dynamic launcher = host.CreateInstance("YourLib.YourClass");
    launcher.ShowModal();
  });
}
```

`InvokeAsync` does not block the caller. `await using` calls `DisposeAsync`.

## Options

```csharp
ComHostOptions options = new ComHostOptions
{
  EventContext = SynchronizationContext.Current, // WinForms/WPF UI thread
  ShutdownTimeout = TimeSpan.FromSeconds(2),
  Logger = logger,
};

using (ComHost host = ComHost.Open(options))
{
  foreach (HostWindowInfo window in host.GetWindows())
  {
    _ = window.Title;
  }
}
```

Activation failures throw `ComHostException` (ProgID/CLSID, HRESULT, `regsvr32` hint). A 64-bit process throws `PlatformNotSupportedException` unless `ThrowIf64BitProcess` is false.

## Requirements

- Windows, .NET Standard 2.0 or .NET 8+ Windows
- 32-bit process when loading 32-bit VB6 DLLs in-process
- The ActiveX DLL registered (`regsvr32` from SysWOW64) or registration-free COM

Sample app and VB6 sources: [github.com/jeromydean/VB6ComHost](https://github.com/jeromydean/VB6ComHost).
