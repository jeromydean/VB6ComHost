ActiveXLibrary (VB6)
=====================

Sample 32-bit ActiveX DLL used by VB6ComHost.ConsoleSample. It exposes class
WindowLauncher (ProgID ActiveXLibrary.WindowLauncher) with ShowNonModal, which
shows Form1 modelessly.

Build in Visual Basic 6 (File > Make ActiveXLibrary.dll). The checked-in binary
for convenience lives at:

  lib/com/x86/ActiveXLibrary.dll

Registration (elevation usually required on modern Windows):

  regsvr32 "%REPO_ROOT%\lib\com\x86\ActiveXLibrary.dll"

The console sample copies that DLL to its output folder on build, but COM still
requires registration (or registration-free COM via manifest) for
ProgID-based activation.
