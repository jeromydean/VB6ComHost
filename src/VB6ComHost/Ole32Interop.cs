using System;
using System.Runtime.InteropServices;

namespace VB6ComHost
{
  internal static class Ole32Interop
  {
    [DllImport("ole32.dll", PreserveSig = true)]
    internal static extern int CoRegisterMessageFilter(IntPtr lpMessageFilter, out IntPtr lplpMessageFilter);
  }
}