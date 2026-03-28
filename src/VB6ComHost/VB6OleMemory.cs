using System;
using System.Runtime.InteropServices;

namespace VB6ComHost
{
  internal static class VB6OleMemory
  {
    private const uint MEM_COMMIT = 0x1000;
    private const uint PAGE_NOACCESS = 0x01;
    private const uint PAGE_READONLY = 0x02;
    private const uint PAGE_READWRITE = 0x04;
    private const uint PAGE_WRITECOPY = 0x08;
    private const uint PAGE_EXECUTE_READ = 0x20;
    private const uint PAGE_EXECUTE_READWRITE = 0x40;
    private const uint PAGE_EXECUTE_WRITECOPY = 0x80;

    [StructLayout(LayoutKind.Sequential)]
    private struct MEMORY_BASIC_INFORMATION
    {
      public IntPtr BaseAddress;
      public IntPtr AllocationBase;
      public uint AllocationProtect;
      public UIntPtr RegionSize;
      public uint State;
      public uint Protect;
      public uint Type;
    }

    [DllImport("kernel32.dll")]
    private static extern UIntPtr VirtualQuery(IntPtr lpAddress, ref MEMORY_BASIC_INFORMATION lpBuffer, UIntPtr dwLength);

    private static bool TryQuery(IntPtr p, out MEMORY_BASIC_INFORMATION mbi)
    {
      mbi = default;
      if (p == IntPtr.Zero)
      {
        return false;
      }

      int sz = Marshal.SizeOf(typeof(MEMORY_BASIC_INFORMATION));
      return VirtualQuery(p, ref mbi, (UIntPtr)sz) != UIntPtr.Zero;
    }

    internal static bool IsReadable(IntPtr p, int byteCount)
    {
      if (!TryQuery(p, out MEMORY_BASIC_INFORMATION mbi))
      {
        return false;
      }

      if ((mbi.State & MEM_COMMIT) == 0)
      {
        return false;
      }

      long offset = p.ToInt64() - mbi.BaseAddress.ToInt64();
      if (offset < 0)
      {
        return false;
      }

      ulong need = (ulong)offset + (ulong)byteCount;
      if (need > mbi.RegionSize.ToUInt64())
      {
        return false;
      }

      uint prot = mbi.Protect & 0xFF;
      if (prot == PAGE_NOACCESS)
      {
        return false;
      }

      return prot == PAGE_READONLY
        || prot == PAGE_READWRITE
        || prot == PAGE_WRITECOPY
        || prot == PAGE_EXECUTE_READ
        || prot == PAGE_EXECUTE_READWRITE
        || prot == PAGE_EXECUTE_WRITECOPY;
    }

    internal static bool IsPointerSlotWritable(IntPtr ppv)
    {
      if (!TryQuery(ppv, out MEMORY_BASIC_INFORMATION mbi))
      {
        return false;
      }

      if ((mbi.State & MEM_COMMIT) == 0)
      {
        return false;
      }

      long offset = ppv.ToInt64() - mbi.BaseAddress.ToInt64();
      if (offset < 0)
      {
        return false;
      }

      ulong need = (ulong)offset + (ulong)IntPtr.Size;
      if (need > mbi.RegionSize.ToUInt64())
      {
        return false;
      }

      uint prot = mbi.Protect & 0xFF;
      return prot == PAGE_READWRITE
        || prot == PAGE_WRITECOPY
        || prot == PAGE_EXECUTE_READWRITE
        || prot == PAGE_EXECUTE_WRITECOPY;
    }
  }
}
