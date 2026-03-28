using System;
using System.Collections.Generic;

namespace VB6ComHost
{
  internal static class ComKnownIidNames
  {
    private static readonly Dictionary<Guid, string> Map = new Dictionary<Guid, string>
    {
      [new Guid("00000016-0000-0000-c000-000000000046")] = "IMessageFilter",
      [new Guid("6d5140c1-7436-11ce-8034-00aa006009fa")] = "IServiceProvider",
      [new Guid("000c0601-0000-0000-c000-000000000046")] = "SID_SMsoComponentManager",
      [new Guid("000c060b-0000-0000-c000-000000000046")] = "IMsoComponentManager",
    };

    public static string Describe(Guid iid)
    {
      return Map.TryGetValue(iid, out string? name)
        ? name + " {" + iid.ToString("D") + "}"
        : "{" + iid.ToString("D") + "}";
    }
  }
}