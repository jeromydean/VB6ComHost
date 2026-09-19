using System;
using System.Dynamic;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace Vb6.ActiveX.Hosting
{
  /// <summary>
  /// Late-bound wrapper so <c>dynamic</c> calls run on the host STA thread.
  /// Calling the raw RCW from MTA makes <c>Show vbModal</c> an incoming COM call;
  /// the nested pump then cannot dismiss the form.
  /// </summary>
  internal sealed class StaComProxy : DynamicObject
  {
    private const BindingFlags InvokeFlags =
      BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase;

    private readonly ComHost _host;
    private readonly object _instance;

    private StaComProxy(ComHost host, object instance)
    {
      _host = host;
      _instance = instance;
    }

    internal static object Wrap(ComHost host, object instance)
    {
      if (host == null)
      {
        throw new ArgumentNullException(nameof(host));
      }

      if (instance == null)
      {
        throw new ArgumentNullException(nameof(instance));
      }

      if (instance is StaComProxy || !Marshal.IsComObject(instance))
      {
        return instance;
      }

      return new StaComProxy(host, instance);
    }

    internal static object Unwrap(object instance)
    {
      StaComProxy? proxy = instance as StaComProxy;
      return proxy != null ? proxy._instance : instance;
    }

    public override bool TryInvokeMember(InvokeMemberBinder binder, object?[]? args, out object? result)
    {
      result = InvokeOnSta(binder.Name, BindingFlags.InvokeMethod | InvokeFlags, args);
      return true;
    }

    public override bool TryGetMember(GetMemberBinder binder, out object? result)
    {
      result = InvokeOnSta(binder.Name, BindingFlags.GetProperty | InvokeFlags, null);
      return true;
    }

    public override bool TrySetMember(SetMemberBinder binder, object? value)
    {
      object?[] args = { value };
      _ = InvokeOnSta(binder.Name, BindingFlags.SetProperty | InvokeFlags, args);
      return true;
    }

    public override bool TryConvert(ConvertBinder binder, out object? result)
    {
      if (binder.Type == typeof(object) || binder.Type.IsInstanceOfType(this))
      {
        result = this;
        return true;
      }

      result = null;
      return false;
    }

    private object? InvokeOnSta(string name, BindingFlags flags, object?[]? args)
    {
      return _host.Invoke(() =>
      {
        try
        {
          object? value = _instance.GetType().InvokeMember(name, flags, null, _instance, args);
          return value != null && Marshal.IsComObject(value) ? Wrap(_host, value) : value;
        }
        catch (TargetInvocationException ex) when (ex.InnerException != null)
        {
          ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
          return null;
        }
      });
    }
  }
}
