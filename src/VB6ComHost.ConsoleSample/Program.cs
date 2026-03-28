using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CSharp.RuntimeBinder;

namespace VB6ComHost.ConsoleSample
{
  internal static class Program
  {
    /// <summary>STA required for VB6 COM + message filter; x86 for 32-bit ActiveX DLL.</summary>
    [STAThread]
    private static void Main()
    {
      using (CancellationTokenSource quit = new CancellationTokenSource())
      {
        _ = Task.Run(() =>
        {
          Console.ReadKey(intercept: true);
          quit.Cancel();
        });

        Console.WriteLine("STA console host — message pump from VB6StaMessagePump (no WinForms).");
        Console.WriteLine("Press any key to exit the pump.");

        try
        {
          using (VB6ComHost host = VB6ComHost.Open())
          {
            Console.WriteLine("CoRegisterMessageFilter HRESULT: 0x" + host.MessageFilterRegistrationHResult.ToString("X8"));

            dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
            launcher.ShowNonModal();
            Console.WriteLine("ShowNonModal() via dynamic; pumping...");

            VB6StaMessagePump.Run(quit.Token);
          }
        }
        catch (RuntimeBinderException ex)
        {
          Console.WriteLine("Dynamic: " + ex.Message);
        }
      }
    }
  }
}
