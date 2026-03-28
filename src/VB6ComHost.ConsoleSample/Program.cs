using System;
using System.Collections.Concurrent;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CSharp.RuntimeBinder;

namespace VB6ComHost.ConsoleSample
{
  internal static class Program
  {
    private static readonly ConcurrentQueue<MenuCommand> s_commands = new ConcurrentQueue<MenuCommand>();
    private static volatile bool s_quit;
    private static int s_spawnCount;

    private enum MenuCommand
    {
      SpawnNonModal,
      SpawnModal,
      Quit,
    }

    /// <summary>STA required for VB6 COM + message filter; x86 for 32-bit ActiveX DLL.</summary>
    [STAThread]
    private static void Main()
    {
      Console.OutputEncoding = Encoding.UTF8;
      Console.InputEncoding = Encoding.UTF8;
      Console.Title = "🖥️ VB6 COM host — console sample";

      PrintBanner();

      try
      {
        using (VB6ComHost host = VB6ComHost.Open())
        {
          Console.WriteLine("CoRegisterMessageFilter HRESULT: 0x" + host.MessageFilterRegistrationHResult.ToString("X8"));
          Console.WriteLine();

          _ = Task.Run((Action)KeyboardLoop);

          VB6StaMessagePump.RunWhile(() =>
          {
            DrainCommands(host);
            return !s_quit;
          });
        }
      }
      catch (RuntimeBinderException ex)
      {
        Console.WriteLine("Dynamic binding: " + ex.Message);
      }
    }

    private static void PrintBanner()
    {
      Console.WriteLine();
      Console.WriteLine("  🪟 🪟 🪟  VB6 ActiveX DLL host (STA + message pump)  🪟 🪟 🪟");
      Console.WriteLine("  \u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015");
      Console.WriteLine();
      Console.WriteLine("  Keys (while pumping):");
      Console.WriteLine("    🪟 N  Open another form  —  non-modal (needs this message pump)");
      Console.WriteLine("    🔒 M  Open another form  —  modal (VB6 runs its own loop until closed)");
      Console.WriteLine("    ❌ Q  Quit");
      Console.WriteLine();
      Console.WriteLine("  While a modal form is open, this thread is inside VB6 and other non-modal");
      Console.WriteLine("  forms may not get messages until the modal form closes.");
      Console.WriteLine();
      Console.WriteLine("  Tip: If you open a modeless window (N) and then a modal (M), closing the modal");
      Console.WriteLine("  prints \"returned\" while the modeless window can still be open — check titles.");
      Console.WriteLine();
    }

    private static void KeyboardLoop()
    {
      while (!s_quit)
      {
        ConsoleKeyInfo ki = Console.ReadKey(intercept: true);
        switch (ki.Key)
        {
          case ConsoleKey.N:
            s_commands.Enqueue(MenuCommand.SpawnNonModal);
            break;
          case ConsoleKey.M:
            s_commands.Enqueue(MenuCommand.SpawnModal);
            break;
          case ConsoleKey.Q:
            s_commands.Enqueue(MenuCommand.Quit);
            break;
          default:
            break;
        }
      }
    }

    private static void DrainCommands(VB6ComHost host)
    {
      while (s_commands.TryDequeue(out MenuCommand cmd))
      {
        switch (cmd)
        {
          case MenuCommand.SpawnNonModal:
            SpawnNonModal(host);
            break;
          case MenuCommand.SpawnModal:
            SpawnModal(host);
            break;
          case MenuCommand.Quit:
            s_quit = true;
            Console.WriteLine();
            Console.WriteLine("  ❌ Quit requested.");
            break;
        }
      }
    }

    private static void SpawnNonModal(VB6ComHost host)
    {
      int n = Interlocked.Increment(ref s_spawnCount);
      dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
      launcher.ShowNonModal(n);
      Console.WriteLine("  🪟 Non-modal #" + n + "  (title: \"VB6 modeless #" + n + "\")");
      Console.Out.Flush();
    }

    private static void SpawnModal(VB6ComHost host)
    {
      int n = Interlocked.Increment(ref s_spawnCount);
      dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
      Console.WriteLine("  🔒 Modal #" + n + "  opening (title: \"VB6 modal #" + n + "\") — next line appears only after you close that window.");
      Console.Out.Flush();
      launcher.ShowModal(n);
      Console.WriteLine("  ✅ Modal #" + n + " closed (only the modal #" + n + " window should be gone).");
      Console.Out.Flush();
    }
  }
}
