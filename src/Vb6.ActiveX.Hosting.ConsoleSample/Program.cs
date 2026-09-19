using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CSharp.RuntimeBinder;

namespace Vb6.ActiveX.Hosting.ConsoleSample
{
  internal static class Program
  {
    private static readonly ConcurrentQueue<MenuCommand> s_commands = new ConcurrentQueue<MenuCommand>();
    private static readonly List<Task> s_asyncOps = new List<Task>();
    private static readonly object s_asyncOpsLock = new object();
    private static volatile bool s_quit;
    private static int s_spawnCount;
    private static ComHost? s_host;

    private enum MenuCommand
    {
      SpawnNonModalSync,
      SpawnModalSync,
      SpawnNonModalAsync,
      SpawnModalAsync,
      Quit,
    }

    /// <summary>MTA is fine: ComHost starts its own STA thread and message pump.</summary>
    private static async Task Main()
    {
      Console.OutputEncoding = Encoding.UTF8;
      Console.InputEncoding = Encoding.UTF8;
      Console.Title = "🖥️ VB6 COM host — console sample";

      PrintBanner();

      try
      {
        await using (ComHost host = ComHost.Open())
        {
          s_host = host;
          host.ModalWindowOpened += OnModalWindowOpened;
          host.ModalWindowClosed += OnModalWindowClosed;
          host.ModelessWindowOpened += OnModelessWindowOpened;
          host.ModelessWindowClosed += OnModelessWindowClosed;

          Console.WriteLine("Apartment: " + Thread.CurrentThread.GetApartmentState() + " (caller) / host STA thread is internal");
          Console.WriteLine("CoRegisterMessageFilter HRESULT: 0x" + host.MessageFilterRegistrationHResult.ToString("X8"));
          Console.WriteLine();

          _ = Task.Run((Action)KeyboardLoop);

          while (!s_quit)
          {
            DrainCommands(host);
            await Task.Delay(s_commands.IsEmpty ? 50 : 0);
          }

          await WaitForAsyncOps();
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
      Console.WriteLine("  🪟 🪟 🪟  VB6 ActiveX DLL host (ComHost owns the STA + pump)  🪟 🪟 🪟");
      Console.WriteLine("  \u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015\u2015");
      Console.WriteLine();
      Console.WriteLine("  Sync (CreateInstance + late-bound call; M blocks this menu until the form closes):");
      Console.WriteLine("    🪟 N  ShowNonModal");
      Console.WriteLine("    🔒 M  ShowModal");
      Console.WriteLine();
      Console.WriteLine("  Async (InvokeAsync; A does not block this menu — press Q to close the form):");
      Console.WriteLine("    🪟 I  ShowNonModal");
      Console.WriteLine("    🔒 A  ShowModal");
      Console.WriteLine();
      Console.WriteLine("    ❌ Q  Close all VB6 forms, then quit (await using → DisposeAsync)");
      Console.WriteLine();
      Console.WriteLine("  ComHost runs VB6 on its own STA thread. This process can be MTA.");
      Console.WriteLine("  Q posts WM_CLOSE to tracked modal and modeless forms so ShowModal can return.");
      Console.WriteLine();
      Console.WriteLine("  Tip: If you open a modeless window (N or I) and then a modal (M or A), closing the");
      Console.WriteLine("  modal prints \"returned\" while the modeless window can still be open — check titles.");
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
            s_commands.Enqueue(MenuCommand.SpawnNonModalSync);
            break;
          case ConsoleKey.M:
            s_commands.Enqueue(MenuCommand.SpawnModalSync);
            break;
          case ConsoleKey.I:
            s_commands.Enqueue(MenuCommand.SpawnNonModalAsync);
            break;
          case ConsoleKey.A:
            s_commands.Enqueue(MenuCommand.SpawnModalAsync);
            break;
          case ConsoleKey.Q:
            s_host?.CloseWindows();
            s_commands.Enqueue(MenuCommand.Quit);
            break;
          default:
            break;
        }
      }
    }

    private static void DrainCommands(ComHost host)
    {
      while (s_commands.TryDequeue(out MenuCommand cmd))
      {
        switch (cmd)
        {
          case MenuCommand.SpawnNonModalSync:
            SpawnNonModalSync(host);
            break;
          case MenuCommand.SpawnModalSync:
            SpawnModalSync(host);
            break;
          case MenuCommand.SpawnNonModalAsync:
            TrackAsync(SpawnNonModalAsync(host));
            break;
          case MenuCommand.SpawnModalAsync:
            TrackAsync(SpawnModalAsync(host));
            break;
          case MenuCommand.Quit:
            s_quit = true;
            Console.WriteLine();
            Console.WriteLine("  ❌ Quit requested.");
            break;
        }
      }
    }

    private static void SpawnNonModalSync(ComHost host)
    {
      int n = Interlocked.Increment(ref s_spawnCount);
      dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
      launcher.ShowNonModal(n);
      Console.WriteLine("  🪟 Sync non-modal #" + n + "  CreateInstance + ShowNonModal");
      Console.Out.Flush();
    }

    private static void SpawnModalSync(ComHost host)
    {
      int n = Interlocked.Increment(ref s_spawnCount);
      Console.WriteLine("  🔒 Sync modal #" + n + "  Invoke(CreateInstance + ShowModal) — menu waits; title-bar X should close it.");
      Console.Out.Flush();
      host.Invoke(() =>
      {
        dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
        launcher.ShowModal(n);
      });
      Console.WriteLine("  ✅ Sync modal #" + n + " closed.");
      Console.Out.Flush();
    }

    private static async Task SpawnNonModalAsync(ComHost host)
    {
      int n = Interlocked.Increment(ref s_spawnCount);
      Console.WriteLine("  🪟 Async non-modal #" + n + "  InvokeAsync(CreateInstance + ShowNonModal)");
      Console.Out.Flush();
      await host.InvokeAsync(() =>
      {
        dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
        launcher.ShowNonModal(n);
      });
    }

    private static async Task SpawnModalAsync(ComHost host)
    {
      int n = Interlocked.Increment(ref s_spawnCount);
      Console.WriteLine("  🔒 Async modal #" + n + "  InvokeAsync(ShowModal) — menu keeps running (Q still works).");
      Console.Out.Flush();
      await host.InvokeAsync(() =>
      {
        dynamic launcher = host.CreateInstance("ActiveXLibrary.WindowLauncher");
        launcher.ShowModal(n);
      });
      Console.WriteLine("  ✅ Async modal #" + n + " closed.");
      Console.Out.Flush();
    }

    private static void TrackAsync(Task task)
    {
      lock (s_asyncOpsLock)
      {
        s_asyncOps.Add(Observe(task));
      }
    }

    private static async Task Observe(Task task)
    {
      try
      {
        await task;
      }
      catch (Exception ex)
      {
        Console.WriteLine("  ⚠ Async operation: " + ex.Message);
        Console.Out.Flush();
      }
    }

    private static Task WaitForAsyncOps()
    {
      Task[] pending;
      lock (s_asyncOpsLock)
      {
        pending = s_asyncOps.ToArray();
      }

      if (pending.Length == 0)
      {
        return Task.CompletedTask;
      }

      return Task.WhenAll(pending);
    }

    private static void OnModalWindowOpened(object? sender, HostWindowEventArgs e)
    {
      Console.WriteLine("  📡 Modal opened  hwnd=0x" + e.Handle.ToString("X") + "  \"" + e.Title + "\"");
      Console.Out.Flush();
    }

    private static void OnModalWindowClosed(object? sender, HostWindowEventArgs e)
    {
      Console.WriteLine("  📡 Modal closed  hwnd=0x" + e.Handle.ToString("X") + "  \"" + e.Title + "\"");
      Console.Out.Flush();
    }

    private static void OnModelessWindowOpened(object? sender, HostWindowEventArgs e)
    {
      Console.WriteLine("  📡 Modeless opened  hwnd=0x" + e.Handle.ToString("X") + "  \"" + e.Title + "\"");
      Console.Out.Flush();
    }

    private static void OnModelessWindowClosed(object? sender, HostWindowEventArgs e)
    {
      Console.WriteLine("  📡 Modeless closed  hwnd=0x" + e.Handle.ToString("X") + "  \"" + e.Title + "\"");
      Console.Out.Flush();
    }
  }
}
