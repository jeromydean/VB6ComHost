using System;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace Vb6.ActiveX.Hosting.Tests
{
  public sealed class ComHostTests
  {
    [Fact]
    public void ThrowIfUnsupportedProcess_throws_on_64_bit_when_enabled()
    {
      ComHostOptions options = new ComHostOptions { ThrowIf64BitProcess = true };
      if (Environment.Is64BitProcess)
      {
        Assert.Throws<PlatformNotSupportedException>(() => ComHost.ThrowIfUnsupportedProcess(options));
      }
      else
      {
        ComHost.ThrowIfUnsupportedProcess(options);
      }
    }

    [Fact]
    public void ThrowIfUnsupportedProcess_can_be_disabled()
    {
      ComHost.ThrowIfUnsupportedProcess(new ComHostOptions { ThrowIf64BitProcess = false });
    }

    [Fact]
    public void ComHostException_includes_progid_and_hresult()
    {
      ComHostException ex = new ComHostException("missing", "No.Such.ProgID", ComHostException.ClassNotRegistered);
      Assert.Equal("No.Such.ProgID", ex.ProgId);
      Assert.Equal(ComHostException.ClassNotRegistered, ex.HResult);
      Assert.Contains("regsvr32", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Open_and_GetWindows_empty()
    {
      using ComHost host = ComHost.Open();
      Assert.Empty(host.GetWindows());
    }

    [Fact]
    public void CreateInstance_unknown_progid_throws_ComHostException()
    {
      using ComHost host = ComHost.Open();
      ComHostException ex = Assert.Throws<ComHostException>(() => host.CreateInstance("Vb6.ActiveX.Hosting.Tests.Missing"));
      Assert.Equal("Vb6.ActiveX.Hosting.Tests.Missing", ex.ProgId);
      Assert.Equal(ComHostException.ClassNotRegistered, ex.HResult);
    }

    [Fact]
    public async Task InvokeAsync_runs_on_host_thread()
    {
      using ComHost host = ComHost.Open();
      ApartmentState apartment = await host.InvokeAsync(() => Thread.CurrentThread.GetApartmentState());
      Assert.Equal(ApartmentState.STA, apartment);
    }

    [Fact]
    public void Invoke_cancelled_throws()
    {
      using ComHost host = ComHost.Open();
      using CancellationTokenSource cts = new CancellationTokenSource();
      cts.Cancel();
      Assert.Throws<OperationCanceledException>(() => host.Invoke(() => 1, cts.Token));
    }

    [Fact]
    public void CreateInstance_generic_unknown_progid_throws()
    {
      using ComHost host = ComHost.Open();
      Assert.Throws<ComHostException>(() => host.CreateInstance<object>("Vb6.ActiveX.Hosting.Tests.Missing"));
    }
  }
}
