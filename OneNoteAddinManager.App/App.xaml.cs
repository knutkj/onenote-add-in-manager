using System.Configuration;
using System.Data;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OneNoteAddinManager.Lib.Services;

namespace OneNoteAddinManager.App;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private IHost? _host;

    protected override void OnStartup(StartupEventArgs e)
    {
        // Call base.OnStartup first to ensure Application.Resources are initialized
        base.OnStartup(e);

        // Build dependency injection container
        _host = Host.CreateDefaultBuilder()
            .ConfigureServices((context, services) =>
            {
                // Register the Windows registry implementation
                services.AddSingleton<DotNetWindowsRegistry.IRegistry, DotNetWindowsRegistry.WindowsRegistry>();
                
                // Register registry service
                services.AddSingleton<IRegistryService, RegistryService>();
                
                // Register other services
                services.AddSingleton<AddinManager>();
                
                // Register the main window
                services.AddSingleton<MainWindow>();
            })
            .Build();

        // Start the host
        _host.Start();

        // Get the main window from DI container
        var mainWindow = _host.Services.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _host?.Dispose();
        base.OnExit(e);
    }
}

