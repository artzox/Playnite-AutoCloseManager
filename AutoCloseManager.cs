// AutoCloseManager.cs
using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Collections.Generic;
using Playnite.SDK;
using Playnite.SDK.Events;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;

namespace AutoCloseManagerPlugin
{
    public class AutoCloseManagerSettings : ObservableObject, ISettings
    {
        public bool EnableAutoClose { get; set; } = true;
        public int CloseTimeoutSeconds { get; set; } = 5;
        public bool ShowNotifications { get; set; } = true;

        // Playnite serializes settings object to a JSON object and saves it as text file.
        // If you want to exclude some property from being saved then use `JsonDontSerialize` ignore attribute.
        public AutoCloseManagerSettings()
        {
        }

        public AutoCloseManagerSettings(AutoCloseManagerSettings source)
        {
            EnableAutoClose = source.EnableAutoClose;
            CloseTimeoutSeconds = source.CloseTimeoutSeconds;
            ShowNotifications = source.ShowNotifications;
        }

        public void BeginEdit()
        {
            // Code executed when settings view is opened and user starts editing them.
        }

        public void CancelEdit()
        {
            // Code executed when user decides to cancel any changes made since BeginEdit was called.
            // This method should revert any changes made to Option1 and Option2.
        }

        public void EndEdit()
        {
            // Code executed when user decides to confirm changes made since BeginEdit was called.
            // This method should save settings made to Option1 and Option2.
        }

        public bool VerifySettings(out List<string> errors)
        {
            // Code execute when user decides to confirm changes made since BeginEdit was called.
            // Executed before EndEdit is called and EndEdit is not called if false is returned.
            // List of errors is presented to user if verification fails.
            errors = new List<string>();
            return true;
        }
    }

    public class AutoCloseManager : GenericPlugin
    {
        private static readonly ILogger logger = LogManager.GetLogger();
        private AutoCloseManagerSettings settings { get; set; }

        public override Guid Id { get; } = Guid.Parse("12345678-1234-5678-9012-123456789012"); // Generate a unique GUID

        public AutoCloseManager(IPlayniteAPI api) : base(api)
        {
            settings = new AutoCloseManagerSettings();
            Properties = new GenericPluginProperties
            {
                HasSettings = true
            };
        }

        public override void OnGameInstalled(OnGameInstalledEventArgs args)
        {
            // Add code to be executed when game is finished installing.
        }

        public override void OnGameStarted(OnGameStartedEventArgs args)
        {
            // Add code to be executed when game is started running.
        }

        public override void OnGameStarting(OnGameStartingEventArgs args)
        {
            // Add code to be executed when game is preparing to be started.
            if (settings.EnableAutoClose)
            {
                OnGameStartingHandler(args.Game);
            }
        }

        public override void OnGameStopped(OnGameStoppedEventArgs args)
        {
            // Add code to be executed when game is stopped running.
        }

        public override void OnGameUninstalled(OnGameUninstalledEventArgs args)
        {
            // Add code to be executed when game is uninstalled.
        }

        public override void OnApplicationStarted(OnApplicationStartedEventArgs args)
        {
            // Add code to be executed when Playnite is initialized.
            logger.Info("AutoCloseManager plugin started");
        }

        public override void OnApplicationStopped(OnApplicationStoppedEventArgs args)
        {
            // Add code to be executed when Playnite is shutting down.
            logger.Info("AutoCloseManager plugin stopped");
        }

        public override void OnLibraryUpdated(OnLibraryUpdatedEventArgs args)
        {
            // Add code to be executed when library is updated.
        }

        private async void OnGameStartingHandler(Game newGame)
        {
            try
            {
                logger.Info($"Game starting: {newGame.Name}");

                // Check if there's already a running game
                var runningGames = PlayniteApi.Database.Games
                    .Where(g => g.IsRunning && g.Id != newGame.Id)
                    .ToList();

                if (runningGames.Any())
                {
                    logger.Info($"Found {runningGames.Count} running game(s). Attempting to close them.");

                    if (settings.ShowNotifications)
                    {
                        PlayniteApi.Notifications.Add(new NotificationMessage(
                            "auto-close-info",
                            $"Closing {runningGames.Count} running game(s) to start {newGame.Name}",
                            NotificationType.Info
                        ));
                    }

                    foreach (var runningGame in runningGames)
                    {
                        await CloseRunningGame(runningGame);
                    }
                }

                // Wait a moment for processes to fully close
                await Task.Delay(1000);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in OnGameStartingHandler");
            }
        }

        private async Task CloseRunningGame(Game game)
        {
            try
            {
                logger.Info($"Attempting to close running game: {game.Name}");

                // Try multiple methods to close the game
                bool closed = false;

                // Method 1: Try to find and close by window title
                closed = await CloseByWindowTitle(game);

                // Method 2: Try to close by process name if window method failed
                if (!closed)
                {
                    closed = await CloseByProcessName(game);
                }

                // Method 3: Force close by executable path
                if (!closed)
                {
                    closed = await ForceCloseByExecutable(game);
                }

                if (closed)
                {
                    logger.Info($"Successfully closed game: {game.Name}");

                    if (settings.ShowNotifications)
                    {
                        PlayniteApi.Notifications.Add(new NotificationMessage(
                            "auto-close-success",
                            $"Closed {game.Name}",
                            NotificationType.Info
                        ));
                    }
                }
                else
                {
                    logger.Warn($"Failed to close game: {game.Name}");

                    if (settings.ShowNotifications)
                    {
                        PlayniteApi.Notifications.Add(new NotificationMessage(
                            "auto-close-failed",
                            $"Failed to close {game.Name}",
                            NotificationType.Error
                        ));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error closing game: {game.Name}");
            }
        }

        private async Task<bool> CloseByWindowTitle(Game game)
        {
            try
            {
                var processes = Process.GetProcesses()
                    .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle) &&
                               (p.MainWindowTitle.IndexOf(game.Name, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                game.Name.IndexOf(p.MainWindowTitle, StringComparison.OrdinalIgnoreCase) >= 0))
                    .ToList();

                foreach (var process in processes)
                {
                    try
                    {
                        logger.Info($"Closing process by window title: {process.ProcessName} ({process.MainWindowTitle})");

                        // Try graceful close first
                        process.CloseMainWindow();

                        // Wait for graceful close
                        if (await WaitForProcessExit(process, settings.CloseTimeoutSeconds * 1000))
                        {
                            return true;
                        }

                        // Force kill if still running
                        if (!process.HasExited)
                        {
                            process.Kill();
                            await WaitForProcessExit(process, 3000);
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Error closing process: {process.ProcessName}");
                    }
                }

                return processes.Any();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in CloseByWindowTitle");
                return false;
            }
        }

        private async Task<bool> CloseByProcessName(Game game)
        {
            try
            {
                if (game.GameActions?.Any() == true)
                {
                    foreach (var action in game.GameActions)
                    {
                        if (!string.IsNullOrEmpty(action.Path))
                        {
                            var processName = System.IO.Path.GetFileNameWithoutExtension(action.Path);
                            var processes = Process.GetProcessesByName(processName);

                            foreach (var process in processes)
                            {
                                try
                                {
                                    logger.Info($"Closing process by name: {process.ProcessName}");

                                    process.CloseMainWindow();
                                    if (await WaitForProcessExit(process, settings.CloseTimeoutSeconds * 1000))
                                    {
                                        return true;
                                    }

                                    if (!process.HasExited)
                                    {
                                        process.Kill();
                                        await WaitForProcessExit(process, 3000);
                                        return true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    logger.Error(ex, $"Error closing process: {process.ProcessName}");
                                }
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in CloseByProcessName");
                return false;
            }
        }

        private async Task<bool> ForceCloseByExecutable(Game game)
        {
            try
            {
                if (game.GameActions?.Any() == true)
                {
                    foreach (var action in game.GameActions.Where(a => !string.IsNullOrEmpty(a.Path)))
                    {
                        var processes = Process.GetProcesses()
                            .Where(p =>
                            {
                                try
                                {
                                    return p.MainModule?.FileName?.Equals(action.Path, StringComparison.OrdinalIgnoreCase) == true;
                                }
                                catch
                                {
                                    return false;
                                }
                            })
                            .ToList();

                        foreach (var process in processes)
                        {
                            try
                            {
                                logger.Info($"Force closing process: {process.ProcessName}");
                                process.Kill();
                                await WaitForProcessExit(process, 3000);
                                return true;
                            }
                            catch (Exception ex)
                            {
                                logger.Error(ex, $"Error force closing process: {process.ProcessName}");
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in ForceCloseByExecutable");
                return false;
            }
        }

        private async Task<bool> WaitForProcessExit(Process process, int timeoutMs)
        {
            try
            {
                return await Task.Run(() =>
                {
                    return process.WaitForExit(timeoutMs);
                });
            }
            catch
            {
                return true; // Assume it exited if we can't check
            }
        }

        public override ISettings GetSettings(bool firstRunSettings)
        {
            return settings;
        }
    }
}