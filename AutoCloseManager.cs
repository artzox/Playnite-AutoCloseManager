using System;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using Playnite.SDK;
using Playnite.SDK.Events;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace AutoCloseManagerPlugin
{
    public class AutoCloseManagerSettings : ObservableObject, ISettings
    {
        private readonly AutoCloseManager plugin;

        public bool EnableAutoClose { get; set; } = true;
        public int GracefulCloseTimeoutSeconds { get; set; } = 10;
        public bool ShowNotifications { get; set; } = true;
        public int PreCloseDelayMs { get; set; } = 500;
        public int SteamGameStartDelayMs { get; set; } = 50;
        public int NonSteamGameStartDelayMs { get; set; } = 400;
        public int ProcessCleanupWaitMs { get; set; } = 1000;

        public AutoCloseManagerSettings()
        {
        }

        public AutoCloseManagerSettings(AutoCloseManager plugin, AutoCloseManagerSettings source)
        {
            this.plugin = plugin;
            EnableAutoClose = source.EnableAutoClose;
            GracefulCloseTimeoutSeconds = source.GracefulCloseTimeoutSeconds;
            ShowNotifications = source.ShowNotifications;
            PreCloseDelayMs = source.PreCloseDelayMs;
            SteamGameStartDelayMs = source.SteamGameStartDelayMs;
            NonSteamGameStartDelayMs = source.NonSteamGameStartDelayMs;
            ProcessCleanupWaitMs = source.ProcessCleanupWaitMs;
        }

        public void BeginEdit() { }
        public void CancelEdit() { }

        public void EndEdit()
        {
            // Now, we call a method in the main plugin instance to handle the save and update
            plugin.UpdateSettings(this);
        }

        public bool VerifySettings(out List<string> errors)
        {
            errors = new List<string>();
            return true;
        }
    }

    public class AutoCloseManager : GenericPlugin
    {
        private static readonly ILogger logger = LogManager.GetLogger();
        private AutoCloseManagerSettings settings { get; set; }
        private readonly ProcessFinder processFinder;

        public override Guid Id { get; } = Guid.Parse("9999e785-f30e-4506-bc63-a703fa80a59d");

        public AutoCloseManager(IPlayniteAPI api) : base(api)
        {
            logger.Info("AutoCloseManager plugin has been initialized.");
            settings = LoadPluginSettings<AutoCloseManagerSettings>() ?? new AutoCloseManagerSettings();
            Properties = new GenericPluginProperties
            {
                HasSettings = true
            };
            this.processFinder = new ProcessFinder();
        }

        // ADDED METHOD: This method will be called to update the internal settings
        public void UpdateSettings(AutoCloseManagerSettings newSettings)
        {
            settings = newSettings;
            SavePluginSettings(settings);
            logger.Info("AutoCloseManager settings have been updated in memory and saved to disk.");
        }

        public override void OnGameStarting(OnGameStartingEventArgs args)
        {
            logger.Info($"AutoClose: OnGameStarting event triggered for '{args.Game.Name}'.");
            if (!settings.EnableAutoClose)
            {
                return;
            }

            var task = DoAutoCloseLogicAsync(args.Game);
            var frame = new DispatcherFrame();
            task.ContinueWith(t =>
            {
                frame.Continue = false;
            }, TaskScheduler.Default);
            Dispatcher.PushFrame(frame);
        }

        public override void OnApplicationStarted(OnApplicationStartedEventArgs args)
        {
            logger.Info("AutoCloseManager plugin started");
        }

        public override void OnApplicationStopped(OnApplicationStoppedEventArgs args)
        {
            logger.Info("AutoCloseManager plugin stopped");
        }

        private async Task DoAutoCloseLogicAsync(Game newGame)
        {
            try
            {
                var runningGames = PlayniteApi.Database.Games
                    .Where(g => g.IsRunning && g.Id != newGame.Id)
                    .ToList();

                if (!runningGames.Any())
                {
                    logger.Info("AutoClose: No other games detected as running. No action needed.");
                    return;
                }

                logger.Info($"AutoClose: New game starting: {newGame.Name}");
                await Task.Delay(settings.PreCloseDelayMs);

                logger.Info($"AutoClose: Detected {runningGames.Count} other games still running. Closing them now.");
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

                logger.Info($"AutoClose: Waiting {settings.ProcessCleanupWaitMs}ms for process cleanup...");
                await Task.Delay(settings.ProcessCleanupWaitMs);

                bool isSteamGame = IsSteamGame(newGame);
                int gameTypeDelay = isSteamGame ? settings.SteamGameStartDelayMs : settings.NonSteamGameStartDelayMs;

                logger.Info($"AutoClose: {(isSteamGame ? "Steam" : "Non-Steam")} game detected. Delaying game launch for {gameTypeDelay}ms...");
                await Task.Delay(gameTypeDelay);

                logger.Info($"AutoClose: Finished closing games and waiting. Now allowing {newGame.Name} to launch.");

                try
                {
                    newGame.LastActivity = DateTime.Now;
                    PlayniteApi.Database.Games.Update(newGame);
                    logger.Info($"AutoClose: Nudged NowPlaying by updating LastActivity for {newGame.Name}.");
                }
                catch (Exception ex)
                {
                    logger.Warn(ex, $"AutoClose: Failed to update LastActivity for {newGame.Name}.");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in DoAutoCloseLogicAsync");
            }
        }

        private bool IsSteamGame(Game game)
        {
            try
            {
                if (game.Source?.Name.IndexOf("Steam", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
                if (game.GameActions?.Any(a => a.Path?.IndexOf("steam://", StringComparison.OrdinalIgnoreCase) >= 0) == true)
                {
                    return true;
                }
                if (!string.IsNullOrEmpty(game.InstallDirectory) && game.InstallDirectory.IndexOf("steamapps", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                logger.Debug($"Error determining if game is Steam-based: {ex.Message}");
                return false;
            }
        }

        private async Task CloseRunningGame(Game game)
        {
            try
            {
                logger.Info($"Attempting to close running game: {game.Name}");
                bool closed = false;

                var processToClose = processFinder.FindRunningGameProcess(game, null);

                if (processToClose != null)
                {
                    await Task.Run(() =>
                    {
                        logger.Info($"Attempting graceful close for process: {processToClose.ProcessName} (ID: {processToClose.Id})");
                        processToClose.CloseMainWindow();

                        if (WaitForProcessExit(processToClose, settings.GracefulCloseTimeoutSeconds * 1000))
                        {
                            closed = true;
                            logger.Info($"Process {processToClose.ProcessName} closed gracefully.");
                        }

                        if (!closed && !processToClose.HasExited)
                        {
                            logger.Info("Graceful close failed. Forcefully killing the process.");
                            processToClose.Kill();
                            WaitForProcessExit(processToClose, 3000);
                            closed = true;
                        }
                    });

                    try
                    {
                        game.IsRunning = false;
                        PlayniteApi.Database.Games.Update(game);
                        logger.Info($"Updated Playnite database: {game.Name} is no longer running.");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to update game state in database for {game.Name}");
                    }
                }

                if (closed)
                {
                    logger.Info($"Successfully closed game: {game.Name}");
                    if (settings.ShowNotifications)
                    {
                        PlayniteApi.Notifications.Add(new NotificationMessage(
                            "auto-close-success", $"Closed {game.Name}", NotificationType.Info));
                    }
                }
                else
                {
                    logger.Warn($"Failed to close game: {game.Name}");
                    if (settings.ShowNotifications)
                    {
                        PlayniteApi.Notifications.Add(new NotificationMessage(
                            "auto-close-failed", $"Failed to close {game.Name}", NotificationType.Error));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error closing game: {game.Name}");
            }
        }

        private bool WaitForProcessExit(Process process, int timeoutMs)
        {
            try
            {
                return process.WaitForExit(timeoutMs);
            }
            catch
            {
                return true;
            }
        }

        public override ISettings GetSettings(bool firstRunSettings)
        {
            return new AutoCloseManagerSettings(this, settings);
        }

        public override System.Windows.Controls.UserControl GetSettingsView(bool firstRunSettings)
        {
            return new AutoCloseManagerSettingsView();
        }
    }

    public class ProcessFinder
    {
        private static readonly ILogger logger = LogManager.GetLogger();

        public Process FindRunningGameProcess(Game game, int? pid)
        {
            if (game == null) return null;
            try
            {
                var candidateProcesses = Process.GetProcesses().Where(p =>
                {
                    try
                    {
                        return !p.HasExited &&
                               p.MainWindowHandle != IntPtr.Zero &&
                               !string.IsNullOrEmpty(p.MainWindowTitle) &&
                               p.WorkingSet64 > 100 * 1024 * 1024;
                    }
                    catch { return false; }
                }).ToList();
                logger.Debug($"Found {candidateProcesses.Count} candidate processes with windows");
                var gameExecutables = GetGameExecutables(game);
                string[] gameNameWords = game.Name.ToLower().Split(new char[] { ' ', '-', '_', ':', '.', '(', ')', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                logger.Debug($"Game name for matching: {game.Name}, split into {gameNameWords.Length} words");
                var candidates = new List<Process>();
                var nameMatchCandidates = new List<Process>();
                var titleMatchCandidates = new List<(Process Process, int MatchCount)>();
                var inaccessibleCandidates = new List<Process>();
                foreach (var p in candidateProcesses)
                {
                    try
                    {
                        bool nameMatches = gameExecutables.Any(exe => string.Equals(exe, p.ProcessName, StringComparison.OrdinalIgnoreCase));
                        int titleMatchScore = CalculateTitleMatchScore(p, gameNameWords);
                        string modulePath = null;
                        try { modulePath = p.MainModule.FileName; }
                        catch { modulePath = GetProcessPath(p); }
                        if (!string.IsNullOrEmpty(modulePath) && !string.IsNullOrEmpty(game.InstallDirectory) && modulePath.ToLower().IndexOf(game.InstallDirectory.ToLower()) >= 0)
                        {
                            candidates.Add(p);
                            logger.Debug($"Found process with matching path: {p.ProcessName} -> {modulePath}");
                        }
                        else if (nameMatches) { nameMatchCandidates.Add(p); }
                        else if (titleMatchScore > 0) { titleMatchCandidates.Add((p, titleMatchScore)); }
                        else if (string.IsNullOrEmpty(modulePath)) { inaccessibleCandidates.Add(p); }
                    }
                    catch { /* Skip processes we can't access at all */ }
                }
                var bestMatch = FindBestMatch(candidates, titleMatchCandidates, nameMatchCandidates, inaccessibleCandidates, pid);
                return bestMatch;
            }
            catch (Exception ex)
            {
                logger.Error($"Error finding game process: {ex.Message}");
                return null;
            }
        }

        private string GetProcessPath(Process process)
        {
            try
            {
                IntPtr handle = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, false, process.Id);
                if (handle == IntPtr.Zero) return null;
                try
                {
                    var buffer = new StringBuilder(1024);
                    uint size = (uint)buffer.Capacity;
                    if (QueryFullProcessImageName(handle, 0, buffer, ref size))
                    {
                        return buffer.ToString();
                    }
                }
                finally { CloseHandle(handle); }
            }
            catch (Exception ex) { logger.Debug($"Error getting process path for {process.ProcessName}: {ex.Message}"); }
            return null;
        }

        private static int CalculateTitleMatchScore(Process process, string[] gameNameWords)
        {
            if (process.MainWindowHandle == IntPtr.Zero || string.IsNullOrEmpty(process.MainWindowTitle)) return 0;
            string[] windowTitleWords = process.MainWindowTitle.ToLower().Split(new char[] { ' ', '-', '_', ':', '.', '(', ')', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
            return gameNameWords.Count(gameWord => windowTitleWords.Any(titleWord => titleWord.IndexOf(gameWord) >= 0 || gameWord.IndexOf(titleWord) >= 0));
        }

        private static Process FindBestMatch(List<Process> candidates, List<(Process Process, int MatchCount)> titleMatchCandidates, List<Process> nameMatchCandidates, List<Process> inaccessibleCandidates, int? originalPid)
        {
            if (candidates.Count > 0)
            {
                var bestMatch = candidates.OrderByDescending(p => p.WorkingSet64).First();
                logger.Debug($"Found process with matching path: {bestMatch.ProcessName} (ID: {bestMatch.Id})");
                return bestMatch;
            }
            if (nameMatchCandidates.Count > 0)
            {
                var bestMatch = nameMatchCandidates.OrderByDescending(p => p.WorkingSet64).First();
                logger.Debug($"Found process with matching name: {bestMatch.ProcessName} (ID: {bestMatch.Id})");
                return bestMatch;
            }
            if (titleMatchCandidates.Count > 0)
            {
                var bestMatch = titleMatchCandidates.OrderByDescending(t => t.MatchCount).ThenByDescending(t => t.Process.WorkingSet64).First().Process;
                logger.Debug($"Found process with matching window title: {bestMatch.ProcessName} (ID: {bestMatch.Id}, Title: {bestMatch.MainWindowTitle})");
                return bestMatch;
            }
            if (inaccessibleCandidates.Count > 0)
            {
                var processFromPlaynite = inaccessibleCandidates.FirstOrDefault(p => p.Id == originalPid);
                if (processFromPlaynite != null)
                {
                    logger.Debug($"Found original process: {processFromPlaynite.ProcessName} (ID: {processFromPlaynite.Id})");
                    return processFromPlaynite;
                }
                var bestGuess = inaccessibleCandidates.OrderByDescending(p => p.WorkingSet64).First();
                logger.Debug($"Best guess from inaccessible processes: {bestGuess.ProcessName} (ID: {bestGuess.Id})");
                return bestGuess;
            }
            return null;
        }

        private static List<string> GetGameExecutables(Game game)
        {
            var gameExecutables = new List<string>();
            try
            {
                if (!string.IsNullOrEmpty(game.InstallDirectory) && Directory.Exists(game.InstallDirectory))
                {
                    gameExecutables = Directory.GetFiles(game.InstallDirectory, "*.exe", SearchOption.AllDirectories).Select(path => Path.GetFileNameWithoutExtension(path)).ToList();
                    var additionalNames = new List<string>();
                    foreach (var exe in gameExecutables)
                    {
                        additionalNames.Add(exe.Replace("-", ""));
                        additionalNames.Add(exe.Replace("_", ""));
                        additionalNames.Add(exe.Replace(" ", ""));
                        additionalNames.Add(exe.Replace("-", " "));
                        additionalNames.Add(exe.Replace("_", " "));
                    }
                    gameExecutables.AddRange(additionalNames);
                    logger.Debug($"Found {gameExecutables.Count} potential game executables in {game.InstallDirectory}");
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Error scanning game directory: {ex.Message}");
            }
            return gameExecutables;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(ProcessAccessFlags processAccess, bool bInheritHandle, int processId);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool QueryFullProcessImageName(IntPtr hProcess, uint dwFlags, StringBuilder lpExeName, ref uint lpdwSize);
        [Flags]
        private enum ProcessAccessFlags : uint
        {
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x1000
        }
    }
}