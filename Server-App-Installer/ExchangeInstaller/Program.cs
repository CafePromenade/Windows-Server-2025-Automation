using Microsoft.Win32;
using OpenAI;
using OpenAI.Chat;
using OpenAI.Images;
using OpenAI.Responses;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static OpenQA.Selenium.PrintOptions;

namespace YourCompany.ExchangeDeployer
{
    class Program
    {
        // Disable close button and Ctrl+C
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        private static extern int GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        private static extern bool EnableMenuItem(int hMenu, int uIDEnableItem, int uEnable);
        private const int SC_CLOSE = 0xF060;
        private const int MF_GRAYED = 0x00000001;

        static async Task<int> Main(string[] args)
        {
            // Prevent console closure
            var handle = GetConsoleWindow();
            var menu = GetSystemMenu(handle, false);
            EnableMenuItem(menu, SC_CLOSE, MF_GRAYED);
            Console.TreatControlCAsInput = true;
            Console.CancelKeyPress += (s, e) => e.Cancel = true;

            // Args validation & logging setup
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: ExchangeDeployer <domainName> <netbiosName> <dsrmPassword> [exchangeSetupPath] <organizationName>");
                return -1;
            }
            var domainName = args[0];
            var netbiosName = args[1];
            var dsrmPassword = args[2];
            var exchangeSetupPath = args.Length >= 5 ? args[3] : @"C:\Setup-Software\Exchange";
            var organizationName = args.Length >= 5 ? args[4] : args[3];

            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var logDir = Path.Combine(baseDir, "Logs");
            Directory.CreateDirectory(logDir);
            var logFile = Path.Combine(logDir, "ExchangeDeploy.log");

            // AI-generated wallpaper
            Console.WriteLine("Generating AI wallpaper...");
            await Helpers.GenerateAndSetWallpaperAsync(logFile);

            // Progress reporter
            var progress = new Progress<ProgressReport>(r =>
                Console.Write($"[{r.Stage}] {r.Percent}% - {r.Message}\r"));

            var fixer = new OpenAIFixer(logFile);
            var deployer = new TouchlessExchangeAsyncDeployer(logFile, fixer, progress);

            int code = await deployer.DeployAsync(
                domainName, netbiosName, dsrmPassword, exchangeSetupPath, organizationName);

            Console.WriteLine();
            Console.WriteLine(code == 0 ? "✅ Deployment succeeded!" : "❌ Deployment failed.");
            return code;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    }

    public static class Helpers
    {
        /// <summary>
        /// Generates AI wallpaper using OpenAI DALL·E via the official SDK and sets it as desktop background.
        /// </summary>
        public static async Task GenerateAndSetWallpaperAsync(string logFile)
        {
            try
            {
                var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                if (string.IsNullOrWhiteSpace(apiKey))
                    throw new InvalidOperationException("OPENAI_API_KEY not set");

                var client = new OpenAIClient(apiKey);
                var imageClient = client.GetImageClient("dall-e-3");

                var options = new ImageGenerationOptions
                {
                    Quality = GeneratedImageQuality.High
                };

                // Async generation
                var imageResponse = await imageClient.GenerateImageAsync("A futuristic data center background with abstract glowing network lines",options);
                var image = imageResponse.Value;
                var path = Path.Combine(Path.GetTempPath(), "wallpaper.png");

                // Save resulting image bytes
                await File.WriteAllBytesAsync(path, image.ImageBytes.ToArray());

                SystemParametersInfo(20, 0, path, 3);
            }
            catch (Exception ex)
            {
                await File.AppendAllTextAsync(logFile, $"{DateTime.Now}: Wallpaper generation failed: {ex.Message}\r\n");
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    }

    public class ProgressReport
    {
        public string Stage { get; set; }
        public int Percent { get; set; }
        public string Message { get; set; }
    }

    public static class StateManager
    {
        private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "state.txt");
        public static int GetStage()
        {
            if (!File.Exists(FilePath))
                File.WriteAllText(FilePath, "0");
            return int.TryParse(File.ReadAllText(FilePath), out var s) ? s : 0;
        }
        public static void SetStage(int s) => File.WriteAllText(FilePath, s.ToString());
    }

    public class OpenAIFixer
    {
        private readonly OpenAIClient _client;
        private readonly string _log;

        public OpenAIFixer(string logFile)
        {
            _log = logFile;
            var key = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrWhiteSpace(key))
                throw new InvalidOperationException("OPENAI_API_KEY not set");
            _client = new OpenAIClient(key);
        }

        /// <summary>
        /// Uses GPT-3.5-turbo via ChatClient from SDK to auto-fix errors.
        /// </summary>
        public async Task<bool> TryAutoFixAsync(string stage, string error)
        {
            await File.AppendAllTextAsync(_log, $"{DateTime.Now}: Error at {stage}: {error}\r\n");

            var chatClient = _client.GetChatClient("o4-mini-2025-04-16");
            var messages = new List<ChatMessage>
            {
                new SystemChatMessage("You are a DevOps assistant."),
                new UserChatMessage($"Stage '{stage}' failed: {error}. Provide a PowerShell script to fix and exit code 0.")
            };

            OpenAIResponseClient client = new(
    model: "gpt-4o-mini",
    apiKey: Environment.GetEnvironmentVariable("OPENAI_API_KEY"));

            OpenAIResponse response = await client.CreateResponseAsync(
                userInputText: "What's a happy news headline from today?",
                new ResponseCreationOptions()
                {
                    Tools = { ResponseTool.CreateWebSearchTool() },
                });


            var options = new ChatCompletionOptions { MaxOutputTokenCount = 512 };
            //var response = await chatClient.CompleteChatAsync(messages, options);
            var script = "";
            var fixPath = Path.Combine(Path.GetDirectoryName(_log), $"fix_{stage}.ps1");

            foreach (ResponseItem item in response.OutputItems)
            {
                if (item is WebSearchCallResponseItem webSearchCall)
                {
                    Console.WriteLine($"[Web search invoked]({webSearchCall.Status}) {webSearchCall.Id}");
                }
                else if (item is MessageResponseItem message)
                {
                    script = message.Content?.FirstOrDefault()?.Text;
                    Console.WriteLine($"[{message.Role}] {message.Content?.FirstOrDefault()?.Text}");
                }
            }

            await File.WriteAllTextAsync(fixPath, script);

            var psi = new ProcessStartInfo("powershell.exe", $"-NoProfile -ExecutionPolicy Bypass -File \"{fixPath}\"")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            using var proc = Process.Start(psi);
            var output = await proc.StandardOutput.ReadToEndAsync();
            var err = await proc.StandardError.ReadToEndAsync();
            proc.WaitForExit();
            await File.AppendAllTextAsync(_log, output + err);

            if (proc.ExitCode == 0)
                Process.Start("shutdown", "/r /t 5 /f");

            return proc.ExitCode == 0;
        }
    }

    public class TouchlessExchangeAsyncDeployer
    {
        private readonly string _log;
        private readonly OpenAIFixer _fixer;
        private readonly IProgress<ProgressReport> _progress;

        public TouchlessExchangeAsyncDeployer(string log, OpenAIFixer fixer, IProgress<ProgressReport> progress)
        {
            _log = log;
            _fixer = fixer;
            _progress = progress;
        }

        public async Task<int> DeployAsync(string domain, string netbios, string pwd, string setupPath, string org)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_log));
            var stages = new List<Func<Task>>
            {
                () => StageAsync("OptimizeSystem", OptimizeSystemAsync),
                () => StageAsync("SystemConfig", ConfigureSystemAsync),
                () => StageAsync("Prereqs", InstallPrereqsAsync),
                () => StageAsync("ADDSForest", () => ConfigureADAsync(domain, netbios, pwd)),
                () => StageAsync("ExplorerSetup", ConfigureExplorerAsync),
                () => StageAsync("UserGeneration", GenerateUsersAsync),
                () => StageAsync("ExchangeInstall", () => InstallExchangeAsync(setupPath, org, domain)),
                () => StageAsync("MailboxTasks", () => ConfigureMailboxesAsync(domain)),
                () => StageAsync("ChromeConfig", ConfigureChromeAsync),
                () => StageAsync("Finalize", FinalizeAsync)
            };
            int s = StateManager.GetStage();
            try
            {
                for (; s < stages.Count; s++)
                {
                    StateManager.SetStage(s);
                    await stages[s]();
                }
            }
            catch
            {
                return -1;
            }
            return 0;
        }

        private async Task StageAsync(string name, Func<Task> action)
        {
            try { _progress.Report(new ProgressReport { Stage = name, Percent = 0, Message = "Starting" }); await action(); }
            catch (Exception ex)
            {
                bool fixedOk = await _fixer.TryAutoFixAsync(name, ex.ToString());
                if (!fixedOk) throw;
            }
        }

        private Task OptimizeSystemAsync()
        {
            _progress.Report(new ProgressReport { Stage = "OptimizeSystem", Percent = 10, Message = "Applying AI optimizations" });
            return Task.CompletedTask;
        }
        private Task ConfigureSystemAsync()
        {
            _progress.Report(new ProgressReport { Stage = "SystemConfig", Percent = 20, Message = "Configuring system settings" });
            var cmds = new List<(string, string)>
            {
                ("tzutil", "/s \"Eastern Standard Time\""),
                ("reg",    "add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Terminal Server\" /v fDenyTSConnections /t REG_DWORD /d 0 /f")
            };
            return RunCommandsAsync("SystemConfig", cmds);
        }
        private Task InstallPrereqsAsync()
        {
            _progress.Report(new ProgressReport { Stage = "Prerequisites", Percent = 30, Message = "Installing prerequisites" });
            var cmds = new List<(string, string)>
            {
                ("powershell.exe", "-NoProfile -ExecutionPolicy Bypass -Command iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')); choco install netfx-4.8 vcredist2012 vcredist2013 UCMA4 -y --ignore-checksums")
            };
            return RunCommandsAsync("Prerequisites", cmds);
        }
        private Task ConfigureADAsync(string domain, string netbios, string pwd)
        {
            _progress.Report(new ProgressReport { Stage = "ADDSForest", Percent = 40, Message = "Setting up AD" });
            string ps = $"Install-WindowsFeature AD-Domain-Services -IncludeManagementTools; Import-Module ADDSDeployment; $dsrm=ConvertTo-SecureString '{pwd}' -AsPlainText -Force; Install-ADDSForest -DomainName '{domain}' -DomainNetbiosName '{netbios}' -SafeModeAdministratorPassword $dsrm -InstallDns -NoRebootOnCompletion -Force;";
            return RunCommandsAsync("ADDSForest", new List<(string, string)> { ("powershell.exe", $"-NoProfile -ExecutionPolicy Bypass -Command \"{ps}\"") });
        }
        private Task ConfigureExplorerAsync()
        {
            _progress.Report(new ProgressReport { Stage = "ExplorerSetup", Percent = 50, Message = "Configuring Explorer" });
            string key = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced";
            var cmds = new List<(string, string)>
            {
                ("reg", $"add \"{key}\" /v Hidden /t REG_DWORD /d 1 /f"),
                ("powercfg","/change standby-timeout-ac 0"), ("powercfg","-hibernate off")
            };
            return RunCommandsAsync("ExplorerSetup", cmds);
        }
        private Task GenerateUsersAsync()
        {
            _progress.Report(new ProgressReport { Stage = "UserGeneration", Percent = 60, Message = "Creating users" });
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var file = Path.Combine(desktop, "users.txt");
            var sb = new StringBuilder();
            for (int i = 1; i <= 10; i++) { var u = $"user{i}"; var p = Guid.NewGuid().ToString("n").Substring(0, 12); sb.AppendLine($"{u}:{p}"); Process.Start("net", $"user {u} {p} /add"); }
            File.WriteAllText(file, sb.ToString());
            return Task.CompletedTask;
        }
        private Task InstallExchangeAsync(string path, string org, string dc)
        {
            _progress.Report(new ProgressReport { Stage = "ExchangeInstall", Percent = 70, Message = "Installing Exchange" });
            return RunCommandsAsync("ExchangeInstall", new List<(string, string)> { (Path.Combine(path, "Setup.exe"), $"/mode:Install /roles:Mailbox,ClientAccess /OrganizationName:{org} /DomainController:{dc} /IAcceptExchangeServerLicenseTerms") });
        }
        private Task ConfigureMailboxesAsync(string domain) => Task.CompletedTask;
        private Task ConfigureChromeAsync() => RunCommandsAsync("ChromeConfig", new List<(string, string)> { ("reg", "add ...") });
        private Task FinalizeAsync() => RunCommandsAsync("Finalize", new List<(string, string)> { ("shutdown", "/r /t 5 /f") });
        private async Task RunCommandsAsync(string stage, List<(string Cmd, string Args)> cmds)
        {
            int total = cmds.Count;
            for (int i = 0; i < total; i++)
            {
                var (c, a) = cmds[i];
                _progress.Report(new ProgressReport { Stage = stage, Percent = 10 + (i + 1) * (80 / total), Message = $"{c} {a}" });
                var psi = new ProcessStartInfo(c, a) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
                using var proc = Process.Start(psi);
                var outp = await proc.StandardOutput.ReadToEndAsync(); var err = await proc.StandardError.ReadToEndAsync(); proc.WaitForExit();
                await File.AppendAllTextAsync(_log, $"{DateTime.Now}[{stage}]{c} {a}\r\n{outp}\r\n{err}\r\n");
                if (proc.ExitCode != 0) throw new InvalidOperationException($"{c} failed {proc.ExitCode}");
            }
        }
    }
}
