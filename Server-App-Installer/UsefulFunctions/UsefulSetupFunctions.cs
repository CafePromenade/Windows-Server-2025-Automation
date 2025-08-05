using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace UsefulFunctions
{
    public static class UsefulSetupFunctions
    {
        /// <summary>
        /// Fully automates Exchange Server deployment:
        /// - Chocolatey bootstrap & prerequisites
        /// - AD DS role + new forest promotion
        /// - Scheduled Exchange Setup at startup
        /// - Automatic reboot
        /// </summary>
        public static class ExchangeDeployment
        {
            /// <summary>
            /// Runs a single, touchless setup for Exchange on a fresh server.
            /// </summary>
            /// <param name="domainName">FQDN of the new AD forest (e.g. "corp.contoso.com")</param>
            /// <param name="netbiosName">NetBIOS name (≤15 chars, e.g. "CONTOSO")</param>
            /// <param name="dsrmPassword">DSRM (Safe Mode) password in plain text</param>
            /// <param name="exchangeSetupPath">Local path to Exchange Setup.exe (e.g. "C:\\ExchangeSetup")</param>
            /// <param name="organizationName">Exchange organization name</param>
            /// <returns>PowerShell exit code (0 if scheduled & rebooting)</returns>
            public static int SetupExchangeServer(
                string domainName,
                string netbiosName,
                string dsrmPassword,
                string exchangeSetupPath,
                string organizationName)
            {
                if (string.IsNullOrWhiteSpace(domainName))
                    throw new ArgumentException("Domain name is required", nameof(domainName));
                if (string.IsNullOrWhiteSpace(netbiosName))
                    throw new ArgumentException("NetBIOS name is required", nameof(netbiosName));
                if (dsrmPassword == null)
                    throw new ArgumentNullException(nameof(dsrmPassword));
                if (string.IsNullOrWhiteSpace(exchangeSetupPath))
                    throw new ArgumentException("Exchange setup path is required", nameof(exchangeSetupPath));
                if (string.IsNullOrWhiteSpace(organizationName))
                    throw new ArgumentException("Organization name is required", nameof(organizationName));

                // Build the PowerShell script
                var psScript = $@"
# ☁️ Bootstrap Chocolatey & install prerequisites
Set-ExecutionPolicy Bypass -Scope Process -Force;
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072;
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {{
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'));
}}
choco install netfx-4.8 vcredist2012 vcredist2013 UCMA4 -y --ignore-checksums;

# 🌳 Install AD DS role & promote to new forest
Install-WindowsFeature AD-Domain-Services -IncludeManagementTools;
Import-Module ADDSDeployment;
$dsrmPwd = ConvertTo-SecureString '{dsrmPassword}' -AsPlainText -Force;
Install-ADDSForest -DomainName '{domainName}' -DomainNetbiosName '{netbiosName}' `
                   -SafeModeAdministratorPassword $dsrmPwd `
                   -InstallDns -NoRebootOnCompletion -Force;

# 📅 Schedule Exchange Setup on next boot
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument @(
    '-NoProfile','-ExecutionPolicy','Bypass','-Command',
    'Start-Process -FilePath ""{exchangeSetupPath}\\Setup.exe"" `
      -ArgumentList ""/mode:Install"",""/roles:Mailbox,ClientAccess"",""/OrganizationName:{organizationName}"",""/DomainController:{domainName}"",""/IAcceptExchangeServerLicenseTerms"" `
      -Wait -Verb runAs; `
     Unregister-ScheduledTask -TaskName ""InstallExchange"" -Confirm:$false; `
     Restart-Computer -Force;'
);
$trigger = New-ScheduledTaskTrigger -AtStartup;
Register-ScheduledTask -TaskName 'InstallExchange' `
                       -Action $action -Trigger $trigger `
                       -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable) `
                       -RunLevel Highest -Force;

# 🔄 Reboot now to kick off Exchange installation
Restart-Computer -Force;
";

                // Launch elevated PowerShell
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psScript}\"",
                    Verb = "runas",              // Admin elevation
                    UseShellExecute = true,
                    CreateNoWindow = true
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                        throw new InvalidOperationException("Could not start PowerShell process.");
                    proc.WaitForExit();
                    return proc.ExitCode;
                }
            }
        }

        /// <summary>
        /// Automates installation of Exchange Server prerequisites via Chocolatey.
        /// </summary>
        public static class ExchangePrerequisitesInstaller
        {
            /// <summary>
            /// Bootstraps Chocolatey (if needed) and installs all required packages
            /// with auto-confirmation and checksum ignoring.
            /// </summary>
            /// <returns>
            /// Exit code from the PowerShell process (0 = success).
            /// </returns>
            public static int InstallExchangePrerequisites()
            {
                // 1️⃣ Chocolatey bootstrap script
                string bootstrap = @"
Set-ExecutionPolicy Bypass -Scope Process -Force;
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072;
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'));
}
";

                // 2️⃣ Exchange prerequisite packages
                string prerequisites = @"
choco install netfx-4.8 vcredist2012 vcredist2013 UCMA4 -y --ignore-checksums;
";

                // 3️⃣ Combine into full PowerShell command
                string fullScript = bootstrap + prerequisites;

                // 4️⃣ Prepare elevated PowerShell process
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{fullScript}\"",
                    Verb = "runas",      // prompts UAC for admin rights
                    UseShellExecute = true,
                    CreateNoWindow = true
                };

                // 5️⃣ Execute and return the exit code
                using (var process = Process.Start(psi))
                {
                    if (process == null)
                        throw new InvalidOperationException("Failed to start PowerShell process.");
                    process.WaitForExit();
                    return process.ExitCode;
                }
            }
        }

        /// <summary>
        /// Contains methods to programmatically promote a Windows Server
        /// to the first domain controller in a brand-new AD forest.
        /// </summary>
        public static class ActiveDirectoryPromoter
        {
            /// <summary>
            /// Promotes the current server to the first domain controller in a new forest.
            /// </summary>
            /// <param name="domainName">
            /// The FQDN of the new forest (e.g. "corp.contoso.com").
            /// </param>
            /// <param name="netbiosName">
            /// The NetBIOS name for the domain (single label ≤15 chars, e.g. "CONTOSO").
            /// </param>
            /// <param name="safeModePassword">
            /// The DSRM (Safe Mode) administrator password, as plain text.
            /// </param>
            /// <param name="installDns">
            /// Whether to install and configure DNS on this DC (default: true).
            /// </param>
            /// <param name="noRebootOnCompletion">
            /// Suppress auto-reboot after promotion (default: false).
            /// </param>
            /// <returns>
            /// The PowerShell process exit code (0 = success; non-zero = error).
            /// </returns>
            public static int PromoteToNewForest(
                string domainName,
                string netbiosName,
                string safeModePassword,
                bool installDns = true,
                bool noRebootOnCompletion = false)
            {
                if (string.IsNullOrWhiteSpace(domainName))
                    throw new ArgumentException("Domain name must be provided", nameof(domainName));
                if (string.IsNullOrWhiteSpace(netbiosName))
                    throw new ArgumentException("NetBIOS name must be provided", nameof(netbiosName));
                if (safeModePassword is null)
                    throw new ArgumentNullException(nameof(safeModePassword));

                // 1️⃣ Install AD-DS role
                string roleScript = @"
Install-WindowsFeature AD-Domain-Services -IncludeManagementTools
Import-Module ADDSDeployment
";

                // 2️⃣ Create secure string for DSRM password
                string pwdScript =
                    $"$dsrmPwd = ConvertTo-SecureString '{safeModePassword}' -AsPlainText -Force;";

                // 3️⃣ Build forest installation command
                string forestCmd =
                    $"Install-ADDSForest -DomainName \"{domainName}\" " +
                    $"-DomainNetbiosName \"{netbiosName}\" " +
                    "-SafeModeAdministratorPassword $dsrmPwd " +
                    (installDns ? "-InstallDns " : "") +
                    (noRebootOnCompletion ? "-NoRebootOnCompletion " : "") +
                    "-Force;";  // suppresses confirmation :contentReference[oaicite:1]{index=1}

                // 4️⃣ Combine into one PowerShell script
                string fullScript = roleScript + pwdScript + forestCmd;

                // 5️⃣ Setup elevated PowerShell process
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{fullScript}\"",
                    Verb = "runas",
                    UseShellExecute = true,
                    CreateNoWindow = true
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                        throw new InvalidOperationException("Failed to start PowerShell");
                    proc.WaitForExit();
                    return proc.ExitCode;
                }
            }
        }
        public static class Chocolatey
        {
            /// <summary>
            /// Installs Chocolatey (if not already installed) and then installs the specified packages
            /// with ignore-checksums and auto-confirm.
            /// </summary>
            /// <param name="packageIds">
            /// A collection of Chocolatey package IDs (e.g. "git", "7zip", "nodejs").
            /// Pass an empty collection to only bootstrap Chocolatey.
            /// </param>
            /// <returns>
            /// The exit code from the PowerShell process (0 = success; non-zero = error).
            /// </returns>
            public static int InstallPackages(IEnumerable<string> packageIds)
            {
                if (packageIds == null)
                    throw new ArgumentNullException(nameof(packageIds));

                // 1️⃣ Bootstrap Chocolatey if missing
                string bootstrapScript = @"
Set-ExecutionPolicy Bypass -Scope Process -Force;
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072;
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'));
}
";

                // 2️⃣ Build package install command only if any IDs were passed
                string installCommand = string.Empty;
                var list = packageIds.Where(id => !string.IsNullOrWhiteSpace(id))
                                     .Select(id => id.Trim())
                                     .ToList();

                if (list.Any())
                {
                    string joined = string.Join(" ", list);
                    installCommand = $"choco install {joined} -y --ignore-checksums;";
                }

                // 3️⃣ Combine into one PowerShell command
                string fullCommand = bootstrapScript + installCommand;

                // 4️⃣ Configure elevated PowerShell process
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{fullCommand}\"",
                    Verb = "runas",           // prompts for admin elevation
                    UseShellExecute = true,
                    CreateNoWindow = true
                };

                // 5️⃣ Execute and return exit code
                using (var process = Process.Start(psi))
                {
                    if (process == null)
                        throw new InvalidOperationException("Failed to start PowerShell process.");

                    process.WaitForExit();
                    return process.ExitCode;
                }
            }
        }
    }
}
