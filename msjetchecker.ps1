#region Script Parameters and File Manifest
[CmdletBinding()]
# The param block must be the first executable statement in the script.
param(
    # Specifies the target Striim version to determine correct dependencies. Required for --downloadonly or initial install.
    [Parameter()]
    [string]$Version,

    # If specified, the script will only download all potential dependencies into the 'downloads' folder and then exit.
    # Requires --version and either --agent or --node.
    [Parameter()]
    [switch]$DownloadOnly,

    # Used with --downloadonly to specify that Agent-specific files should be downloaded.
    [Parameter()]
    [switch]$Agent,

    # Used with --downloadonly to specify that Node-specific files should be downloaded.
    [Parameter()]
    [switch]$Node
)

# Set strict mode to catch common errors. This is placed after the param() block as required.
Set-StrictMode -Version Latest

# Central manifest of all downloadable files
# This allows for easy management and enables the --downloadonly feature.
# NodeType 'A' = Agent, 'N' = Node. Files without NodeType are common to both.
$AllDownloads = @(
    # Dependencies
    [pscustomobject]@{ Name = "icudt72.dll"; Url = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/main/MSJet/Dlls/icudt72.dll"; Category = "Dependency"; MinVersion = "0.0"; MaxVersion = "99.9" }
    [pscustomobject]@{ Name = "icuuc72.dll"; Url = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/main/MSJet/Dlls/icuuc72.dll"; Category = "Dependency"; MinVersion = "0.0"; MaxVersion = "99.9" }
    [pscustomobject]@{ Name = "MSSQLNative.dll"; Url = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/main/MSJet/FixesFor4.2.0.20/MSSQLNative.dll"; Category = "Dependency"; MinVersion = "0.0"; MaxVersion = "99.9" }

    # Java Versions
    [pscustomobject]@{ Name = "openlogic-openjdk-8u422-b05-windows-x64.msi"; Url = "https://builds.openlogic.com/downloadJDK/openlogic-openjdk/8u422-b05/openlogic-openjdk-8u422-b05-windows-x64.msi"; Category = "Java"; MinVersion = "0.0"; MaxVersion = "4.99" }
    [pscustomobject]@{ Name = "microsoft-jdk-11.0.26-windows-x64.msi"; Url = "https://aka.ms/download-jdk/microsoft-jdk-11.0.26-windows-x64.msi"; Category = "Java"; MinVersion = "5.0"; MaxVersion = "99.9" }

    # Security
    [pscustomobject]@{ Name = "sqljdbc_auth.dll"; Url = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/main/MSJet/sqljdbc_auth.dll"; Category = "Security"; MinVersion = "0.0"; MaxVersion = "99.9" }

    # Prerequisites
    [pscustomobject]@{ Name = "vc_redist.x64.exe"; Url = "https://aka.ms/vs/17/release/vc_redist.x64.exe"; Category = "Prereq"; MinVersion = "0.0"; MaxVersion = "99.9" }
    [pscustomobject]@{ Name = "msoledbsql.msi"; Url = "https://go.microsoft.com/fwlink/?linkid=2278907"; Category = "Prereq"; MinVersion = "0.0"; MaxVersion = "99.9" }

    # JDBC Drivers
    [pscustomobject]@{ Name = "mariadb-java-client-2.4.3.jar"; Url = "https://repo1.maven.org/maven2/org/mariadb/jdbc/mariadb-java-client/2.4.3/mariadb-java-client-2.4.3.jar"; Category = "Driver"; MinVersion = "0.0"; MaxVersion = "99.9"; NodeType = "A" }
    [pscustomobject]@{ Name = "mysql-connector-j-8.0.30.zip"; Url = "https://cdn.mysql.com/archives/mysql-connector-j/mysql-connector-j-8.0.30.zip"; Category = "Driver"; MinVersion = "0.0"; MaxVersion = "99.9"; NodeType = "A" }
    [pscustomobject]@{ Name = "instantclient-basic-windows.x64-21.6.0.0.0dbru.zip"; Url = "https://download.oracle.com/otn_software/nt/instantclient/216000/instantclient-basic-windows.x64-21.6.0.0.0dbru.zip"; Category = "Driver"; MinVersion = "0.0"; MaxVersion = "99.9"; NodeType = "A" }
    [pscustomobject]@{ Name = "postgresql-42.2.27.jar"; Url = "https://jdbc.postgresql.org/download/postgresql-42.2.27.jar"; Category = "Driver"; MinVersion = "0.0"; MaxVersion = "99.9"; NodeType = "A" }

    # Patches for 4.2.0.20
    [pscustomobject]@{ Name = "Platform_48036_v4.2.0.20_27_Sep_2024.jar"; Url = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/Platform_48036_v4.2.0.20_27_Sep_2024.jar"; Category = "Patch"; TargetFile = "Platform-4.2.0.20.jar"; MinVersion = "4.2.0.20"; MaxVersion = "4.2.0.20" }
    [pscustomobject]@{ Name = "MSJet_48036_v4.2.0.20_27_Sep_2024.jar"; Url = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/MSJet_48036_v4.2.0.20_27_Sep_2024.jar"; Category = "Patch"; TargetFile = "MSJet-4.2.0.20.jar"; MinVersion = "4.2.0.20"; MaxVersion = "4.2.0.20" }
    [pscustomobject]@{ Name = "SourceCommons_48036_v4.2.0.20_27_Sep_2024.jar"; Url = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/SourceCommons_48036_v4.2.0.20_27_Sep_2024.jar"; Category = "Patch"; TargetFile = "SourceCommons-4.2.0.20.jar"; MinVersion = "4.2.0.20"; MaxVersion = "4.2.0.20" }

    # Main Striim Application Installers
    [pscustomobject]@{ Name = "Striim_Agent_{VERSION}.zip"; Url = "https://striim-downloads.striim.com/Releases/{VERSION}/Striim_Agent_{VERSION}.zip"; Category = "MainInstaller"; NodeType = "A"; MinVersion = "0.0"; MaxVersion = "99.9" }
    [pscustomobject]@{ Name = "Striim_{VERSION}.zip"; Url = "https://striim-downloads.striim.com/Releases/{VERSION}/Striim_{VERSION}.zip"; Category = "MainInstaller"; NodeType = "N"; MinVersion = "0.0"; MaxVersion = "99.9" }

    # Windows Service Installers
    [pscustomobject]@{ Name = "Striim_windowsAgent_{VERSION}.zip"; Url = "https://striim-downloads.striim.com/Releases/{VERSION}/Striim_windowsAgent_{VERSION}.zip"; Category = "ServiceInstaller"; NodeType = "A"; MinVersion = "0.0"; MaxVersion = "99.9" }
    [pscustomobject]@{ Name = "Striim_windowsService_{VERSION}.zip"; Url = "https://striim-downloads.striim.com/Releases/{VERSION}/Striim_windowsService_{VERSION}.zip"; Category = "ServiceInstaller"; NodeType = "N"; MinVersion = "0.0"; MaxVersion = "99.9" }
)
#endregion

# --- Global Variables & Initial Path Setup ---
$striimInstallPath = Get-Location
$downloadDir = Join-Path -Path $striimInstallPath -ChildPath "downloads"
$global:AgentRestartNeeded = $false

# Create downloads folder if it doesn't exist. This is needed early for all modes.
if (-not (Test-Path $downloadDir)) {
    New-Item -ItemType Directory -Force -Path $downloadDir | Out-Null
}

# Wrap entire script in a try/catch for robust error handling
try {
    #region Helper Functions

    # Function to check if the script is running with Administrator privileges
    function Test-IsAdmin {
        return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    # Standardized function for user confirmation prompts
    function Confirm-UserChoice {
        param(
            [string]$Prompt,
            [string]$DefaultChoice = 'y'
        )
        $validChoices = 'y', 'n'
        if ($DefaultChoice.ToLower() -notin $validChoices) { $DefaultChoice = 'y' }

        $promptHint = if ($DefaultChoice.ToLower() -eq 'y') { '(Y/n)' } else { '(y/N)' }
        $fullPrompt = "[Confirm] $Prompt $promptHint"

        while ($true) {
            $response = Read-Host -Prompt $fullPrompt
            if ([string]::IsNullOrWhiteSpace($response)) {
                return $DefaultChoice.ToLower() -eq 'y'
            }
            if ($response.ToLower() -in $validChoices) {
                return $response.ToLower() -eq 'y'
            }
            Write-Warning "Invalid input. Please enter 'y' or 'n'."
        }
    }

    # Function to execute a command in a new, elevated PowerShell process
    function Invoke-AsAdmin {
        param(
            [string]$ArgumentList,
            [string]$WorkingDirectory
        )
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "powershell.exe"
        $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"$ArgumentList`""
        $psi.Verb = "RunAs"
        $psi.WorkingDirectory = $WorkingDirectory
        $process = [System.Diagnostics.Process]::Start($psi)
        $process.WaitForExit()
        return $process.ExitCode
    }

    # Robust function to download a file, with a fallback mechanism
    function Download-File {
        param(
            [string]$Uri,
            [string]$OutFilePath
        )
        if (Test-Path -Path $OutFilePath) {
            Write-Host "[Download] File already exists: $OutFilePath"
            return $true
        }

        Write-Host "[Download] Downloading from $Uri to $OutFilePath..."
        $httpClient = $null
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $httpClient = New-Object System.Net.Http.HttpClient
            $httpClient.Timeout = [TimeSpan]::FromMinutes(10)
            $response = $httpClient.GetAsync($Uri).Result
            if ($response.IsSuccessStatusCode) {
                $content = $response.Content.ReadAsByteArrayAsync().Result
                [System.IO.File]::WriteAllBytes($OutFilePath, $content)
                Write-Host "[Download] Success: File downloaded using HttpClient."
                return $true
            } else {
                Write-Warning "[Download] HttpClient download failed. HTTP Status: $($response.StatusCode). Attempting fallback..."
                throw "HttpClient failed"
            }
        }
        catch {
            try {
                Invoke-WebRequest -Uri $Uri -OutFile $OutFilePath -ErrorAction Stop -UseBasicParsing
                Write-Host "[Download] Success: File downloaded using fallback (Invoke-WebRequest)."
                return $true
            }
            catch {
                Write-Host "[Download] Error: Failed to download file using both methods. $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
        finally {
            if ($httpClient) { $httpClient.Dispose() }
        }
    }

    # Function to detect an existing Striim installation in default paths
    function Detect-ExistingStriim {
        $potentialPaths = @{
            "Agent" = "C:\striim\Agent";
            "Node"  = "C:\striim"
        }

        # Check Agent path first as it's more specific, then Node path
        foreach ($type in @("Agent", "Node")) {
            $path = $potentialPaths[$type]
            $libPath = Join-Path -Path $path -ChildPath "lib" -ErrorAction SilentlyContinue
            if ($libPath -and (Test-Path $libPath)) {
                $jarFile = Get-ChildItem -Path $libPath -Filter "Platform-*.jar" | Select-Object -First 1
                if ($jarFile) {
                    $version = $jarFile.Name -replace '^Platform-([\d\.]+)\.jar$', '$1'
                    return [pscustomobject]@{
                        Path    = $path
                        Version = $version
                        Type    = $type
                    }
                }
            }
        }
        return $null # Nothing found
    }

    # Function to handle the entire Striim installation process if not detected
    function Invoke-StriimInstallation {
        Write-Host "[Installer] Striim installation not detected in the current directory." -ForegroundColor Yellow
        if (-not (Confirm-UserChoice -Prompt "Would you like to download and install Striim now?" -DefaultChoice 'y')) {
            Write-Host "[Installer] Aborting script. Please run this script from a valid Striim installation directory." -ForegroundColor Red
            exit 1
        }

        # Get the current script's name to avoid deleting it and for later use
        $thisScriptPath = $PSCommandPath
        $thisScriptName = Split-Path $thisScriptPath -Leaf

        # 1. Get Version from parameter or prompt
        $targetVersion = $script:Version
        while ([string]::IsNullOrWhiteSpace($targetVersion)) {
            $targetVersion = Read-Host "[Installer] Please enter the Striim version to install (e.g., 5.0.6)"
        }
        try {
            [version]$targetVersion | Out-Null
        } catch {
            Write-Host "[Installer] Invalid version format provided. Aborting." -ForegroundColor Red
            exit 1
        }

        # 2. Get Node Type
        $installType = ""
        while ($installType -notin @("N", "A")) {
            $installType = (Read-Host "[Installer] Install Striim 'N'ode or 'A'gent? (N/A)").ToUpper()
        }
        $nodeName = if ($installType -eq "A") { "Agent" } else { "Node" }
        Write-Host "[Installer] Preparing to install Striim $nodeName version $targetVersion."

        # 3. Determine Paths and Handle Existing Installation
        $installRoot = "C:\striim"
        $installPath = if ($installType -eq "A") { Join-Path $installRoot "Agent" } else { $installRoot }

        $performInstall = $true
        if (Test-Path $installPath) {
            if (-not (Confirm-UserChoice -Prompt "Installation path '$installPath' already exists. Overwrite?" -DefaultChoice 'n')) {
                Write-Host "[Installer] Skipping installation, using existing directory at '$installPath'."
                $performInstall = $false
            } else {
                # MODIFICATION: Instead of deleting the whole directory, delete its contents except for 'downloads', the script itself, and msjetchecker.ps1
                Write-Host "[Installer] Removing contents of existing directory (excluding 'downloads', '$thisScriptName', and 'msjetchecker.ps1')..."
                Get-ChildItem -Path $installPath -Force | Where-Object { $_.Name -ne 'downloads' -and $_.Name -ne $thisScriptName -and $_.Name -ne 'msjetchecker.ps1' } | Remove-Item -Recurse -Force
            }
        }

        if ($performInstall) {
            New-Item -Path $installPath -ItemType Directory -Force | Out-Null
            Write-Host "[Installer] Created installation directory: $installPath"

            # 4. Download the MAIN application
            $installerInfo = $AllDownloads | Where-Object { $_.psobject.Properties['NodeType'] -and $_.NodeType -eq $installType -and $_.Category -eq "MainInstaller" } | Select-Object -First 1
            if (-not $installerInfo) {
                throw "[Installer] Could not find MAIN installer information for type '$nodeName' in the script manifest."
            }
            $installerFileName = $installerInfo.Name.Replace("{VERSION}", $targetVersion)
            $installerUrl = $installerInfo.Url.Replace("{VERSION}", $targetVersion)
            $zipPath = Join-Path $downloadDir $installerFileName

            if (-not (Download-File -Uri $installerUrl -OutFilePath $zipPath)) {
                throw "[Installer] Failed to download the Striim installer. Please check the URL and your internet connection."
            }

            # 5. Extract
            Write-Host "[Installer] Extracting files to $installPath..."
            try {
                Expand-Archive -Path $zipPath -DestinationPath $installPath -Force

                # MODIFICATION: Exclude the 'downloads' directory when checking for a single subdirectory.
                $subDirs = @(Get-ChildItem -Path $installPath -Directory -Force | Where-Object { $_.Name -ne 'downloads' })

                if ($subDirs.Count -eq 1) {
                    $singleSubDir = $subDirs[0]
                    Write-Host "[Installer] Moving contents from single subdirectory '$($singleSubDir.Name)' up one level."
                    # Move all items (files and folders) from the subdirectory to the parent installation path.
                    Get-ChildItem -Path $singleSubDir.FullName -Force | Move-Item -Destination $installPath -Force
                    # Remove the now-empty subdirectory.
                    Remove-Item -Path $singleSubDir.FullName -Force
                }
            } catch {
                throw "[Installer] Failed to extract the installer zip file. It may be corrupt or you may lack permissions. Error: $($_.Exception.Message)"
            }
            Write-Host "[Installer] Extraction complete."
        }

        # 6. Copy self and re-launch
        $newScriptPath = Join-Path $installPath $thisScriptName
        Write-Host "[Installer] Copying this script to '$newScriptPath'..."
        Copy-Item -Path $thisScriptPath -Destination $newScriptPath -Force

        # Reconstruct arguments for the new process
        $arguments = ""
        $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
            if ($_.Value -is [System.Management.Automation.SwitchParameter]) {
                if ($_.Value.IsPresent) { $arguments += "-$($_.Key) " }
            } else {
                $arguments += "-$($_.Key) `"$($_.Value)`" "
            }
        }

        Write-Host "[Installer] Relaunching script from the new location. A new PowerShell window will open."
        $relaunchCommand = "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File `"$newScriptPath`" $arguments' -WorkingDirectory `"$installPath`""
        Invoke-Expression $relaunchCommand

        Write-Host "[Installer] Installation handed off to new process. This script instance will now exit."
        exit 0
    }

    # Function to get a file, checking the local 'downloads' cache first before prompting for internet download.
    function Get-LocalOrDownload {
        param(
            [string]$FileName,
            [string]$Url,
            [string]$PromptMessage,
            [string]$LogPrefix
        )
        $localDownloadPath = Join-Path $downloadDir $FileName
        if (Test-Path $localDownloadPath) {
            Write-Host "[$LogPrefix] Found '$FileName' in local downloads folder."
            return $localDownloadPath
        }

        if (Confirm-UserChoice -Prompt "$PromptMessage" -DefaultChoice 'y') {
            if (Download-File -Uri $Url -OutFilePath $localDownloadPath) {
                return $localDownloadPath
            }
        }
        return $null
    }

    # Function to read a specific property from a config file
    function Get-ConfigProperty {
        param(
            [string]$ConfigPath,
            [string]$PropertyName
        )
        if (-not (Test-Path $ConfigPath)) { return $null }

        $configContent = Get-Content $ConfigPath
        $propertyLine = $configContent | Where-Object { $_ -match "^\s*$PropertyName\s*=\s*(.*)" } | Select-Object -First 1

        if ($propertyLine -and -not $propertyLine.StartsWith("#")) {
            return $propertyLine.Split("=", 2)[1].Trim()
        }
        return $null
    }

    # Reusable function to update properties in a configuration file
    function Update-ConfigFile {
        param(
            [string]$ConfigPath,
            [string[]]$RequiredProps,
            [string[]]$OptionalProps
        )
        if (-not (Test-Path $ConfigPath)) {
            Write-Host "[Config ] Fail***: Configuration file not found at $ConfigPath" -ForegroundColor Red
            return
        }

        $configLines = Get-Content $ConfigPath
        $allProps = $RequiredProps + $OptionalProps
        $fileChanged = $false
        $updatedValues = @{}

        foreach ($prop in $allProps) {
            $propFound = $false
            foreach ($lineIndex in 0..($configLines.Length - 1)) {
                $line = $configLines[$lineIndex]
                if ($line -match "^#?\s*$prop\s*=\s*(.*)") {
                    $propFound = $true
                    $propValue = $matches[1].Trim()
                    $originalLine = $line
                    $currentValue = $propValue

                    if ($line.StartsWith("#")) {
                        if (Confirm-UserChoice -Prompt "'$prop' is commented out. Uncomment and set a value?" -DefaultChoice 'y') {
                            $newValue = Read-Host "[Input  ] Enter a value for $prop"
                            $configLines[$lineIndex] = "$prop=$newValue"
                            $currentValue = $newValue
                        }
                    } elseif ([string]::IsNullOrEmpty($propValue) -and ($prop -in $RequiredProps)) {
                        Write-Host "[Config ] '$prop' is empty and required."
                        $newValue = Read-Host "[Input  ] Please provide a value for $prop"
                        $configLines[$lineIndex] = "$prop=$newValue"
                        $currentValue = $newValue
                    } else {
                        Write-Host "[Config ] Success: '$prop' found with value: $propValue"
                        if (Confirm-UserChoice -Prompt "Current value is '$propValue'. Update it?" -DefaultChoice 'n') {
                            $newValue = Read-Host "[Input  ] Enter the new value for $prop"
                            $configLines[$lineIndex] = "$prop=$newValue"
                            $currentValue = $newValue
                        }
                    }

                    if ($configLines[$lineIndex] -ne $originalLine) {
                        $fileChanged = $true
                        Write-Host "[Config ] Updated '$prop'."
                    } else {
                        Write-Host "[Config ] '$prop' remains unchanged."
                    }
                    $updatedValues[$prop] = $currentValue
                    break # Property found, move to next prop
                }
            }
            if (-not $propFound -and ($prop -in $RequiredProps)) {
                 Write-Warning "[Config ] Required property '$prop' not found in file. Adding it."
                 $newValue = Read-Host "[Input  ] Please provide a value for $prop"
                 $configLines += "$prop=$newValue"
                 $updatedValues[$prop] = $newValue
                 $fileChanged = $true
            }
        }

        if ($fileChanged) {
            Write-Host "[Config ] Saving changes to $(Split-Path $ConfigPath -Leaf)..."
            Set-Content -Path $ConfigPath -Value $configLines
        }
        return $updatedValues
    }

    #endregion

    #region Driver and Firewall Functions
    # Function to check for a specific firewall rule
    function Check-FirewallRule {
        param(
            [string]$DisplayName,
            [string]$Direction,
            [string]$Protocol,
            [string]$LocalPort,
            [string]$RemoteAddress = "Any",
            [string]$RemotePort = "Any"
        )
        $rule = Get-NetFirewallRule -Direction $Direction -Protocol $Protocol -ErrorAction SilentlyContinue | Where-Object {
            $_.Enabled -eq "True" -and
            ($_.LocalPort -contains $LocalPort) -and
            ($_.RemoteAddress -contains $RemoteAddress -or $_.RemoteAddress -contains "Any") -and
            ($_.RemotePort -contains $RemotePort)
        }

        if ($rule) {
            Write-Host "[Firewall] Success: Found enabled rule '$($rule.DisplayName | Select -First 1)' for $DisplayName." -ForegroundColor Green
            return $true
        } else {
            Write-Host "[Firewall] Fail***: No enabled rule found for $DisplayName (Direction: $Direction, Protocol: $Protocol, Port(s): $LocalPort, Remote Address: $RemoteAddress)." -ForegroundColor Red
            return $false
        }
    }

    # Main function to test all required firewall rules
    function Test-FirewallRules {
        param([string]$ServerAddress)

        Write-Host "[Firewall] Checking firewall rules... (Requires Administrator privileges)"
        if (-not (Test-IsAdmin)) {
            Write-Host "[Firewall] This check requires elevation. Please re-run the script as an Administrator to check firewall rules." -ForegroundColor Yellow
            return
        }

        # FIX: Explicitly import the NetSecurity module to ensure firewall cmdlets are available.
        Import-Module NetSecurity

        # This part only runs if the script is already elevated
        if ([string]::IsNullOrEmpty($ServerAddress)) {
            Write-Host "[Firewall] Fail***: Server address is not defined. Cannot perform firewall checks." -ForegroundColor Red
            return
        }

        Write-Host "[Firewall] Checking rules for server address: $ServerAddress"
        # 1. Inbound Hazelcast (5701-5703)
        Check-FirewallRule -DisplayName "Hazelcast Inbound" -Direction Inbound -Protocol TCP -LocalPort "5701-5703"

        # 2. Outbound HTTPS (9081 or 9080)
        $agentConfPath = Join-Path $striimInstallPath "conf\agent.conf"
        $httpsEnabled = Get-ConfigProperty -ConfigPath $agentConfPath -PropertyName "striim.cluster.https.enabled"
        $authPort = if ($httpsEnabled -eq "false") { "9080" } else { "9081" }
        Check-FirewallRule -DisplayName "Striim Auth Outbound" -Direction Outbound -Protocol TCP -RemotePort $authPort -RemoteAddress $ServerAddress

        # 3. Outbound ZeroMQ (49152-65535)
        Check-FirewallRule -DisplayName "ZeroMQ Outbound" -Direction Outbound -Protocol TCP -RemotePort "49152-65535" -RemoteAddress $ServerAddress
    }

    # Function to set a property in agent.conf non-interactively
    function Set-AgentConfProperty {
        param(
            [string]$PropertyName,
            [string]$PropertyValue
        )
        $agentConfPath = Join-Path $striimInstallPath "conf\agent.conf"
        $configLines = Get-Content $agentConfPath
        # Remove existing lines for this property (commented or not)
        $filteredLines = $configLines | Where-Object { $_ -notmatch "^#?\s*$PropertyName\s*=" }
        # Add the new line
        $newLines = $filteredLines + "$PropertyName=$PropertyValue"
        Set-Content -Path $agentConfPath -Value $newLines
        Write-Host "[Drivers] Success: Set '$PropertyName' in agent.conf."
    }

    # Main function to handle JDBC driver installations
    function Install-JdbcDrivers {
        if (-not (Confirm-UserChoice -Prompt "`nWould you like to check for and install any JDBC drivers for the agent?" -DefaultChoice 'y')) {
            return
        }

        $striimLibPath = Join-Path -Path $striimInstallPath -ChildPath "lib"

        $driverMenu = [ordered]@{
            '1' = 'HP NonStop (Manual Path)'
            '2' = 'MariaDB (v2.4.3)'
            '3' = 'MySQL / MemSQL (Connector/J 8.0.30)'
            '4' = 'Oracle Instant Client (for OJet)'
            '5' = 'PostgreSQL (v42.2.27)'
            '6' = 'Teradata (Manual Path)'
            '7' = 'Vertica (Manual Path)'
        }

        $exitMenu = $false
        while (-not $exitMenu) {
            Write-Host "`n--- JDBC Driver Installation Menu ---"
            $driverMenu.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Name). $($_.Value)" }
            Write-Host "  0. Exit Menu"
            $choice = Read-Host "[Input  ] Select a driver to install (or 0 to exit)"

            switch ($choice) {
                '1' { Install-HpNonStopDriver -LibPath $striimLibPath }
                '2' { Install-MariaDbDriver -LibPath $striimLibPath }
                '3' { Install-MySqlDriver -LibPath $striimLibPath }
                '4' { Install-OracleInstantClient -LibPath $striimLibPath }
                '5' { Install-PostgresDriver -LibPath $striimLibPath }
                '6' { Install-TeradataDriver -LibPath $striimLibPath }
                '7' { Install-VerticaDriver -LibPath $striimLibPath }
                '0' { $exitMenu = $true }
                default { Write-Warning "Invalid selection." }
            }
        }
    }

    # Individual driver installation functions
    function Install-HpNonStopDriver {
        param($LibPath)
        Write-Host "`n--- Installing HP NonStop Driver ---"
        Write-Host "This requires a manual step. Please follow your HPE documentation to get the 't4sqlmx.jar' file from your NonStop system."
        $jarPath = Read-Host "[Input  ] Enter the full path to the 't4sqlmx.jar' file"
        if (Test-Path $jarPath -Filter "t4sqlmx.jar") {
            Copy-Item -Path $jarPath -Destination $LibPath -Force
            Write-Host "[Drivers] Success: Copied t4sqlmx.jar to $LibPath"
            $global:AgentRestartNeeded = $true
        } else {
            Write-Host "[Drivers] Error: File not found at the specified path." -ForegroundColor Red
        }
    }

    function Install-MariaDbDriver {
        param($LibPath)
        Write-Host "`n--- Installing MariaDB Driver ---"
        $driverInfo = $AllDownloads | Where-Object { $_.Name -eq "mariadb-java-client-2.4.3.jar" }
        $downloadPath = Get-LocalOrDownload -FileName $driverInfo.Name -Url $driverInfo.Url -PromptMessage "MariaDB driver not found in downloads. Download now?" -LogPrefix "Drivers"
        if ($downloadPath) {
            Copy-Item -Path $downloadPath -Destination $LibPath -Force
            Write-Host "[Drivers] Success: Copied $($driverInfo.Name) to $LibPath"
            $global:AgentRestartNeeded = $true
        }
    }

    function Install-MySqlDriver {
        param($LibPath)
        Write-Host "`n--- Installing MySQL/MemSQL Driver ---"
        $driverInfo = $AllDownloads | Where-Object { $_.Name -eq "mysql-connector-j-8.0.30.zip" }
        $jarFile = "mysql-connector-java-8.0.30.jar"

        $zipPath = Get-LocalOrDownload -FileName $driverInfo.Name -Url $driverInfo.Url -PromptMessage "MySQL driver zip not found in downloads. Download now?" -LogPrefix "Drivers"
        if ($zipPath) {
            $extractPath = Join-Path $downloadDir "mysql-temp"
            if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
            Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
            $sourceJar = Join-Path $extractPath "mysql-connector-j-8.0.30\$jarFile"
            if (Test-Path $sourceJar) {
                Copy-Item -Path $sourceJar -Destination $LibPath -Force
                Write-Host "[Drivers] Success: Copied $jarFile to $LibPath"
                $global:AgentRestartNeeded = $true
            } else {
                Write-Host "[Drivers] Error: Could not find $jarFile in the extracted archive." -ForegroundColor Red
            }
            Remove-Item $extractPath -Recurse -Force
        }
    }

    function Install-OracleInstantClient {
        param($LibPath)
        Write-Host "`n--- Installing Oracle Instant Client ---"
        $driverInfo = $AllDownloads | Where-Object { $_.Name -like "instantclient-basic-windows*.zip" }

        $zipPath = Get-LocalOrDownload -FileName $driverInfo.Name -Url $driverInfo.Url -PromptMessage "Oracle Instant Client zip not found in downloads. Download now?" -LogPrefix "Drivers"
        if ($zipPath) {
            $defaultInstallDir = Join-Path $striimInstallPath "oracle_instant_client"
            $installDir = Read-Host "[Input  ] Enter path to install Oracle Instant Client (default: $defaultInstallDir)"
            if ([string]::IsNullOrWhiteSpace($installDir)) { $installDir = $defaultInstallDir }

            Expand-Archive -Path $zipPath -DestinationPath $installDir -Force
            $clientDir = Get-ChildItem -Path $installDir -Directory | Select-Object -First 1
            if ($clientDir) {
                $nativeLibsPath = $clientDir.FullName
                Write-Host "[Drivers] Oracle Instant Client extracted to $nativeLibsPath"
                Set-AgentConfProperty -PropertyName "NATIVE_LIBS" -PropertyValue $nativeLibsPath
                $global:AgentRestartNeeded = $true
            } else {
                Write-Host "[Drivers] Error: Could not determine client directory after extraction." -ForegroundColor Red
            }
        }
    }

    function Install-PostgresDriver {
        param($LibPath)
        Write-Host "`n--- Installing PostgreSQL Driver ---"
        $driverInfo = $AllDownloads | Where-Object { $_.Name -eq "postgresql-42.2.27.jar" }
        $downloadPath = Get-LocalOrDownload -FileName $driverInfo.Name -Url $driverInfo.Url -PromptMessage "PostgreSQL driver not found in downloads. Download now?" -LogPrefix "Drivers"
        if ($downloadPath) {
            Copy-Item -Path $downloadPath -Destination $LibPath -Force
            Write-Host "[Drivers] Success: Copied $($driverInfo.Name) to $LibPath"
            $global:AgentRestartNeeded = $true
        }
    }

    function Install-TeradataDriver {
        param($LibPath)
        Write-Host "`n--- Installing Teradata Driver ---"
        Write-Host "This requires a manual step. Please download and extract the Teradata JDBC driver from the Teradata website."
        $dirPath = Read-Host "[Input  ] Enter the full path to the directory containing 'terajdbc4.jar' and 'tdgssconfig.jar'"
        $teraJar = Join-Path $dirPath "terajdbc4.jar"
        $tdgssJar = Join-Path $dirPath "tdgssconfig.jar"
        if ((Test-Path $teraJar) -and (Test-Path $tdgssJar)) {
            Copy-Item -Path $teraJar -Destination $LibPath -Force
            Copy-Item -Path $tdgssJar -Destination $LibPath -Force
            Write-Host "[Drivers] Success: Copied Teradata JARs to $LibPath"
            $global:AgentRestartNeeded = $true
        } else {
            Write-Host "[Drivers] Error: Could not find one or both required JAR files in the specified directory." -ForegroundColor Red
        }
    }

    function Install-VerticaDriver {
        param($LibPath)
        Write-Host "`n--- Installing Vertica Driver ---"
        Write-Host "This requires a manual step. Please download and extract the Vertica JDBC driver from the OpenText website."
        $jarPath = Read-Host "[Input  ] Enter the full path to the 'vertica-jdbc-24.2.0-1.jar' file"
        if (Test-Path $jarPath -Filter "vertica-jdbc-*.jar") {
            Copy-Item -Path $jarPath -Destination $LibPath -Force
            Write-Host "[Drivers] Success: Copied Vertica JAR to $LibPath"
            $global:AgentRestartNeeded = $true
        } else {
            Write-Host "[Drivers] Error: File not found at the specified path." -ForegroundColor Red
        }
    }
    #endregion

    # --- Main Logic Branching ---
    if ($DownloadOnly) {
        # --- Download-Only Mode ---
        Write-Host "--- Download-Only Mode Activated ---" -ForegroundColor Cyan

        # Validate parameters for this mode
        if ([string]::IsNullOrWhiteSpace($Version)) {
            throw "The --version parameter is required when using --downloadonly. Example: --version 5.0.6"
        }
        if (($Agent -and $Node) -or (-not $Agent -and -not $Node)) {
            throw "When using --downloadonly, you must specify exactly one of --agent or --node."
        }

        try {
            $TargetVersion = [version]$Version
        }
        catch {
            throw "Invalid format for --version. Please use a format like '5.0.6'."
        }

        $nodeTypeLetter = if ($Agent) { 'A' } else { 'N' }
        $nodeTypeName = if ($Agent) { 'Agent' } else { 'Node' }

        Write-Host "Downloading all dependencies for Striim $nodeTypeName version $TargetVersion to '$downloadDir'..."

        # Filter files based on version and node type.
        # A file is included if its NodeType property does not exist (common file) or if it matches the requested type.
        $filesToDownload = $AllDownloads | Where-Object {
            ($TargetVersion -ge [version]$_.MinVersion -and $TargetVersion -le [version]$_.MaxVersion) -and
            (-not $_.psobject.Properties['NodeType'] -or $_.NodeType -eq $nodeTypeLetter)
        }

        if (-not $filesToDownload) {
            Write-Host "No specific downloads found for version $TargetVersion in the manifest."
        } else {
            foreach ($file in $filesToDownload) {
                Write-Host "`nChecking for $($file.Name) (Category: $($file.Category))..."
                $fileName = $file.Name.Replace("{VERSION}", $TargetVersion)
                $url = $file.Url.Replace("{VERSION}", $TargetVersion)
                $downloadPath = Join-Path $downloadDir $fileName
                Download-File -Uri $url -OutFilePath $downloadPath
            }
        }

        Write-Host "`n--- Download-Only Mode Finished ---" -ForegroundColor Cyan
        Write-Host "All applicable files have been downloaded to the '$downloadDir' folder."
        Write-Host "You can now transfer this entire directory to an offline machine and run the script normally."
        exit 0 # Exit successfully after downloads are complete.
    }
    else {
        # --- Standard / Interactive Mode ---

        # --- Initial Environment Check ---
        $agentConfPath = Join-Path $striimInstallPath "conf\agent.conf"
        $startUpPropsPath = Join-Path $striimInstallPath "conf\startUp.properties"
        $isCurrentDirStriim = (Test-Path $agentConfPath) -or (Test-Path $startUpPropsPath)

        if (-not $isCurrentDirStriim) {
            # Not running from a Striim directory, let's check default locations
            $existingInstall = Detect-ExistingStriim

            if ($existingInstall) {
                Write-Host "[Envrnmt] Detected existing Striim $($existingInstall.Type) version $($existingInstall.Version) at $($existingInstall.Path)"
                if (-not (Confirm-UserChoice -Prompt "Do you want to download a different version?" -DefaultChoice 'n')) {
                    # User wants to use the existing version.
                    # Copy self and re-launch from that location.
                    $installPath = $existingInstall.Path
                    $thisScriptPath = $PSCommandPath
                    $thisScriptName = Split-Path $thisScriptPath -Leaf
                    $newScriptPath = Join-Path $installPath $thisScriptName
                    Write-Host "[Envrnmt] Using existing installation. Copying this script to '$newScriptPath'..."
                    Copy-Item -Path $thisScriptPath -Destination $newScriptPath -Force

                    $arguments = ""
                    $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
                        if ($_.Value -is [System.Management.Automation.SwitchParameter]) {
                            if ($_.Value.IsPresent) { $arguments += "-$($_.Key) " }
                        } else {
                            $arguments += "-$($_.Key) `"$($_.Value)`" "
                        }
                    }

                    Write-Host "[Envrnmt] Relaunching script from the Striim directory. A new PowerShell window will open."
                    $relaunchCommand = "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File `"$newScriptPath`" $arguments' -WorkingDirectory `"$installPath`""
                    Invoke-Expression $relaunchCommand
                    exit 0 # Exit the current script
                } else {
                    # User wants to download a new version, proceed to the full installation flow.
                    Invoke-StriimInstallation
                }
            } else {
                # No existing install found in default locations, and not in a Striim dir.
                # Proceed to the full installation flow.
                Invoke-StriimInstallation
            }
        }

        # --- Startup Summary (Standard Mode) ---
        Write-Host "`nStarting Striim Configuration Check..."
        Write-Host "------------------------------------------"
        Write-Host "[Checks Summary] The following validations will occur:"
        Write-Host "  • [System]        Verify ≥15GB free space on C: drive"
        Write-Host "  • [Node Type]     Auto-detect Agent/Node or prompt for selection"
        Write-Host "  • [Config Files]  Validate agent.conf/startUp.properties settings"
        Write-Host "  • [Firewall]      AGENT ONLY: Check required inbound/outbound firewall rules (requires Admin) $([char]0x27a4)" -ForegroundColor Green
        Write-Host "  • [System PATH]   Add Striim lib directory to system PATH (requires Admin) $([char]0x27a4)" -ForegroundColor Yellow
        Write-Host "  • [Dependencies]  Verify required DLLs and offer to download them"
        Write-Host "  • [Java Check]    Ensure compatible Java version is installed and offer download if needed"
        Write-Host "  • [Security]      Setup Integrated Security sqljdbc_auth.dll (requires Admin) $([char]0x27a4)" -ForegroundColor Yellow
        Write-Host "  • [Prereqs]       Verify required software (VC++ Redist, OLE DB Driver)" -ForegroundColor Green
        Write-Host "  • [Patches]       Apply critical fixes for Striim v4.2.0.20 (if applicable)"
        Write-Host "  • [Service]       Configure Striim as a Windows Service (requires Admin) $([char]0x27a4)" -ForegroundColor Yellow
        Write-Host "  • [JKS Config]    Prompt to generate/recreate JKS and key files (requires Admin) $([char]0x27a4)" -ForegroundColor Yellow
        Write-Host "  • [Drivers]       AGENT ONLY: Prompt to install common JDBC drivers" -ForegroundColor Green
        Write-Host "------------------------------------------"
        Read-Host "`nPress Enter to continue or Ctrl+C to abort..."

        # --- Environment Validation ---
        # 1. Check Disk Space
        if ((Get-PSDrive -Name C).Free / 1GB -lt 15) {
            throw "[Envrnmt] Insufficient disk space on C: drive. At least 15 GB is required."
        } else {
            Write-Host "[Envrnmt]        Sufficient disk space available on C: drive."
        }

        # 2. Detect Node Type
        $nodeType = ""
        if (Test-Path $agentConfPath) {
            $nodeType = "A"
            Write-Host "[Envrnmt]        -> AGENT Environment detected."
        } elseif (Test-Path $startUpPropsPath) {
            $nodeType = "N"
            Write-Host "[Envrnmt]        -> NODE environment detected."
        } else {
            # This case should technically not be reached due to the initial check, but it's good practice.
            throw "Striim not found. Please place the script in a Striim installation directory and re-run."
        }
        Write-Host "[Envrnmt] Success: Striim Install Path set to: $striimInstallPath"

        # Define core paths now that install path is confirmed
        $striimLibPath = Join-Path -Path $striimInstallPath -ChildPath "lib"
        $striimConfPath = Join-Path -Path $striimInstallPath -ChildPath "conf"
        $striimBinPath = Join-Path -Path $striimInstallPath -ChildPath "bin"

        # Get Striim Version
        $striimJarFile = Get-ChildItem -Path $striimLibPath -Filter "Platform-*.jar" | Select-Object -First 1
        if (-not $striimJarFile) { throw "[Striim ] Fail***: Could not find Striim JAR file to determine version." }
        $striimVersionString = $striimJarFile.Name -replace '^Platform-([\d\.]+)\.jar$', '$1'
        $majorVersion = $striimVersionString.Split('.')[0]
        Write-Host "[Striim  ] Success: Found Striim version $striimVersionString (Major version $majorVersion)"


        # --- Configuration Checks ---
        # 4. Update Configuration Files
        if ($nodeType -eq "A") {
            Write-Host "[Config ]        -> AGENT -> Specific Tests for configuration:"
            $configValues = Update-ConfigFile -ConfigPath $agentConfPath -RequiredProps "striim.cluster.clusterName", "striim.node.servernode.address" -OptionalProps "MEM_MAX"
            $jksFileName = "aks.jks"; $pwdFileName = "aksKey.pwd"; $configScriptName = "aksConfig.bat"; $configType = "AKS"

            # --- Firewall Check ---
            if ($configValues -and $configValues.ContainsKey("striim.node.servernode.address")) {
                Test-FirewallRules -ServerAddress $configValues["striim.node.servernode.address"]
            } else {
                Write-Host "[Firewall] Skipping check because 'striim.node.servernode.address' could not be read." -ForegroundColor Yellow
            }

        } elseif ($nodeType -eq "N") {
            Write-Host "[Config ]        -> NODE -> Specific Tests for configuration:"
            Update-ConfigFile -ConfigPath $startUpPropsPath -RequiredProps "CompanyName", "LicenceKey", "ProductKey", "WAClusterName" -OptionalProps "MEM_MAX"
            $jksFileName = "sks.jks"; $pwdFileName = "sksKey.pwd"; $configScriptName = "sksConfig.bat"; $configType = "SKS"
        }

        # 5. Check System PATH
        Write-Host "[Config ]        -> Striim Lib Path set to: $striimLibPath"
        $normalizedEnvPath = ($env:Path -split ';').TrimEnd('\')
        $normalizedStriimLibPath = $striimLibPath.TrimEnd('\')
        if ($normalizedEnvPath -icontains $normalizedStriimLibPath) { # Use -icontains for case-insensitivity
            Write-Host "[Config ] Success: Striim lib directory found in PATH."
        } else {
            Write-Host "[Config ] Fail***: Striim lib directory not found in PATH."
            if (Confirm-UserChoice -Prompt "Add it to the system PATH (requires Admin)?" -DefaultChoice 'y') {
                if (Test-IsAdmin) {
                    [Environment]::SetEnvironmentVariable("Path", ($env:Path + ";" + $striimLibPath), "Machine")
                    Write-Host "[Config ] Success: PATH updated (already elevated)."
                    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") # Refresh current session
                } else {
                    Write-Host "[Config ]  Elevating permissions to update PATH..."
                    $command = "[Environment]::SetEnvironmentVariable('Path', ([Environment]::GetEnvironmentVariable('Path', 'Machine') + ';$striimLibPath'), 'Machine')"
                    if ((Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $striimInstallPath) -eq 0) {
                        Write-Host "[Config ] Success: Striim lib directory added to PATH. Please restart your terminal to see the change."
                    } else {
                        Write-Host "[Config ] Error: Failed to add Striim lib directory to PATH." -ForegroundColor Red
                    }
                }
            }
        }


        # 6. Check for required DLLs
        $requiredDlls = $AllDownloads | Where-Object { $_.Category -eq "Dependency" }
        foreach ($dllInfo in $requiredDlls) {
            $dllPath = Join-Path $striimLibPath $dllInfo.Name
            if (Test-Path $dllPath) {
                Write-Host "[DLLs    ] Success: $($dllInfo.Name) found."
            } else {
                Write-Host "[DLLs    ] Fail***: $($dllInfo.Name) not found."
                $downloadedDllPath = Get-LocalOrDownload -FileName $dllInfo.Name -Url $dllInfo.Url -PromptMessage "Download it now?" -LogPrefix "DLLs"
                if ($downloadedDllPath) {
                    Copy-Item $downloadedDllPath $striimLibPath -Force
                    Write-Host "[DLLs    ] Success: $($dllInfo.Name) copied to $striimLibPath"
                } else {
                    Write-Host "[DLLs    ] Info: Skipping $($dllInfo.Name)."
                }
            }
        }

        # 7. Check Java Version
        $requiredJavaVersion = if ([int]$majorVersion -le 4) { 8 } else { 11 }
        $javaDownloadInfo = $AllDownloads | Where-Object { $_.Category -eq "Java" -and (($requiredJavaVersion -eq 8 -and $_.MaxVersion -lt 5) -or ($requiredJavaVersion -eq 11 -and $_.MinVersion -ge 5)) } | Select-Object -First 1

        $installedJavaVersion = 0
        $javaCheckCmd = Get-Command java -ErrorAction SilentlyContinue
        if ($javaCheckCmd) {
            $versionOutput = java -version 2>&1
            $versionString = ($versionOutput | Select-String 'version "([^"]+)"').Matches.Groups[1].Value
            if ($versionString) {
                Write-Host "[Java    ] Success: Found installed Java version string: `"$versionString`""
                if ($versionString -match "^11\.") {
                    $installedJavaVersion = 11
                } elseif ($versionString -match "^1\.8") {
                    $installedJavaVersion = 8
                }
            } else {
                 Write-Host "[Java    ] Warning: Could not parse version from 'java -version' output." -ForegroundColor Yellow
            }
        }

        if ($installedJavaVersion -eq $requiredJavaVersion) {
            Write-Host "[Java    ] Success: Required Java version $requiredJavaVersion is installed."
        } else {
            $foundVersionString = if ($installedJavaVersion -gt 0) { $installedJavaVersion } else { "None" }
            Write-Host "[Java    ] Fail***: Java version mismatch. Required: $requiredJavaVersion, Found: $foundVersionString." -ForegroundColor Yellow
            $javaInstallerPath = Get-LocalOrDownload -FileName $javaDownloadInfo.Name -Url $javaDownloadInfo.Url -PromptMessage "Download required Java version $requiredJavaVersion now?" -LogPrefix "Java"
            if ($javaInstallerPath) {
                if (Confirm-UserChoice -Prompt "Java installer available at $javaInstallerPath. Run installer now (as administrator)?" -DefaultChoice 'y') {
                    $command = "Start-Process -FilePath '$javaInstallerPath' -Wait -PassThru | Wait-Process"
                    if ((Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $downloadDir) -eq 0) {
                        Write-Host "[Java    ] Installer launched successfully."
                    } else {
                        Write-Host "[Java    ] Installer launch failed." -ForegroundColor Red
                    }
                }
            }
        }

        # 8. Integrated Security Setup
        $sqljdbcAuthDllPath = "C:\Windows\System32\sqljdbc_auth.dll"
        if (Test-Path $sqljdbcAuthDllPath) {
            Write-Host "[Int Sec] Success: sqljdbc_auth.dll already exists in C:\Windows\System32."
        } else {
            if (Confirm-UserChoice -Prompt "Plan to use Integrated Security (requires sqljdbc_auth.dll)?" -DefaultChoice 'y') {
                $sourceDllPath = Join-Path $striimLibPath "sqljdbc_auth.dll"
                if (Test-Path $sourceDllPath) {
                    if (Confirm-UserChoice -Prompt "Found sqljdbc_auth.dll in lib. Copy to System32 (requires Admin)?" -DefaultChoice 'y') {
                         $command = "Copy-Item -Path '$sourceDllPath' -Destination '$sqljdbcAuthDllPath' -Force"
                         if ((Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $striimInstallPath) -eq 0) {
                              Write-Host "[Int Sec] Success: Copied sqljdbc_auth.dll to System32."
                         } else {
                              Write-Host "[Int Sec] Error: Failed to copy DLL." -ForegroundColor Red
                         }
                    }
                } else {
                    $authDllInfo = $AllDownloads | Where-Object { $_.Name -eq "sqljdbc_auth.dll" }
                    $downloadedDllPath = Get-LocalOrDownload -FileName $authDllInfo.Name -Url $authDllInfo.Url -PromptMessage "sqljdbc_auth.dll not found. Download it now?" -LogPrefix "Int Sec"
                    if ($downloadedDllPath) {
                        if (Confirm-UserChoice -Prompt "Download complete. Copy to System32 now (requires Admin)?" -DefaultChoice 'y') {
                            $command = "Copy-Item -Path '$downloadedDllPath' -Destination '$sqljdbcAuthDllPath' -Force"
                            if ((Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $downloadDir) -eq 0) {
                                Write-Host "[Int Sec] Success: Copied sqljdbc_auth.dll to System32."
                            } else {
                                Write-Host "[Int Sec] Error: Failed to copy downloaded DLL." -ForegroundColor Red
                            }
                        }
                    }
                }
            }
        }

        # 9. Version-specific Patches
        $patches = $AllDownloads | Where-Object { $_.Category -eq "Patch" -and $striimVersionString -eq $_.MinVersion }
        if ($patches) {
            Write-Host "[Patches] Patches are available for version $striimVersionString."
            $baseFilesExist = $patches | ForEach-Object { Test-Path (Join-Path $striimLibPath $_.TargetFile) }
            if ($baseFilesExist -contains $true) {
                if (Confirm-UserChoice -Prompt "One or more files that can be patched exist. Apply patches now?" -DefaultChoice 'y') {
                    foreach ($patch in $patches) {
                        Write-Host "[Patches] Applying patch for $($patch.TargetFile)..."
                        $patchFilePath = Get-LocalOrDownload -FileName $patch.Name -Url $patch.Url -PromptMessage "Download patch $($patch.Name)?" -LogPrefix "Patches"
                        if ($patchFilePath) {
                            $destinationPath = Join-Path $striimLibPath $patch.TargetFile
                            try {
                                Move-Item -Path $patchFilePath -Destination $destinationPath -Force -ErrorAction Stop
                                Write-Host "[Patches] Success: Patched $($patch.TargetFile)."
                            } catch {
                                Write-Host "[Patches] Error: Failed to apply patch for $($patch.TargetFile). $($_.Exception.Message)" -ForegroundColor Red
                            }
                        }
                    }
                }
            }
        } else {
            Write-Host "[Patches] No specific patches identified for Striim version $striimVersionString in this script."
        }


        # 10. Check Software Requirements
        Write-Host "[Softwre] Checking for required software... (This may take a moment)"
        # Using Get-WmiObject as it was previously working, though it can be slow.
        $installedSoftware = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue
        if (-not $installedSoftware) {
            Write-Host "[Softwre] Warning: Could not retrieve list of installed software via WMI. Skipping this check." -ForegroundColor Yellow
        } else {

            # --- Requirement 1: Visual C++ Redistributable ---
            $vcRedistInfo = $AllDownloads | Where-Object { $_.Name -eq "vc_redist.x64.exe" }
            $vcRedistName = "Microsoft Visual C++" # Generic name to match multiple versions
            $vcRedistMinVersion = "14.0"

            # Note: Get-WmiObject uses 'Name' property, not 'DisplayName'
            $foundVcRedist = $installedSoftware | Where-Object {
                $_.Name -like "*$vcRedistName*"
            } | Where-Object {
                $_.Version -match '^\d' -and [version]($_.Version -replace ',','.') -ge [version]$vcRedistMinVersion
            }

            if ($foundVcRedist) {
                $vcRedist = $foundVcRedist | Select-Object -First 1
                Write-Host "[Softwre] Success: Found compatible $($vcRedist.Name) version $($vcRedist.Version)."
            } else {
                Write-Host "[Softwre] Fail***: Did not find required version of 'Microsoft Visual C++ 2015-2022 Redistributable'." -ForegroundColor Yellow
                $installerPath = Get-LocalOrDownload -FileName $vcRedistInfo.Name -Url $vcRedistInfo.Url -PromptMessage "Download and install '$($vcRedistInfo.Name)' now?" -LogPrefix "Softwre"
                if ($installerPath) {
                    if (Confirm-UserChoice -Prompt "Installer available at '$installerPath'. Run it now (as administrator)?" -DefaultChoice 'y') {
                        $command = "Start-Process -FilePath '$installerPath' -Wait -PassThru | Wait-Process"
                        $exitCode = Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $downloadDir
                        if ($exitCode -ne 0 -and $exitCode -ne 3010) { # 3010 is success, restart required
                            Write-Host "[Softwre] Warning: Installer process exited with code $exitCode." -ForegroundColor Yellow
                        } else {
                            Write-Host "[Softwre] Installer finished. A system restart may be required."
                        }
                    }
                }
            }

            # --- Requirement 2: Microsoft OLE DB Driver for SQL Server ---
            $oleDbInfo = $AllDownloads | Where-Object { $_.Name -eq "msoledbsql.msi" }
            $oleDbDriverName = "Microsoft OLE DB Driver for SQL Server"
            $oleDbRequiredVersions = @("18.2.3.0", "18.7.4.0") # Add any other acceptable versions here
            $installedDriver = $installedSoftware | Where-Object { $_.Name -like "*$oleDbDriverName*" }

            $foundCorrectVersion = $false
            if ($installedDriver) {
                foreach ($driver in $installedDriver) {
                    if ($oleDbRequiredVersions -contains $driver.Version) {
                        Write-Host "[Softwre] Success: $oleDbDriverName version $($driver.Version) found. Meets requirement."
                        $foundCorrectVersion = $true
                        break
                    }
                }
            }

            if (-not $foundCorrectVersion) {
                $installedVersions = if ($installedDriver) { ($installedDriver.Version | ForEach-Object { $_.ToString() }) -join ", " } else { "None" }
                Write-Host "[Softwre] Fail***: A required version of $oleDbDriverName was not found (Found: $installedVersions)." -ForegroundColor Yellow
                $installerPath = Get-LocalOrDownload -FileName $oleDbInfo.Name -Url $oleDbInfo.Url -PromptMessage "Download and install '$oleDbDriverName' now?" -LogPrefix "Softwre"
                if ($installerPath) {
                    if (Confirm-UserChoice -Prompt "Installer available at '$installerPath'. Run it now (as administrator)?" -DefaultChoice 'y') {
                        $command = "Start-Process -FilePath '$installerPath' -Wait -PassThru | Wait-Process"
                        $exitCode = Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $downloadDir
                        if ($exitCode -ne 0 -and $exitCode -ne 3010) {
                            Write-Host "[Softwre] Warning: Installer process exited with code $exitCode." -ForegroundColor Yellow
                        } else {
                            Write-Host "[Softwre] Installer finished. A system restart may be required."
                        }
                    }
                }
            }
        }

        # 11. Check Striim Service
        $serviceName = if ($nodeType -eq "A") { "Striim Agent" } else { "Striim" }
        if (Get-Service $serviceName -ErrorAction SilentlyContinue) {
            Write-Host "[Service] Success: Striim service '$serviceName' is installed."
        } else {
            Write-Host "[Service] Striim service '$serviceName' is not installed."
            if (Confirm-UserChoice -Prompt "Do you want to set up the service now (requires Admin)?" -DefaultChoice 'y') {
                $serviceInfo = $AllDownloads | Where-Object { $_.psobject.Properties['NodeType'] -and $_.NodeType -eq $nodeType -and $_.Category -eq "ServiceInstaller" } | Select-Object -First 1
                if ($serviceInfo) {
                    $serviceFileName = $serviceInfo.Name.Replace("{VERSION}", $striimVersionString)
                    $serviceUrl = $serviceInfo.Url.Replace("{VERSION}", $striimVersionString)
                    $zipPath = Get-LocalOrDownload -FileName $serviceFileName -Url $serviceUrl -PromptMessage "Download service installer?" -LogPrefix "Service"

                    if ($zipPath) {
                        $serviceConfigFolder = if ($nodeType -eq "A") { Join-Path $striimInstallPath "conf\windowsAgent" } else { Join-Path $striimInstallPath "conf\windowsService" }
                        if (Test-Path $serviceConfigFolder) { Remove-Item $serviceConfigFolder -Recurse -Force }
                        New-Item -ItemType Directory -Path $serviceConfigFolder -Force | Out-Null

                        Write-Host "[Service] Extracting service files..."
                        $tempExtractPath = Join-Path $downloadDir "service-temp"
                        if (Test-Path $tempExtractPath) { Remove-Item $tempExtractPath -Recurse -Force }
                        Expand-Archive -Path $zipPath -DestinationPath $tempExtractPath -Force

                        $subItems = @(Get-ChildItem -Path $tempExtractPath -Force)
                        if ($subItems.Count -eq 1 -and $subItems[0].PSIsContainer) {
                            Write-Host "[Service] Moving contents from single subdirectory '$($subItems[0].Name)'..."
                            Get-ChildItem -Path $subItems[0].FullName -Force | Move-Item -Destination $serviceConfigFolder -Force
                        } else {
                            $subItems | Move-Item -Destination $serviceConfigFolder -Force
                        }

                        Remove-Item $tempExtractPath -Recurse -Force
                        Write-Host "[Service] Extraction and move complete."

                        $setupScript = if ($nodeType -eq "A") { "setupWindowsAgent.ps1" } else { "setupWindowsService.ps1" }
                        $setupScriptPath = Join-Path $serviceConfigFolder $setupScript

                        if (Test-Path $setupScriptPath) {
                            Write-Host "[Service] Running service setup script: $setupScriptPath"
                            $scriptExecutionCommand = "Set-Location -Path '$serviceConfigFolder'; & '$setupScriptPath'"
                            $exitCode = Invoke-AsAdmin -ArgumentList $scriptExecutionCommand -WorkingDirectory $serviceConfigFolder
                            if ($exitCode -eq 0) {
                                Write-Host "[Service] Service setup script completed successfully."
                            } else {
                                Write-Host "[Service] Error: Service setup script failed (exit code: $exitCode)." -ForegroundColor Red
                            }
                        } else {
                            Write-Host "[Service] Error: Setup script '$setupScript' not found after extraction." -ForegroundColor Red
                        }
                    }
                } else {
                    Write-Host "[Service] Error: Could not find service installer information for node type '$nodeType'." -ForegroundColor Red
                }
            }
        }

        # 12. JKS/Key File Generation
        Write-Host "[JKS Cfg] Checking $configType Key Store (JKS) and Password files..."
        if ($jksFileName) { # Ensure node type was determined
            $jksPath = Join-Path -Path $striimConfPath -ChildPath $jksFileName
            $pwdPath = Join-Path -Path $striimConfPath -ChildPath $pwdFileName
            $configScriptPath = Join-Path -Path $striimBinPath -ChildPath $configScriptName

            $filesExist = (Test-Path $jksPath) -and (Test-Path $pwdPath)
            $runScript = $false
            if ($filesExist) {
                if (Confirm-UserChoice -Prompt "JKS/PWD files exist. Re-create them (requires Admin)?" -DefaultChoice 'n') {
                    Write-Host "[JKS Cfg] Deleting existing files..."
                    Remove-Item $jksPath, $pwdPath -Force -ErrorAction SilentlyContinue
                    $runScript = $true
                }
            } else {
                if (Confirm-UserChoice -Prompt "JKS/PWD files do not exist. Create them now (requires Admin)?" -DefaultChoice 'y') {
                    $runScript = $true
                }
            }

            if ($runScript) {
                if (-not (Test-Path $configScriptPath)) {
                    Write-Host "[JKS Cfg] Error: Config script not found at $configScriptPath" -ForegroundColor Red
                } else {
                    Write-Host "[JKS Cfg] Running '$configScriptName' as Administrator..."
                    $scriptDir = Split-Path $configScriptPath -Parent
                    # Use cmd /c to properly execute the batch file from PowerShell
                    $command = "cmd /c `"`"$configScriptPath`"`""
                    if ((Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $scriptDir) -eq 0) {
                        Write-Host "[JKS Cfg] Success: Script executed. Verifying files..."
                        if ((Test-Path $jksPath) -and (Test-Path $pwdPath)) {
                            Write-Host "[JKS Cfg] Success: JKS and PWD files created."
                        } else {
                            Write-Host "[JKS Cfg] Fail***: Script ran but files were not created. Check for errors." -ForegroundColor Red
                        }
                    } else {
                        Write-Host "[JKS Cfg] Error: Script execution failed." -ForegroundColor Red
                    }
                }
            }
        }

        # --- JDBC Driver Installation ---
        if ($nodeType -eq "A") {
            Install-JdbcDrivers
        }

        # --- Restart Service if needed ---
        if ($global:AgentRestartNeeded) {
            if ((Get-Service $serviceName -ErrorAction SilentlyContinue) -and (Confirm-UserChoice -Prompt "`nDrivers were installed/updated. Restart the '$serviceName' service now (requires Admin)?" -DefaultChoice 'y')) {
                if (Test-IsAdmin) {
                    Restart-Service -Name $serviceName -Force
                    Write-Host "[Service] Success: Service '$serviceName' restarted."
                } else {
                    $command = "Restart-Service -Name '$serviceName' -Force"
                    if ((Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $striimInstallPath) -eq 0) {
                        Write-Host "[Service] Success: Service '$serviceName' restarted (elevated)."
                    } else {
                        Write-Host "[Service] Error: Failed to restart service." -ForegroundColor Red
                    }
                }
            } else {
                Write-Host "`n[Service] Drivers were installed/updated. Please restart the Striim Agent service manually for changes to take effect." -ForegroundColor Yellow
            }
        }


        # --- Final Summary ---
        Write-Host "`n------------------------------------------"
        Write-Host "Striim Configuration Check Finished." -ForegroundColor Green
        Write-Host "------------------------------------------"
        Write-Host "Please review the output for any warnings or errors."
        Write-Host "Some changes (like System PATH or Java installation) may require restarting your terminal or system."
    }
}
catch {
    Write-Host "`n------------------------------------------" -ForegroundColor Red
    Write-Host "An unexpected error occurred:" -ForegroundColor Red
    # Check for specific error records for more detail
    if ($_.Exception) {
        Write-Host $_.Exception.Message -ForegroundColor Red
    } else {
        Write-Host $_ -ForegroundColor Red
    }
    Write-Host "Script execution halted." -ForegroundColor Red
    Write-Host "------------------------------------------"
    # Exit with a non-zero status code to indicate failure
    exit 1
}
finally {
    # Pause at the end of the script if running in a console, so the user can see the output
    if ($Host.Name -eq "ConsoleHost" -and -not $DownloadOnly) {
        Read-Host "Press Enter to exit the script."
    }
}
