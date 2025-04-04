# Display startup summary with checks and admin requirements
Write-Host "`nStarting Striim Configuration Check..."
Write-Host "------------------------------------------"
Write-Host "[Checks Summary] The following validations will occur:"
Write-Host "  • [System]        Verify ≥15GB free space on C: drive"
Write-Host "  • [Node Type]     Auto-detect Agent/Node or prompt for selection"
Write-Host "  • [Striim Setup]  Install Striim if missing"
Write-Host "  • [Config Files]  Validate agent.conf/startUp.properties settings"
Write-Host "  • [System PATH]   Add Striim lib directory to system PATH (requires Admin) $([char]0x27a4)" -ForegroundColor Yellow
Write-Host "  • [Dependencies]  Verify DLLs (icudt72.dll, icuuc72.dll, MSSQLNative.dll) or download them"
Write-Host "  • [Java Check]    Ensure compatible Java version installed. Offer download if missing"
Write-Host "  • [Security]      Setup Integrated Security sqljdbc_auth.dll (requires Admin) $([char]0x27a4)" -ForegroundColor Yellow
Write-Host "  • [Patches]       Apply critical fixes for Striim v4.2.0.20 (if applicable)"
Write-Host "  • [Service]       Configure Striim as Windows Service (requires Admin) $([char]0x27a4)" -ForegroundColor Yellow
Write-Host "  • [Config ]       Prompt to generate or recreate jks and key files (if applicable)"
Write-Host "------------------------------------------"
Read-Host "`nPress Enter to continue or Ctrl+C to abort..."

# Define default values
$striimInstallPath = Get-Location
$downloadDir = -join ($striimInstallPath, "\downloads")

# Get the free space of the C: drive in GB
$freeSpaceGB = (Get-PSDrive -Name C).Free / 1GB

# Check if the free space is less than 15 GB
if ($freeSpaceGB -lt 15) {
    Write-Host "[Envrnmt]       Insufficient disk space on C: drive. At least 15 GB of free space is required."
    exit 1  # Stop the script and indicate failure
} else {
    Write-Host "[Envrnmt]       Sufficient disk space available on C: drive."
}

# Step 1: Create downloads folder if it doesn't exist
if (-not (Test-Path $downloadDir)) {
    New-Item -ItemType Directory -Force -Path $downloadDir
}

function Download-FileIfNotExists {
    param(
        [string]$Uri,
        [string]$DownloadPath
    )

    if (Test-Path -Path $DownloadPath) {
        Write-Host "[Download] File already exists: $DownloadPath"
        return $true # Exit the function, don't download
    }

    Write-Host "[Download] Downloading from $Uri to $DownloadPath..."

    try {
        # Try using HttpClient (faster)
        $httpClient = New-Object System.Net.Http.HttpClient

        # Set a reasonable timeout (e.g., 60 seconds) - IMPORTANT!
        $httpClient.Timeout = New-TimeSpan -Seconds 60

        $response = $httpClient.GetAsync($Uri).Result

        if ($response.IsSuccessStatusCode) {
            $content = $response.Content.ReadAsByteArrayAsync().Result
            [System.IO.File]::WriteAllBytes($DownloadPath, $content)
            Write-Host "[Download] Success: File downloaded using HttpClient."
        } else {
            Write-Host "[Download] HttpClient download failed. HTTP Status: $($response.StatusCode)" -ForegroundColor Yellow
            throw "HttpClient failed with status code: $($response.StatusCode)" # Throw an exception to trigger the catch block
        }

        $httpClient.Dispose()

    }
    catch {
        Write-Host "[Download] Attempting fallback download using Invoke-WebRequest (slower)..." -ForegroundColor Yellow
        try {
            # Fallback to Invoke-WebRequest (slower, but more compatible)
            Invoke-WebRequest -Uri $Uri -OutFile $DownloadPath -ErrorAction Stop -UseBasicParsing
            Write-Host "[Download] Success: File downloaded using Invoke-WebRequest."
        }
        catch {
            Write-Host "[Download] Error: Failed to download file using both methods. $($_.Exception.Message)" -ForegroundColor Red
            #  Critical error:  Could not download with either method.
            #  Consider exiting the script here, or taking other drastic action.
            return $false  # Indicate failure
        }
    }
    finally {
        #  Dispose of HttpClient *always*, even if there's an exception
        if ($httpClient) {
            $httpClient.Dispose()
        }
    }
    return $true #Indicate Success
}


# Function to check if the script is running as administrator
function Test-IsAdmin {
  ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to execute a command as administrator in a new PowerShell process
function Invoke-AsAdmin {
  param(
    [string]$ArgumentList,
    [string]$WorkingDirectory
  )

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = "powershell.exe"
  #$psi.Arguments = '-NoExit -NoProfile -ExecutionPolicy Bypass -Command "' + $ArgumentList + '"'  # Single quotes
  $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -Command "' + $ArgumentList + '"'  # Single quotes
  $psi.Verb = "RunAs"
  $psi.WorkingDirectory = $WorkingDirectory
  #$psi.RedirectStandardOutput = $true
  #$psi.RedirectStandardError = $true
  #$psi.UseShellExecute = $false

  $process = [System.Diagnostics.Process]::Start($psi)

  #$stdOut = $process.StandardOutput.ReadToEnd()
  #$stdErr = $process.StandardError.ReadToEnd()

  $process.WaitForExit()

  #$stdOut | Out-File -FilePath "$env:TEMP\setup_stdout.txt" -Append -Encoding UTF8
  #$stdErr | Out-File -FilePath "$env:TEMP\setup_stderr.txt" -Append -Encoding UTF8

  return $process.ExitCode
}

# Step 2: Check for agent.conf or startUp.properties to determine node type
$agentConfPath = -join ($striimInstallPath, "\conf\agent.conf")
$startUpPropsPath = -join ($striimInstallPath, "\conf\startUp.properties")

if (Test-Path $agentConfPath) {
    $nodeType = "A"  # Agent
    Write-Host "[Envrnmt]       -> AGENT Environment based on agent.conf"
} elseif (Test-Path $startUpPropsPath) {
    $nodeType = "N"  # Node
    Write-Host "[Envrnmt]       -> NODE environment based on startUp.properties"
} else {
    # If neither file is found, check if Striim is installed, and offer to download/install
    $striimLibPath = -join ($striimInstallPath, "\lib")
    if (-not (Test-Path $striimLibPath -PathType Container)) {
        Write-Host "[Envrnmt] Striim installation not found in current directory."
        $downloadStriim = Read-Host "[Envrnmt]  Do you want to automatically download and install Striim? (Y/N)"

        if ($downloadStriim.ToUpper() -eq "Y") {
            $nodeType = Read-Host "[Envrnmt]  Do you want to install Striim Node (N) or Agent (A)? (Enter 'N' or 'A')"
            $nodeType = $nodeType.ToUpper()

            $urlAddAgent = ""

             while ($nodeType -ne "N" -and $nodeType -ne "A")
            {
                Write-Host "Invalid input. Please enter 'N' for Node or 'A' for Agent."
                $nodeType = Read-Host "[Envrnmt]  Do you want to install Striim Node (N) or Agent (A)? (Enter 'N' or 'A')"
                $nodeType = $nodeType.ToUpper()
            }

            if ($nodeType -eq "N") {
                $defaultStriimInstallPath = "C:\striim"
                $defaultStriimExtractPath = "C:\"
            } else {
                $defaultStriimInstallPath = "C:\striim\Agent"
                $defaultStriimExtractPath = "C:\striim\"
                $urlAddAgent = "Agent_"
            }

            $striimVersion = ""

            $downloadUrl = ""

            do {
                $striimVersion = Read-Host "[Envrnmt] Enter the Striim version you want to install (e.g., 4.2.0.20 or 5.0.6)"
                Write-Host "[Envrnmt] Valid version detected: $striimVersion"

                $downloadUrl = "https://striim-downloads.striim.com/Releases/$striimVersion/Striim_$urlAddAgent$striimVersion.zip"

                Write-Host "[Envrnmt] Checking download path: $downloadUrl"

                try {
                    # Check URL existence using HEAD request without downloading
                    Invoke-WebRequest -Uri $downloadUrl -Method Head -UseBasicParsing -ErrorAction Stop | Out-Null

                    Write-Host "[Envrnmt] Success! Striim Download path is valid: $downloadUrl"
                    $valid = $true  # Exit loop if URL exists
                }
                catch {
                    Write-Warning "[Envrnmt] INVALID version detected: $striimVersion"
                    $valid = $false
                }

            } while (-not $valid)

            Write-Host "[Envrnmt]  Striim will be downloaded from: $downloadUrl"

            $striimInstallPathInput = Read-Host "[Envrnmt]  Enter Striim installation path (Default: '$defaultStriimInstallPath')"
            if ($striimInstallPathInput -ne "") {
                $striimInstallPath = $striimInstallPathInput
                $defaultStriimExtractPath = $striimInstallPath
            } else {
                $striimInstallPath = $defaultStriimInstallPath
            }

            # Create the Striim install directory if it does not exist.
            if (-not (Test-Path -Path $striimInstallPath))
            {
                try {
                    New-Item -ItemType Directory -Path $striimInstallPath -Force -ErrorAction Stop | Out-Null
                    Write-Host "[Envrnmt] Success: Created Striim installation directory at '$striimInstallPath'."
                }
                catch {
                    Write-Host "[Envrnmt] Error: Failed to create directory '$striimInstallPath': $($_.Exception.Message)"
                    exit 1
                }
            }

            # Mock Download URL (replace with actual logic later)
            # https://striim-downloads.striim.com/Releases/5.0.6/Striim_Agent_5.0.6.tgz
            # https://striim-downloads.striim.com/Releases/5.0.6/Striim_5.0.6.zip

            if (!(Test-Path -Path $striimInstallPath -PathType Container)) {
                try {
                    New-Item -ItemType Directory -Path $striimInstallPath -Force -ErrorAction Stop
                    Write-Host "[Envrnmt] Created installation directory: $striimInstallPath"
                }
                catch {
                    Write-Error "[Error] Failed to create directory: $($_.Exception.Message)"
                    exit 1  # Exit with a non-zero exit code to indicate failure
                }
            }

            # Check if the ZIP file already exists before downloading
            $zipFilePath = Join-Path -Path $striimInstallPath -ChildPath "Striim_$urlAddAgent$striimVersion.zip"
            if (Test-Path -Path $zipFilePath) {
                Write-Host "[Envrnmt] ZIP file already exists: $zipFilePath. Skipping download."
            } else {
                Write-Host "[Envrnmt] Downloading Striim from $downloadUrl to $zipFilePath... (note: this may take a few minutes depending on your internet speed; there is no progress bar)"

                $downloadSuccess = Download-FileIfNotExists -Uri $downloadUrl -DownloadPath $zipFilePath

                if ($downloadSuccess) {
                    Write-Host "[Envrnmt] Download complete!"
                } else {
                    Write-Host "[Envrnmt] Download failed!" -ForegroundColor Red
                    #  Handle the download failure (e.g., exit the script, retry, log an error)
                }
            }

            # --- Extract the ZIP File ---
            try {
                Write-Host "[Envrnmt] Extracting Striim to $striimInstallPath..."
                Expand-Archive -Path $zipFilePath -DestinationPath $defaultStriimExtractPath -Force -ErrorAction Stop
            }
            catch {
                Write-Error "[Error] Extraction failed: $($_.Exception.Message)"
                  # Attempt to clean up the potentially partially-extracted files.
                if (Test-Path -Path $striimInstallPath) {
                   try {
                        Remove-Item -Path $striimInstallPath\* -Recurse -Force -ErrorAction SilentlyContinue # try deleting
                      }
                  catch {
                      Write-Warning "[Warning] Could not completley clean up failed extraction in $striimInstallPath. Manual cleanup may be needed"
                  }
                }
                exit 1
            }

            # Remove the downloaded ZIP file after successful extraction.
            Write-Host "[Service]   Striim Install saved here: $zipFilePath"
            Write-Host "[Service]   Striim Install         *** Clean up manually if required ***"

        } else {
            # If neither file is found, ask the user
            $nodeType = Read-Host "[Envrnmt]  Is this Agent (default) or Node? (Enter 'A' for Agent or 'N' for Node)"
            if ($nodeType -eq "") { $nodeType = "A" }
            $nodeType = $nodeType.ToUpper()
        }
    }
    else {
        # If neither file is found, ask the user
        $nodeType = Read-Host "[Envrnmt]  Is this Agent (default) or Node? (Enter 'A' for Agent or 'N' for Node)"
        if ($nodeType -eq "") { $nodeType = "A" }
        $nodeType = $nodeType.ToUpper()
    }
}

# Ask user for Striim install path only if node type couldn't be auto-detected
if ($nodeType -eq "" -or (-not (Test-Path $striimInstallPath -PathType Container))) {
    # Determine default path based on node type, only if not already set by download process
    if(-not $striimInstallPath) {
        if ($nodeType -eq "N") {
            $defaultStriimInstallPath = "C:\striim"
        }
        else {
            $defaultStriimInstallPath = "C:\striim\agent"
        }
    } else {
        $defaultStriimInstallPath = $striimInstallPath
    }
    $striimInstallPathInput = Read-Host "[Envrnmt] Provide Striim install path or press Enter to default to $($defaultStriimInstallPath):"
    if ($striimInstallPathInput -ne "") {
        $striimInstallPath = $striimInstallPathInput
    } else {
        $striimInstallPath = $defaultStriimInstallPath
    }
    Write-Host "[Envrnmt] Success: User set Striim Install Path set to: $striimInstallPath"
} else {
    Write-Host "[Envrnmt] Success: Striim Install Path set to: $striimInstallPath"
}

if ($nodeType -eq "A") {
    $jksFileName = "aks.jks"
    $pwdFileName = "aksKey.pwd"
    $configScriptName = "aksConfig.bat"
    $configType = "AKS"
    Write-Host "Node type set to 'A' (AKS)" -ForegroundColor Cyan
}
elseif ($nodeType -eq "N") {
    $jksFileName = "sks.jks"
    $pwdFileName = "sksKey.pwd"
    $configScriptName = "sksConfig.bat"
    $configType = "SKS"
     Write-Host "Node type set to 'N' (SKS)" -ForegroundColor Cyan
}

# Agent-specific checks
if ($nodeType -eq "A") {
    Write-Host "[Config ]       -> AGENT -> Specific Tests for configuration:"
    $agentConfPath = -join ($striimInstallPath, "\conf\agent.conf")

    # Check if agent.conf exists
    if (Test-Path $agentConfPath) {
        $agentConfLines = Get-Content $agentConfPath

        # Define required and optional properties
        $requiredProps = "striim.cluster.clusterName", "striim.node.servernode.address"
        $optionalProps = "MEM_MAX"  # Example optional property
        $propsFound = @{}  # Dictionary to track found properties

        # Initialize required properties as not found
        foreach ($prop in $requiredProps) {
            $propsFound[$prop] = $false
        }

        # Process each line and check for required properties
        foreach ($lineIndex in 0..($agentConfLines.Length - 1)) {
            $line = $agentConfLines[$lineIndex]

            foreach ($prop in $requiredProps + $optionalProps) {
                if ($line -match "^#?\s*$prop\s*=\s*(.*)") {
                    $propValue = $matches[1]

                    if ($line.StartsWith("#")) {
                        # Property is commented out
                        Write-Host "[Config ] '$prop' is commented out. Would you like to uncomment it and set a value? (y/n)"
                        $response = Read-Host
                        if ($response -eq "y") {
                            $newValue = Read-Host "Enter a value for $prop"
                            $agentConfLines[$lineIndex] = "$prop=$newValue"  # Update the line
                            Write-Host "[Config ] Updated '$prop' with value: $newValue"
                        } else {
                            Write-Host "[Config ] '$prop' remains commented out."
                        }
                    } elseif ($propValue -eq "") {
                        # Property is empty, ask the user for a value
                        Write-Host "[Config ] '$prop' is empty. Please provide a value."
                        $newValue = Read-Host "Enter a value for $prop"
                        $agentConfLines[$lineIndex] = "$prop=$newValue"  # Update the line
                        Write-Host "[Config ] Updated '$prop' with value: $newValue"
                    } else {
                        # Property has a valid value
                        Write-Host "[Config ] Success: '$prop' found in agent.conf with value: $propValue"
                        Write-Host "[Config ] Current value: '$prop' is '$propValue'. Do you want to update it? (y/n, default is no)"
                        $response = Read-Host
                        if ($response -eq "y") {
                            $newValue = Read-Host "Enter the new value for $prop"
                            $agentConfLines[$lineIndex] = "$prop=$newValue"  # Update the line
                            Write-Host "[Config ] Updated '$prop' with new value: $newValue"
                        } else {
                            Write-Host "[Config ] '$prop' remains unchanged."
                        }
                        $propsFound[$prop] = $true
                    }
                    break  # Exit the inner loop once a property is found on a line
                }
            }
        }

        # Save the updated content to the file
        Write-Host "[Config ] Saving changes to agent.conf"
        Set-Content $agentConfPath $agentConfLines
    } else {
        Write-Host "[Config ] Fail***: agent.conf not found in $($striimInstallPath)\conf"
    }
}

# Node-specific checks
if ($nodeType -eq "N") {
    Write-Host "[Config ]       -> NODE -> Specific Tests for configuration:"
    $startUpPropsPath = -join ($striimInstallPath, "\conf\startUp.properties")

    # Check if startUp.properties exists
    if (Test-Path $startUpPropsPath) {
        $startUpPropsLines = Get-Content $startUpPropsPath

        # Define required and optional properties
        $requiredProps = "CompanyName", "LicenceKey", "ProductKey", "WAClusterName"
        $optionalProps = "MEM_MAX"
        $propsFound = @{}  # Dictionary to track found properties

        # Initialize required properties as not found
        foreach ($prop in $requiredProps) {
            $propsFound[$prop] = $false
        }

        # Process each line and check for required properties
        foreach ($lineIndex in 0..($startUpPropsLines.Length - 1)) {
            $line = $startUpPropsLines[$lineIndex]

            foreach ($prop in $requiredProps + $optionalProps) {
                if ($line -match "^#?\s*$prop\s*=\s*(.*)") {
                    $propValue = $matches[1]

                    if ($line.StartsWith("#")) {
                        # Property is commented out
                        Write-Host "[Config ] '$prop' is commented out. Would you like to uncomment it and set a value? (y/n)"
                        $response = Read-Host
                        if ($response -eq "y") {
                            $newValue = Read-Host "Enter a value for $prop"
                            $startUpPropsLines[$lineIndex] = "$prop=$newValue"  # Update the line
                            Write-Host "[Config ] Updated '$prop' with value: $newValue"
                        } else {
                            Write-Host "[Config ] '$prop' remains commented out."
                        }
                    } elseif ($propValue -eq "") {
                        # Property is empty, ask the user for a value
                        Write-Host "[Config ] '$prop' is empty. Please provide a value."
                        $newValue = Read-Host "Enter a value for $prop"
                        $startUpPropsLines[$lineIndex] = "$prop=$newValue"  # Update the line
                        Write-Host "[Config ] Updated '$prop' with value: $newValue"
                    } else {
                        # Property has a valid value
                        Write-Host "[Config ] Success: '$prop' found in startUp.properties with value: $propValue"
                        Write-Host "[Config ] Current value: '$prop' is '$propValue'. Do you want to update it? (y/n, default is no)"
                        $response = Read-Host
                        if ($response -eq "y") {
                            $newValue = Read-Host "Enter the new value for $prop"
                            $startUpPropsLines[$lineIndex] = "$prop=$newValue"  # Update the line
                            Write-Host "[Config ] Updated '$prop' with new value: $newValue"
                        } else {
                            Write-Host "[Config ] '$prop' remains unchanged."
                        }
                        $propsFound[$prop] = $true
                    }
                    break  # Exit the inner loop once a property is found on a line
                }
            }
        }

        # Save the updated content to the file
        Write-Host "[Config ] Saving changes to startUp.properties"
        Set-Content $startUpPropsPath $startUpPropsLines
    } else {
        Write-Host "[Config ] Fail***: startUp.properties not found in $($striimInstallPath)\conf"
    }
}


# Check if Striim lib directory is in PATH
$striimLibPath = Join-Path $striimInstallPath "\lib" # Use Join-Path, it's more reliable than string concatenation
Write-Host "[Config ]       -> Striim Lib Path set to: $striimLibPath"

if ($env:Path -split ";" -icontains $striimLibPath) {
    Write-Host "[Config ] Success: Striim lib directory found in PATH."
} else {
    Write-Host "[Config ] Fail***: Striim lib directory not found in PATH."
    Write-Host "[Config ]  (Requires Running Powershell as Administrator) -"
	  $addToPathChoice = Read-Host "[Config ]  Do you want to add it to the system PATH? (Y/N)"
    if ($addToPathChoice.ToUpper() -eq "Y") {
        # --- ELEVATED SECTION (within the if) ---
        if (Test-IsAdmin) {
            # We are already running as admin, so we can directly modify the environment variable.
            $newPath = $env:Path + ";" + $striimLibPath
            [Environment]::SetEnvironmentVariable("Path", $newPath, "Machine") # Set for all users
            Write-Host "[Config ] Success: Striim lib directory added to PATH (already elevated)."

        } else {
            # Not running as admin, so re-launch this part as admin
            Write-Host "[Config ]  Elevating permissions to add Striim to PATH..."

            # Build the command to execute as admin.  This needs to be a *single* string.
            # We use single quotes (') around the inner command to avoid variable expansion *here*.
            # We want the variables expanded *inside* the elevated process.
            $commandToRunAsAdmin = "[Environment]::SetEnvironmentVariable('Path', ((Get-Item Env:Path).Value + ';$striimLibPath'), 'Machine')"

            # Call Invoke-AsAdmin with the command.
            $currentWorkingDirectory = Get-Location
            $exitCode = Invoke-AsAdmin -ArgumentList $commandToRunAsAdmin -WorkingDirectory $currentWorkingDirectory

            if ($exitCode -eq 0) {
                Write-Host "[Config ] Success: Striim lib directory added to PATH (elevated)."
            } else {
                Write-Host "[Config ] Error: Failed to add Striim lib directory to PATH (exit code: $exitCode)." -ForegroundColor Red
            }
        }
        # --- END ELEVATED SECTION ---

        # Refresh the current session's environment variables *only if running as admin*.  Otherwise, it won't see the changes.
        if(Test-IsAdmin){
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
        }
    }
}

# Check for required DLLs
$requiredDlls = "icudt72.dll", "icuuc72.dll", "MSSQLNative.dll"
foreach ($dll in $requiredDlls) {
    if (Test-Path "$striimLibPath\$dll") {
        Write-Host "[DLLs   ] Success: $dll found in $striimLibPath"
    } else {
        # Offer to download DLLs
        $downloadChoice = Read-Host "[DLLs   ]  $striimLibPath\$dll not found. Download it? (Y/N)"
        if ($downloadChoice.ToUpper() -eq "Y") {
            # Create download directory if it doesn't exist
            if (-not (Test-Path $downloadDir)) {
                New-Item -ItemType Directory -Force -Path $downloadDir
            }

            if ($dll -eq "MSSQLNative.dll") {
                $downloadUrl = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/MSSQLNative.dll"
            } else {
                $downloadUrl = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/Dlls.zip"
            }

			$finalPath = -join ($striimLibPath, "\Dlls\", $dll)

			if (-not (Test-Path $finalPath)) {

				$downloadPath = -join ($downloadDir,  "\", $downloadUrl.Split("/")[-1])
				Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath

				if ($downloadUrl.EndsWith(".zip")) {
					Expand-Archive -Path $downloadPath -DestinationPath $striimLibPath -Force
					Copy-Item $finalPath $striimLibPath
					Remove-Item $finalPath -Force
				} else {
					Copy-Item $downloadPath $striimLibPath
					Remove-Item $downloadPath -Force
				}
			} else {
				Copy-Item $finalPath $striimLibPath
				Remove-Item $finalPath -Force
			}

            Write-Host "[DLLs   ] Success: $dll downloaded and extracted to $striimLibPath"
        }
    }
}

$tempFolders = "$striimLibPath\__MACOSX", "$striimLibPath\Dlls"
foreach ($folder in $tempFolders) {
	if (Test-Path $folder) {
		Remove-Item $folder -Recurse -Force
		Write-Host "[Cleanup] Success: Deleted temporary folder: $folder"
	}
}

# Get the Striim version by parsing the filename 'Platform-4.2.0.20.jar'
$striimJarFile = Get-ChildItem -Path $striimLibPath -Filter "Platform-*.jar" | Select-Object -First 1
if ($striimJarFile) {
    # Extract the version number from the filename (e.g., 4.2.0.20)
    $striimVersion = $striimJarFile.Name -replace '^Platform-(\d+\.\d+\.\d+\.\d+)\.jar$', '$1'
    Write-Host "[Striim  ] Success: Found Striim version: $striimVersion"

    # Extract the major version (e.g., 4 from 4.2.0.20)
    $majorVersion = $striimVersion.Split('.')[0]
    Write-Host "[Striim  ] Major version: $majorVersion"
} else {
    Write-Host "[Striim  ] Fail***: Could not find Striim JAR file."
    exit
}

# Determine the required Java version based on Striim's major version
if ($majorVersion -le 4) {
    # Striim major version 4 or less requires Java 8
    $requiredJavaVersion = "Java 8"
    $javaDownloadUrl = "https://builds.openlogic.com/downloadJDK/openlogic-openjdk/8u422-b05/openlogic-openjdk-8u422-b05-windows-x64.msi"
} elseif ($majorVersion -ge 5) {
    # Striim major version 5 or greater requires Java 11
    $requiredJavaVersion = "Java 11"
    $javaDownloadUrl = "https://aka.ms/download-jdk/microsoft-jdk-11.0.26-windows-x64.msi"
} else {
    Write-Host "[Striim  ] Error: Unknown Striim version. Exiting..."
    exit
}

# Check if Java is installed
if (Get-Command java -ErrorAction SilentlyContinue) {
    $javaVersionOutput = java -version 2>&1 | Select-String -Pattern 'java version "(.*)"'
    Write-Host "[Java   ] Success: Java version: $javaVersionOutput"

    if ($javaVersionOutput) {
        $javaVersion = $javaVersionOutput.Matches.Groups[1].Value
        if ($javaVersion -match "1\.8" -or $javaVersion -match "18\.0\.\d+\.\d+") {
            Write-Host "[Java   ] Success: Java 8 found."
            if ($requiredJavaVersion -eq "Java 11") {
                Write-Host "[Java   ] Java 8 detected, but Java 11 is required for Striim version $majorVersion. Please install Java 11."
                # Offer to download Java 11 if Java 8 is detected but Java 11 is required
                $downloadJavaChoice = Read-Host "  Download Java 11? (Y/N)"
                if ($downloadJavaChoice.ToUpper() -eq "Y") {
                    $javaDownloadPath = Join-Path $downloadDir ($javaDownloadUrl.Split("/")[-1])
                    $downloadSuccess = Download-FileIfNotExists -Uri $javaDownloadUrl -DownloadPath $javaDownloadPath
                    # Invoke-WebRequest -Uri $javaDownloadUrl -OutFile $javaDownloadPath
                    # Write-Host "[Java   ] Success: Java 11 installer downloaded to $javaDownloadPath. Please install it."
                    if ($downloadSuccess) {
                        # Ask the user if they want to run the installer
                        $runInstallerChoice = Read-Host "Java installer downloaded to $javaDownloadPath.  Run installer now (as administrator)? (Y/N)"

                        if ($runInstallerChoice.ToUpper() -eq "Y") {
                            if(Test-IsAdmin){
                                #Already admin
                                Start-Process -FilePath $javaDownloadPath
                            }
                            else{
                                #Run as Admin
                                 $exitCode = Invoke-AsAdmin -ArgumentList "Start-Process -FilePath '$javaDownloadPath'" -WorkingDirectory $downloadDir
                                 if ($exitCode -eq 0) {
                                    Write-Host "Installer launched successfully."
                                 }
                                 else{
                                    Write-Host "Installer launch failed"
                                 }
                            }
                        }
                    } else {
                        Write-Host "Java download failed.  Cannot proceed with installation." -ForegroundColor Red
                    }
                }
            }
        } elseif ($javaVersion -match "11\.0") {
            Write-Host "[Java   ] Success: Java 11 found."
            if ($requiredJavaVersion -eq "Java 8") {
                Write-Host "[Java   ] Java 11 detected, but Java 8 is required for Striim version $majorVersion. Please install Java 8."
                # Offer to download Java 8 if Java 11 is detected but Java 8 is required
                $downloadJavaChoice = Read-Host "  Download Java 8? (Y/N)"
                if ($downloadJavaChoice.ToUpper() -eq "Y") {
                    $javaDownloadUrl = "https://builds.openlogic.com/downloadJDK/openlogic-openjdk/8u422-b05/openlogic-openjdk-8u422-b05-windows-x64.msi"
                    $javaDownloadPath = Join-Path $downloadDir ($javaDownloadUrl.Split("/")[-1])
                    $downloadSuccess = Download-FileIfNotExists -Uri $javaDownloadUrl -DownloadPath $javaDownloadPath
                    #Invoke-WebRequest -Uri $javaDownloadUrl -OutFile $javaDownloadPath
                    #Write-Host "[Java   ] Success: Java 8 installer downloaded to $javaDownloadPath. Please install it."
                    if ($downloadSuccess) {
                        # Ask the user if they want to run the installer
                        $runInstallerChoice = Read-Host "Java installer downloaded to $javaDownloadPath.  Run installer now (as administrator)? (Y/N)"

                        if ($runInstallerChoice.ToUpper() -eq "Y") {
                            if(Test-IsAdmin){
                                #Already admin
                                Start-Process -FilePath $javaDownloadPath
                            }
                            else{
                                #Run as Admin
                                 $exitCode = Invoke-AsAdmin -ArgumentList "Start-Process -FilePath '$javaDownloadPath'" -WorkingDirectory $downloadDir
                                 if ($exitCode -eq 0) {
                                    Write-Host "Installer launched successfully."
                                 }
                                 else{
                                    Write-Host "Installer launch failed"
                                 }
                            }
                        }
                    } else {
                        Write-Host "Java download failed.  Cannot proceed with installation." -ForegroundColor Red
                    }
                }
            }
        } else {
            Write-Host "[Java   ] Fail***: Unsupported Java version detected."
            # Offer to download the required Java version
            $downloadJavaChoice = Read-Host "  Download $requiredJavaVersion? (Y/N)"
            if ($downloadJavaChoice.ToUpper() -eq "Y") {
                $javaDownloadPath = Join-Path $downloadDir ($javaDownloadUrl.Split("/")[-1])
                $downloadSuccess = Download-FileIfNotExists -Uri $javaDownloadUrl -DownloadPath $javaDownloadPath
                #Invoke-WebRequest -Uri $javaDownloadUrl -OutFile $javaDownloadPath
                #Write-Host "[Java   ] Success: $requiredJavaVersion installer downloaded to $javaDownloadPath. Please install it."
                if ($downloadSuccess) {
                    # Ask the user if they want to run the installer
                    $runInstallerChoice = Read-Host "Java installer downloaded to $javaDownloadPath.  Run installer now (as administrator)? (Y/N)"

                    if ($runInstallerChoice.ToUpper() -eq "Y") {
                        if(Test-IsAdmin){
                            #Already admin
                            Start-Process -FilePath $javaDownloadPath
                        }
                        else{
                            #Run as Admin
                             $exitCode = Invoke-AsAdmin -ArgumentList "Start-Process -FilePath '$javaDownloadPath'" -WorkingDirectory $downloadDir
                             if ($exitCode -eq 0) {
                                Write-Host "Installer launched successfully."
                             }
                             else{
                                Write-Host "Installer launch failed"
                             }
                        }
                    }
                } else {
                    Write-Host "Java download failed.  Cannot proceed with installation." -ForegroundColor Red
                }
            }
        }
    } else {
        Write-Host "[Java   ] Fail***: Could not determine Java version."
    }
} else {
    # Java is not installed, offer to download the required version
    Write-Host "[Java   ] Java not found. Download $requiredJavaVersion? (Y/N)"
    $downloadJavaChoice = Read-Host
    if ($downloadJavaChoice.ToUpper() -eq "Y") {
        $javaDownloadPath = Join-Path $downloadDir ($javaDownloadUrl.Split("/")[-1])
        $downloadSuccess = Download-FileIfNotExists -Uri $javaDownloadUrl -DownloadPath $javaDownloadPath
        #Invoke-WebRequest -Uri $javaDownloadUrl -OutFile $javaDownloadPath
        #Write-Host "[Java   ] Fail***: $requiredJavaVersion installer downloaded to $javaDownloadPath. Please install it."
        if ($downloadSuccess) {
            # Ask the user if they want to run the installer
            $runInstallerChoice = Read-Host "Java installer downloaded to $javaDownloadPath.  Run installer now (as administrator)? (Y/N)"

            if ($runInstallerChoice.ToUpper() -eq "Y") {
                if(Test-IsAdmin){
                    #Already admin
                    Start-Process -FilePath $javaDownloadPath
                }
                else{
                    #Run as Admin
                     $exitCode = Invoke-AsAdmin -ArgumentList "Start-Process -FilePath '$javaDownloadPath'" -WorkingDirectory $downloadDir
                     if ($exitCode -eq 0) {
                        Write-Host "Installer launched successfully."
                     }
                     else{
                        Write-Host "Installer launch failed"
                     }
                }
            }
        } else {
            Write-Host "Java download failed.  Cannot proceed with installation." -ForegroundColor Red
        }
    }
}

# --- Integrated Security Section ---
# Determine if Integrated Security is needed
$sqljdbcAuthDllPath = "C:\Windows\System32\sqljdbc_auth.dll"
if (Test-Path $sqljdbcAuthDllPath) {
    Write-Host "[Int Sec] Success: Integrated Security: sqljdbc_auth.dll found in C:\Windows\System32"
} else {
    # Ask about Integrated Security only if the DLL is missing
    $useIntegratedSecurity = Read-Host "[Int Sec] Plan to use Integrated Security? (Y/N)"
    if ($useIntegratedSecurity.ToUpper() -eq "Y") {
        # Check if it exists in /lib
        $sourceDllPath = Join-Path $striimLibPath "sqljdbc_auth.dll"
        if (Test-Path $sourceDllPath) {
            $copyChoice = Read-Host "[Int Sec] Integrated Security:  sqljdbc_auth.dll found in $striimLibPath. Copy it to C:\Windows\System32? (Y/N)"
            if ($copyChoice.ToUpper() -eq "Y") {
                # --- ELEVATED COPY ---
                if (Test-IsAdmin) {
                    Copy-Item -Path $sourceDllPath -Destination $sqljdbcAuthDllPath -Force
                    Write-Host "[Int Sec] Success: Integrated Security: sqljdbc_auth.dll copied to C:\Windows\System32 (already elevated)"
                } else {
                    $copyCommand = "Copy-Item -Path '$sourceDllPath' -Destination '$sqljdbcAuthDllPath' -Force"
                    $currentWorkingDirectory = Get-Location
                    $exitCode = Invoke-AsAdmin -ArgumentList $copyCommand -WorkingDirectory $currentWorkingDirectory
                    if ($exitCode -eq 0) {
                        Write-Host "[Int Sec] Success: Integrated Security: sqljdbc_auth.dll copied to C:\Windows\System32 (elevated)"
                    } else {
                        Write-Host "[Int Sec] Error: Failed to copy sqljdbc_auth.dll (exit code: $exitCode)." -ForegroundColor Red
                    }
                }
                # --- END ELEVATED COPY ---
            }
        } else {
            # Offer to download
            $downloadChoice = Read-Host "[Int Sec] Integrated Security:  sqljdbc_auth.dll not found. Download it? (Y/N)"
            if ($downloadChoice.ToUpper() -eq "Y") {
                $downloadUrl = "https://github.com/daniel-striim/StriimQueryAutoLoader/raw/main/MSJet/sqljdbc_auth.dll" # Update with the correct URL if needed
                $downloadPath = Join-Path $downloadDir "sqljdbc_auth.dll"
                try {
                    Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -ErrorAction Stop  # Use -ErrorAction Stop for proper error handling
                    # --- ELEVATED COPY (after download) ---
                    if (Test-IsAdmin) {
                        Copy-Item -Path $downloadPath -Destination $sqljdbcAuthDllPath -Force
                        Write-Host "[Int Sec] Success: Integrated Security: sqljdbc_auth.dll downloaded and copied to C:\Windows\System32 (already elevated)"
                    } else {
                        $copyCommand = "Copy-Item -Path '$downloadPath' -Destination '$sqljdbcAuthDllPath' -Force"
                        $currentWorkingDirectory = Get-Location
                        $exitCode = Invoke-AsAdmin -ArgumentList $copyCommand -WorkingDirectory $currentWorkingDirectory
                        if ($exitCode -eq 0) {
                            Write-Host "[Int Sec] Success: Integrated Security: sqljdbc_auth.dll downloaded and copied to C:\Windows\System32 (elevated)"
                        } else {
                             Write-Host "[Int Sec] Error: Failed to copy downloaded sqljdbc_auth.dll (exit code: $exitCode)." -ForegroundColor Red
                        }
                    }
                    # --- END ELEVATED COPY ---

                } catch {
                    Write-Host "[Int Sec] Error: Failed to download sqljdbc_auth.dll. $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
    }
}

# Define the file paths and URLs
$files = @(
    @{
        Name = 'Platform-4.2.0.20.jar'
	OName = 'Platform_48036_v4.2.0.20_27_Sep_2024.jar'
        Url = 'https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/Platform_48036_v4.2.0.20_27_Sep_2024.jar'
    },
    @{
        Name = 'MSJet-4.2.0.20.jar'
	OName = 'MSJet_48036_v4.2.0.20_27_Sep_2024.jar'
        Url = 'https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/MSJet_48036_v4.2.0.20_27_Sep_2024.jar'
    },
    @{
        Name = 'SourceCommons-4.2.0.20.jar'
	OName = 'SourceCommons_48036_v4.2.0.20_27_Sep_2024.jar'
        Url = 'https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/SourceCommons_48036_v4.2.0.20_27_Sep_2024.jar'
    }
)

# Check if the files exist in the specified directory
$filesExist = $files | ForEach-Object {
    Test-Path (Join-Path $striimLibPath $_.Name)
}

# If any of the files exist, prompt the user
if ($filesExist -contains $true) {
    $response = Read-Host -Prompt "The specified directory contains files that may require patching. Do you want to install the patches? (y/n)"

    if ($response -eq 'y') {
        # Download and replace the files
        $files | ForEach-Object {
            Invoke-WebRequest $_.Url -OutFile (Join-Path $striimLibPath $_.OName)
            Remove-Item (Join-Path $striimLibPath $_.Name.Replace('_48036', '')) -Force
        }

        # Download and replace MSSQLNative.dll
        Invoke-WebRequest 'https://github.com/daniel-striim/StriimQueryAutoLoader/raw/refs/heads/main/MSJet/FixesFor4.2.0.20/MSSQLNative.dll' -OutFile (Join-Path $striimLibPath 'MSSQLNative.dll')

        Write-Host "Patches installed successfully."
    } else {
        Write-Host "Patching skipped."
    }
} else {
    Write-Host "No files found that require patching."
}

# Check Software Requirements
Write-Host "[Softwre] Success: Checking for installed requirements..."

# Get the list of installed software once
$installedSoftwareList = Get-WmiObject -Class Win32_Product

Write-Host "[Softwre] Success: Checking for installed requirements...Software list gathered."

function CheckAndDownloadSoftware {
    param(
        [string]$softwareName,
        [string]$requiredVersion,
        [string]$downloadUrl
    )

    $matchingSoftware = $installedSoftwareList |
                        Where-Object {
                            $_.Name -like "*$softwareName*"
                        }

    if ($matchingSoftware) {
        foreach ($software in $matchingSoftware) {
            if ([version]$software.Version -ge [version]$requiredVersion) {
                Write-Host "[Softwre] Success: $($software.Name) version $($software.Version) found. Meets requirement."
            } else {
                Write-Host "[Softwre] Fail***: $($software.Name) version $($software.Version) found, but it's too old."
                DownloadAndInstallSoftware $softwareName $downloadUrl
            }
        }
    } else {
        Write-Host "[Softwre] Fail***: $softwareName not found."
        DownloadAndInstallSoftware $softwareName $downloadUrl
    }
}

# Function to download and provide instructions for software installation
function DownloadAndInstallSoftware {
    param(
        [string]$softwareName,
        [string]$downloadUrl
    )

    $downloadChoice = Read-Host "[Softwre]  Do you want to download and install $softwareName? (Y/N)"

    if ($downloadChoice.ToUpper() -eq "Y") {
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

		# Resolve the redirection link to get the actual download URL
		if ($downloadUrl -eq "https://aka.ms/vs/16/release/14.29.30133/VC_Redist.x64.exe") {
			$downloadPath = -join ($downloadDir, "\VC_Redist.x64.exe" )
		} else {
			$downloadPath = -join ($downloadDir, "\msoledbsql.msi" )
		}
		Write-Host "[Softwre] downloadPath $downloadPath"

        $downloadResult = Download-FileIfNotExists -Uri $downloadUrl -DownloadPath $downloadPath

		if ($downloadResult)
        {
            Write-Host "[Softwre] Success: $softwareName installer downloaded to $downloadPath. "
            # Prompt to run as administrator
            $runAsAdminChoice = Read-Host "Run the installer as administrator now? (Y/N)"
            if ($runAsAdminChoice.ToUpper() -eq "Y") {

                if (Test-IsAdmin) {
                    # Already running as admin
                    Start-Process -FilePath $downloadPath
                }
                else{
                    # Run as administrator using Invoke-AsAdmin
                    $command = "Start-Process -FilePath '$downloadPath'"
                    $exitCode = Invoke-AsAdmin -ArgumentList $command -WorkingDirectory $downloadDir

                    if ($exitCode -eq 0) {
                        Write-Host "[Softwre] Installer launched successfully (elevated)."
                    } else {
                        Write-Host "[Softwre] Error: Failed to launch installer (elevated). Exit code: $exitCode" -ForegroundColor Red
                    }
                }
            }
        }

    }
}

# Check for Microsoft Visual C++ 2015-2019 Redistributable
CheckAndDownloadSoftware "Microsoft Visual C++ 2019 X64 Minimum Runtime" "14.28.29914" "https://aka.ms/vs/16/release/14.29.30133/VC_Redist.x64.exe"

# Check for Microsoft OLE DB Driver for SQL Server
CheckAndDownloadSoftware "Microsoft OLE DB Driver for SQL Server" "18.2.3.0" "https://go.microsoft.com/fwlink/?linkid=2119554"

# Check for fixes for 4.2.0.20

# Check if Striim service is installed
if ($nodeType -eq "A") {
	$serviceName = "Striim Agent"
} else {
	$serviceName = "Striim"
}
if (Get-Service $serviceName -ErrorAction SilentlyContinue) {
        Write-Host "[Service] Success: Striim service is installed."
		Write-Host "[Service] * Note : If your Striim service is using Integrated Security, you may need to change the user the service runs as."
} else {
	$runAsService = Read-Host "[Service] Plan to run Striim as a service? (Y/N)"
	if ($runAsService.ToUpper() -eq "Y") {

		# Check for windowsService/windowsAgent folder
		if ($nodeType -eq "A") {
			$serviceConfigFolder = -join ($striimInstallPath, "\conf\windowsAgent")
		} else {
			$serviceConfigFolder = -join ($striimInstallPath, "\conf\windowsService")
		}

		Write-Host "[Service] Service path searched: $serviceConfigFolder"

		if (Test-Path $serviceConfigFolder) {
			# Check if the folder is empty
			if ((Get-ChildItem $serviceConfigFolder | Measure-Object).Count -eq 0) {
				# Delete the empty folder
				Remove-Item $serviceConfigFolder -Recurse -Force
				Write-Host "[Service] Success Empty $serviceConfigFolder deleted."
			} else {
				Write-Host "[Service] Success: Striim service configuration found in $serviceConfigFolder. It is not empty, so it will not be deleted."
			}
		}

		if (-not (Test-Path $serviceConfigFolder)) {
			# Find Striim version
			$platformJar = Get-ChildItem $striimInstallPath\lib -Filter "DatabaseReader-*.jar" | Select-Object -First 1
			if ($platformJar) {
				$versionMatch = $platformJar.Name -match "DatabaseReader-(.*)\.jar"
				if ($versionMatch) {
					$striimVersion = $matches[1]

					# Download and extract service/agent file
					$downloadUrl = if ($nodeType -eq "A") {
						"https://striim-downloads.striim.com/Releases/$striimVersion/Striim_windowsAgent_$striimVersion.zip"
					} else {
						"https://striim-downloads.striim.com/Releases/$striimVersion/Striim_windowsService_$striimVersion.zip"
					}

					[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

					$downloadPath = -join ($downloadDir, "\", $downloadUrl.Split("/")[-1])
					Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath

					# Extract to a temporary directory to avoid nested folders
					$tempExtractPath = Join-Path $env:TEMP ([Guid]::NewGuid().ToString())
					Expand-Archive -Path $downloadPath -DestinationPath $tempExtractPath

					# Move the contents of the extracted folder to the desired location
					$extractedContent = Get-ChildItem $tempExtractPath
					Move-Item $extractedContent.FullName -Destination $serviceConfigFolder -Force

					# Clean up temporary directory and downloaded ZIP
					Remove-Item $downloadPath -Force

				} else {
					Write-Host "[Service] Fail***: Could not determine Striim version from Platform jar file."
				}
			} else {
				Write-Host "[Service] Fail***: Could not find Platform jar file in $striimInstallPath\lib to determine Striim version."
			}
		} else {
			Write-Host "[Service] Fail***: Striim service configuration found in $serviceConfigFolder"
		}

		# Ask if user wants to set up the service
		$setupService = Read-Host "[Service]  Do you want to set up the Striim service now? (Y/N)"
		if ($setupService.ToUpper() -eq "Y") {
			# Execute setup script (assuming it's in the extracted folder)
			if ($nodeType -eq "A") {
				$setupScriptPath = Join-Path $serviceConfigFolder "setupWindowsAgent.ps1"
			} else {
				$setupScriptPath = Join-Path $serviceConfigFolder "setupWindowsService.ps1"
			}

			Write-Host "[Service] Runing the service setup located here: $setupScriptPath"

            # --- ELEVATED SCRIPT EXECUTION ---
            if (Test-IsAdmin) {
                # Already running as admin, execute directly
                Write-Host "[Service] Running service setup script (already elevated): $setupScriptPath"
                & $setupScriptPath  # Use the call operator (&) to execute the script
            } else {
                # Not running as admin, elevate
                Write-Host "[Service] Elevating permissions to run service setup script: $setupScriptPath"
                # Construct the command string.  Note the use of single quotes.
                $currentWorkingDirectory = Join-Path -Path (Get-Location).Path -ChildPath "conf\windowsAgent"
                $scriptExecutionCommand = "Set-Location -Path '$currentWorkingDirectory'; & '$setupScriptPath'"
                $exitCode = Invoke-AsAdmin -ArgumentList $scriptExecutionCommand -WorkingDirectory $currentWorkingDirectory

                if ($exitCode -eq 0) {
                    Write-Host "[Service] Service setup script completed successfully (elevated)."
                } else {
                    Write-Host "[Service] Error: Service setup script failed (exit code: $exitCode)." -ForegroundColor Red
                }
            }
		}

		Write-Host "[Service] * Note : If your Striim service is using Integrated Security, you may need to change the user the service runs as."
	}
}

# Construct core paths AFTER $striimInstallPath is confirmed
$striimLibPath = Join-Path -Path $striimInstallPath -ChildPath "lib"
$striimConfPath = Join-Path -Path $striimInstallPath -ChildPath "conf"
$striimBinPath = Join-Path -Path $striimInstallPath -ChildPath "bin"

# --- [Config JKS] JKS/Key File Generation ---
Write-Host "[Config JKS] Checking $configType Key Store (JKS) and Password files..."

# Use paths constructed earlier ($striimConfPath, $striimBinPath)
# Check base paths again just before use
if (!(Test-Path -Path $striimConfPath -PathType Container)) {
    Write-Error "[Config JKS] Error: Configuration directory '$striimConfPath' not found. Cannot check/manage JKS/PWD files."
} elseif (!(Test-Path -Path $striimBinPath -PathType Container)) {
    Write-Error "[Config JKS] Error: Bin directory '$striimBinPath' not found. Cannot run configuration scripts."
} else {
    # Proceed only if base paths are valid

    # Variables $jksFileName, $pwdFileName, $configScriptName, $configType should be set based on $nodeType earlier
    if ([string]::IsNullOrEmpty($jksFileName) -or [string]::IsNullOrEmpty($pwdFileName) -or [string]::IsNullOrEmpty($configScriptName) -or [string]::IsNullOrEmpty($configType)) {
         Write-Error "[Config JKS] Internal Error: Node type variables (jksFileName, etc.) not set correctly. NodeType was '$nodeType'."
    } else {
        $jksPath = Join-Path -Path $striimConfPath -ChildPath $jksFileName
        $pwdPath = Join-Path -Path $striimConfPath -ChildPath $pwdFileName
        $configScriptPath = Join-Path -Path $striimBinPath -ChildPath $configScriptName

        Write-Host "[Config JKS]  JKS File Path: '$jksPath'"
        Write-Host "[Config JKS]  PWD File Path: '$pwdPath'"
        Write-Host "[Config JKS]  Config Script: '$configScriptPath'"

        # Check File Existence
        $filesExist = (Test-Path -Path $jksPath -PathType Leaf) -and (Test-Path -Path $pwdPath -PathType Leaf)

        # Prompt User
        $runConfigScript = $false
        $deleteFilesFirst = $false

        if ($filesExist) {
            Write-Host "[Config JKS] Configuration files '$jksFileName' and '$pwdFileName' already exist." -ForegroundColor Yellow
            $response = ""
            while ($response -notin ('y','n')) {
                $response = Read-Host "[Config JKS] Do you want to re-run '$configScriptName' AS ADMINISTRATOR? This will DELETE the existing files first. (y/n)"
                $response = $response.ToLower()
            }
            if ($response -eq 'y') {
                $runConfigScript = $true
                $deleteFilesFirst = $true
            }
        } else {
            Write-Host "[Config JKS] Configuration files '$jksFileName' and/or '$pwdFileName' do not exist." -ForegroundColor Yellow
            $response = ""
            while ($response -notin ('y','n')) {
                $response = Read-Host "[Config JKS] Do you want to run '$configScriptName' AS ADMINISTRATOR now to create them? (y/n)"
                $response = $response.ToLower()
            }
             if ($response -eq 'y') {
                $runConfigScript = $true
            }
        }

        # Execute Actions
        if ($runConfigScript) {
            # Deletion (Attempt direct first, then elevated if needed via Invoke-AsAdmin)
            if ($deleteFilesFirst) {
                Write-Host "[Config JKS] Attempting to delete existing configuration files..."
                $deletionSuccess = $true
                try {
                    # Try direct deletion first
                    if (Test-Path -Path $jksPath -PathType Leaf) { Remove-Item -Path $jksPath -Force -ErrorAction Stop }
                    if (Test-Path -Path $pwdPath -PathType Leaf) { Remove-Item -Path $pwdPath -Force -ErrorAction Stop }
                    Write-Host "[Config JKS]  Direct deletion attempt finished."
                    # Verify deletion
                    if ((Test-Path -Path $jksPath -PathType Leaf) -or (Test-Path -Path $pwdPath -PathType Leaf)) {
                         throw "Direct deletion failed or files still exist."
                    } else {
                         Write-Host "[Config JKS]  Successfully deleted files directly." -ForegroundColor Green
                    }
                } catch {
                    Write-Warning "[Config JKS] Direct deletion failed: $($_.Exception.Message). Attempting with elevation via Invoke-AsAdmin..."
                    # Construct command for elevated deletion
                    # Escape paths for PowerShell command string
                    $escapedJksPath = $jksPath -replace "'", "''"
                    $escapedPwdPath = $pwdPath -replace "'", "''"
                    $deleteCommand = "if(Test-Path -Path '$escapedJksPath'){ Remove-Item -Path '$escapedJksPath' -Force -ErrorAction SilentlyContinue }; if(Test-Path -Path '$escapedPwdPath'){ Remove-Item -Path '$escapedPwdPath' -Force -ErrorAction SilentlyContinue }; exit 0" # Exit 0 even if files didn't exist

                    $exitCode = Invoke-AsAdmin -ArgumentList $deleteCommand -WorkingDirectory $striimConfPath
                    # Verify deletion after elevation attempt
                    if ((Test-Path -Path $jksPath -PathType Leaf) -or (Test-Path -Path $pwdPath -PathType Leaf)) {
                        Write-Error "[Config JKS] Failed to delete files even with elevation (Exit Code: $exitCode). Check permissions on '$striimConfPath'."
                        $deletionSuccess = $false
                    } else {
                        Write-Host "[Config JKS]  Successfully deleted files (elevated)." -ForegroundColor Green
                    }
                }
                if (-not $deletionSuccess) {
                    Write-Warning "[Config JKS] Failed to delete necessary files. Stopping execution of '$configScriptName'."
                    $runConfigScript = $false # Prevent script run if deletion failed
                }
            } # End deletion block

            # Run Config Script (Only if $runConfigScript is still true)
            if ($runConfigScript) {
                if (!(Test-Path -Path $configScriptPath -PathType Leaf)) {
                     Write-Error "[Config JKS] Configuration script '$configScriptPath' not found! Cannot execute."
                } else {
                    Write-Host "[Config JKS] Preparing to run '$configScriptName' as Administrator..." -ForegroundColor Cyan
                    Write-Host "[Config JKS]   Script Path: $configScriptPath" -ForegroundColor Cyan
                    $scriptDir = Split-Path -Path $configScriptPath -Parent
                    Write-Host "[Config JKS]   Target Directory for elevated process: $scriptDir" -ForegroundColor Cyan

                    # Construct the PowerShell command string to be executed elevated via Invoke-AsAdmin.
                    # Use a HEREDOC string (@"..."@) for clarity.
                    # 1. Explicitly Set-Location IN the elevated process.
                    # 2. Execute the batch file using cmd /c for better error handling/exit codes.
                    # 3. Capture and exit with the batch file's exit code.
                    # Escape literal quotes (\") needed for the cmd /c command.
                    # Escape single quotes ('') inside paths if they could occur, for robustness within the outer single quotes of Set-Location.
                    $commandToRunElevated = @"
# Ensure paths with potential single quotes are handled
`$scriptDirForElevated = '$($scriptDir -replace "'", "''")'
`$configScriptPathForElevated = '$($configScriptPath -replace "'", "''")'

Write-Host "Elevated Shell: Setting location to `"`$scriptDirForElevated`"..."
Set-Location -Path "`$scriptDirForElevated" -ErrorAction Stop
Write-Host "Elevated Shell: Current Directory = `$(Get-Location)"
Write-Host "Elevated Shell: Executing '`$configScriptPathForElevated' via cmd /c..."

# Execute the batch file using cmd /c (quotes around path handle spaces)
cmd /c "\`"`$configScriptPathForElevated\`""

# Capture the exit code from cmd.exe/batch file
`$exitCode = `$LASTEXITCODE
Write-Host "Elevated Shell: cmd /c finished with Exit Code: `$exitCode"

# Exit the elevated PowerShell script with the captured exit code
exit `$exitCode
"@
                    # For debugging: Print the command that will be run elevated
                    # Write-Host "[Config JKS] Command to run elevated:"
                    # Write-Host $commandToRunElevated -ForegroundColor DarkGray

                    try
                    {
                        Write-Host "[Config JKS] Attempting elevation using Invoke-AsAdmin..."
                        # Use the existing Invoke-AsAdmin function
                        # Pass the constructed PowerShell command block as the ArgumentList
                        # WorkingDirectory here sets the initial CWD for the elevated powershell.exe process
                        $processExitCode = Invoke-AsAdmin -ArgumentList $commandToRunElevated -WorkingDirectory $scriptDir

                        Write-Host "[Config JKS] Elevated process finished with Exit Code: $processExitCode" -ForegroundColor Cyan

                        # --- Verification --- (Runs after elevation attempt returns)
                        Write-Host "[Config JKS] Verifying output file existence..."
                        if (Test-Path -Path $jksPath -PathType Leaf -ErrorAction SilentlyContinue) {
                            Write-Host "[Config JKS] Verified: '$jksFileName' exists after script execution." -ForegroundColor Green
                        } else {
                            Write-Warning "[Config JKS] Verification Failed: '$jksFileName' does NOT exist after script execution. Elevated script likely failed internally (Exit Code: $processExitCode)."
                        }
                         if (Test-Path -Path $pwdPath -PathType Leaf -ErrorAction SilentlyContinue) {
                            Write-Host "[Config JKS] Verified: '$pwdFileName' exists after script execution." -ForegroundColor Green
                        } else {
                            Write-Warning "[Config JKS] Verification Failed: '$pwdFileName' does NOT exist after script execution. Elevated script likely failed internally (Exit Code: $processExitCode)."
                        }
                         # --- End Verification ---

                        if ($processExitCode -ne 0) {
                            Write-Warning "[Config JKS] Elevated process for '$configScriptName' finished with a non-zero exit code ($processExitCode)."
                            Write-Warning "[Config JKS] This likely indicates errors within the batch script '$configScriptName'. Review logs or output if the script creates any."
                        } else {
                            Write-Host "[Config JKS] Elevated process for '$configScriptName' completed successfully (Exit Code 0)." -ForegroundColor Green
                        }
                    }
                    catch # Catches errors from Invoke-AsAdmin itself (e.g., UAC denial) or if ErrorAction Stop is hit inside $commandToRunElevated
                    {
                        Write-Error "[Config JKS] Failed to execute Invoke-AsAdmin or error within elevated script: $($_.Exception.Message)"
                        Write-Error "[Config JKS] Full Error Record: $($_.ToString())"
                    }
                } # End script execution block
            }# End if($runConfigScript) after deletion check
        } else { # if not $runConfigScript initially
             Write-Host "[Config JKS] Configuration script '$configScriptName' will not be run based on user input." -ForegroundColor Yellow
        }
    } # End check for valid node type variables
} # End check for valid base paths

Write-Host "[Config JKS] Finished JKS/Key File check."
# --- END JKS/Key File Generation ---

# --- Final Summary ---
Write-Host "`n------------------------------------------"
Write-Host "Striim Configuration Check Finished."
Write-Host "------------------------------------------"
Write-Host "Please review the output above for any warnings or errors."
Write-Host "Some changes (like System PATH or Java installation) may require restarting your terminal or system."
if ($Host.Name -eq "ConsoleHost") {
    Read-Host "Press Enter to exit the script."
}
