# Striim Query AutoLoader Script

This PowerShell script automates the configuration and verification of a Striim environment, ensuring proper setup and necessary dependencies are in place.  It checks for crucial files and configurations, downloads missing DLLs and patches, and guides the user through essential setup steps.

## Prerequisites

*   **PowerShell Version:** This script is designed to work with PowerShell, run as Administrator. 
*   **Striim Installation:** Place this Powershell Script in a valid Striim or Striim Agent base installation directory.

## Usage

2.  **Run the script:** Execute the script from the PowerShell (as Striim created user, not as Administrator):
    ```powershell
    REM Navigate to your Striim base install directory
    cd C:\striim\Agent
    
    REM Run the following to auto-download and install msjetchcker:
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/daniel-striim/msjet-assist/refs/heads/main/msjetchecker.ps1" -OutFile "msjetchecker.ps1"
    Unblock-File -Path "msjetchecker.ps1"
    .\msjetchecker.ps1
    ```

## Script Logic

The script performs the following actions:

1.  **Environment Detection:**
    *   Determines the Striim node type (Agent or Node). If neither is found, user is prompted to provide the Striim install path.
2.  **Agent-Specific Checks (if Agent):**
    *   Verifies the existence and content of `agent.conf`, ensuring `striim.cluster.clusterName` and `striim.node.servernode.address` are defined.
3.  **Node-Specific Checks (if Node):**
    *   Verifies the existence and content of `startUp.properties`, ensuring `CompanyName`, `LicenceKey`, `ProductKey`, and `WAClusterName` are defined.
4.  **Striim Library Path in PATH:**
    *   Checks if the Striim library path is included in the system PATH environment variable.  If not, prompts the user to add it and automatically adds the path if the user confirms.
5.  **DLL Verification and Download:**
    *   Checks for required DLLs (`icudt72.dll`, `icuuc72.dll`, `MSSQLNative.dll`).
    *   Offers to download missing DLLs.
6.  **Java Version Check:**
    *   Detects the installed Java version.
    *   Offers to download Java 8 if Java is missing or an incompatible version is found.
7.  **Integrated Security Setup:**
    *   Checks for the `sqljdbc_auth.dll` file.
    *   Guides the user through setting up Integrated Security if needed.
8.  **Versioned Patches:**
    *   Checks for and optionally downloads versioned patches.
9.  **Striim Service Setup:**
     *  Check Striim Service, if it doesn't exits, then guide the user through the installation process.

## Output

The script provides detailed output to the console, indicating the status of each check and any actions taken.  Success messages are prefixed with `[Config ] Success:`, while failures are indicated with `[Config ] Fail***:`.

## Configuration

*   **Striim Install Path:**  The script automatically detects the Striim install path based on the presence of configuration files.
*   **Download Directory:** The download directory is automatically created within the Striim install directory.

## Troubleshooting

*   **Permissions Errors:** Ensure you are running the script with administrative privileges.
*   **DLL Download Failures:**  Check your internet connection and try again.  Consider using a proxy if necessary.
*   **Java Version Compatibility:**  Ensure that the installed Java version is compatible with Striim.
*   **Check DLL Path:** Manually check the libraries are updated and accessible.

## Notes

*   The script is designed to be as automated as possible, but may require user interaction for certain steps.
*   The script attempts to handle common configuration issues, but may not cover all possible scenarios.
*   Review the script's output carefully to ensure that all checks pass and any necessary actions are taken.

## Sample Output

```
PS C:\striim\Agent> .\msjetchecker.ps1

Starting Striim Configuration Check...
------------------------------------------
[Checks Summary] The following validations will occur:
  • [System]        Verify ≥15GB free space on C: drive
  • [Node Type]     Auto-detect Agent/Node or prompt for selection
  • [Striim Setup]  Install Striim if missing
  • [Config Files]  Validate agent.conf/startUp.properties settings
  • [System PATH]   Add Striim lib directory to system PATH (requires Admin) ➤
  • [Dependencies]  Verify DLLs (icudt72.dll, icuuc72.dll, MSSQLNative.dll) or download them
  • [Java Check]    Ensure compatible Java version installed. Offer download if missing
  • [Security]      Setup Integrated Security sqljdbc_auth.dll (requires Admin) ➤
  • [Patches]       Apply critical fixes for Striim v4.2.0.20 (if applicable)
  • [Service]       Configure Striim as Windows Service (requires Admin) ➤
------------------------------------------

Press Enter to continue or Ctrl+C to abort...:

[Envrnmt]       Sufficient disk space available on C: drive.
[Envrnmt]       -> AGENT Environment based on agent.conf
[Envrnmt] Success: Striim Install Path set to: C:\striim\Agent
[Config ]       -> AGENT -> Specific Tests for configuration:
[Config ] Success: 'striim.cluster.clusterName' found in agent.conf with value: ASCENSION_ahgcplapp0071_dev
[Config ] Current value: 'striim.cluster.clusterName' is 'ASCENSION_ahgcplapp0071_dev'. Do you want to update it? (y/n, default is no)

[Config ] 'striim.cluster.clusterName' remains unchanged.
[Config ] Success: 'striim.node.servernode.address' found in agent.conf with value: 10.246.246.59
[Config ] Current value: 'striim.node.servernode.address' is '10.246.246.59'. Do you want to update it? (y/n, default is no)

[Config ] 'striim.node.servernode.address' remains unchanged.
[Config ] Success: 'MEM_MAX' found in agent.conf with value: 2048m
[Config ] Current value: 'MEM_MAX' is '2048m'. Do you want to update it? (y/n, default is no)
[Config ] 'MEM_MAX' remains unchanged.
[Config ] Saving changes to agent.conf
[Config ]       -> Striim Lib Path set to: C:\striim\Agent\lib
[Config ] Success: Striim lib directory found in PATH.
[DLLs   ] Success: icudt72.dll found in C:\striim\Agent\lib
[DLLs   ] Success: icuuc72.dll found in C:\striim\Agent\lib
[DLLs   ] Success: MSSQLNative.dll found in C:\striim\Agent\lib
[Striim  ] Success: Found Striim version: Platform-5.0.6.jar
[Striim  ] Major version: Platform-5
[Java   ] Success: Java version: java version "1.8.0_202"
[Java   ] Success: Java 8 found.
[Java   ] Java 8 detected, but Java 11 is required for Striim version Platform-5. Please install Java 11.
  Download Java 11? (Y/N): Y
[Download] File already exists: C:\striim\Agent\downloads\microsoft-jdk-11.0.26-windows-x64.msi
Java installer downloaded to C:\striim\Agent\downloads\microsoft-jdk-11.0.26-windows-x64.msi.  Run installer now (as administrator)? (Y/N): Y
Installer launched successfully.
[Int Sec] Success: Integrated Security: sqljdbc_auth.dll found in C:\Windows\System32
No files found that require patching.
[Softwre] Success: Checking for installed requirements...
[Softwre] Success: Checking for installed requirements...Software list gathered.
[Softwre] Fail***: Microsoft Visual C++ 2019 X64 Minimum Runtime not found.
[Softwre]  Do you want to download and install  (Y/N): Y
[Softwre] downloadPath C:\striim\Agent\downloads\VC_Redist.x64.exe
[Download] Downloading from https://aka.ms/vs/16/release/14.29.30133/VC_Redist.x64.exe to C:\striim\Agent\downloads\VC_Redist.x64.exe...
[Download] Success: File downloaded using HttpClient.
[Softwre] Success: Microsoft Visual C++ 2019 X64 Minimum Runtime installer downloaded to C:\striim\Agent\downloads\VC_Redist.x64.exe.
Run the installer as administrator now? (Y/N): Y
[Softwre] Installer launched successfully (elevated).
[Softwre] Success: Microsoft OLE DB Driver for SQL Server version 18.7.4.0 found. Meets requirement.
[Service] Success: Striim service is installed.
[Service] * Note : If your Striim service is using Integrated Security, you may need to change the user the service runs as.
```
