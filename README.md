# Striim MSJet Environment Configuration Checker

This PowerShell script automates the validation, configuration, and setup of a Striim environment on Windows. It's designed to be a comprehensive tool to ensure a Striim Node or Agent is ready for operation, whether on a machine with or without internet access.

The script intelligently detects the environment, validates configurations, checks for dependencies, installs prerequisites, applies patches, and assists with service and security setup, prompting for administrative rights only when necessary.

## Prerequisites

- **PowerShell:** The script requires PowerShell to run. It is recommended to use PowerShell 5.1 or later.
- **Striim Installation:** The script must be placed in the root directory of an existing Striim Node or Striim Agent installation (e.g., `C:\striim` or `C:\striim\Agent`).

## Usage

### Getting the Script

Run the following commands from PowerShell within your Striim installation directory to download and unblock the script:

```
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/daniel-striim/msjet-assist/main/msjetchecker.ps1" -OutFile "msjetchecker.ps1"
Unblock-File -Path "msjetchecker.ps1"
```

### Standard Interactive Mode

Execute the script from a standard (non-elevated) PowerShell prompt. The script will guide you through all checks and will request administrative elevation automatically for specific actions.

```
.\msjetchecker.ps1
```

### Offline Installation & Download-Only Mode

The script fully supports offline environments. To prepare for this, you must first run the script on a machine with internet access to download all required files.

**1. Download All Files (Online Machine):**

On a machine with an internet connection, run the script with the `--downloadonly` and `--version` flags. This will create a `downloads` subfolder and populate it with all installers, libraries, and patches needed for the specified Striim version.

```
# Example: Download all files needed for Striim v5.0.6
.\msjetchecker.ps1 --version 5.0.6 --downloadonly

# Example: Download all files needed for Striim v4.2.0.20
.\msjetchecker.ps1 --version 4.2.0.20 --downloadonly
```

**2. Transfer to Offline Machine:**

Copy the entire directory, including the script (`msjetchecker.ps1`) and the newly populated `downloads` folder, to the offline target machine's Striim installation directory.

**3. Run on Offline Machine:**

Execute the script on the offline machine as you would in standard mode. It will automatically detect and use the files in the `downloads` folder instead of attempting to connect to the internet.

```
.\msjetchecker.ps1
```

## Features

The script performs the following actions:

- **Environment Validation:**
    - Verifies sufficient disk space on the C: drive.
    - Auto-detects the node type (Agent or Node) based on `agent.conf` or `startUp.properties`.
- **Configuration File Management:**
    - Interactively validates and helps you set required values in `agent.conf` or `startUp.properties`.
- **System PATH:**
    - Checks if the Striim `lib` directory is in the system PATH and offers to add it (requires elevation).
- **Dependency & Patch Management:**
    - Verifies required DLLs (`icudt72.dll`, etc.) are present in the `lib` folder.
    - Checks for and offers to apply version-specific patches (e.g., for version 4.2.0.20).
- **Java Version Check:**
    - Detects the installed Java version and verifies it's correct for your Striim version (Java 8 for Striim v4, Java 11 for Striim v5+).
- **Software Prerequisites:**
    - Checks for necessary software like **Microsoft Visual C++ Redistributable** and the **Microsoft OLE DB Driver for SQL Server**.
- **Security & Keystores:**
    - Assists with setting up **Integrated Security** by placing `sqljdbc_auth.dll` in `System32` (requires elevation).
    - Prompts to generate or re-create `sks.jks`/`aks.jks` keystore files by running the appropriate configuration script (requires elevation).
- **Windows Service Management:**
    - Checks if the Striim service is installed and guides you through the setup process if it's missing (requires elevation).
- **Agent-Specific Features:**
    - **Firewall Validation:** Checks for required inbound and outbound firewall rules for Hazelcast, ZeroMQ, and the Striim server connection (requires elevation).
    - **JDBC Driver Installation:** Provides an interactive menu to help install common JDBC drivers for sources like Oracle, MySQL, PostgreSQL, and more.

## Troubleshooting

- **Permissions Errors:** Do not run the entire script as an administrator. It is designed to request elevation for specific tasks. If a UAC prompt appears, you must accept it for that action to succeed.
- **Download Failures:** In online mode, ensure the machine has a stable internet connection and can reach GitHub, Microsoft, and Oracle download sites.
- **Offline Mode Issues:** Double-check that the `downloads` folder and its contents were copied correctly to the offline machine's directory alongside the script.
- **Script Execution Policy:** If you receive an error about scripts being disabled, you may need to adjust your execution policy. Run `Get-ExecutionPolicy` to check. If it's `Restricted`, you can change it for the current process by running `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process`.

## Sample Output

```
PS C:\striim\Agent> .\msjetchecker.ps1

Starting Striim Configuration Check...
------------------------------------------
[Checks Summary] The following validations will occur:
  • [System]        Verify ≥15GB free space on C: drive
  • [Node Type]     Auto-detect Agent/Node or prompt for selection
  • [Config Files]  Validate agent.conf/startUp.properties settings
  • [Firewall]      AGENT ONLY: Check required inbound/outbound firewall rules (requires Admin) ➤
  • [System PATH]   Add Striim lib directory to system PATH (requires Admin) ➤
  • [Dependencies]  Verify required DLLs and offer to download them
  • [Java Check]    Ensure compatible Java version is installed and offer download if needed
  • [Security]      Setup Integrated Security sqljdbc_auth.dll (requires Admin) ➤
  • [Prereqs]       Verify required software (VC++ Redist, OLE DB Driver)
  • [Patches]       Apply critical fixes for Striim v4.2.0.20 (if applicable)
  • [Service]       Configure Striim as a Windows Service (requires Admin) ➤
  • [JKS Config]    Prompt to generate/recreate JKS and key files (requires Admin) ➤
  • [Drivers]       AGENT ONLY: Prompt to install common JDBC drivers
------------------------------------------

Press Enter to continue or Ctrl+C to abort...:

[Envrnmt]       Sufficient disk space available on C: drive.
[Envrnmt]       -> AGENT Environment detected.
[Envrnmt] Success: Striim Install Path set to: C:\striim\Agent
[Striim  ] Success: Found Striim version 5.0.6 (Major version 5)
[Config ]       -> AGENT -> Specific Tests for configuration:
[Config ] Success: 'striim.cluster.clusterName' found with value: MyCluster
[Config ] Current value is 'MyCluster'. Update it? (y/n, default is no): n
[Config ] 'striim.cluster.clusterName' remains unchanged.
[Config ] Success: 'striim.node.servernode.address' found with value: 192.168.1.100
[Config ] Current value is '192.168.1.100'. Update it? (y/n, default is no): n
[Config ] 'striim.node.servernode.address' remains unchanged.
[Firewall] Checking firewall rules... (Requires Administrator privileges)
[Firewall] This check requires elevation. Please re-run the script as an Administrator to check firewall rules.
[Config ]       -> Striim Lib Path set to: C:\striim\Agent\lib
[Config ] Fail***: Striim lib directory not found in PATH.
[Config ]  Add it to the system PATH (requires Admin)? (Y/N): y
[Config ]  Elevating permissions to update PATH...
[Config ] Success: Striim lib directory added to PATH. Please restart your terminal to see the change.
[DLLs   ] Success: icudt72.dll found.
[DLLs   ] Success: icuuc72.dll found.
[DLLs   ] Fail***: MSSQLNative.dll not found.
[DLLs   ]  Download it now? (Y/N): y
[Download] Downloading from https://.../MSSQLNative.dll to C:\striim\Agent\downloads\MSSQLNative.dll...
[Download] Success: File downloaded using HttpClient.
[DLLs   ] Success: MSSQLNative.dll copied to C:\striim\Agent\lib
[Java   ] Success: Found installed Java version: "11.0.21"
[Java   ] Success: Required Java version 11 is installed.
[Softwre] Checking for required software...
[Softwre] Success: Found compatible Microsoft Visual C++ 2015-2022 Redistributable (x64) version 14.38.33130.0.
[Softwre] Fail***: A required version of Microsoft OLE DB Driver for SQL Server was not found.
[Softwre]  Download and install 'Microsoft OLE DB Driver for SQL Server' now? (Y/N): y
...
[Service] Striim service 'Striim Agent' is not installed.
[Service]  Do you want to set up the service now (requires Admin)? (Y/N): y
...
[JKS Cfg] Checking AKS Key Store (JKS) and Password files...
[JKS Cfg] JKS/PWD files do not exist. Create them now (requires Admin)? (y/n): y
...
[Drivers] Would you like to check for and install any JDBC drivers for the agent? (y/n): y

--- JDBC Driver Installation Menu ---
  1. HP NonStop (Manual Path)
  2. MariaDB (v2.4.3)
  3. MySQL / MemSQL (Connector/J 8.0.30)
  4. Oracle Instant Client (for OJet)
  5. PostgreSQL (v42.2.27)
  6. Teradata (Manual Path)
  7. Vertica (Manual Path)
  0. Exit Menu
Select a driver to install (or 0 to exit):
```