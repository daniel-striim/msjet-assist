# Striim Query AutoLoader Script

This PowerShell script automates the configuration and verification of a Striim environment, ensuring proper setup and necessary dependencies are in place.  It checks for crucial files and configurations, downloads missing DLLs and patches, and guides the user through essential setup steps.

## Prerequisites

*   **PowerShell Version:** This script is designed to work with PowerShell, run as Administrator. 
*   **Striim Installation:** Place this Powershell Script in a valid Striim or Striim Agent base installation directory.

## Usage

2.  **Run the script:** Execute the script from the PowerShell command line As Administrator:
    ```powershell
    REM Navigate to your Striim base install directory
    cd C:\striim\agent
    
    REM Run the following to auto-download and install msjetchcker:
    Invote-WebRequest -Uri "https://raw.githubusercontent.com/daniel-striim/msjet-assist/refs/heads/main/msjetchecker.ps1" -OutFile "msjetchecker.ps1"
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