# CDO Local Agent

This agent enables the Web Application (running in AWS/Docker) to interact with your local computer securely.

## Features
- **File System Access**: Allows the Web App to browse local drives and folders to select download/upload locations.
- **Local Browser Automation**: Allows the Web App to launch a Selenium browser on your machine to perform tasks (e.g., OVIDA downloads) that require local credentials or access.
- **Secure Connection**: Uses a self-signed SSL certificate to allow secure communication with the Web App (`https://localhost:12345`).

## Installation

1. Open PowerShell as **Administrator**.
2. Navigate to the project root directory.
3. Run the installation script:
   ```powershell
   Set-ExecutionPolicy Bypass -Scope Process -Force; ./src/local_agent/install_agent.ps1
   ```
4. The script will:
   - Install required Python dependencies.
   - Generate a self-signed SSL certificate.
   - Trust the certificate in your system's Root Store (to prevent browser warnings).
   - Create a System Tray application ("CDO Agent").
   - Add the agent to your Windows Startup folder.

## Usage

- The agent runs in the background. You will see a blue/white icon in your System Tray.
- **Right-click** the icon to:
  - Open **Dashboard** (API Docs).
  - **Exit** the agent.
- When using the CDO Web App, enable "Local Mode" (if available) to use these features.
