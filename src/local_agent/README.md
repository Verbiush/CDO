# CDO Local Agent

This agent connects your local PC to the CDO Web Cloud, enabling native features like:
- File System Browsing (for selecting upload/download folders)
- Local Browser Automation (Selenium)

## Installation

1. Unzip the installer package.
2. Run `install_agent.ps1` with PowerShell (Administrator).
3. A shortcut "CDO_Agent" will be created on your Desktop.

## Configuration

1. Double-click the "CDO_Agent" shortcut.
2. Enter the Server URL (e.g., `https://cdo-aws.com` or `http://localhost:8000`).
3. Enter your CDO Username and Password.
4. The agent will verify the connection and save your configuration.

## Usage

- Keep the agent window open (or minimized) while using the Web App.
- When you perform actions in the Web App that require local access (e.g., "Browse Folder"), the agent will execute them on your PC.
