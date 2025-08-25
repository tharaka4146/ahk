# AutoHotkey Script for 60% Keyboards

This AutoHotkey script provides keyboard shortcuts and automation for 60% keyboards, with a configurable system for easy portability across different PCs.

## Setup Instructions

### For New PC Setup:

1. **Copy the script files** to your new PC:
   - `ahk script.ahk` (main AutoHotkey script)
   - `config_home.json` (template configuration file)
   - `sensitive_home.json` (template for sensitive data)
   - `.gitignore` (Git ignore file for security)

2. **Create your configuration files**:
   - Copy `config_home.json` to `config_office.json`
   - Copy `sensitive_home.json` to `sensitive_office.json`
   - Edit `config_office.json` with your PC-specific settings:
     - Replace `YOUR_USERNAME` with your actual Windows username
     - Update file paths to match your system
   - Edit `sensitive_office.json` with your sensitive data:
     - Add your email, passwords, access tokens
     - Update personal URLs and development commands
     - **NEVER commit this file to Git!**

3. **Run the script**:
   - Double-click `ahk script.ahk` to start
   - The script will automatically load settings from both configuration files

### Security Features:

- **Sensitive data isolation**: Passwords, tokens, and personal URLs are stored separately in `sensitive_office.json`
- **Git protection**: The `.gitignore` file prevents sensitive data from being committed to version control
- **Safe sharing**: You can safely share `config_home.json` and `sensitive_home.json` as they contain no real sensitive data

### Configuration Options:

The configuration is split between two files for security:

**`config_office.json`** (safe to commit to Git):
- **User Settings**: Username and common folder paths (Desktop, Downloads, Screenshots)
- **Applications**: Default browser and application paths
- **Browser Profiles**: Chrome profile directories for different shortcuts
- **Public URLs**: Non-sensitive website links

**`sensitive_office.json`** (NEVER commit to Git):
- **Credentials**: Email addresses, API tokens, passwords
- **Access Codes**: Authentication codes and keys
- **Personal URLs**: Private Jira boards, Notion workspaces, etc.
- **Development**: Project-specific commands and paths

### Example Configuration:

**`config_office.json`** example:
```json
{
    "user": {
        "username": "JohnDoe",
        "downloads_folder": "C:\\Users\\JohnDoe\\Downloads",
        "desktop_folder": "C:\\Users\\JohnDoe\\Desktop"
    },
    "applications": {
        "default_browser": "chrome.exe",
        "vpn_client_path": "C:\\Program Files\\Your VPN\\vpn.exe"
    },
    "clipboard_shortcuts": {
        "email": "[loaded from sensitive_office.json]",
        "access_code": "[loaded from sensitive_office.json]"
    }
}
```

**`sensitive_office.json`** example:
```json
{
    "credentials": {
        "email": "john.doe@company.com",
        "git_token": "ghp_xxxxxxxxxxxxxxxxxxxx"
    },
    "access_codes": {
        "primary_access_code": "ABCDE-FGHIJ-KLMNO-PQRST"
    },
    "additional_passwords": {
        "vpn_password": "your_secure_password",
        "database_password": "another_secure_password"
    }
}
```

### Important Notes:

- Always use double backslashes (`\\`) in Windows file paths within JSON files
- Both `config_office.json` and `sensitive_office.json` must be in the same directory as the script
- **NEVER commit `sensitive_office.json` to Git** - it contains passwords and private data
- If either configuration file is missing or has errors, the script will show an error message and exit
- Always backup your `sensitive_office.json` file safely before making changes
- When sharing your script setup, only share the template files, never the actual config files

## Keyboard Shortcuts

The script provides various keyboard shortcuts for:
- Media controls (play/pause, next/previous, volume)
- Quick app launching (browser, notepad, VPN)
- Folder shortcuts (Desktop, Downloads, Screenshots)
- Website shortcuts (Outlook, YouTube, etc.)
- Clipboard shortcuts for frequently used text
- Window management and desktop switching

Refer to the script comments for specific key combinations.