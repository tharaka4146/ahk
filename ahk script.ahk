; #NoTrayIcon
#Persistent ; Keeps the script running
#SingleInstance force
DetectHiddenWindows, on

; ============== Configuration Loading ==============
; Email counter functionality
emailCounterFile := A_ScriptDir . "\email_counter.txt"
emailCounter := LoadEmailCounter()

; Detect current username and load appropriate configuration files
EnvGet, CurrentUsername, USERNAME

if (CurrentUsername = "UBANDT1") {
    ; Office configuration for UBANDT1
    configFile := A_ScriptDir . "\config_office.json"
    sensitiveFile := A_ScriptDir . "\sensitive_office.json"
} else if (CurrentUsername = "HanSolo") {
    ; Home configuration for HanSolo
    configFile := A_ScriptDir . "\config_home.json"
    sensitiveFile := A_ScriptDir . "\sensitive_home.json"
} else {
    ; Unknown username - show error and exit
    MsgBox, Error: Unknown username '%CurrentUsername%'`n`nThis script is configured for:`n- UBANDT1 (office configuration)`n- HanSolo (home configuration)`n`nPlease update the script to support your username.
        ExitApp
}
; Debug output to show which configuration is being loaded
OutputDebug, Loading configuration for user: %CurrentUsername%
    OutputDebug, Config file: %configFile%
OutputDebug, Sensitive file: %sensitiveFile%

config := LoadConfig(configFile)
sensitive := LoadSensitiveData(sensitiveFile)

; Extract configuration values
defaultBrowser := config.applications.default_browser
userDownloadsFolder := config.user.downloads_folder
userDesktopFolder := config.user.desktop_folder
userScreenshotsFolder := config.user.screenshots_folder
vpnClientPath := config.applications.vpn_client_path
chromeDefaultProfile := config.browser_profiles.chrome_default_profile
chromeMusicProfile := config.browser_profiles.chrome_music_profile

; Clipboard shortcuts (loaded from sensitive data)
clipboardEmail := sensitive.credentials.email
clipboardAccessCode := sensitive.access_codes.primary_access_code
clipboardGitToken := sensitive.credentials.git_token
clipboardCursorCommand := sensitive.development.cursor_command

; URLs (mix of public and sensitive)
outlookUrl := config.urls.outlook
jiraUrl := sensitive.personal_urls.jira_board
youtubeUrl := config.urls.youtube
youtubeMusicUrl := config.urls.youtube_music
copilotUrl := config.urls.copilot
chatgptUrl := config.urls.chatgpt
notionUrl := sensitive.personal_urls.notion_workspace
sltUsageUrl := config.urls.slt_usage

; ============== Configuration Functions ==============
LoadConfig(filePath) {
    FileRead, jsonContent, %filePath%
    if ErrorLevel {
        MsgBox, Error: Could not read config file: %filePath%
        ExitApp
    }
    return ParseJSON(jsonContent)
}

LoadSensitiveData(filePath) {
    FileRead, jsonContent, %filePath%
    if ErrorLevel {
        MsgBox, Error: Could not read sensitive data file: %filePath%`n`nPlease ensure you have created this file from the template.
        ExitApp
    }
    return ParseSensitiveJSON(jsonContent)
}

ParseJSON(jsonStr) {
    ; Simple JSON parser for our specific config structure
    config := {}

    ; Parse user section
    config.user := {}
    config.user.username := ExtractJSONValue(jsonStr, "username")
    config.user.downloads_folder := ExtractJSONValue(jsonStr, "downloads_folder")
    config.user.desktop_folder := ExtractJSONValue(jsonStr, "desktop_folder")
    config.user.screenshots_folder := ExtractJSONValue(jsonStr, "screenshots_folder")

    ; Parse applications section
    config.applications := {}
    config.applications.default_browser := ExtractJSONValue(jsonStr, "default_browser")
    config.applications.alternative_browser := ExtractJSONValue(jsonStr, "alternative_browser")
    config.applications.vpn_client_path := ExtractJSONValue(jsonStr, "vpn_client_path")

    ; Parse browser profiles section
    config.browser_profiles := {}
    config.browser_profiles.chrome_default_profile := ExtractJSONValue(jsonStr, "chrome_default_profile")
    config.browser_profiles.chrome_music_profile := ExtractJSONValue(jsonStr, "chrome_music_profile")

    ; Parse clipboard shortcuts section
    config.clipboard_shortcuts := {}
    config.clipboard_shortcuts.email := ExtractJSONValue(jsonStr, "email")
    config.clipboard_shortcuts.access_code := ExtractJSONValue(jsonStr, "access_code")
    config.clipboard_shortcuts.git_token := ExtractJSONValue(jsonStr, "git_token")
    config.clipboard_shortcuts.cursor_command := ExtractJSONValue(jsonStr, "cursor_command")

    ; Parse URLs section
    config.urls := {}
    config.urls.outlook := ExtractJSONValue(jsonStr, "outlook")
    config.urls.jira := ExtractJSONValue(jsonStr, "jira")
    config.urls.youtube := ExtractJSONValue(jsonStr, "youtube")
    config.urls.youtube_music := ExtractJSONValue(jsonStr, "youtube_music")
    config.urls.copilot := ExtractJSONValue(jsonStr, "copilot")
    config.urls.chatgpt := ExtractJSONValue(jsonStr, "chatgpt")
    config.urls.notion := ExtractJSONValue(jsonStr, "notion")
    config.urls.slt_usage := ExtractJSONValue(jsonStr, "slt_usage")

    return config
}

ExtractJSONValue(jsonStr, key) {
    ; Simple regex to extract values from JSON
    ; Handles both regular strings and escaped backslashes
    RegexMatch(jsonStr, """" . key . """\s*:\s*""([^""\\]*(?:\\.[^""\\]*)*)""", match)
    value := match1
    ; Convert escaped backslashes back to single backslashes
    StringReplace, value, value, \\, \, All
    return value
}

ParseSensitiveJSON(jsonStr) {
    ; Simple JSON parser for sensitive data structure
    sensitive := {}

    ; Parse credentials section
    sensitive.credentials := {}
    sensitive.credentials.email := ExtractJSONValue(jsonStr, "email")
    sensitive.credentials.git_token := ExtractJSONValue(jsonStr, "git_token")

    ; Parse access codes section
    sensitive.access_codes := {}
    sensitive.access_codes.primary_access_code := ExtractJSONValue(jsonStr, "primary_access_code")

    ; Parse personal URLs section
    sensitive.personal_urls := {}
    sensitive.personal_urls.jira_board := ExtractJSONValue(jsonStr, "jira_board")
    sensitive.personal_urls.notion_workspace := ExtractJSONValue(jsonStr, "notion_workspace")

    ; Parse development section
    sensitive.development := {}
    sensitive.development.cursor_command := ExtractJSONValue(jsonStr, "cursor_command")

    ; Parse additional passwords section (for future use)
    sensitive.additional_passwords := {}
    sensitive.additional_passwords.example_service := ExtractJSONValue(jsonStr, "example_service")
    sensitive.additional_passwords.another_service := ExtractJSONValue(jsonStr, "another_service")
    sensitive.additional_passwords.vpn_password := ExtractJSONValue(jsonStr, "vpn_password")
    sensitive.additional_passwords.database_password := ExtractJSONValue(jsonStr, "database_password")

    return sensitive
}

; ============== Email Counter Functions ==============
LoadEmailCounter() {
    global emailCounterFile
    ; Try to read the counter from file
    FileRead, counterValue, %emailCounterFile%
    if ErrorLevel {
        ; File doesn't exist or can't be read, start with 960
        return 960
    }
    ; Parse the counter value, default to 960 if not valid
    if counterValue is integer
        return counterValue
    else
        return 960
}

SaveEmailCounter(counter) {
    global emailCounterFile
    FileDelete, %emailCounterFile%
    FileAppend, %counter%, %emailCounterFile%
}

GetNextEmail() {
    global emailCounter
    currentEmail := "tb" . emailCounter . "@mailinator.com"
    emailCounter++
    SaveEmailCounter(emailCounter)
    return currentEmail
}

GetLastEmail() {
    global emailCounter
    ; Return the previous email (current counter - 1)
    lastEmailNumber := emailCounter - 1
    return "tb" . lastEmailNumber . "@mailinator.com"
}

; ============== Window Management Functions ==============

; Function to run an application and maximize it (fullscreen without F11)
RunAndCenter(command, workingDir := "", windowState := "", waitTime := 500) {
    Run, %command%, %workingDir%, %windowState%, newPID
    
    ; Wait for the window to appear and become active
    Sleep, %waitTime%
    
    ; Get the active window (should be the newly opened one)
    WinGet, activeWindow, ID, A
    
    ; Maximize the window
    WinMaximize, ahk_id %activeWindow%
    
    return newPID
}

; Function to center a specific window by ID
CenterWindow(windowID) {
    WinGetPos, , , winW, winH, ahk_id %windowID%
    SysGet, screenW, 78
    SysGet, screenH, 79
    newX := (screenW - winW) // 2
    newY := (screenH - winH) // 2
    WinMove, ahk_id %windowID%, , newX, newY
}

; ============== Resolution Change Functions ==============
ChangeResolution(width, height) {
    VarSetCapacity(DeviceMode, 156, 0)
    NumPut(156, DeviceMode, 36) ; dmSize
    NumPut(0x5c0000, DeviceMode, 40) ; dmFields (DM_BITSPERPEL | DM_PELSWIDTH | DM_PELSHEIGHT | DM_DISPLAYFLAGS | DM_DISPLAYFREQUENCY)
    NumPut(32, DeviceMode, 104) ; dmBitsPerPel
    NumPut(width, DeviceMode, 108) ; dmPelsWidth
    NumPut(height, DeviceMode, 112) ; dmPelsHeight
    NumPut(100, DeviceMode, 120) ; dmDisplayFrequency (60Hz)

    result := DllCall("ChangeDisplaySettingsA", "Ptr", &DeviceMode, "UInt", 0, "Int")

    if (result = 0) {
        ; Success
        OutputDebug, Resolution changed to %width%x%height% successfully
        ; Optional: Show a brief notification
        ; TrayTip, Resolution Changed, Changed to %width%x%height%, 2
    } else if (result = 1) {
        ; Restart required
        MsgBox, 48, Resolution Change, Resolution changed to %width%x%height%.`nA restart may be required for some applications.
            OutputDebug, Resolution changed to %width%x%height% - restart may be required
    } else {
        ; Error occurred
        MsgBox, 16, Resolution Change Error, Failed to change resolution to %width%x%height%.`nError code: %result%`n`nPossible causes:`n- Resolution not supported by display`n- Graphics driver issues
        OutputDebug, Failed to change resolution to %width%x%height% - error code: %result%
    }

    return result
}

; ============== MEDIA CONTROLS ==============

; ctrl + win + s: Next track
<^#s::Media_Next
return

; ctrl + win + a: Previous track
<^#a::Media_Prev
return

; ctrl + win + space: Play/Pause
<^#Space::
    Sleep 100
    Send {Media_Play_Pause}
return

; ctrl + alt + x: Volume up
<^<!x::
    Send {Volume_Up}
return

; ctrl + alt + z: Volume down
<^<!z::
    Send {Volume_Down}
return

; ============== SYSTEM UTILITIES ==============

; ctrl + alt + s: Screenshot
<^<!s::
    Send #{PrintScreen}
return

; alt + 4: Alt+F4 alternative (works without Fn key)
<!4::<!f4
return

; win + w: Close active application
#w::
    WinClose, A
return

; ctrl + win + alt: Toggle fullscreen (F11)
^#Alt::
    Send, {F11}
return

; win + ctrl + 1: Change resolution to 1920x1080
#^1::
    ChangeResolution(1920, 1080)
return

; win + ctrl + 2: Change resolution to 3440x1440 (ultrawide)
#^2::
    ChangeResolution(3440, 1440)
return

; Double alt: Center active window on screen
~Alt::
    if (A_PriorHotkey = "~Alt" and A_TimeSincePriorHotkey < 300) {
        WinGet, activeWindow, ID, A
        WinGetPos, , , winW, winH, ahk_id %activeWindow%
        SysGet, screenW, 78
        SysGet, screenH, 79
        newX := (screenW - winW) // 2
        newY := (screenH - winH) // 2
        WinMove, ahk_id %activeWindow%, , newX, newY
    }
return

; ============== WINDOW NAVIGATION ==============

; ctrl + win + x: Next window (virtual desktop)
<^#x::<^#Right
return

; ctrl + win + z: Previous window (virtual desktop)
<^#z::<^#left
return

; ============== TEXT SELECTION ==============

; alt + a: Select text to beginning of line
!a::
    Send, +{Home}
return

; alt + s: Select text to end of line
!s::
    Send, +{End}
return

; ============== KEY REMAPPING ==============

; Escape key: Keep as escape
Esc::Esc
return

; CapsLock: Remap to Escape
CapsLock::Esc
return

; ============== APPLICATION LAUNCHERS ==============

; shift + alt + c: Default browser
<+<!c::
    RunAndCenter(defaultBrowser, "", "Max")
return

; ctrl + alt + c: Chrome with default profile
<^<!c::
    RunAndCenter("chrome.exe --profile-directory=""" . chromeDefaultProfile . """")
return

; win + n: Notepad
#N::
    RunAndCenter("notepad.exe")
return

; shift + alt + w: VPN client
<+<!w::
    RunAndCenter(vpnClientPath)
return

; ============== FOLDER SHORTCUTS ==============

; shift + alt + d: Downloads folder
<+<!d::
    RunAndCenter(userDownloadsFolder)
return

; shift + alt + s: Desktop folder
<+<!s::
    RunAndCenter(userDesktopFolder)
return

; shift + alt + p: Screenshots folder
<+<!p::
    RunAndCenter(userScreenshotsFolder)
return

; ============== WEB APPLICATIONS ==============

; shift + alt + e: Outlook (same tab)
<+<!e::
    RunAndCenter(defaultBrowser . " """ . outlookUrl . """")
return

; shift + alt + ctrl + e: Outlook (new window)
<+<!<^e::
    RunAndCenter(defaultBrowser . " """ . outlookUrl . """ --new-window")
return

; win + j: Jira board
#j::
    RunAndCenter(defaultBrowser . " """ . jiraUrl . """")
return

; win + y: YouTube (same tab)
#y::
    RunAndCenter(defaultBrowser . " """ . youtubeUrl . """")
return

; win + shift + y: YouTube (new window)
<+#y::
    RunAndCenter(defaultBrowser . " """ . youtubeUrl . """ --new-window")
return

; win + m: YouTube Music (dedicated Chrome profile)
#M::
    RunAndCenter("chrome.exe --profile-directory=""" . chromeMusicProfile . """ """ . youtubeMusicUrl . """")
return

; win + a: GitHub Copilot
#A::
    RunAndCenter(defaultBrowser . " --app=" . copilotUrl)
return

; win + shift + a: ChatGPT
<+#A::
    RunAndCenter(defaultBrowser . " --app=" . chatgptUrl)
return

; shift + alt + n: Notion
<+<!n::
    RunAndCenter(defaultBrowser . " --app=" . notionUrl)
return

; win + u: SLT usage dashboard
#u::
    RunAndCenter(defaultBrowser . " """ . sltUsageUrl . """")
return

; ============== CLIPBOARD AUTOMATION ==============

; shift + alt + ctrl + v: Next incremented email address
<+<!^v::
    Clipboard := GetNextEmail()
    SendInput, ^v
return

; shift + alt + v: Last used email address (no increment)
<+<!v::
    Clipboard := GetLastEmail()
    SendInput, ^v
return

; shift + alt + b: Primary access code
<+<!b::
    Clipboard := clipboardAccessCode
    SendInput, ^v
return

; shift + alt + g: Git token
<+<!g::
    Clipboard := clipboardGitToken
    SendInput, ^v
return

; shift + alt + z: Cursor command
<+<!z::
    Clipboard := clipboardCursorCommand
    SendInput, ^v
return

; shift + alt + a: Generate console.log with selected text
<+<!a::
    ; Store original clipboard content
    originalClipboard := Clipboard

    ; Copy the selected text
    Send, ^c
    ClipWait, 0.5 ; Wait for clipboard to contain the copied data

    ; Get the selected text
    selectedText := Clipboard

    ; Move to the end of the current line and go to next line
    Send, {End}{Enter}

    ; Create the console.log string based on whether text was selected
    if (selectedText != "" && selectedText != originalClipboard) {
        ; Text was selected - use full format with variable name and value
        ; Format: console.log('\x1b[33m%s\x1b[0m', 'selectedText', selectedText)
        consoleLogString := "console.log('\x1b[33m%s\x1b[0m', '" . selectedText . "', " . selectedText . ")"
    } else {
        ; No text was selected - use simple empty format
        ; Format: console.log('\x1b[33m%s\x1b[0m', '')
        consoleLogString := "console.log('\x1b[33m%s\x1b[0m', '')"
    }

    ; Put the console.log string in clipboard
    Clipboard := consoleLogString

    ; Paste the console.log string
    SendInput, ^v

    ; Restore original clipboard after a short delay
    Sleep, 100
    Clipboard := originalClipboard
return

; ============== VIRTUAL DESKTOP MANAGEMENT ==============

; Virtual desktop globals
DesktopCount = 2 ; Windows starts with 2 desktops at boot
CurrentDesktop = 1 ; Desktop count is 1-indexed (Microsoft numbers them this way)

; Initialize virtual desktop mapping
SetKeyDelay, 75
mapDesktopsFromRegistry()
OutputDebug, [loading] desktops: %DesktopCount% current: %CurrentDesktop%

; Virtual desktop shortcuts (Esc + number keys)
Esc & 1::switchDesktopByNumber(1)
Esc & 2::switchDesktopByNumber(2)
Esc & 3::switchDesktopByNumber(3)
Esc & 4::switchDesktopByNumber(4)
Esc & 5::switchDesktopByNumber(5)
Esc & 6::switchDesktopByNumber(6)
Esc & 7::switchDesktopByNumber(7)
Esc & 8::switchDesktopByNumber(8)
Esc & 9::switchDesktopByNumber(9)

; ============== VIRTUAL DESKTOP FUNCTIONS ==============

; This function examines the registry to build an accurate list of the current virtual desktops and which one we're currently on.
; Current desktop UUID appears to be in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\1\VirtualDesktops
; List of desktops appears to be in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops
mapDesktopsFromRegistry() {
    global CurrentDesktop, DesktopCount
    ; Get the current desktop UUID. Length should be 32 always, but there's no guarantee this couldn't change in a later Windows release so we check.
    IdLength := 32
    SessionId := getSessionId()
    if (SessionId) {
        RegRead, CurrentDesktopId, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\%SessionId%\VirtualDesktops, CurrentVirtualDesktop
        if (CurrentDesktopId) {
            IdLength := StrLen(CurrentDesktopId)
        }
    }
    ; Get a list of the UUIDs for all virtual desktops on the system
    RegRead, DesktopList, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops, VirtualDesktopIDs
    if (DesktopList) {
        DesktopListLength := StrLen(DesktopList)
        ; Figure out how many virtual desktops there are
        DesktopCount := DesktopListLength / IdLength
    }
    else {
        DesktopCount := 1
    }
    ; Parse the REG_DATA string that stores the array of UUID's for virtual desktops in the registry.
    i := 0
    while (CurrentDesktopId and i < DesktopCount) {
        StartPos := (i * IdLength) + 1
        DesktopIter := SubStr(DesktopList, StartPos, IdLength)
        OutputDebug, The iterator is pointing at %DesktopIter% and count is %i%.
        ; Break out if we find a match in the list. If we didn't find anything, keep the
        ; old guess and pray we're still correct :-D.
        if (DesktopIter = CurrentDesktopId) {
            CurrentDesktop := i + 1
            OutputDebug, Current desktop number is %CurrentDesktop% with an ID of %DesktopIter%.
            break
        }
        i++
    }
}

; This functions finds out ID of current session.
getSessionId() {
    ProcessId := DllCall("GetCurrentProcessId", "UInt")
    if ErrorLevel {
        OutputDebug, Error getting current process id: %ErrorLevel%
        return
    }
    OutputDebug, Current Process Id: %ProcessId%
    DllCall("ProcessIdToSessionId", "UInt", ProcessId, "UInt*", SessionId)
    if ErrorLevel {
        OutputDebug, Error getting session id: %ErrorLevel%
        return
    }
    OutputDebug, Current Session Id: %SessionId%
return SessionId
}

; This function switches to the desktop number provided.
switchDesktopByNumber(targetDesktop) {
    global CurrentDesktop, DesktopCount
    ; Re-generate the list of desktops and where we fit in that. We do this because
    ; the user may have switched desktops via some other means than the script.
    mapDesktopsFromRegistry()
    ; Don't attempt to switch to an invalid desktop
    if (targetDesktop > DesktopCount || targetDesktop < 1) {
        OutputDebug, [invalid] target: %targetDesktop% current: %CurrentDesktop%
        return
    }
    ; Go right until we reach the desktop we want
    while(CurrentDesktop < targetDesktop) {
        Send ^#{Right}
        CurrentDesktop++
        OutputDebug, [right] target: %targetDesktop% current: %CurrentDesktop%
    }
    ; Go left until we reach the desktop we want
    while(CurrentDesktop > targetDesktop) {
        Send ^#{Left}
        CurrentDesktop--
        OutputDebug, [left] target: %targetDesktop% current: %CurrentDesktop%
    }
}

; This function creates a new virtual desktop and switches to it
createVirtualDesktop() {
    global CurrentDesktop, DesktopCount
    Send, #^d
    DesktopCount++
    CurrentDesktop = %DesktopCount%
    OutputDebug, [create] desktops: %DesktopCount% current: %CurrentDesktop%
}

; This function deletes the current virtual desktop
deleteVirtualDesktop() {
    global CurrentDesktop, DesktopCount
    Send, #^{F4}
    DesktopCount--
    CurrentDesktop--
    OutputDebug, [delete] desktops: %DesktopCount% current: %CurrentDesktop%
}

; ============== COMMENTED VIRTUAL DESKTOP ALTERNATIVES ==============
; Alternative virtual desktop shortcuts (currently disabled)
; Uncomment these if you prefer different key combinations:
;
; LWin shortcuts:
; LWin & 1::switchDesktopByNumber(1)
; LWin & 2::switchDesktopByNumber(2)
; LWin & 3::switchDesktopByNumber(3)
; LWin & 4::switchDesktopByNumber(4)
; LWin & 5::switchDesktopByNumber(5)
; LWin & 6::switchDesktopByNumber(6)
; LWin & 7::switchDesktopByNumber(7)
; LWin & 8::switchDesktopByNumber(8)
; LWin & 9::switchDesktopByNumber(9)
;
; CapsLock shortcuts:
; CapsLock & 1::switchDesktopByNumber(1)
; CapsLock & 2::switchDesktopByNumber(2)
; CapsLock & 3::switchDesktopByNumber(3)
; CapsLock & 4::switchDesktopByNumber(4)
; CapsLock & 5::switchDesktopByNumber(5)
; CapsLock & 6::switchDesktopByNumber(6)
; CapsLock & 7::switchDesktopByNumber(7)
; CapsLock & 8::switchDesktopByNumber(8)
; CapsLock & 9::switchDesktopByNumber(9)
; CapsLock & n::switchDesktopByNumber(CurrentDesktop + 1)
; CapsLock & p::switchDesktopByNumber(CurrentDesktop - 1)
; CapsLock & s::switchDesktopByNumber(CurrentDesktop + 1)
; CapsLock & a::switchDesktopByNumber(CurrentDesktop - 1)
; CapsLock & c::createVirtualDesktop()
; CapsLock & d::deleteVirtualDesktop()
;
; Ctrl+Alt shortcuts:
; ^!1::switchDesktopByNumber(1)
; ^!2::switchDesktopByNumber(2)
; ^!3::switchDesktopByNumber(3)
; ^!4::switchDesktopByNumber(4)
; ^!5::switchDesktopByNumber(5)
; ^!6::switchDesktopByNumber(6)
; ^!7::switchDesktopByNumber(7)
; ^!8::switchDesktopByNumber(8)
; ^!9::switchDesktopByNumber(9)
; ^!n::switchDesktopByNumber(CurrentDesktop + 1)
; ^!p::switchDesktopByNumber(CurrentDesktop - 1)
; ^!s::switchDesktopByNumber(CurrentDesktop + 1)
; ^!a::switchDesktopByNumber(CurrentDesktop - 1)
; ^!c::createVirtualDesktop()
; ^!d::deleteVirtualDesktop()