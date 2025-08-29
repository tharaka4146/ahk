#Persistent ; Keeps the script running
; #NoTrayIcon
#SingleInstance force
DetectHiddenWindows, on

; ============== Configuration Loading ==============
; Read configuration from config_office.json and sensitive_office.json files
configFile := A_ScriptDir . "\config_office.json"
sensitiveFile := A_ScriptDir . "\sensitive_office.json"
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
teamsChatUrl := config.urls.teams_chat
chatgptUrl := config.urls.chatgpt
notionUrl := sensitive.personal_urls.notion_workspace
sltUsageUrl := config.urls.slt_usage
testfullyUrl := config.urls.testfully

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
    config.urls.teams_chat := ExtractJSONValue(jsonStr, "teams_chat")
    config.urls.chatgpt := ExtractJSONValue(jsonStr, "chatgpt")
    config.urls.notion := ExtractJSONValue(jsonStr, "notion")
    config.urls.slt_usage := ExtractJSONValue(jsonStr, "slt_usage")
    config.urls.testfully := ExtractJSONValue(jsonStr, "testfully")
    
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

; ============== media shortcuts ==============

; ctrl + win + x
<^#s::Media_Next ; next track
return

; ctrl + win + z
<^#a::Media_Prev ; previous track
return

; ctrl + win + space
<^#Space::
    ;Run, C:\Users\UBANDT1\Downloads\nircmd-x64\nircmd.exe mediapause
    ; Send {Media_Stop}
    Sleep 100
    Send {Media_Play_Pause} 
    ;Run, curl -X POST http://localhost:9222/json/execute -d "{""expression"": ""document.querySelector('video')?.pause()""}"
    ;Send ^+p
    ;Send {Media_Play_Pause}
    ;Send {Media_Play_Pause down}
    ;Sleep 100
    ;Send {Media_Play_Pause up}
return

; ctrl + alt + x
<^<!x:: ; increase volume by 2
    send {Volume_Up}
return

; ctrl + alt + z
<^<!z:: ; decrease volume by 2
    send {Volume_Down}
return

; ctrl + alt + s
<^<!s:: ; screenshot
    send #{PrintScreen}
return

; ============== apps shortcuts ==============

<+<!c::
    Run, %defaultBrowser%,, Max ; run msedge browser
return

; ctrl + alt + c
<^<!c::
    ; Run, chrome.exe ; run msedge browser
    Run, chrome.exe --profile-directory="%chromeDefaultProfile%"
return

; shift + alt + n
#N::
    Run, notepad.exe ; run notepad
return

; ; shift + alt + m
; <+<!m::
;     Run, AppleMusic.exe ; run apple music (not itunes)
; return

; shift + alt + a
; <+<!a:: ;
;     Run, Teams.exe ; run teams
; return

; shift + alt + w
<+<!w::
    Run, %vpnClientPath% ; run vpn
return

; ============== folder shortcuts ==============

; shift + alt + d
<+<!d:: ; open downloads folder
    Run, %userDownloadsFolder%
return

; shift + alt + s
<+<!s:: ; open desktop folder
    Run, %userDesktopFolder%
return

; shift + alt + p
<+<!p:: ; open screenshots folder
    Run, %userScreenshotsFolder%
return

; ============== browser shortcuts ==============

; shift + alt + e
<+<!e:: ; open outlook in new window
    Run, %defaultBrowser% "%outlookUrl%"
return

; win J
#j:: ; open jira in new window
    Run, %defaultBrowser% "%jiraUrl%"
return

; shift + alt + ctrl + e
<+<!<^e:: ; open outlook in new window
    Run, %defaultBrowser% "%outlookUrl%" " --new-window "
return

; win Y
#y:: ; open youtube in new window
    Run, %defaultBrowser% "%youtubeUrl%"
return

; win A
#A:: ; open teams chat in new window
    Run, %defaultBrowser% --app=%teamsChatUrl%
return

; win A
<+#A:: ; open chatgpt in new window
    Run, %defaultBrowser% --app=%chatgptUrl%
return

<+<!n:: ; open notion in new window
    Run, %defaultBrowser% --app=%notionUrl%
return

; win shift Y
<+#y:: ; open youtube in new window
    Run, %defaultBrowser% "%youtubeUrl%" " --new-window "
return

; win M
; #m:: ; open youtube in new window
;     Run, %defaultBrowser% "https://www.youtube.com/results?search_query=music" " --new-window "
; return

; ; shift + alt + m
#M:: ; open youtube music in new window
    ; Run, chrome.exe --profile-directory="%chromeMusicProfile%" --app="%youtubeMusicUrl%"
    Run, chrome.exe --profile-directory="%chromeMusicProfile%" "%youtubeMusicUrl%"
return

; ; shift + alt + ctrl + m
; <+<!<^m:: ; open youtube music in new window
;     Run, chrome.exe --profile-directory="profile 3" "https://www.youtube.com/watch?v=MoN9ql6Yymw&list=RDMM&start_radio=1" " --new-window "
; return

; shift + alt + u
#u:: ; open slt usage in new window
    Run, %defaultBrowser% --app=%sltUsageUrl%
return

; shift + alt + t
<+<!t:: ; open testfully in new window
    Run, %defaultBrowser% "%testfullyUrl%"
return

; ============== alt f4 shortcut ==============

; alt + 4
<!4::<!f4 ; alt + f4 works with or withtout the fn button

; win + w - close active application
#w::
    WinClose, A
return

; alt + a - highlight text to the left
!a::
    Send, +{Home}
return

!s::
    Send, +{End}
return

; ctrl + win + x
<^#x::<^#Right ; next window
return

; ctrl + win + z
<^#z::<^#left ; previous window
return

; ctrl + win + alt
^#Alt::
    Send, {F11}
return

; ctrl + win + alt
; #z::
;     Send, {F11}
; return

; esc to esc
Esc::Esc
return

; capslock to esc
CapsLock::Esc
return

;endofscript

; ============== mouse button shortcut ==============

; copy, mouse forward
; XButton2::^v
; return

; paste, mouse back
; XButton1::^c
; return

;endofscript

; ----------------------------------------- desktop switching scritp ----------------------------------------

; Globals
DesktopCount = 2 ; Windows starts with 2 desktops at boot
CurrentDesktop = 1 ; Desktop count is 1-indexed (Microsoft numbers them this way)
;
; This function examines the registry to build an accurate list of the current virtual desktops and which one we're currently on.
; Current desktop UUID appears to be in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\1\VirtualDesktops
; List of desktops appears to be in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops
;
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
;
; This functions finds out ID of current session.
;
getSessionId()
{
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
;
; This function switches to the desktop number provided.
;
switchDesktopByNumber(targetDesktop)
{
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
;
; This function creates a new virtual desktop and switches to it
;
createVirtualDesktop()
{
    global CurrentDesktop, DesktopCount
    Send, #^d
    DesktopCount++
    CurrentDesktop = %DesktopCount%
    OutputDebug, [create] desktops: %DesktopCount% current: %CurrentDesktop%
}
;
; This function deletes the current virtual desktop
;
deleteVirtualDesktop()
{
    global CurrentDesktop, DesktopCount
    Send, #^{F4}
    DesktopCount--
    CurrentDesktop--
    OutputDebug, [delete] desktops: %DesktopCount% current: %CurrentDesktop%
}
; Main
SetKeyDelay, 75
mapDesktopsFromRegistry()
OutputDebug, [loading] desktops: %DesktopCount% current: %CurrentDesktop%
; User config!
; This section binds the key combo to the switch/create/delete actions
; LWin & 1::switchDesktopByNumber(1)
; LWin & 2::switchDesktopByNumber(2)
; LWin & 3::switchDesktopByNumber(3)
; LWin & 4::switchDesktopByNumber(4)
; LWin & 5::switchDesktopByNumber(5)
; LWin & 6::switchDesktopByNumber(6)
; LWin & 7::switchDesktopByNumber(7)
; LWin & 8::switchDesktopByNumber(8)
; LWin & 9::switchDesktopByNumber(9)
;CapsLock & 1::switchDesktopByNumber(1)
;CapsLock & 2::switchDesktopByNumber(2)
;CapsLock & 3::switchDesktopByNumber(3)
;CapsLock & 4::switchDesktopByNumber(4)
;CapsLock & 5::switchDesktopByNumber(5)
;CapsLock & 6::switchDesktopByNumber(6)
;CapsLock & 7::switchDesktopByNumber(7)
;CapsLock & 8::switchDesktopByNumber(8)
;CapsLock & 9::switchDesktopByNumber(9)
;CapsLock & n::switchDesktopByNumber(CurrentDesktop + 1)
;CapsLock & p::switchDesktopByNumber(CurrentDesktop - 1)
;CapsLock & s::switchDesktopByNumber(CurrentDesktop + 1)
;CapsLock & a::switchDesktopByNumber(CurrentDesktop - 1)
;CapsLock & c::createVirtualDesktop()
;CapsLock & d::deleteVirtualDesktop()
; Alternate keys for this config. Adding these because DragonFly (python) doesn't send CapsLock correctly.
;^!1::switchDesktopByNumber(1)
;^!2::switchDesktopByNumber(2)
;^!3::switchDesktopByNumber(3)
;^!4::switchDesktopByNumber(4)
;^!5::switchDesktopByNumber(5)
;^!6::switchDesktopByNumber(6)
;^!7::switchDesktopByNumber(7)
;^!8::switchDesktopByNumber(8)
;^!9::switchDesktopByNumber(9)
;^!n::switchDesktopByNumber(CurrentDesktop + 1)
;^!p::switchDesktopByNumber(CurrentDesktop - 1)
;^!s::switchDesktopByNumber(CurrentDesktop + 1)
;^!a::switchDesktopByNumber(CurrentDesktop - 1)
;^!c::createVirtualDesktop()
;^!d::deleteVirtualDesktop()
Esc & 1::switchDesktopByNumber(1)
Esc & 2::switchDesktopByNumber(2)
Esc & 3::switchDesktopByNumber(3)
Esc & 4::switchDesktopByNumber(4)
Esc & 5::switchDesktopByNumber(5)
Esc & 6::switchDesktopByNumber(6)
Esc & 7::switchDesktopByNumber(7)
Esc & 8::switchDesktopByNumber(8)
Esc & 9::switchDesktopByNumber(9)

; paste keys

;  Clipboard := "PPELOE-AHEAD-ASPEN-BOWAN-PATHS-SIRES"; revel
;  Clipboard := "WASU-QESHM-OSMIC-VARNA-ABASH-ROUSE"  ; SMS

; shift + alt + v
<+<!v::
    ; Clipboard := "{{EXCHANGE_URL_ELB}}"
    Clipboard := clipboardEmail
    SendInput, ^v
return

; shift + alt + b
<+<!b::
    ; Clipboard := "PIMS-QESHM-OSMIC-VARNA-RETOT-LOOSE" ; MLM new
    ; Clipboard := "TBATRE-AHEAD-ASPEN-BOWAN-SAPIR-NONES" ; revel old
    ; Clipboard := "PPELOE-AHEAD-ASPEN-BOWAN-PATHS-SIRES" ; revel
    ; Clipboard := "PPELOE-AHEAD-ASPEN-BOWAN-PATHS-SIRES" ; revel combo
    ; Clipboard := "TBATRE-BAEDA-BEECH-ELSAN-LOBBY-LINES" ; revel with related product
    ; Clipboard := "TBATRE-BAEDA-BEECH-ELSAN-LOBBY-LINES" ; revel new stg
    Clipboard := clipboardAccessCode ; MLM
    ; Clipboard := "DSWBJN-PRANK-BEECH-ELSAN-HELOT-FLEES" ; 1015 bug
    ; Clipboard := "472d0a3-df52-4baa-9a06-97d936e74a5a" ; 404 product on exchange ui
    ; Clipboard := "TIERN-SCOFF-TOXIC-HOOGH-TRAWL-LUTES" ; REVEL no-offer-id stg
    ; Clipboard := "TTIERN-SCOFF-TOXIC-HOOGH-TRAWL-LUTEB" ; REVEL invalid AC
    ; Clipboard := "WASU-SCOFF-OSMIC-VARNA-CHOIR-JUTES" ; SFC, OCC validation error
    SendInput, ^v
return

; shift + alt + g
<+<!g::
    ; Clipboard := "glpat-LFkF9ziN9yPepsuAU9W6"
    ; Clipboard := "glpat--rQSUNYkRkk_J64kkid8"
    Clipboard := clipboardGitToken
    SendInput, ^v
return

; shift + alt + z
<+<!z::
    Clipboard := clipboardCursorCommand
    SendInput, ^v
return

; ; shift + alt + a
; <+<!a::
;     ; Clipboard := "glpat-LFkF9ziN9yPepsuAU9W6"
;     Clipboard := "console.log('\x1b[33m%s\x1b[0m', '', )"
;     SendInput, ^v
; return

; shift + alt + a
; <+<!a::
;     ; Store original clipboard content
;     originalClipboard := Clipboard

;     ; Copy the selected text
;     Send, ^c
;     ClipWait, 0.5 ; Wait for clipboard to contain the copied data

;     ; Get the selected text
;     selectedText := Clipboard

;     ; Create the console.log string with the selected text
;     ; Format: console.log('\x1b[33m%s\x1b[0m', 'selectedText', selectedText)
;     consoleLogString := "console.log('\x1b[33m%s\x1b[0m', '" . selectedText . "', " . selectedText . ")"

;     ; Put the console.log string in clipboard
;     Clipboard := consoleLogString

;     ; Paste the console.log string
;     SendInput, ^v

;     ; Restore original clipboard after a short delay
;     Sleep, 100
;     Clipboard := originalClipboard
; return

; shift + alt + a
; <+<!a::
;     ; Get the text that's already in clipboard (manually copied)
;     selectedText := Clipboard

;     ; Create the console.log string with the copied text
;     ; Format: console.log('\x1b[33m%s\x1b[0m', 'selectedText', selectedText)
;     consoleLogString := "console.log('\x1b[33m%s\x1b[0m', '" . selectedText . "', " . selectedText . ")"

;     ; Put the console.log string in clipboard and paste it
;     Clipboard := consoleLogString
;     SendInput, ^v
; return

; shift + alt + a
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

    ; Create the console.log string with the selected text
    ; Format: console.log('\x1b[33m%s\x1b[0m', 'selectedText', selectedText)
    consoleLogString := "console.log('\x1b[33m%s\x1b[0m', '" . selectedText . "', " . selectedText . ")"

    ; Put the console.log string in clipboard
    Clipboard := consoleLogString

    ; Paste the console.log string
    SendInput, ^v

    ; Restore original clipboard after a short delay
    Sleep, 100
    Clipboard := originalClipboard
return
; ----------------------------------------- center windows ----------------------------------------

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

; #Persistent
; SetTimer, HideTaskbar, 1000

; HideTaskbar() {
;     WinHide, ahk_class Shell_TrayWnd
;     WinHide, ahk_class NotifyIconOverflowWindow
; }

; #IfWinActive ; Applies to all windows
; ~LWin::
;     WinShow, ahk_class Shell_TrayWnd
;     WinShow, ahk_class NotifyIconOverflowWindow
;     SetTimer, HideTaskbar, -3000 ; Hide again after 3 seconds
; return
