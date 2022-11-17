;~ ;ChemHelp v.2.0 - updated 2022-01-16
;Written by Dieter van der Westhuizen 2018-2022
;Inspired from TrakHelper by Chad Centner

/* Changes from previous version 1.11 added and to be incorporated into Github
1. added 2 columns to the extraction script: branch and user_site which should be updated on Excel template and uploaded.
2. Added Button AutoA (need to uncomment it to get it active - use with caution)
3. Increased FPSA test set maintenance opening wait time
4. Added mousewheel settings for Cumulative view window - useful when doing SPE's by "F3" function viewing previous results (on Tuesdays).
5. Added comment: ::IMMFIX::Immunofixation will be performed.  See result below, to follow.
6. Added comment: ::CTT::Close to target. Stable.
7. Added rSPE: to place SPE on hold, then refer the result to Chem Registrar VQ list so it is visible.
8. Added UPE and UIFE buttons in the more menu (lines)
9. Removed messagebox after dCDUM button pressed.
10. Changed EPR URL - removed port number.
11. Changed EPR URL - https for secure connection.

Changes in v2.0 (coded from 2022-01-11)
-Alt+B and Alt+V shortcuts works more efficiently in Result Entry - Single and Verify windows - Used the new If WinActive() - syntax
-Alt+Q and Alt+Z time delay shortened to 50ms from 100ms to page up / down faster.
-Alt+Q and Alt+Z added for Cumulative View - useful when viewing SPE histories. It has also been made active for non-TrakCare windows such as Chrome / Internet explorer.
-Tried to change mouse click position for login (adding to the variable lengths - unable to know if it will work)
-Added Equation-ip to the configuration file in ChemHelp-directory to make the form viewing compatible with other sites' Equation server IP. Your user site's Equation server IP should be the 6th item in chemhelp-config.txt file in ChemHelp-directory (comma separated from the other fields): 172.22.4.40 (GSH), 172.22.8.50 (RXH), 172.22.12.40 (TBH), 172.22.16.63 (?Green Point Lab)
-Changed the sleep (wait) times in the Alt+F and Alt+W functions - original sleep times are in comments at the EpMedVal() and MRNMedVal() function definitions.
-Added a way of logging the current episode number with Ctrl+Alt+N or Ctrl+Alt+L - it saves the episode number and the time in ChemHelp-directory\ChemHelp\episode_log.txt - may be useful for case logging.
-Added Window Capture button. Note you need to add a Screenshots folder in "ChemHelp-directory" - may be useful for case logging.
-Alt+W and Alt+F now works from within "Result Entry (Total Count)" and "Result Entry - Verify" window - no need to open the particular episode first - you can just select it by clicking and highlighting it.
-Keep-open button now doesn't cause the active window and mouse position to be lost after keeping TrakCare alive - thus more user-friendly.
-_i (info) link updated to point to Github.
-Directory structure changed: ChemHelp.ahk MUST now be in a separate folder along with its dependency files and folders else it will produce various errors.
-Configuration file structure changed. It now has fields for the Equation server IP and your own name (Hotstring d1)
-d1 Hotstring should now type your name (as in config file).
-Added 2 new hotkeys: Ctrl + Alt + F: Adds a Free-PSA and removes the CDUM.  Ctrl + Alt + D: Removes a CDUM.
-Changed 1 hotkey: Ctrl + Alt + W (from Shift + Alt + W) - to be more consistent with above new hotkeys.
-Removed shortcut for Shift+Alt+R - escape works better.
-Added shortcut to navigate to Results Entry - by using Ctrl+Alt+P - similar to Alt+P but without the "PHONC"-part.
-Added HYPOA and INFLAMOLIGO comments for SPE reporting.

Changes in v2.1 (coded from 2022-06-30)
-Faster Episode and MRN number requisition (from F7 in Medical Validation window and episode number from Audit trail or title of the window when applicable). When applicable, these also do not rely on the copying to the clipboard. It reads the title of the window rather than clicking and copying.
-Automatically click the close button when "Show all content" in Internet Explorer appears, since the https (SSL) EPR server upgrade - requires a new file, show_all_content.PNG in the trakcare_icons folder.
-A fix or two in the Alt+N shortcut to copy the episode to the clipboard.
-Tooltip now shows for 1 second with the saved MRN or episode when any of the functions run which uses it.
-Error checking when opening / waiting for the episode or MRN acquisition windows.

To-do in v2.2 and beyond
-https://www.autohotkey.com/board/topic/81915-solved-gui-control-tooltip-on-hover/ - to show hotkeys when hovering mouse over the app
-API which sends the size of VQ list over Telegram / Whatsapp (use tooltip scrape or an OCR program, such as command: C:\Program Files\Capture2Text_v4.6.0_64bit\Capture2Text>Capture2Text_CLI.exe --screen-rect "931 79 957 91" -o "C:\test.txt" --trim-capture --blacklist abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.! --scale-factor 2.5
-Add uFix button & changed rFix to sFix.
-Do something cool with OCR using Capture-2-text's CLI.

Servers failed to copy:
17
32
24
29


Servers copying fine:
39
15 x2
2
37



*/

;;;;;;;;;;;;;;;;;;;;;;;;;;;; AHK syntax notes;;;;;;;;;;;;;;;;;;
; ;just creates a comment so that code is not active
;code = hotkey
;! = alt
;^ = ctrl
;+ = shift
;# = windows button


#SingleInstance, force
;#NoTrayIcon
#NoEnv
;#Persistent
#MaxThreadsPerHotkey 2
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 85
SendMode, Input
SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
SetMouseDelay, 0
SetWinDelay, 500


;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Add Buttons ;;;;;;;;;;;;;;;;;;
Gui, Add, button, x2 y2 w40 h20, Log-on
Gui, Add, button, x44 y2 w40 h20, Mobile
Gui, add, button,x2 y22 w30 h20, Form
;Gui, add, button,x34 y22 w35 h20, AutoA ;Hidden by default
Gui, add, button,x34 y22 w45 h20, Capture
Gui, Add, button, x2 y42 w28 h20, EPR
Gui, Add, button, x30 y42 w45 h20, EPRxlsx
Gui, Add, button, x2 y64 w33 h20, FPSA
Gui, Add, button, x36 y64 w45 h20, dCDUM
Gui, Add, button, x2 y86 w43 h20, Verified
Gui, Add, button, x46 y86 w35 h20, rFix
Gui, Add, button, x2 y108 w57 h20, KeepOpen
Gui, Add, button, x60 y108 w20 h20, Ex
Gui, Add, button, x2 y130 w36 h20, More
Gui, Add, button, x40 y130 w36 h20, VPN
;Gui, Add, button, x40 y130 w45 h20, ExMRN
Gui, Add, button, x2 y152 w32 h20, Close
Gui, Add, button, x40 y152 w20 h20, _i
Gui, Add, button, x60 y152 w20 h20, _t
;Gui, Add, Button, x6 y17 w100 h30 , Ok
;Gui, Add, Button, x116 y17 w100 h30 , Cancel
;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Set Window Options   ;;;;;;;;;;;;;;
;Gui, +AlwaysOnTop
Gui, -sysmenu +AlwaysOnTop
Gui, Show, , ChemHelp2.1
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
height := A_ScreenHeight-400
width := A_ScreenWidth-98
Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%width% y%height% w84

ToolTip, ChemHelp (re)started...
SetTimer, RemoveToolTip, -1000

;The following part of initialisation reads a configuration file located at the user's "ChemHelp-directory", and saves the variables,
;saved as comma separated values into the various variables below when the script is loaded.
FileRead, chemhelpsettingsraw, %A_ScriptDir%\chemhelp_settings.txt
if ErrorLevel
    MsgBox, Failed to load the chemhelp_settings.txt from "ChemHelp-directory" folder. `nMake sure you place a file chemhelp_settings.txt formatted as the following with the commas: citrix.username,citrixpassword,trakcareusername,trakcarepassword,amnt-of-tabs-to-get-to-citrix-icon,equationserver-ip
else
settingsarray := StrSplit(chemhelpsettingsraw, ",")
citrix_username := settingsarray[1]
citrix_password := settingsarray[2]
trakcare_username := settingsarray[3]
trakcare_password := settingsarray[4]
full_name := settingsarray[5]
tabs_citrix := settingsarray[6]
equation_ip := settingsarray[7]
return

/*
;Scheduling a task:
Loop,
{
Sleep, 60000; check time every minet
if %A_Hour% = 17 ; 5 o clock french time
    {
        ; your script here
    }
}
*/

/*
;                                                          Generic script to close a pop-up window.
office_activation_watch:
IfWinExist, %titleoffice% ahk_class %classoffice%
{
  WinActivate, %titleoffice%
  WinWaitActive, %titleoffice%
  sleep, 200
  send, {AltDown}c{AltUp}
  sleep, 100
settimer, office_activation_watch, off
}
else
  tooltip    ; remove the tooltip
return
*/

;--                                                              Enabling and disabling the proxy
;Write to the registry; set autoproxy to this value
+!p::
RegWrite, REG_DWORD, HKCU, Software\Microsoft\Windows\CurrentVersion\Internet Settings, ProxyEnable, 1
if errorlevel
	MsgBox % A_LastError
else
    MsgBox, Proxy enabled
return

;Write to the registry; set autoproxy to this value
+!r::
RegWrite, REG_DWORD, HKCU, Software\Microsoft\Windows\CurrentVersion\Internet Settings, ProxyEnable, 0
if errorlevel
	MsgBox % A_LastError
else
    MsgBox, Proxy disabled
return

; inetcpl.cpl <- Run this command in Run prompt to set proxy settings


/*
RemoveToolTip:
ToolTip
return
*/

buttonVPN:
run, "C:\Program Files\ShrewSoft\VPN Client\ipseca.exe"
WinWaitActive, VPN Access Manager
WinActivate, VPN Access Manager
MouseClick, left, 46, 121, 2, 100
WinWaitActive, VPN Connect
send, %citrix_username%
sleep, 200
send, {TAB down}{TAB up}
sleep, 200
send, %citrix_password%
sleep, 300
send, {TAB down}{TAB up}
send, {Return}
sleep, 2000
WinMinimize, VPN Connect
WinMinimize, VPN Access Manager
return


#c::
Run, Calc.exe
return


ButtonMobile:
run, ChemHelpMobile.ahk
Return

;---------------------------------             This is to LOG-ON with username and password.

ButtonLog-on:
settimer, win_proxy_logon, 500
settimer, citrix_proxy_logon, 500
settimer, miscellaneous_user_config_file_notice, 500

url :="https://lisctxstorefnt.nhls.ac.za/Citrix/StoreWeb/"
run, %url%
WinActivate, Citrix Workspace
WinWaitActive, Citrix Workspace
sleep, 2500
send, %citrix_username%
sleep, 500
send, {TAB down}{TAB up}
sleep, 200
send, %citrix_password%
sleep, 300
send, {Enter}
WinWaitActive, Citrix Workspace
WinWaitActive, Citrix Workspace
sleep, 3000
Remove_IE_notice()
/*
Loop %tabs_citrix%
{
    Send {Tab down}
	Send {Tab up}
    Sleep 80  ;      The number of milliseconds between keystrokes (or use SetKeyDelay).
}
*/
ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big100.PNG
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for the TrakCare Icon. Possibly the files in the folder ChemHelp\trakcare_icons is missing in your "Documents" folder.
else if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big095.PNG
    if (ErrorLevel = 1)
        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big090.PNG
        if (ErrorLevel = 1)
            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big085.PNG
            if (ErrorLevel = 1)
                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big080.PNG
                if (ErrorLevel = 1)
                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big105.PNG
                    if (ErrorLevel = 1)
                        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big110.PNG
                        if (ErrorLevel = 1)
                            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big115.PNG
                            if (ErrorLevel = 1)
                                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big120.PNG
                                if (ErrorLevel = 1)
                                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_big125.PNG
if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small100.PNG
if (ErrorLevel = 1)
    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small095.PNG
    if (ErrorLevel = 1)
        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small090.PNG
        if (ErrorLevel = 1)
            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small085.PNG
            if (ErrorLevel = 1)
                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small080.PNG
                if (ErrorLevel = 1)
                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small105.PNG
                    if (ErrorLevel = 1)
                        ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small110.PNG
                        if (ErrorLevel = 1)
                            ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small115.PNG
                            if (ErrorLevel = 1)
                                ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small120.PNG
                                if (ErrorLevel = 1)
                                    ImageSearch, FoundX, FoundY, 0,0,3000,2000, %A_ScriptDir%\trakcare_icons\trakcarelab_small125.PNG
else if (ErrorLevel = 1)
MsgBox TrakCare Icon could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x`;%FoundY%y. `nIt will now try clicking on it.
    FoundX += 20
    FoundY += 20
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
sleep, 300

send, {Enter}
sleep, 2000

settimer, labtrakstart_login, 1000

;;;;Windows Blue Proxy logon-screen sign-in
win_proxy_logon:
IfWinExist, Windows sign-in
{
WinActivate, Windows sign-in
send, {tab}
sleep, 200
send, %citrix_password%
sleep, 200
send, {enter}
settimer, win_proxy_logon, Off
Remove_IE_notice()
}
else
    tooltip ;nothing
    return

citrix_proxy_logon:
IfWinExist, Proxy authentication required
{
WinActivate, Proxy authentication required
    sleep, 100
    send, %citrix_username%
    sleep, 200
    send, {Tab}
    sleep, 100
    send, %citrix_password%
    sleep, 200
    send, {Enter}
    sleep, 1000
    send, %citrix_password%
    sleep, 200
    Remove_IE_notice()
    ;settimer, citrix_proxy_logon, Off
}
else
    return

labtrakstart_login:
IfWinExist,  (Current Site : NHLS)
{
    settimer, labtrakstart_login, Off
    settimer, win_proxy_logon, Off
    ;settimer, citrix_proxy_logon, Off
    settimer, miscellaneous_user_config_file_notice, Off
    WinActivate,  (Current Site : NHLS)
    WinWaitActive,  (Current Site : NHLS)
    sleep, 200
    send, %trakcare_username%
    sleep, 500
    send, {TAB down}{TAB up}
    send, %trakcare_password%
    sleep, 50
send, {TAB down}{TAB up}
sleep, 300
SendInput, {down}
sleep, 1000
sendinput, {Enter}
return
;Loop, %trakcare_workarea_tabs%
;{
;send, {TAB down}{TAB up}
;sleep, 500
;}
;send, {Down}
;sleep, 200
;send, {Enter}
;Return
}
else
  tooltip Waiting for TrakCare Window
return

miscellaneous_user_config_file_notice:
IfWinExist, LabTrakStart
{
WinActivate, LabTrakStart
WinWaitActive, LabTrakStart
    Send, {Enter}
    settimer, miscellaneous_user_config_file_notice, Off
    return
}
else
return

;;;;;;;;;;;;;;;;;;;;;                                                                           CREF error C900 entered but test is registered on TrakCare
!y::
txt:=EpResultSingle()
sleep, 100
FileAppend, CREF-error`n%txt%`n, C:\TrakCare-Errors\TrakCare-Entry-Errors.txt
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                         Button KeepOpen
ButtonKeepOpen:
;MsgBox, This script clicks the refresh button on "Results Entry - Verify" window every 10 minutes to keep TrakCare open.`nPlease make sure that "Results Entry - Verify" window is open when you leave your computer.`nClick OK to start.`nHit Esc when you are back on your computer.

MsgBox, 1, , This script clicks the refresh button on "Results Entry - Verify" window every 10 minutes to keep TrakCare open.`nPlease make sure that "Results Entry - Verify" window is open when you leave your computer.`nClick OK to start.`nHit Esc when you are back on your computer.`nThis MsgBox will time out in 5 seconds.  Continue?, 5
sleep, 300
SetTimer, refreshtimer, 600000 ;every 10 minutes
;SetTimer, refreshtimer, 300000 ;every 5 minutes
;SetTimer, refreshtimer, 10000 ;every 10 sec
;MsgBox, 1, , Timer has been set.`nThis Message will disappear automatically in 2 seconds, 2
Return

;This is the script to click refresh on Lab Results verify window.
refreshtimer:
if WinExist("Result Entry - Verify ahk_class Transparent Windows Client")
{
  WinGetTitle, last_window_title, A
  MouseGetPos, xmousepos, ymousepos
  WinActivate, Result Entry - Verify ahk_class Transparent Windows Client
  WinWaitActive, Result Entry - Verify ahk_class Transparent Windows Client
  sleep, 100
  WinGetPos,,,winw, winh
  mouseclick, left, winw-60 , 125, 1
  sleep, 50
    WinActivate, %last_window_title%
    MouseMove, %xmousepos%, %ymousepos%
  sleep, 150
  WinMinimize, Result Entry - Verify ahk_class Transparent Windows Client
  sleep, 500
  ToolTip, Trak is being kept open automatically. `n Hit ESC to stop.  ;This tooltip doesn't want to work well for some reason. I only get it working with a while loop.
  SetTimer, RemoveToolTip, -1000
  sleep, 1000
  return
}
else if WinExist(" NHLS")
{
    If !WinActive(" NHLS")
    WinGetTitle, last_window_title, A
    MouseGetPos, xmousepos, ymousepos
    WinActivate,  NHLS
    WinWaitActive,  NHLS
    Sleep, 50
    send, {Alt}
    sleep, 50
    Send, v
    sleep, 50
    send, u
    WinWaitActive, User Audit Trail
    sleep, 100
    send, {Enter}
    sleep, 50
    WinActivate, %last_window_title%
    MouseMove, %xmousepos%, %ymousepos%
    sleep, 100
    WinMinimize,  NHLS
    sleep, 500
  ToolTip, Trak is being kept open automatically. `n Hit ESC to stop.   ;This tooltip doesn't want to work well for some reason
  SetTimer, RemoveToolTip, -1000
  sleep, 1000
}
return


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                           Alt+N Shortcut to copy Episode
!n::
if WinActive("Medical Validation :   (Authorise By Episode)")
    {
    episode := EpMedVal()
    Clipboard := episode
    Return
    }
else
    if WinActive("Result Entry - Single -")
    {
    episode:=EpResultSingle()
    Clipboard := episode
    Return
    }
else
    if WinActive("Result Verify - Single -")
    {
    episode:=EpResultVerSingle()
    Clipboard := episode
    Return
    }
else
    if WinActive("Result Entry - Verify")
    {
        episode:=EpResEntrVer()
        Clipboard := episode
        Return
    }
else
    if WinActive("Result Entry (Total Count")
    {
        episode:=EpResEntrTot()
        Clipboard := episode
        Return
    }
else
    MsgBox, No relevant TrakCare Window active.
    Return

^!n:: ;                                                                                 Ctrl+Alt+N or Ctrl+Alt+L to log episode in a text file - useful for logging of episodes.
^!l::
log_episode:
if WinActive("Medical Validation :   (Authorise By Episode)")
    {
    episode := EpMedVal()
    LogEpisode(episode)
    }
else
    if WinActive("Result Entry - Single -")
    {
    episode := EpResultSingle()
    LogEpisode(episode)
    }
else
    if WinActive("Result Entry - Verify")
    {
    EpResEntrVer()
    episode := Clipboard
    LogEpisode(episode)
    }
else
    if WinActive("Result Verify - Single -")
    {
    EpResultVerSingle()
    episode := Clipboard
    LogEpisode(episode)
    }
else
    if WinActive("Result Entry (Total")
    {
    EpResEntrTot()
    episode:= Clipboard
    LogEpisode(episode)
    }
else
    MsgBox,1,, No relevant TrakCare Window active.,3
    Return

LogEpisode(episode)
{
    FormatTime, timevar , YYYYMMDDHH24MISS, yyyy-MM-dd HH:mm
    FileAppend, %episode%`,%timevar%`n, %A_ScriptDir%\episode_log.txt
    if ErrorLevel
        MsgBox, Unable to write to file Documents\ChemHelp\episode_log.txt.  Make sure the directory exists in the "My Documents" folder.
    else
        MsgBox, 1,,Episode %episode% logged at %timevar% `n at ChemHelp-directory\episode_log.txt,1.8
}

;-------------------------                                                                  Alt+P Enter PHONC
ButtonPHONC:
WinActivate, Medical Validation :   (Authorise By Episode)

#IfWinActive, Medical Validation :   (Authorise By Episode)
!p::
episode := EpMedVal()
sleep, 200
WinActivate Medical Validation :   (Authorise By Episode)
sleep, 200
   send, {altdown}<{altup}
sleep, 200

if !WinExist("Result Entry, , Result Entry - Verify")
{
	If !WinActive(" NHLS,")
    WinActivate,  NHLS,
	WinWaitActive,  NHLS,
	Sleep, 100
	Send, a{TAB}
	Send, Result Entry{Enter}
	WinWaitActive, Result Entry
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 100
if WinExist("Search Unsuccessful")
{
    WinClose, Search Unsuccessful
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 200
send, {AltDown}l{AltUp}
sleep, 200
send, %episode%
sleep, 200 ; changed down from 300
send, {Enter}
sleep, 150 ; changed down from 300
MouseClick, Left, 340, 133, 2, 100
sleep, 100
Send, PHONC
sleep, 100
send, {Enter}
sleep, 400
if WinExist("Search Unsuccessful")
{
    MsgBox, 4,,  PHONC does not appear for this Episode. `nDo you want to add a PHONC to this episode?
    IfMsgBox Yes
    {
        WinClose, Search Unsuccessful
        sleep, 200
        WinActivate, Result Entry, , Result Entry - Verify
        WinWaitActive, Result Entry, , Result Entry - Verify
        sleep, 200
        send, {AltDown}l{AltUp}
        sleep, 400
        send, {CtrlDown}v{CtrlUp}
        sleep, 300
        send, {Enter}
        sleep, 300
        MouseClick, Left, 215, 300
        sleep, 200
        send, {altdown}d{altup}
        sleep, 200
        send, s
        sleep, 200
        WinWaitActive, Test Set Maintenance
        sleep, 200
        send, PHONC
        sleep, 200
        send, {tab}
        sleep, 100
        send, {enter}
        return
    }
    else IfMsgBox No
    {
        sleep, 300
        WinClose, Search Unsuccessful
        Return
    }
}
else
MouseClick, left, 70, 300, 1
sleep, 200
send, {AltDown}e{AltUp}
WinWaitActive, Result Entry - Single
sleep, 200
send, t{tab}
sleep, 100
send, n{tab}
sleep, 100
send, {tab}
SendInput, %full_name%
send, {ShiftDown}{tab}{ShiftUp}

return
#IfWinActive

;-------------------------                                                                  Ctrl + Alt + P Enter PHONC

#IfWinActive, Medical Validation :   (Authorise By Episode)
^!p::
episode := EpMedVal()
sleep, 200
WinActivate Medical Validation :   (Authorise By Episode)
sleep, 200
   send, {altdown}<{altup}
sleep, 200

if !WinExist("Result Entry, , Result Entry - Verify")
{
    If !WinActive(" NHLS,")
    WinActivate,  NHLS,
    WinWaitActive,  NHLS,
    Sleep, 100
    Send, a{TAB}
    Send, Result Entry{Enter}
    WinWaitActive, Result Entry
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 100
if WinExist("Search Unsuccessful")
{
    WinClose, Search Unsuccessful
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 200
send, {AltDown}l{AltUp}
sleep, 200
send, %episode%
sleep, 200 ; changed down from 300
send, {Enter}
sleep, 150 ; changed down from 300
MouseClick, Left, 340, 133, 2, 100
sleep, 100

return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                            Ctrl+Alt+R Enter CCOM with Dr. update links


#IfWinActive, Medical Validation :   (Authorise By Episode)
^!r::
episode := EpMedVal()
sleep, 100
WinActivate Medical Validation :   (Authorise By Episode)
sleep, 200
   send, {altdown}<{altup}
sleep, 200

IfWinNotExist, Result Entry, , Result Entry - Verify
{
	IfWinNotActive,  NHLS,
    WinActivate,  NHLS,
	WinWaitActive,  NHLS,
	Sleep, 100
	Send, a{TAB}
	Send, Result Entry{Enter}
	WinWaitActive, Result Entry
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 100
IfWinExist, Search Unsuccessful
{
    WinClose, Search Unsuccessful
    sleep, 200
}
WinActivate, Result Entry, , Result Entry - Verify
WinWaitActive, Result Entry, , Result Entry - Verifiy
sleep, 150
send, {AltDown}l{AltUp}
sleep, 150
send, %episode%
sleep, 200
send, {Enter}
sleep, 200

MouseClick, left, 70, 300, 1
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 200
WinWaitActive, Test Set Maintenance
sleep, 500
send, CCOM
sleep, 500
send, {tab}
sleep, 300
send, {enter}
sleep, 300
WinClose, Test Set Maintenance
WinWaitClose, Test Set Maintenance
sleep, 500

MouseClick, Left, 340, 133, 2, 100
sleep, 100
Send, CCOM
sleep, 100
send, {Enter}
sleep, 250

IfWinExist, Search Unsuccessful
{
    MsgBox, 4,,  CCOM does not appear for this Episode. `nDo you want to add a CCOM to this episode?
    IfMsgBox Yes
    {
        WinClose, Search Unsuccessful
        sleep, 200
        WinActivate, Result Entry, , Result Entry - Verify
        WinWaitActive, Result Entry, , Result Entry - Verify
        sleep, 200
        send, {AltDown}l{AltUp}
        sleep, 400
        send, {CtrlDown}v{CtrlUp}
        sleep, 300
        send, {Enter}
        sleep, 300
        MouseClick, Left, 215, 300
        sleep, 200
        send, {altdown}d{altup}
        sleep, 200
        send, s
        sleep, 200
        WinWaitActive, Test Set Maintenance
        sleep, 150
        send, CCOM
        sleep, 150
        send, {tab}
        sleep, 100
        send, {enter}
        return
    }
    else IfMsgBox No
    {
        sleep, 300
        WinClose, Search Unsuccessful
        Return
    }
}
else

MouseClick, Left, 400, 300, 1
sleep, 150
send, {AltDown}e{AltUp}
WinWaitActive, Result Entry - Single -
sleep, 100
send, {F6}
WinWaitActive, Comments
sleep, 50
send, Clinician contact details may not be coded in our database.  Please follow the link below to update clinician contact details:`ntinyurl.com/nhls-update       ;<<-----Put link in here after the "`n"
sleep, 1800
WinClose, Comments
sleep, 100
send, {AltDown}a{AltUp}
sleep, 800
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation
sleep, 200
send, {AltDown}>{AltUp}

return

#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                 ALT+W for EPR

ButtonEPR:
send, {AltDown}{Tab}{AltUp}
sleep, 200

!w::
sleep, 300
IfWinActive, Result Entry - Single -
{
MRNSingle()
txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
sleep, 300
Remove_IE_notice()
return
}

else IfWinActive, Medical Validation :   (Authorise By Episode)
{
mrn := MRNMedVal()
;txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . mrn
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
sleep, 300
Remove_IE_notice()
return
}

else IfWinActive, Result Verify - Single -
{
MRNResultVerSingle()
txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
sleep, 300
Remove_IE_notice()
return
}

else IfWinActive, Result Entry (Total
{
    MRNResultEntTotal()
    txt := Clipboard
    url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
    run, %url%
    sleep, 800
    WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
    sleep, 300
    Remove_IE_notice()
    return
}

else IfWinActive, Result Entry - Verify
{
    MRNResultEntVer()
    txt := Clipboard
    url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
    run, %url%
    sleep, 800
    WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
    sleep, 300
    Remove_IE_notice()
    return
}

else
{
MsgBox, No TrakCare window active.  `nPlease ensure Result Entry or Result Verify windows are open and visible.
sleep, 200
return
}
Return

;                                                         Alt+F for Form
ButtonForm:
sleep, 200
send, {AltDown}{Tab}{AltUp}
sleep, 200

!F::
;sleep, 300
IfWinActive, Medical Validation :   (Authorise By Episode)
{
episode:=EpMedVal()
;txt := ClipBoard
;sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . episode . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
}
else
IfWinActive,  Result Entry - Single -
{
episode := EpResultSingle()
txt := ClipBoard
;sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . episode . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
}
else
IfWinActive, Result Verify - Single -
{
episode := EpResultVerSingle()
txt := ClipBoard
url := "http://172.22.4.40/multipdfsearch.php?file=" . episode . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return
}
else
IfWinActive, Result Entry - Verify
    {
        episode:=EpResEntrVer()
        url := "http://172.22.4.40/multipdfsearch.php?file=" . episode . ""
        run, %url%
        sleep, 1500
        WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
        Return
    }
else
IfWinActive, Result Entry (Total Count
    {
        EpResEntrTot()
        txt := ClipBoard
        url := "http://172.22.4.40/multipdfsearch.php?file=" . txt . ""
        run, %url%
        sleep, 1500
        WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
        Return
    }
else
{
MsgBox, No TrakCare window active.  `nPlease ensure Result Entry or Result Verify windows are open and visible.
sleep, 200
send, {AltDown}e{AltUp}
return
}

iehalfscreen:
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
return



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                          Button FPSA
ButtonFPSA:
^!f::
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 1000
WinActivate, Test Set Maintenance
WinWaitActive, Test Set Maintenance
sleep, 500
ImageSearch, FoundX, FoundY, 15,86, 523, 379, %A_ScriptDir%\trakcare_icons\cdum.png
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for CDUM. Possibly the file cdum.png is missing in ChemHelp\trakcare_icons folder.
else if (ErrorLevel = 1)
    MsgBox CDUM could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x%FoundY%.
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
    send, {AltDown}d{AltUp}
sleep, 300
WinActivate, Test Set Maintenance
sleep, 100
MouseClick, Left, 830, 308
sleep, 100
send, FPSA
sleep, 200
send, {tab down}{tab up}
send, {enter}
sleep, 200
send, {AltDown}o{AltUp}
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                          Button dCDUM
ButtondCDUM:
^!d::
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, s
sleep, 200
WinWaitActive, Test Set Maintenance
sleep, 800
ImageSearch, FoundX, FoundY, 15,86, 523, 379, %A_ScriptDir%\trakcare_icons\cdum.png
if (ErrorLevel = 2)
    MsgBox Could not conduct the search for CDUM. `nPossibly the file cdum.png is missing in ChemHelp\trakcare_icons folder.`n If this is the issue, find it on github.com/dietervdwes/chemhelp
else if (ErrorLevel = 1)
    MsgBox CDUM could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x%FoundY%.
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
    send, {AltDown}d{AltUp}
sleep, 300
WinClose, Test Set Maintenance
;MsgBox, Please confirm that CDUM2 has been removed.
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                          Button rFix
ButtonrFix:
send, {AltDown}{Escape}{AltUp}
#IfWinActive Result Entry - Single -
    send, {AltDown}u{AltUp}
    sleep, 1200
    send, {AltDown}2{AltUp}
    sleep, 300
    send, s
    sleep, 300
    WinActivate, Test Set Maintenance
    sleep, 300
    MouseClick, Left, 830, 308
    sleep, 300
    send, IFE
    sleep, 300
    send, {tab down}{tab up}
    send, {enter}
    sleep, 700
    WinClose, Test Set Maintenance
    sleep, 1000
    send, {Alt}
    sleep, 700
    send, 2
    sleep, 700
    send, n
    sleep, 700
    send, {AltDown}1{AltUp}
    sleep, 700
    send, f
    sleep, 700
    WinWaitActive, Refer To Verification Queue
    WinActivate, Refer To Verification Queue
    sleep, 700
    MouseClick, Left, 429, 111
    sleep, 700
    MouseClick, Left, 429, 165
    sleep, 700
    send, {AltDown}o{AltUp}
    sleep, 700
    WinClose, Result Entry - Single
    return
#IfWinActive
MsgBox, Result Entry - Single window is not active.
return
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                               Button Verified to insert staff note "Transcription Verified"
ButtonVerified:
sleep, 200
send, {AltDown}{Escape}{AltUp}
sleep, 200
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                Alt+V for Staff note in Med Val or Single: Transcription Verified

!v::
IfWinActive, Medical Validation :   (Authorise By Episode)
{
;WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 100
send, {altdown}d{altup}
sleep, 100
send, f
WinWaitActive, Staff Notes
sleep, 100
send, Transcription Verified.
sleep, 100
send, {altdown}o{altup}
sleep, 50
Return
}
else
IfWinActive, Result Verify - Single
{
WinActivate, Result Verify - Single
WinWaitActive, Result Verify - Single
sleep, 200
MouseClick, left, 904, 164, 1
sleep, 500
MouseClick, left, 19, 292, 1
sleep, 500
send, Transcription Verified.
sleep, 50
send, {altdown}o{altup}
sleep, 50
Return
}
else
IfWinActive, Result Entry - Single
{
WinActivate, Result Entry - Single
WinWaitActive, Result Entry - Single
sleep, 200
MouseClick, left, 904, 164, 1
sleep, 500
MouseClick, left, 19, 292, 1
sleep, 500
send, Transcription Verified.
sleep, 50
send, {altdown}o{altup}
sleep, 50
Return
}
else
return


;;;;;;;;;;;;;;;;;;;                                                                             Alt+/ for Staff note in Med Val or Single: Transcription Verified
!/::
IfWinActive, Medical Validation :   (Authorise By Episode)
{
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)
sleep, 200
send, {altdown}d{altup}
sleep, 200
send, f
sleep, 200
send, Not yet scanned.
sleep, 100
send, {alt down}o{alt up}
sleep, 50
Return
}
else
IfWinActive, Result Verify - Single
{
WinActivate, Result Verify - Single
WinWaitActive, Result Verify - Single
sleep, 200
MouseClick, left, 904, 164, 1
sleep, 500
MouseClick, left, 19, 292, 1
sleep, 500
send, Not yet scanned.
sleep, 50
send, {alt down}o{alt up}
sleep, 50
Return
}
else
{
return
}

;;;;;;;;;;;;;;;;;;;;;;                                                                                     Page Down when Alt+z is pressed.
#IfWinActive Medical Validation :   (Authorise By Episode)
!z::
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgDn down}{PgDn up}
Return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;                                                                                                     Page Up when Alt+q is pressed.
/*
#IfWinActive Medical Validation :   (Authorise By Episode)
!q::
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
Return
#IfWinActive
*/
;;
!q::
   If WinActive("Result Entry - Single - ")
   {
   mouseclick, left, 78,  209, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
   Return
   }
   If WinActive("Cumulative View - ")
   {
    mouseclick, left, 78,  209, 1
   sleep, 50
   send, {PgUp down}{PgUp up}
   sleep, 150
   Return
   }
   Else
   {
   send, {PgUp}
    Return
   }
!z::
   If WinActive("Result Entry - Single - ")
   {
   mouseclick, left, 598, 214, 1
   sleep, 50
   send, Y
   sleep, 100
   MouseClick, left, 254, 511
   sleep, 50
   send, {F6}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   mouseclick, left, 433,  231, 1
   sleep, 50
   send, {PgDn down}{PgDn up}
   Return
   }
   If WinActive("Cumulative View - ")
   {
    mouseclick, left, 78,  209, 1
   sleep, 50
   send, {PgDn down}{PgDn up}
   sleep, 150
   Return
   }
   Else
   {
   send, {PgDn}
    Return
   }


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                   Alt+B for Back
!b::
   If WinActive("Result Entry - Single - ")
   {
   send, {altdown}<{altup}
   sleep, 80
   send, {Enter}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   send, {altdown}<{altup}
   sleep, 500
   Return
   }
   If WinActive("Result Verify - Single - ")
   {
   send, {altdown}<{altup}
    sleep, 80
   send, {Enter}
   sleep, 150
   Return
   }
   Else
   {
   MsgBox, No TrakCare Window selected.
    Return
   }


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                   Alt+X for Next
!x::
   If WinActive("Result Entry - Single - ")
   {
   send, {altdown}>{altup}
   sleep, 80
   send, {Enter}
   sleep, 150
   Return
   }
   If WinActive("Medical Validation :   (Authorise By Episode) - \\Remote")
   {
   send, {altdown}>{altup}
   sleep, 500
   Return
   }
    If WinActive("Result Verify - Single - ")
   {
   send, {altdown}>{altup}
    sleep, 80
   send, {Enter}
   sleep, 150
   Return
   }
   Else
   {
   MsgBox, No TrakCare Window selected.
    Return
   }

;                                                                                                                  Alt+I to invalidate a Potassium
#IfWinActive Result Entry
!i::
  send, {AltDown}d{AltUp}
  sleep, 150
  send, a
  sleep,100
  WinWaitActive, Assign Reason
  sleep, 80
  loop, 9
    {
    send i
    sleep, 70
    }

Return
#IfWinActive

;                                                                                                                  Alt+C to cancel a PHONC
#IfWinActive Result Entry
!c::
  send, {AltDown}d{AltUp}
  sleep, 150
  send, a
  sleep,100
  WinWaitActive, Assign Reason
  sleep, 80
  loop, 10
    {
    send c
    sleep, 70
    }

Return
#IfWinActive

;                                                                                                              Alt+` to type username and password
!`::
 send, %citrix_username%
 send,{tab}
 send, %citrix_password%
 send,{Enter}
 Return


;                                                                                              Mouse Wheel Actions in Verify   UP and DOWN SCROLL
#IfWinActive Medical Validation :   (Authorise By Episode)

WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-17 , winh-41, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-17 , 200, 1
Return
#IfWinActive

#IfWinActive Result Entry - Verify
WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , winh-26, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , 243, 1
Return
#IfWinActive

#IfWinActive Result Entry
WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , winh-26, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-125 , 243, 1
Return
#IfWinActive

#IfWinActive Patient Audit -
WheelDown::
WinGetPos,,,winw, winh
mouseclick, left, winw-22 , 581, 1
Return

WheelUp::
WinGetPos,,,winw, winh
mouseclick, left, winw-22 , 80, 1
Return
#IfWinActive

#IfWinActive Visit Events -
WheelDown::
MouseClick, left, 765 , 305, 1
sleep, 50
Return

WheelUp::
mouseclick, left, 765 , 65, 1
sleep, 50
Return
#IfWinActive

#IfWinActive Test Set Maintenance -
WheelDown::
mouseclick, left, 799 , 367
return

WheelUp::
mouseclick, left, 799 , 62
return
#IfWinActive

#IfWinActive Comment Text -
WheelDown::
mouseclick, left, 547 , 272, 1
Return

WheelUp::
mouseclick, left, 547 , 65, 1
Return
#IfWinActive

#IfWinActive Cumulative View -
WheelDown::
mouseclick, left, 993 , 670, 1
Return

WheelUp::
mouseclick, left, 993 , 208, 1
Return
#IfWinActive

;                                                        ALT-RIGHT CLICK ALT MENU: cobbled together somehow, by Samuel Murray - 26 May 2018, edited by Dieter (18/09/2019)
;;;;;;                                    for testing: Episode Number: SA01912457

!RButton::

MouseGetPos, xpos, ypos, window

Gui, New , , ChemHelp

Gui, Add, Button, gScript1, [&1] Open in Patient Entry
Gui, Add, Button, gScript2, [&2] Open in Result Entry
Gui, Add, Button, gScript3, [&3] Open in Specimen Info
Gui, Add, Button, gScript4, [&4] Open form on Equation
Gui, Add, Button, gExitApp, Reload

Gui, -Border
Gui, Show, x%xpos% y%ypos%
return

; Test Episode Number: SA01912457
Script1:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500

IfWinExist, Patient Entry
{
	WinActivate, Patient Entry
	sleep, 500
	send, {altdown}c{altup}
	sleep, 2000
	send, {enter}
	sleep, 2000
}

Else IfWinNotExist, Patient Entry
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS,
	WinWaitActive,  NHLS,
	Sleep, 100
	Send, pa{ENTER}
	WinWaitActive, Patient Entry
}

sleep, 500
send, {CtrlDown}v{CtrlUp}
sleep, 200
send, {TAB}
return


Script2:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500

IfWinExist, Result Entry
{
	WinActivate, Result Entry
	sleep, 500
	send, {altdown}l{altup}
	sleep, 300
}

Else IfWinNotExist, Result Entry
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS,
	WinWaitActive,  NHLS,
	Sleep, 100
	Send, a{TAB}
	Send, Result Entry{Enter}
	WinWaitActive, Result Entry
}

sleep, 500
send, {CtrlDown}v{CtrlUp}
sleep, 200
send, {Enter}
return


Script3:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500

IfWinExist, Specimen Information
{
	WinActivate, Specimen Information
	sleep, 500
	MouseClick, Left, 77, 82, 2, 100
	sleep, 500
}

Else IfWinNotExist, Specimen Information
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS,
	WinWaitActive,  NHLS,
	Sleep, 100
	Send, a{TAB}
	Send, Specimen Information{Enter}
	WinWaitActive, Specimen Information
}

sleep, 500
send, {CtrlDown}v{CtrlUp}
sleep, 200
send, {Tab}
return

Script4:
Gui, Hide
sleep, 600
send, {CtrlDown}c{CtrlUp}
sleep, 500
txt := Clipboard
sleep, 300
url := "http://" . equation_ip . "/multipdfsearch.php?file=" . txt . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
Return

ExitApp:
Reload
return




;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                       This section defines the different FUNCTIONS

EpMedVal()
{
WinActivate, Medical Validation :   (Authorise By Episode)
WinWaitActive, Medical Validation :   (Authorise By Episode)

send, {altdown}te{altup}
send, {Enter}
sleep, 200
WinWaitActive, Visit Events
if ErrorLevel
    MsgBox, Visit Events window could not be found.
else
sleep,50
WinGetTitle, Title, A
;MsgBox, The active window is "%Title%".
rawarray := StrSplit(Title , "-", " " )
episode := rawarray[2]
;MsgBox, The mrn is "%episode%".
WinClose, %Title%
WinWaitClose, %Title%
ToolTip, %episode%
SetTimer, RemoveToolTip, -2000
return episode
}


EpResEntrTot()
{
WinActivate, Result Entry (Total
WinWaitActive, Result Entry (Total
sleep, 100 ; changed this from 250 (2022-01-12)
send, {altdown}i{altup}
sleep, 100 ; changed this down from 200
send, c{Enter}
sleep, 50 ;changed this down from 100
WinActivate, Clinical History
WinWaitactive, Clinical History
sleep, 100 ; changed this down from 100
MouseClick, left, 122, 75, 2, 100
clipboard=
sleep, 100 ;changed down from 100
send, ^c
ClipWait
sleep, 100
txt := Clipboard
sleep, 100 ; changed this down from 100
WinClose, Clinical History
WinWaitClose, Clinical History
sleep, 50
}

EpResEntrVer()
{
WinActivate, Result Entry - Verify
WinWaitActive, Result Entry - Verify
send, {altdown}tt{altup}
send, {Enter}
send, s
sleep, 200
WinWaitActive, Visit Events
if ErrorLevel
{
    MsgBox, Visit Events window could not be found.
    return
}
else
sleep,50
WinGetTitle, Title, A
;MsgBox, The active window is "%Title%".
rawarray := StrSplit(Title , "-", " " )
episode := rawarray[2]
;MsgBox, The mrn is "%episode%".
WinClose, %Title%
WinWaitClose, %Title%
ToolTip, %episode%
SetTimer, RemoveToolTip, -1500
return episode
}

EpResultSingle()
{
WinActivate, Result Entry - Single -
WinWaitActive, Result Entry - Single -
WinGetTitle, Title, A
;MsgBox, The active window is "%Title%".
rawarray := StrSplit(Title , "-", " " )
episode := rawarray[3]
;MsgBox, The mrn is "%episode%".
ToolTip, %episode%
SetTimer, RemoveToolTip, -1500
return episode
}

EpResultVerSingle()
{
WinActivate, Result Verify - Single -
WinWaitActive, Result Verify - Single -
WinGetTitle, Title, A
;MsgBox, The active window is "%Title%".
rawarray := StrSplit(Title , "-", " " )
episode := rawarray[3]
;MsgBox, The mrn is "%episode%".
ToolTip, %episode%
SetTimer, RemoveToolTip, -1500
return episode
}

MRNSingle()
{
    WinActivate, Result Entry - Single -
WinWaitActive, Result Entry - Single -
sleep, 500 ; changed from 500
send, {AltDown}3
sleep, 300
send, {AltUp}
sleep, 200 ; changed from 500
send, {enter}
sleep, 100 ; changed from 500
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 100 ; changed from 200
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100 ; changed from 100
send, ^c
ClipWait
sleep, 100 ; changed from 200
txt := Clipboard
sleep, 100 ; changed from 200
WinClose, Patient History
WinWaitClose, Patient History
ToolTip, %Clipboard%
SetTimer, RemoveToolTip, -1500
sleep, 50
}

MRNMedVal()
{
mouseclick, left, 433,  231, 1
sleep,150
send, {F7}
WinWaitActive, Test Item Results
if ErrorLevel
    MsgBox, "Test Item Results" window could not be found.
else
sleep,150
WinGetTitle, Title, A
;MsgBox, The active window is "%Title%".
rawarray := StrSplit(Title , "-", " " )
mrn := rawarray[2]
;MsgBox, The mrn is "%mrn%".
WinClose, %Title%
WinWaitClose, %Title%
ToolTip, %mrn%
SetTimer, RemoveToolTip, -1500
return mrn
}

MRNResultVerSingle()
{
WinActivate, Result Verify - Single -
WinWaitActive, Result Verify - Single -
sleep, 200
send, {Altdown}3
sleep, 300
send, {enter}
send, {Altup}
sleep, 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 200
send, ^c
ClipWait
sleep, 200
txt := Clipboard
sleep, 200
WinClose, Patient History
WinWaitClose, Patient History
ToolTip, %Clipboard%
SetTimer, RemoveToolTip, -1500
sleep, 50
}

MRNResultEntTotal()
{
sleep, 200 ;changed down from 500
send, {altdown}i{altup}
sleep, 200 ; changed down from 500
send, {enter}
sleep, 50 ; changed down from 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 50 ; changed down from 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100 ; changed down from 100
send, ^c
ClipWait
sleep, 50 ; changed down from 200
txt := Clipboard
sleep, 50 ; changed down from 200
WinClose, Patient History
WinWaitClose, Patient History
ToolTip, %Clipboard%
SetTimer, RemoveToolTip, -1500
sleep, 50
}

MRNResultEntVer()
{
sleep, 200 ;changed down from 500
send, {altdown}i{altup}
sleep, 200 ; changed down from 500
send, {enter}
sleep, 50 ; changed down from 200
WinActivate, Patient History
WinWaitactive, Patient History
sleep, 50 ; changed down from 300
mouseclick, left, 93, 75, 2, 100
sleep, 50
clipboard=
sleep, 100 ; changed down from 100
send, ^c
ClipWait
sleep, 50 ; changed down from 200
txt := Clipboard
sleep, 50 ; changed down from 200
WinClose, Patient History
WinWaitClose, Patient History
ToolTip, %Clipboard%
SetTimer, RemoveToolTip, -1500
sleep, 50
}

Remove_IE_notice() {
    sleep, 200
    ImageSearch, FoundX, FoundY, 700,700, 1300, 900, %A_ScriptDir%\trakcare_icons\show_all_content.png
if (ErrorLevel = 2)
    ;MsgBox Could not conduct the search for show_all_content notice. Possibly the file show_all_content.png is missing in ChemHelp\trakcare_icons folder.
    return
else if (ErrorLevel = 1)
    return
    ;MsgBox Show_all_content could not be found on the screen.
else
    ;MsgBox The icon was found at %FoundX%x%FoundY%.
    FoundX += 142
    FoundY += 15
    MouseClick, Left, %FoundX%, %FoundY%,
    sleep, 500
}

;---------------------------------------------------------- Script to Loop Extraction by pre-defined Extracting Criteria with a pre-formatted configuration CSV.
#IfWinActive, Monthly Statistics (Not Responding) - \\Remote
!e::
settimer, extraction_watch, 10000
#IfWinActive

extraction_watch:
if WinExist("Monthly Statistics (Not Responding) - \\Remote")
{
    ToolTip, Waiting for Monthly Statistics Window to finish extraction
    sleep, 500
    ToolTip
    Return
}

if !WinExist("Monthly Statistics (Not Responding) - \\Remote")
{
    Extract()
    return
}
Return


ButtonEPRxlsx:
;--------------------------------Saving the patient's cumulative history in an Excel file: requires Node installed with Puppeteer module and the SavePatientEPR.js script.

send, {AltDown}{Tab}{AltUp}
sleep, 200

!+w::
sleep, 300
IfWinActive, Result Entry - Single -
{
MRNSingle()
txt := Clipboard
run node %A_ScriptDir%\SavePatientEPR.js %txt%
if (ErrorLevel = "ERROR")
    MsgBox There was an error, likely Node is not installed.
else
return
}

else IfWinActive, Medical Validation :   (Authorise By Episode)
{
mrn := MRNMedVal()
run node %A_ScriptDir%\SavePatientEPR.js %mrn%
;Make sure Node.js is installed and that the SavePatientEPRs.js file, the node_modules folder, package.json and package-lock.json is present in "My Documents" folder.
if (ErrorLevel != 0)
    MsgBox There was an error, likely Node is not installed.
else
return
}

else IfWinActive, Result Verify - Single -
{
MRNResultVerSingle()
txt := Clipboard
run node %A_ScriptDir%\SavePatientEPR.js %txt%
if (ErrorLevel = "ERROR")
    MsgBox There was an error, likely Node is not installed.
else
return
}

else
{
MsgBox, No TrakCare window active.  `nPlease ensure Result Entry or Result Verify windows are open and visible.
sleep, 200
return
}
Return



ButtonEx:
Extract()
Return

Extract()    {
FileAppend, time_started`,begin_date`,end_date`,loc_prov`,test_set_test_item`,email`,non_reportable`,sort_direction`n, extract_list.txt
settimer, extraction_watch, Off
settimer, refreshtimer, Off
Loop, read, %A_ScriptDir%\extract_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{
    LoopNumber := A_Index ;A_Index stores the number of the current look, i.e. in the case of 'loop, read' it will be the line number, as AHK reads files line by line.
    Looplinecontents := A_LoopReadLine ;This code stores the contents of the current line into the variable "Looplinecontents" or %Looplinecontents%.
    LineArray := StrSplit(Looplinecontents, ",")
    ;This portion configures which columns is present in the CSV
    begin_date := LineArray[1]
    end_date := LineArray[2]
    loc_prov := LineArray[3]
    branch := LineArray[4]
    user_site := LineArray[5]
    test_set_test_item := LineArray[6]
    email := LineArray[7]
    non_reportable := LineArray[8]
    sort_direction := LineArray[9]
  sleep, 2000
   if !WinExist("Monthly Statistics") ;The exclamation mark before the WinExist specifies the negative of the statement, i.e. if win"doesnot"exist, execute the function below it in curly brackets.
    {
        MsgBox, Monthly Statistics Window does not exist.,4 ; Issue here needs to be resolved as sometimes it doesn't see the window, even though it's active.
        Return
    }

    if WinExist("Monthly Statistics (Not Responding)")
    {
        MsgBox, Monthly Statistics Window is "Not Responding" and likely busy extracting.  `nThe timer will now be set so that it can continue extraction when it starts responding.`nHit Esc to abort this process.
        while WinExist("Monthly Statistics (Not Responding)") {
            sleep, 10000
            ToolTip, Extracting...
            sleep, 500
            ToolTip,
        }
    }
  WinActivate, Monthly Statistics - \\Remote
  WinWaitActive, Monthly Statistics - \\Remote
  sleep, 2000
  MouseClick, Left, 18, 87
  sleep, 500
  send, results
  sleep, 500
  send, {Enter}
  sleep, 3000
  WinWaitActive, Report Listing - \\Remote
  sleep, 3000
  PixelGetColor, color, 178, 142
  While !(color = 0xFFFF00)
        {
        PixelGetColor, color, 178, 142
        ToolTip, Color is %color%
        sleep, 300
        Tooltip
        }
  sleep, 1000
  send, %begin_date%
  sleep, 1000
  send, {Tab}
  sleep, 500
  send, %end_date%
  sleep, 500
  send, {Tab}
  sleep, 500
  send, %loc_prov% ; Location code (province)
  sleep, 500
  send, {Down}
  sleep, 1000
  send, {Tab}
  sleep, 500
  send, %branch%
  send, {Down} ;Uncomment this if branch is present
  sleep, 1000
  send, {Tab}
  sleep, 500
  send, %user_site%
  send, {Down}  ;Uncomment this if province is present
  sleep, 1000
  send, {Tab} ;Uncomment this tab if above user site is present.
  sleep, 500
  loop, 6
    {
        send, {Tab}
        sleep, 200
    }
  send, %test_set_test_item%
  sleep, 200
  send, {Tab}
  sleep, 500
  #IfWinActive, CLNSTATLIST
  sleep, 1000
  send, {Enter}
  WinWaitClose, CLNSTATLIST
  #IfWinActive
  sleep, 1000
  MouseClick, Left, 470, 533
  sleep, 500
  send, {Tab}
  sleep, 500
  send, %email%
  sleep, 500
  send, {Tab}
  sleep, 500
  if (non_reportable = 1) {
    send, {space}
    sleep, 1000
    send, {Enter}
    sleep, 500
    send, {Tab}
    }
  else {
    send, {Tab}
       }
  sleep, 200
  send, {Tab}
  sleep, 500
  if (sort_direction = 1) {
    send, {Space}
    sleep, 500
    send, {Tab}
    sleep, 500
    }
  else {
    send, {Tab}
    sleep, 500
    send, {Space}
    sleep, 500
    }
   sleep, 500
   /*
   MsgBox, 4, , Please confirm that the correct parameters are entered. `nThis MsgBox will *auto-destruct* in 007 seconds and continue extracting.  Continue?, 7
        IfMsgBox No
            Return
        else IfMsgBox Timeout
            sleep, 500
  */
  WinActivate, Report Listing
  WinWaitActive, Report Listing
  sleep, 800
  MouseClick, Left, 808, 59 ; Click on "Print" button
  FileAppend, %A_Now%`,%begin_date%`,%end_date%`,%loc_prov%`,%test_set_test_item%`,%email%`,%non_reportable%`,%sort_direction%`n, %A_ScriptDir%\extract_list.txt
  sleep, 10000
  settimer, refreshtimer, 600000 ; This is to keep TrakCare open
  While WinExist("Monthly Statistics (Not Responding)")
  {
    Sleep, 5000
    ToolTip, Extracting
  }
  settimer, refreshtimer, off ; This is to prevent the refreshtimer from interferring in the script.
  sleep, 5000
}
settimer, refreshtimer, 600000
  tooltip    ; remove the tooltip
}
Return

ButtonAutoA: ;Feature hidden by default. Use this with extreme caution - it will auto-authorize samples. Useful for a batch processing but shouldn't be used if you don't know what you're doing.
MsgBox, Do not use this feature if you don't know what you're doing.  Hit ESC immediately if you don't know what this feature is.
WinActivate, Medical Validation
WinWaitActive, Medical Validation
loop, 100
    {
        Send, {AltDown}a{AltUp}
        sleep, 800
        PixelGetColor, color, 105, 229
        While !(color = 0xD1B499)
                {
                PixelGetColor, color, 105, 229
                ToolTip, Color is %color%
                sleep, 200
                Tooltip
                }
        sleep, 1500
    }
Return

;;;;;;;;;;;---------------------------------    How to open phoresis

^+2::
run, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Phoresis\Phoresis.lnk"
Sleep, 5000
sleep, 7500
Send, sebia
send, {enter}
sleep, 2000
send, {altdown}dh
send, {enter}
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                   SPE canned text
::NORM::
SetKeyDelay, 150
send, The SPE is essentially normal. No monoclonal is evident.`n`n
send, If the clinical suspicion of myeloma remains, please send a minimum of 20mL of urine for Bence Jones protein determination with sodium azide preservative (containers available from C17 lab), or serum for free light chain analysis.

::INFLAM::Albumin is decreased. Alpha-1, alpha-2 and beta-2 are increased. There is a polyclonal hypergammaglobulinaemia present measuring  g/L (8 - 14 g/L). There is no monoclonal evident. `n`nThis pattern is in keeping with an inflammatory process. `n`nIf the clinical suspicion of myeloma remains, please send a minimum of 20mL of urine for Bence Jones protein determination with sodium azide preservative (containers available from C17 lab), or serum for free light chain analysis.
::IRREG::Albumin is decreased. Alpha-1, alpha-2 and beta-2 are increased. There is a polyclonal hypergammaglobulinaemia present measuring  g/L (8 - 14 g/L). There is a small irregularity in  gamma (<1 g/L). `n`nThis pattern is in keeping with an inflammatory process. The clinical significance of the small irregularity is uncertain but likely reflects the inflammatory/reactive background. `n`nIf the clinical suspicion of myeloma remains, please send a minimum of 20mL of urine for Bence Jones protein determination with sodium azide preservative (containers available from C17 lab), or serum for free light chain analysis. Repeat SPE is advised in 6-12 months.
::SUS::Albumin is decreased. Alpha-1, alpha-2 and beta-2 are increased. There is a suspicious peak in  gamma measuring  g/L and the rest of gamma measures  g/L (8 - 14 g/L). `n`nImmunofixation will be performed. See result below.
::HYPO::Albumin is decreased. Alpha-1, alpha-2 and beta-2 are increased. There is a hypogammaglobulinaemia measuring  g/L (8 - 14 g/L) with an irregularity in  gamma (<1 g/L). `n`nImmunofixation will be performed to assess monoclonality. See result below.
::BET1::Albumin is decreased. Alpha-1, alpha-2, beta-1 and beta-2 are increased. There is a polyclonal hypergammaglobulinaemia present measuring  g/L (8 - 14 g/L). `n`nImmunofixation will be performed. See result below.
::FEDEF:: The increase in beta-1 (transferrin) may represent iron deficiency anaemia present in this patient. `n`nRepeat SPE is advised once the iron deficiency anaemia has resolved.
::PREV::The previously typed monoclonal IgG kappa peak persists in  gamma measuring  g/L. Immunoparesis is present.
::MCI::The suspicious peak in mid gamma types as monoclonal IgG kappa. `n`nPlease send urine for Bence Jones protein determination, and serum for free light chain analysis and quantitative immunoglobulins.
::NMCI::There is no monoclonal evident. `n`nIf the clinical suspicion of myeloma remains, please send urine for Bence Jones protein determination (with sodium azide preservative), or serum for free light chain analysis.
::PROM::A prominent peak is present in the mid-gamma region measuring g/L. The remainder of the gamma region measures g/L (8-14 g/L).  Immunotyping will be performed. Please see results below.
::SUSP::If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::CNT::Please contact the Chemical Pathology Registrar's office (021 404 4135/5040) if repeat immunofixation is required to confirm the presence or absence of this monoclonal.
::NORMUPE::There is no significant proteinuria. No monoclonal band is evident.
::MILDNON::Mild non-selective proteinuria.
::MILDSEL::Mild selective proteinuria.
::MODNON::Moderate non-selective proteinuria.
::MODSEL::Moderate selective proteinuria.
::SEVNON::Severe non-selective proteinuria.
::SEVSEL::Severe selective proteinuria
::NMCB::No monoclonal band is evident.
::MCB::There is a suspicious band in _ gamma measuring _. Immunofixation will be performed.
::snd::If the clinical suspicion of myeloma remains, please send a minimum of 20mL of urine for Bence Jones protein determination with sodium azide preservative (containers available from C17 lab), or serum for free light chain analysis.
::PREV2::The previously typed monoclonal IgA kappa/lambda peak persists overlying the beta-2 region with the combined monoclonal/complement peak measuring  g/L.
::PREV3::The previously typed monoclonal IgA kappa/lambda peak persists overlying the beta-1 region with the combined monoclonal/transferrin peak measuring  g/L.
::ins::Haemolysis falsely decreases insulin. This H reading is _ which is much higher than the cut-off of 30, where we anticipate seeing significant effects on insulin.
::NEPHR::Hypoalbuminaemia is present.  The alpha-2 (macroglobulin) region is significantly increased at _ g/L (5-9 g/L).  The gamma region measures _ g/L (8-14 g/L). No monoclonal peaks are visible. `nThis pattern suggests nephrotic syndrome. If the clinical suspicion of myeloma remains, urine Bence Jones protein electrophoresis (at least 20ml urine in a container with sodium azide preservative obtainable from the lab) or serum free light chain analysis are recommended.
::A-1::The alpha-1 peak is biphasic, suggesting alpha-1-antitrypsin heterozygosity.
::IMMFIX::Immunofixation will be performed.  See result below, to follow.
::HYPOG::Hypogammaglobulinemia is present at _ g/L (8-14 g/L).`nImmunofixation will be performed. See results below, to follow.
::POLYG::Polyclonal hypergammaglobulinemia is present at _ g/L (normal 8-14 g/L).
::ARR::Unable to calculate the aldosterone:renin ratio due to the upper / lower measuring limit of aldosterone / renin, but the ratio is </> _.
::TGR:: The gamma region measures
::CSFELEC::
(
CSF electrophoresis:
Samples with high total protein concentrations (>16.8 g/L) will not be run due to the increased likelihood of false negative results.
)
::CLINCONT::Clinician contact details may not be coded in our database.  Please go to the link below to update clinician contact details:`ntinyurl.com/nhls-update
::DTFT::These thyroid function test results are discordant.`n`nPlease contact the Chemical Pathology Registrars' Office (021 404 4135/5040) if these results are not in keeping with the clinical scenario.
::text1::
(
Any text
Newline preserved in this way
)
::USPCOMMENT::
(
~BOLD
URINE STEROID PROFILE:
Laboratory information
~NORMAL
SOPs used for analysis: SOP-SPT-003: Analysis of Anabolic Steroids
Samples were analysed according to standard SOPs.  There were no
deviations from the test methods.  Final results were reviewed and
approved by a Technical Signatory(ies).
Key to Coding: GC-MS/MS = Gas Chromatography/Tandem Mass Spectrometry
               LOQ = Limit of Quantification
Compound                               Method        Conc. ng/mL
5a-androstane-2a,17?-diol (5a-diol)    GC-MS/MS
5?-androstane-2a,17?-diol (5?-diol)    GC-MS/MS
Androsterone (Andro)                   GC-MS/MS
Epitestosterone (E)                    GC-MS/MS
Etiocholanolone (Etio)                 GC-MS/MS
Testosterone (T)                       GC-MS/MS
DHT (5a-Dihydrotestosterone)           GC-MS/MS
~BOLD
Ratios
~NORMAL
T/E                                    GC-MS/MS
5a-diol/5?-diol                        GC-MS/MS
5?-diol/EpiT                           GC-MS/MS
5a-diol/EpiT                           GC-MS/MS
Andro/Etio                             GC-MS/MS
Andro/T                                GC-MS/MS
Andro/EpiT                             GC-MS/MS
T/DHT                                  GC-MS/MS
DHT/EpiT                               GC-MS/MS

Comments:
The above steroid profile may not be suitable for diagnostic purposes, as some ratios could not be measured accurately as the concentrations were below the limit of quantification of the assay.

The sample shows signs of extensive degradation as the 5a-androstanedione/Androsterone and/or 5?-androstanedione/Etiocholanolone ratio are >/= 0.1 in the sample.

Signed
Director: JL du Preez
Test performed at a referral laboratory:
South African Doping Control Laboratory - SADoCoL
)

::FAICOMMENT::
(
Note that the free androgen indices are unable to be calculated due to the upper / lower measuring limit of testosterone / SHBG, but if testosterone of 0.1 nmol/L were used for the calculation, the values below are applicable. The values of the free androgen indices below are likely thus likely an over/underestimation of the actual values.
~BOLD
Index                          Result            Reference range
~NORMAL
Calculated free testosterone   <0.001 L nmol/L     0.003 - 0.033
% Free testosterone            <1.00 %              0.70 - 2.19
Calc bioavailable testosterone <0.02 L nmol/L       0.06 - 0.76
% Bioavailable testosterone    <20.0 %              15.3 - 47.7
Free androgen index            <0.1 L                0.3 - 5.6
)

::AVSStudy::
(
~BOLD
Sampling site       Time    Episode      Aldo   Cort   SI   ACR LI
~NORMAL
PIVC2 suprarenal*   15h56   SA05467137   2350   2607        0.9
PIVC1 infrarenal*   15h55   SA05467134   3130   997     3.1
L adrenal**         15h52   SA05467129   129000 26360   14.6    4.9 2.1
L phrenoadrenal**   15h50   SA05467127   67500  6090    3.4 11.1    4.7
L renal*            15h33   SA05467126   16300  1060    0.6 15.4    6.6
R adrenal vein A    15h30   SA05467121   2170   989 0.5 2.2
R adrenal vein B    15h29   SA05467119   1870   749 0.4 2.5
Peripheral femoral  15h57   SA05467115   2150   951 0.5 2.3
Baseline IVC        14h26   SA05467110   610    95  0.1 6.4
~BOLD
Glossary
~NORMAL
Aldo = Aldosterone in pmol/L
Cort = Cortisol in nmol/L
SI = Selectivity Index = Cortisol Adrenal vein : Cortisol Peripheral vein
ACR = Aldosterone : Cortisol ratio
LI = Lateralization Index = Dominant ACR : Non-dominant ACR
PIVC = Peripheral (Inferior Vena Cava)
L = Left
R = Right
* Aldosterone in dilution 1:10
** Aldosterone in dilution 1:50
Please note: It is recommended that AVS studies that are not bilaterally
successful should not be used to establish lateralization.
)

; Work in progress
;sending a hotstring with a variable https://www.autohotkey.com/board/topic/95718-hotstrings-that-have-a-variable-in-the-command/
;These are the default endchars for a hotkey, pressing any of these keys will trigger the hotstring
;You can add or remove characters as you need
endchars = -()[]{}':;/\,.?`n `t!

;'*B0 = do not auto-backspace, and do not wait for a hotstring endchar to run this hostring
:*B0:hotstring::
;::hotstring::
  ;Wait for additional text
  Input, modifier, V, %endchars%
  ;Extract the endkey pressed from ErrorLevel
  RegexMatch(ErrorLevel, "EndKey:\K.*", EndKey)
  ;Send backspaces to clear out the hotstring text
  SendInput % "{bs " strlen(modifier)+strlen(A_ThisHotkey)-4 "}"

  ;Put the text you want to send here
  ;"modifier" contains the text entered after the hotstring
  ;"EndKey" contains the key that terminated the hotkey
  ;Place EndKey between brackets { } as it is not always a literal character, for example, EndKey might be the text "Space"
  SendInput You pressed these keys: %modifier% after the hotstring{%EndKey%}
return


#IfWinNotActive, ahk_class XLMAIN ahk_exe EXCEL.EXE

::STN::Stable negative bias.
::STP::Stable positive bias.
::CTM::Close to mean.  Stable.
::CTT::Close to target. Stable.
::FreeT4comment::Unfortunately we were unable to get hold of the attending clinician for the free-T4 result which has changed after re-analysis.

;If you have a long name like me, you can enter a shortcode like below, without the preceeding ";".
;Type "d1" then space or tab and the text as coded will be typed automatically.

::D1::
SendInput, %full_name%
return
#IfWinNotActive


;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                        Button Info

Button_i:
run, https://github.com/dietervdwes/chemhelp/
return

;!RButton::
;MouseGetPos, xpos, ypos, window
;Gui, New , , ChemHelp
;Gui, Add, Button, gScript1, [&1] Open in Patient Entry
;Gui, Add, Button, gScript2, [&2] Open in Result Entry
;Gui, Add, Button, gScript3, [&3] Open in Specimen Info
;Gui, Add, Button, gScript4, [&4] Open form on Equation
;Gui, Add, Button, gExitApp, Reload
;
;Gui, -Border
;Gui, Show, x%xpos% y%ypos%
;return

;Script1:
;Gui, Hide
;sleep, 600
;send, {CtrlDown}c{CtrlUp}
;sleep, 500


;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                            Button More
ButtonMore:

Gui, New , , ChemHelp
Gui, Add, Edit, r1 vEpisode w135, EpisodeOrMRN
Gui, Add, Button, gMoreScript01, [&01] Open form on Equation
Gui, Add, Button, gMoreScript02, [&02] Open MRN in EPR view
Gui, Add, Button, gMoreScript03, [&03] Locate Episode Storage position
Gui, Add, Button, gMoreScript04, [&04] Open in Patient Entry
Gui, Add, Button, gMoreScript05, [&05] Extract MRN to ExcelFile
Gui, Add, Button, gMoreScript06, [&06] SPE Comment
Gui, Add, Button, gMoreScript07, [&07] IFE Comment
Gui, Add, Button, gMoreScript08, [&08] Histo Comment
Gui, Add, Button, gMoreScript09, [&09] UOA Comment
Gui, Add, Button, glog_episode, [&12] Log Episode ;see Ctrl+Alt+N or Ctrl+Alt+L ~line 450
Gui, Add, Button, gMoreScript10 gMoreScript10, [&10] UPE Comment
Gui, Add, Button, gMoreScript11 gMoreScript11, [&11] UIFE Comment
Gui, Add, Button, gMoreScript13 gMoreScript13, [&13] Porphyria Comment
Gui, Add, Button, gMoreScript14, [&14] Bone Marrow Report
Gui, Add, Button, gMoreScript15, [&15] Film interpretation
Gui, Add, Button, x150 y265 w45 h20 gExitApp,  Restart
;Gui, -Border
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
ypos := A_ScreenHeight-600
xpos := A_ScreenWidth-320
Gui, Margin, 0, 0
Gui, Show, x%xpos% y%ypos% w220 h500, ChemHelp - More
return

MoreScript01:
;Open form on Equation
        Gui, Submit
		url := "http://" . equation_ip . "/multipdfsearch.php?file=" . Episode . ""
		run, %url%
		sleep, 1500
		WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
		Return

MoreScript02:
; Open MRN in EPR view
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . Episode . ""
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
		Return
MoreScript03:
;Locate Episode Storage position
		Gui, Submit
		sleep, 100
        Specinfolookup(Episode)
		Return
MoreScript04:
;Open in Patient Entry
        Gui, Submit
		sleep, 100
        OpenEpInPatientEntry(Episode)
		Return
MoreScript05:
;Extract MRN to ExcelFile
        Gui, Submit
        run node %A_ScriptDir%\SavePatientEPR.js %Episode%
        if (ErrorLevel = "ERROR")
            MsgBox There was an error, likely Node is not installed.
        else
        Return
MoreScript06:
; SPE Comment
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^C112^1^CF112^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript07:
;IFE Comment
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^C113^1^CF113^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript08:
;Histo Comment
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^A030^1^T^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
        Return
MoreScript09:
;UOA Comment
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^C440^1^C3200^"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript10:
;Log Episode
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "%5EC323%5E1%5EC1245%5E"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript11:
;UIFE Comment
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "%5EC324%5E1%5EC1260%5E"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript13:
;porphyria comment
        Gui, Submit
		url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^C900^1^C9902^BP040"
		run, %url%
		sleep, 800
		WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript14:
;Bone Marrow Report
        Gui, submit
        url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^H445^1^HF445^"
        run, %url%
        sleep, 800
        WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return
MoreScript15:
;Film interpretation
        Gui, submit
        url := "https://trakdb-prod.nhls.ac.za/csp/reporting/eprajax.csp?chunk=" . Episode . "^H800^1^H9916^"
        run, %url%
        sleep, 800
        WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
        Return

ButtonClose2:
sleep, 500
send, {altdown}{F4 down}{F4 up}{altup}
sleep, 100
Return

;SA04972912
;MRN73839380
Specinfolookup(epis)
{
IfWinExist, Specimen Information
{
	WinActivate, Specimen Information
	sleep, 500
	MouseClick, Left, 77, 82, 2, 100
	sleep, 500
}
Else IfWinNotExist, Specimen Information
{
	IfWinNotActive,  NHLS, , WinActivate,  NHLS,
	WinWaitActive,  NHLS,
	Sleep, 100
	Send, a{TAB}
	Send, Specimen Information{Enter}
	WinWaitActive, Specimen Information
}
sleep, 500
send, %epis%
sleep, 200
send, {Tab}
sleep, 200
MouseClick, Left, 271, 271
send, {Altdown}d{AltUp}
return
}

OpenEpInPatientEntry(epis)
{
    IfWinExist, Result Entry - Single
    {
        MsgBox, Result Entry - Single window is already open.  It needs to be closed for this function to work.
        Return
    }
    Else IfWinNotExist, Result Entry - Single
    {
        IfWinNotActive,  NHLS, , WinActivate,  NHLS,
        WinWaitActive,  NHLS,
        Sleep, 200
        Send, {Home}
        sleep, 100
        Send, Result Entry
        sleep, 100
        Send, {Enter}
        WinWaitActive, Result Entry - \\Remote
        sleep, 200
        send, %epis%
        send, {Enter}
        Return
    }
}

;Linear Spoon's Screen capture script from https://www.autohotkey.com/board/topic/121619-screencaptureahk-broken-capturescreen-function-win-81-x64/
/* CaptureScreen(aRect, bCursor, sFileTo, nQuality)
1) If the optional parameter bCursor is True, captures the cursor too.
2) If the optional parameter sFileTo is 0, set the image to Clipboard.
   If it is omitted or "", saves to screen.bmp in the script folder,
   otherwise to sFileTo which can be BMP/JPG/PNG/GIF/TIF.
   sFile should be 0 if you want it to go to the clipboard, leaving it blank makes it save to some default filename in the working directory.
3) The optional parameter nQuality is applicable only when sFileTo is JPG. Set it to the desired quality level of the resulting JPG, an integer between 0 - 100.
4) If aRect is 0/1/2/3, captures the entire desktop/active window/active client area/active monitor.
5) aRect can be comma delimited sequence of coordinates, e.g., "Left, Top, Right, Bottom" or "Left, Top, Right, Bottom, Width_Zoomed, Height_Zoomed".
   In this case, only that portion of the rectangle will be captured. Additionally, in the latter case, zoomed to the new width/height, Width_Zoomed/Height_Zoomed.

Example:
CaptureScreen(0)
CaptureScreen(1)
CaptureScreen(2)
CaptureScreen(3)
CaptureScreen("100, 100, 200, 200")
CaptureScreen("100, 100, 200, 200, 400, 400")   ; Zoomed
*/

/* Convert(sFileFr, sFileTo, nQuality)
Convert("C:\image.bmp", "C:\image.jpg")
Convert("C:\image.bmp", "C:\image.jpg", 95)
Convert(0, "C:\clip.png")   ; Save the bitmap in the clipboard to sFileTo if sFileFr is "" or 0.
*/


CaptureScreen(aRect = 0, bCursor = False, sFile ="", nQuality = "")
{
	If !aRect
	{
		SysGet, nL, 76  ; virtual screen left & top
		SysGet, nT, 77
		SysGet, nW, 78	; virtual screen width and height
		SysGet, nH, 79
	}
	Else If aRect = 1
		WinGetPos, nL, nT, nW, nH, A
	Else If aRect = 2
	{
		WinGet, hWnd, ID, A
		VarSetCapacity(rt, 16, 0)
		DllCall("GetClientRect" , "ptr", hWnd, "ptr", &rt)
		DllCall("ClientToScreen", "ptr", hWnd, "ptr", &rt)
		nL := NumGet(rt, 0, "int")
		nT := NumGet(rt, 4, "int")
		nW := NumGet(rt, 8)
		nH := NumGet(rt,12)
	}
	Else If aRect = 3
	{
		VarSetCapacity(mi, 40, 0)
		DllCall("GetCursorPos", "int64P", pt), NumPut(40,mi,0,"uint")
		DllCall("GetMonitorInfo", "ptr", DllCall("MonitorFromPoint", "int64", pt, "Uint", 2, "ptr"), "ptr", &mi)
		nL := NumGet(mi, 4, "int")
		nT := NumGet(mi, 8, "int")
		nW := NumGet(mi,12, "int") - nL
		nH := NumGet(mi,16, "int") - nT
	}
	Else
	{
		StringSplit, rt, aRect, `,, %A_Space%%A_Tab%
		nL := rt1	; convert the Left,top, right, bottom into left, top, width, height
		nT := rt2
		nW := rt3 - rt1
		nH := rt4 - rt2
		znW := rt5
		znH := rt6
	}

	mDC := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
	hBM := CreateDIBSection(mDC, nW, nH)
	oBM := DllCall("SelectObject", "ptr", mDC, "ptr", hBM, "ptr")
	hDC := DllCall("GetDC", "ptr", 0, "ptr")
	DllCall("BitBlt", "ptr", mDC, "int", 0, "int", 0, "int", nW, "int", nH, "ptr", hDC, "int", nL, "int", nT, "Uint", 0x40CC0020)
	DllCall("ReleaseDC", "ptr", 0, "ptr", hDC)
	If bCursor
		CaptureCursor(mDC, nL, nT)
	DllCall("SelectObject", "ptr", mDC, "ptr", oBM)
	DllCall("DeleteDC", "ptr", mDC)
	If znW && znH
		hBM := Zoomer(hBM, nW, nH, znW, znH)
	If sFile = 0
		SetClipboardData(hBM)
	Else Convert(hBM, sFile, nQuality), DllCall("DeleteObject", "ptr", hBM)
}

CaptureCursor(hDC, nL, nT)
{
	VarSetCapacity(mi, 32, 0), Numput(16+A_PtrSize, mi, 0, "uint")
	DllCall("GetCursorInfo", "ptr", &mi)
	bShow   := NumGet(mi, 4, "uint")
	hCursor := NumGet(mi, 8)
	xCursor := NumGet(mi,8+A_PtrSize, "int")
	yCursor := NumGet(mi,12+A_PtrSize, "int")

	DllCall("GetIconInfo", "ptr", hCursor, "ptr", &mi)
	xHotspot := NumGet(mi, 4, "uint")
	yHotspot := NumGet(mi, 8, "uint")
	hBMMask  := NumGet(mi,8+A_PtrSize)
	hBMColor := NumGet(mi,16+A_PtrSize)

	If bShow
		DllCall("DrawIcon", "ptr", hDC, "int", xCursor - xHotspot - nL, "int", yCursor - yHotspot - nT, "ptr", hCursor)
	If hBMMask
		DllCall("DeleteObject", "ptr", hBMMask)
	If hBMColor
		DllCall("DeleteObject", "ptr", hBMColor)
}

Zoomer(hBM, nW, nH, znW, znH)
{
	mDC1 := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
	mDC2 := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
	zhBM := CreateDIBSection(mDC2, znW, znH)
	oBM1 := DllCall("SelectObject", "ptr", mDC1, "ptr",  hBM, "ptr")
	oBM2 := DllCall("SelectObject", "ptr", mDC2, "ptr", zhBM, "ptr")
	DllCall("SetStretchBltMode", "ptr", mDC2, "int", 4)
	DllCall("StretchBlt", "ptr", mDC2, "int", 0, "int", 0, "int", znW, "int", znH, "ptr", mDC1, "int", 0, "int", 0, "int", nW, "int", nH, "Uint", 0x00CC0020)
	DllCall("SelectObject", "ptr", mDC1, "ptr", oBM1)
	DllCall("SelectObject", "ptr", mDC2, "ptr", oBM2)
	DllCall("DeleteDC", "ptr", mDC1)
	DllCall("DeleteDC", "ptr", mDC2)
	DllCall("DeleteObject", "ptr", hBM)
	Return zhBM
}

Convert(sFileFr = "", sFileTo = "", nQuality = "")
{
	If (sFileTo = "")
		sFileTo := A_ScriptDir . "\screen.bmp"
	SplitPath, sFileTo, , sDirTo, sExtTo, sNameTo

	If Not hGdiPlus := DllCall("LoadLibrary", "str", "gdiplus.dll", "ptr")
		Return	sFileFr+0 ? SaveHBITMAPToFile(sFileFr, sDirTo (sDirTo = "" ? "" : "\") sNameTo ".bmp") : ""
	VarSetCapacity(si, 16, 0), si := Chr(1)
	DllCall("gdiplus\GdiplusStartup", "UintP", pToken, "ptr", &si, "ptr", 0)

	If !sFileFr
	{
		DllCall("OpenClipboard", "ptr", 0)
		If	(DllCall("IsClipboardFormatAvailable", "Uint", 2) && (hBM:=DllCall("GetClipboardData", "Uint", 2, "ptr")))
			DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", hBM, "ptr", 0, "ptr*", pImage)
		DllCall("CloseClipboard")
	}
	Else If	sFileFr Is Integer
		DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", sFileFr, "ptr", 0, "ptr*", pImage)
	Else	DllCall("gdiplus\GdipLoadImageFromFile", "wstr", sFileFr, "ptr*", pImage)
	DllCall("gdiplus\GdipGetImageEncodersSize", "UintP", nCount, "UintP", nSize)
	VarSetCapacity(ci,nSize,0)
	DllCall("gdiplus\GdipGetImageEncoders", "Uint", nCount, "Uint", nSize, "ptr", &ci)
	struct_size := 48+7*A_PtrSize, offset := 32 + 3*A_PtrSize, pCodec := &ci - struct_size
	Loop, %	nCount
		If InStr(StrGet(Numget(offset + (pCodec+=struct_size)), "utf-16") , "." . sExtTo)
			break

	If (InStr(".JPG.JPEG.JPE.JFIF", "." . sExtTo) && nQuality<>"" && pImage && pCodec < &ci + nSize)
	{
		DllCall("gdiplus\GdipGetEncoderParameterListSize", "ptr", pImage, "ptr", pCodec, "UintP", nCount)
		VarSetCapacity(pi,nCount,0), struct_size := 24 + A_PtrSize
		DllCall("gdiplus\GdipGetEncoderParameterList", "ptr", pImage, "ptr", pCodec, "Uint", nCount, "ptr", &pi)
		Loop, %	NumGet(pi,0,"uint")
			If (NumGet(pi,struct_size*(A_Index-1)+16+A_PtrSize,"uint")=1 && NumGet(pi,struct_size*(A_Index-1)+20+A_PtrSize,"uint")=6)
			{
				pParam := &pi+struct_size*(A_Index-1)
				NumPut(nQuality,NumGet(NumPut(4,NumPut(1,pParam+0,"uint")+16+A_PtrSize,"uint")),"uint")
				Break
			}
	}

	If pImage
		pCodec < &ci + nSize	? DllCall("gdiplus\GdipSaveImageToFile", "ptr", pImage, "wstr", sFileTo, "ptr", pCodec, "ptr", pParam) : DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "ptr", pImage, "ptr*", hBitmap, "Uint", 0) . SetClipboardData(hBitmap), DllCall("gdiplus\GdipDisposeImage", "ptr", pImage)

	DllCall("gdiplus\GdiplusShutdown" , "Uint", pToken)
	DllCall("FreeLibrary", "ptr", hGdiPlus)
}


CreateDIBSection(hDC, nW, nH, bpp = 32, ByRef pBits = "")
{
	VarSetCapacity(bi, 40, 0)
	NumPut(40, bi, "uint")
	NumPut(nW, bi, 4, "int")
	NumPut(nH, bi, 8, "int")
	NumPut(bpp, NumPut(1, bi, 12, "UShort"), 0, "Ushort")
	Return DllCall("gdi32\CreateDIBSection", "ptr", hDC, "ptr", &bi, "Uint", 0, "UintP", pBits, "ptr", 0, "Uint", 0, "ptr")
}

SaveHBITMAPToFile(hBitmap, sFile)
{
	VarSetCapacity(oi,104,0)
	DllCall("GetObject", "ptr", hBitmap, "int", 64+5*A_PtrSize, "ptr", &oi)
	fObj := FileOpen(sFile, "w")
	fObj.WriteShort(0x4D42)
	fObj.WriteInt(54+NumGet(oi,36+2*A_PtrSize,"uint"))
	fObj.WriteInt64(54<<32)
	fObj.RawWrite(&oi + 16 + 2*A_PtrSize, 40)
	fObj.RawWrite(NumGet(oi, 16+A_PtrSize), NumGet(oi,36+2*A_PtrSize,"uint"))
	fObj.Close()
}

SetClipboardData(hBitmap)
{
	VarSetCapacity(oi,104,0)
	DllCall("GetObject", "ptr", hBitmap, "int", 64+5*A_PtrSize, "ptr", &oi)
	sz := NumGet(oi,36+2*A_PtrSize,"uint")
	hDIB :=	DllCall("GlobalAlloc", "Uint", 2, "Uptr", 40+sz, "ptr")
	pDIB := DllCall("GlobalLock", "ptr", hDIB, "ptr")
	DllCall("RtlMoveMemory", "ptr", pDIB, "ptr", &oi + 16 + 2*A_PtrSize, "Uptr", 40)
	DllCall("RtlMoveMemory", "ptr", pDIB+40, "ptr", NumGet(oi, 16+A_PtrSize), "Uptr", sz)
	DllCall("GlobalUnlock", "ptr", hDIB)
	DllCall("DeleteObject", "ptr", hBitmap)
	DllCall("OpenClipboard", "ptr", 0)
	DllCall("EmptyClipboard")
	DllCall("SetClipboardData", "Uint", 8, "ptr", hDIB)
	DllCall("CloseClipboard")
}

ButtonCapture:
sleep, 200
send, {AltDown}{Tab}{AltUp}
sleep, 300
^!c::
FormatTime, Timevar,, yyyy-MM-dd_HH-mm-ss
CaptureScreen(1,"","" . A_ScriptDir . "\Screenshots\" . Timevar . ".png","90")
if errorlevel
	MsgBox, Error capturing screen to "ChemHelp-directory\Screenshots"
else
MsgBox, 1,,Window screenshot captured at `nChemHelp-directory\Screenshots,2
    sleep,500
Return

RemoveToolTip:
Tooltip
return

Button_t:
^!'::
CoordMode,ToolTip,Screen
Loop % (20){
    Random,posx,0,% A_ScreenWidth
    Random,posy,0,% A_ScreenHeight
    Random,TimeIn,500,3000
    Random,TimeOut,5000,10000
    ShowToolTip("Random tooltip " A_Index,A_Index,TimeOut/1000,TimeIn/1000,posx,posy)
}
MsgBox, Do not click random buttons... `nNow watch the ToolTips.`nHopefully they disappear soon...
sleep, 5000
return

ShowToolTip(Text, Nr=1,Out=5,In=0,X="",Y=""){
    static
    If (text=""){
        If ((T_Text%Nr% && !T_%Nr% && T_In%Nr%<A_TickCount) || (T_%Nr% && T_Out%Nr%<A_TickCount && !((T_Text%Nr%:="") . (T_X%Nr%:="") . (T_Y%Nr%:="") . (T_In%Nr%:="") . (T_Out%Nr%:="") . (T_%Nr%:=""))))
            ToolTip % T_Text%Nr%,% T_X%Nr%,% T_Y%Nr%,% Nr+(T_%Nr%:=1)-1
        Return (!T_%Nr% ? T_In%Nr% : (T_Out%Nr%>A_TickCount ? T_Out%Nr% : ""))
    }
   T_Text%Nr%:=text,T_X%Nr%:= X,T_Y%Nr%:= Y,timer ? "" : (timer:=1),T_In%Nr%:=Round(A_TickCount + In*1000),T_Out%Nr%:=Round(A_TickCount + (In+Out)*1000), NextTimer ? "" : (NextTimer:=10)
    ShowToolTip:
    NextTimer=
    Loop 20
        If (ErrorLevel:=ShowToolTip("",A_Index))
            If (!NextTimer || NextTimer+A_TickCount>ErrorLevel)
                NextTimer:=ErrorLevel-A_TickCount+10
    SetTimer, ShowToolTip,% NextTimer ? (-1*(NextTimer+10)) : "Off"
    Return
}


/*
;-----------------------------   Logging each Authorize
#IfWinActive, Medical Validation :   (Authorise By Episode) - \\Remote
!a::
FileAppend, %A_now%`n, %A_ScriptDir%\chemhelp_log.txt
sleep, 100
If WinNotActive("Medical Validation :   (Authorise By Episode) - \\Remote")
{
    WinActivate, Medical Validation :   (Authorise By Episode) - \\Remote)
    send, {AltDown}
    sleep, 100
    send, a
    sleep, 100
    send, {AltUp}
    return
}
return
*/


/*
responding to the first post:

close balloon tips on a case by case basis
http://www.autohotkey.com/forum/viewtopic.php?t=55390

Leef_me's old post, rediscovered, yeah!
ControlClick coordinates and Windows Spy question (reads a tooltip in the process)
http://www.autohotkey.com/forum/topic54333.html

Keyboard shortcut to close balloon tips
http://www.autohotkey.com/forum/viewtopic.php?t=15141&highlight=close+balloon
*/

/* This script tried to pull information from a tooltip (aim is to use the tooltip displayed over the VQ and write an API which sends updates as the size of this list increases.
#singleinstance force
CoordMode, ToolTip, screen
CoordMode, Mouse, screen
return

;======== show the information from a single tooltip =========
F1::
ControlGetText, tooltip2,,ahk_class tooltips_class32
;tooltip, ******************`n%tooltip2%`n******************, 1300, 300
;sleep, 1000
;tooltip
msgbox, %tooltip2%
return


;======== show the information from a multiple tooltips =========
F2::
WinGet, ID, LIST,ahk_class tooltips_class32
tt_text=
Loop, %id%
{
  this_id := id%A_Index%
  ControlGetText, tooltip2,,ahk_id %this_id%
  tt_text.= "******************`n" tooltip2 "`n******************`n"
}

;tooltip, %tt_text%, 300, 300
msgbox, %tt_text%
;sleep, 3000
;tooltip

return
*/

;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                        Button Close
ButtonClose:
WinActivate, ChemHelp
WinWaitActive, ChemHelp
ExitApp
sleep, 100
Return
;send, {altdown}{F4 down}{F4 up}{altup}
;sleep, 100
;Return

Escape::Reload
;ExitApp
Return

^SPACE::  Winset, Alwaysontop, , A ;forces a window to remain on top
^6::  Winset, Alwaysontop, , B

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;                                                                       Google Search highlighted text (ctrl+shift+c)
Googlehighlightedtext:
^+c::
{
 Send, ^c
 Sleep 50
 Run, http://www.google.com/search?q=%clipboard%
 Return
}


;------------------------------------------------------------------------------
;Autocorrect script;
; INTRODUCTION
;
; This is an AutoHotKey script that implements AutoCorrect against several
; "Lists of common misspellings":
;
; This does not replace a proper spellchecker such as in Firefox, Word, etc.
; It is usually better to have uncertain typos highlighted by a spellchecker
; than to "correct" them incorrectly so that they are no longer even caught by
; a spellchecker: it is not the job of an autocorrector to correct *all*
; misspellings, but only those which are very obviously incorrect.
;
; From a suggestion by Tara Gibb, you can add your own corrections to any
; highlighted word by hitting Win+H. These will be added to a separate file,
; so that you can safely update this file without overwriting your changes.
;
; Some entries have more than one possible resolution (achive->achieve/archive)
; or are clearly a matter of deliberate personal writing style (wanna, colour)
;
; These have been placed at the end of this file and commented out, so you can
; easily edit and add them back in as you like, tailored to your preferences.
;
; SOURCES
;
; http://en.wikipedia.org/wiki/Wikipedia:Lists_of_common_misspellings
; http://en.wikipedia.org/wiki/Wikipedia:Typo
; Microsoft Office autocorrect list
; Script by jaco0646 http://www.autohotkey.com/forum/topic8057.html
; OpenOffice autocorrect list
; TextTrust press release
; User suggestions.
;
; CONTENTS
;
;   Settings
;   AUto-COrrect TWo COnsecutive CApitals (commented out by default)
;   Win+H code
;   Fix for -ign instead of -ing
;   Word endings
;   Word beginnings
;   Accented English words
;   Common Misspellings - the main list
;   Ambiguous entries - commented out
;------------------------------------------------------------------------------

;------------------------------------------------------------------------------
; Settings
;------------------------------------------------------------------------------
#NoEnv ; For security
#SingleInstance force

;------------------------------------------------------------------------------
; AUto-COrrect TWo COnsecutive CApitals.
; Disabled by default to prevent unwanted corrections such as IfEqual->Ifequal.
; To enable it, remove the /*..*/ symbols around it.
; From Laszlo's script at http://www.autohotkey.com/forum/topic9689.html
;------------------------------------------------------------------------------
/*
; The first line of code below is the set of letters, digits, and/or symbols
; that are eligible for this type of correction.  Customize if you wish:
keys = abcdefghijklmnopqrstuvwxyz
Loop Parse, keys
    HotKey ~+%A_LoopField%, Hoty
Hoty:
    CapCount := SubStr(A_PriorHotKey,2,1)="+" && A_TimeSincePriorHotkey<999 ? CapCount+1 : 1
    if CapCount = 2
        SendInput % "{BS}" . SubStr(A_ThisHotKey,3,1)
    else if CapCount = 3
        SendInput % "{Left}{BS}+" . SubStr(A_PriorHotKey,3,1) . "{Right}"
Return
*/


;------------------------------------------------------------------------------
; Win+H to enter misspelling correction.  It will be added to this script.
;------------------------------------------------------------------------------
#h::
; Get the selected text. The clipboard is used instead of "ControlGet Selected"
; as it works in more editors and word processors, java apps, etc. Save the
; current clipboard contents to be restored later.
AutoTrim Off  ; Retain any leading and trailing whitespace on the clipboard.
ClipboardOld = %ClipboardAll%
Clipboard =  ; Must start off blank for detection to work.
Send ^c
ClipWait 1
if ErrorLevel  ; ClipWait timed out.
    return
; Replace CRLF and/or LF with `n for use in a "send-raw" hotstring:
; The same is done for any other characters that might otherwise
; be a problem in raw mode:
StringReplace, Hotstring, Clipboard, ``, ````, All  ; Do this replacement first to avoid interfering with the others below.
StringReplace, Hotstring, Hotstring, `r`n, ``r, All  ; Using `r works better than `n in MS Word, etc.
StringReplace, Hotstring, Hotstring, `n, ``r, All
StringReplace, Hotstring, Hotstring, %A_Tab%, ``t, All
StringReplace, Hotstring, Hotstring, `;, ```;, All
Clipboard = %ClipboardOld%  ; Restore previous contents of clipboard.
; This will move the InputBox's caret to a more friendly position:
SetTimer, MoveCaret, 10
; Show the InputBox, providing the default hotstring:
InputBox, Hotstring, New Hotstring, Provide the corrected word on the right side. You can also edit the left side if you wish.`n`nExample entry:`n::teh::the,,,,,,,, ::%Hotstring%::%Hotstring%

if ErrorLevel <> 0  ; The user pressed Cancel.
    return
; Otherwise, add the hotstring and reload the script:
FileAppend, `n%Hotstring%, %A_ScriptFullPath%  ; Put a `n at the beginning in case file lacks a blank line at its end.
Reload
Sleep 200 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
MsgBox, 4,, The hotstring just added appears to be improperly formatted.  Would you like to open the script for editing? Note that the bad hotstring is at the bottom of the script.
IfMsgBox, Yes, Edit
return

MoveCaret:
IfWinNotActive, New Hotstring
    return
; Otherwise, move the InputBox's insertion point to where the user will type the abbreviation.
Send {HOME}
Loop % StrLen(Hotstring) + 4
    SendInput {Right}
SetTimer, MoveCaret, Off
return

#Hotstring R  ; Set the default to be "raw mode" (might not actually be relied upon by anything yet).


;------------------------------------------------------------------------------
; Fix for -ign instead of -ing.
; Words to exclude: (could probably do this by return without rewrite)
; From: http://www.morewords.com/e nds-with/gn/
;------------------------------------------------------------------------------
#Hotstring B0  ; Turns off automatic backspacing for the following hotstrings.
::align::
::antiforeign::
::arraign::
::assign::
::benign::
::campaign::
::champaign::
::codesign::
::coign::
::condign::
::consign::
::coreign::
::cosign::
::countercampaign::
::countersign::
::deign::
::deraign::
::design::
::eloign::
::ensign::
::feign::
::foreign::
::indign::
::malign::
::misalign::
::outdesign::
::overdesign::
::preassign::
::realign::
::reassign::
::redesign::
::reign::
::resign::
::sign::
::sovereign::
::unbenign::
::verisign::
return  ; This makes the above hotstrings do nothing so that they override the ign->ing rule below.

#Hotstring B  ; Turn back on automatic backspacing for all subsequent hotstrings.
:?:ign::ing


;------------------------------------------------------------------------------
; Word endings
;------------------------------------------------------------------------------
:?:bilites::bilities
:?:bilties::bilities
:?:blities::bilities
:?:bilty::bility
:?:blity::bility
:?:, btu::, but ; Not just replacing "btu", as that is a unit of heat.
:?:; btu::; but
:?:n;t::n't
:?:;ll::'ll
:?:;re::'re
:?:;ve::'ve
::sice::since  ; Must precede the following line!
:?:sice::sive
:?:t eh:: the
:?:t hem:: them


;------------------------------------------------------------------------------
; Word beginnings
;------------------------------------------------------------------------------
:*:abondon::abandon
:*:abreviat::abbreviat
:*:accomadat::accommodat
:*:accomodat::accommodat
:*:acheiv::achiev
:*:achievment::achievement
:*:acquaintence::acquaintance
:*:adquir::acquir
:*:aquisition::acquisition
:*:agravat::aggravat
:*:allign::align
:*:ameria::America
:*:archaelog::archaeolog
:*:archtyp::archetyp
:*:archetect::architect
:*:arguement::argument
:*:assasin::assassin
:*:asociat::associat
:*:assymetr::asymmet
:*:atempt::attempt
:*:atribut::attribut
:*:avaialb::availab
:*:comision::commission
:*:contien::conscien
:*:critisi::critici
:*:crticis::criticis
:*:critiz::criticiz
:*:desicant::desiccant
:*:desicat::desiccat
::develope::develop  ; Omit asterisk so that it doesn't disrupt the typing of developed/developer.
:*:dissapoint::disappoint
:*:divsion::division
:*:dcument::document
:*:embarass::embarrass
:*:emminent::eminent
:*:empahs::emphas
:*:enlargment::enlargement
:*:envirom::environm
:*:enviorment::environment
:*:excede::exceed
:*:exilerat::exhilarat
:*:extraterrestial::extraterrestrial
:*:faciliat::facilitat
:*:garantee::guaranteed
:*:guerrila::guerrilla
:*:guidlin::guidelin
:*:girat::gyrat
:*:harasm::harassm
:*:immitat::imitat
:*:imigra::immigra
:*:impliment::implement
:*:inlcud::includ
:*:indenpenden::independen
:*:indisputib::indisputab
:*:isntall::install
:*:insitut::institut
:*:knwo::know
:*:lsit::list
:*:mountian::mountain
:*:nmae::name
:*:necassa::necessa
:*:negociat::negotiat
:*:neigbor::neighbour
:*:noticibl::noticeabl
:*:ocasion::occasion
:*:occuranc::occurrence
:*:priveledg::privileg
:*:recie::recei
:*:recived::received
:*:reciver::receiver
:*:recepient::recipient
:*:reccomend::recommend
:*:recquir::requir
:*:requirment::requirement
:*:respomd::respond
:*:repons::respons
:*:ressurect::resurrect
:*:seperat::separat
:*:sevic::servic
:*:smoe::some
:*:supercede::supersede
:*:superceed::supersede
:*:weild::wield


;------------------------------------------------------------------------------
; Word middles
;------------------------------------------------------------------------------
:?*:compatab::compatib  ; Covers incompat* and compat*
:?*:catagor::categor  ; Covers subcatagories and catagories.


;------------------------------------------------------------------------------
; Accented English words, from, amongst others,
; http://en.wikipedia.org/wiki/List_of_English_words_with_diacritics
; I have included all the ones compatible with reasonable codepages, and placed
; those that may often not be accented either from a clash with an unaccented
; word (resume), or because the unaccented version is now common (cafe).
;------------------------------------------------------------------------------
::aesop::Æsop
::a bas::à bas
::a la::à la
::ancien regime::Ancien Régime
::angstrom::Ångström
::angstroms::Ångströms
::anime::animé
::animes::animés
::ao dai::ào dái
::apertif::apértif
::apertifs::apértifs
::applique::appliqué
::appliques::appliqués
::apres::après
::arete::arête
::attache::attaché
::attaches::attachés
::auto-da-fe::auto-da-fé
::belle epoque::belle époque
::bete noire::bête noire
::betise::bêtise
::Bjorn::Bjørn
::blase::blasé
::boite::boîte
::boutonniere::boutonnière
::canape::canapé
::canapes::canapés
::celebre::célèbre
::celebres::célèbres
::chaines::chaînés
::cinema verite::cinéma vérité
::cinemas verite::cinémas vérité
::cinema verites::cinéma vérités
::champs-elysees::Champs-Élysées
::charge d'affaires::chargé d'affaires
::chateau::château
::chateaux::châteaux
::chateaus::châteaus
::cliche::cliché
::cliched::clichéd
::cliches::clichés
::cloisonne::cloisonné
::consomme::consommé
::consommes::consommés
::communique::communiqué
::communiques::communiqués
::confrere::confrère
::confreres::confrères
::cortege::cortège
::corteges::cortèges
::coup d'etat::coup d'état
::coup d'etats::coup d'états
::coup de tat::coup d'état
::coup de tats::coup d'états
::coup de grace::coup de grâce
::creche::crèche
::creches::crèches
::coulee::coulée
::coulees::coulées
::creme brulee::crème brûlée
::creme brulees::crème brûlées
::creme caramel::crème caramel
::creme caramels::crème caramels
::creme de cacao::crème de cacao
::creme de menthe::crème de menthe
::crepe::crêpe
::crepes::crêpes
::creusa::Creüsa
::crouton::croûton
::croutons::croûtons
::crudites::crudités
::curacao::curaçao
::dais::daïs
::daises::daïses
::debacle::débâcle
::debacles::débâcles
::debutante::débutante
::debutants::débutants
::declasse::déclassé
::decolletage::décolletage
::decollete::décolleté
::decor::décor
::decors::décors
::decoupage::découpage
::degage::dégagé
::deja vu::déjà vu
::demode::démodé
::denoument::dénoument
::derailleur::dérailleur
::derriere::derrière
::deshabille::déshabillé
::detente::détente
::diamante::diamanté
::discotheque::discothèque
::discotheques::discothèques
::divorcee::divorcée
::divorcees::divorcées
::doppelganger::doppelgänger
::doppelgangers::doppelgängers
::eclair::éclair
::eclairs::éclairs
::eclat::éclat
::el nino::El Niño
::elan::élan
::emigre::émigré
::emigres::émigrés
::entree::entrée
::entrees::entrées
::entrepot::entrepôt
::entrecote::entrecôte
::epee::épée
::epees::épées
::etouffee::étouffée
::facade::façade
::facades::façades
::fete::fête
::fetes::fêtes
::faience::faïence
::fiance::fiancé
::fiances::fiancés
::fiancee::fiancée
::fiancees::fiancées
::filmjolk::filmjölk
::fin de siecle::fin de siècle
::flambe::flambé
::flambes::flambés
::fleche::flèche
::Fohn wind::Föhn wind
::folie a deux::folie à deux
::folies a deux::folies à deux
::fouette::fouetté
::frappe::frappé
::frappes::frappés
:?*:fraulein::fräulein
:?*:fuhrer::Führer
::garcon::garçon
::garcons::garçons
::gateau::gâteau
::gateaus::gâteaus
::gateaux::gâteaux
::gemutlichkeit::gemütlichkeit
::glace::glacé
::glogg::glögg
::gewurztraminer::Gewürztraminer
::gotterdammerung::Götterdämmerung
::grafenberg spot::Gräfenberg spot
::habitue::habitué
::ingenue::ingénue
::jager::jäger
::jalapeno::jalapeño
::jalapenos::jalapeños
::jardiniere::jardinière
::krouzek::kroužek
::kummel::kümmel
::kaldolmar::kåldolmar
::landler::ländler
::langue d'oil::langue d'oïl
::la nina::La Niña
::litterateur::littérateur
::lycee::lycée
::macedoine::macédoine
::macrame::macramé
::maitre d'hotel::maître d'hôtel
::malaguena::malagueña
::manana::mañana
::manege::manège
::manque::manqué
::materiel::matériel
::matinee::matinée
::matinees::matinées
::melange::mélange
::melee::mêlée
::melees::mêlées
::menage a trois::ménage à trois
::menages a trois::ménages à trois
::mesalliance::mésalliance
::metier::métier
::minaudiere::minaudière
::mobius strip::Möbius strip
::mobius strips::Möbius strips
::moire::moiré
::moireing::moiréing
::moires::moirés
::motley crue::Mötley Crüe
::motorhead::Motörhead
::naif::naïf
::naifs::naïfs
::naive::naïve
::naiver::naïver
::naives::naïves
::naivete::naïveté
::nee::née
::negligee::negligée
::negligees::negligées
::neufchatel cheese::Neufchâtel cheese
::nez perce::Nez Percé
::noël::Noël
::noëls::Noëls
::número uno::número uno
::objet trouve::objet trouvé
::objets trouve::objets trouvé
::ombre::ombré
::ombres::ombrés
::omerta::omertà
::opera bouffe::opéra bouffe
::operas bouffe::opéras bouffe
::opera comique::opéra comique
::operas comique::opéras comique
::outre::outré
::papier-mache::papier-mâché
::passe::passé
::piece de resistance::pièce de résistance
::pied-a-terre::pied-à-terre
::plisse::plissé
::pina colada::Piña Colada
::pina coladas::Piña Coladas
::pinata::piñata
::pinatas::piñatas
::pinon::piñon
::pinons::piñons
::pirana::piraña
::pique::piqué
::piqued::piquéd
::più::più
::plie::plié
::precis::précis
::polsa::pölsa
::pret-a-porter::prêt-à-porter
::protoge::protégé
::protege::protégé
::proteged::protégéd
::proteges::protégés
::protegee::protégée
::protegees::protégées
::protegeed::protégéed
::puree::purée
::pureed::puréed
::purees::purées
::Quebecois::Québécois
::raison d'etre::raison d'être
::recherche::recherché
::reclame::réclame
::résume::résumé
::resumé::résumé
::résumes::résumés
::resumés::résumés
::retrousse::retroussé
::risque::risqué
::riviere::rivière
::roman a clef::roman à clef
::roue::roué
::saute::sauté
::sauted::sautéd
::seance::séance
::seances::séances
::senor::señor
::senors::señors
::senora::señora
::senoras::señoras
::senorita::señorita
::senoritas::señoritas
::sinn fein::Sinn Féin
::smorgasbord::smörgåsbord
::smorgasbords::smörgåsbords
::smorgastarta::smörgåstårta
::soigne::soigné
::soiree::soirée
::soireed::soiréed
::soirees::soirées
::souffle::soufflé
::souffles::soufflés
::soupcon::soupçon
::soupcons::soupçons
::surstromming::surströmming
::tete-a-tete::tête-à-tête
::tete-a-tetes::tête-à-têtes
::touche::touché
::tourtiere::tourtière
::ubermensch::Übermensch
::ubermensches::Übermensches
::ventre a terre::ventre à terre
::vicuna::vicuña
::vin rose::vin rosé
::vins rose::vins rosé
::vis a vis::vis à vis
::vis-a-vis::vis-à-vis
::voila::voilà

;------------------------------------------------------------------------------
; Common Misspellings - the main list
;------------------------------------------------------------------------------
::htp:::http:
::http:\\::http://
::httpL::http:
::herf::href

::avengence::a vengeance
::adbandon::abandon
::abandonned::abandoned
::aberation::aberration
::aborigene::aborigine
::abortificant::abortifacient
::abbout::about
::abotu::about
::baout::about
::abouta::about a
::aboutit::about it
::aboutthe::about the
::abscence::absence
::absense::absence
::abcense::absense
::absolutly::absolutely
::asorbed::absorbed
::absorbsion::absorption
::absorbtion::absorption
::abundacies::abundances
::abundancies::abundances
::abundunt::abundant
::abutts::abuts
::acadmic::academic
::accademic::academic
::acedemic::academic
::acadamy::academy
::accademy::academy
::accelleration::acceleration
::acceptible::acceptable
::acceptence::acceptance
::accessable::accessible
::accension::accession
::accesories::accessories
::accesorise::accessorise
::accidant::accident
::accidentaly::accidentally
::accidently::accidentally
::acclimitization::acclimatization
::accomdate::accommodate
::accomodate::accommodate
::acommodate::accommodate
::acomodate::accommodate
::accomodated::accommodated
::accomodates::accommodates
::accomodating::accommodating
::accomodation::accommodation
::accomodations::accommodations
::accompanyed::accompanied
::acomplish::accomplish
::acomplished::accomplished
::acomplishment::accomplishment
::acomplishments::accomplishments
::accoring::according
::acording::according
::accordingto::according to
::acordingly::accordingly
::accordeon::accordion
::accordian::accordion
::acocunt::account
::acuracy::accuracy
::acccused::accused
::accussed::accused
::acused::accused
::acustom::accustom
::acustommed::accustomed
::achive::achieve
::achivement::achievement
::achivements::achievements
::acknowldeged::acknowledged
::acknowledgeing::acknowledging
::accoustic::acoustic
::acquiantence::acquaintance
::aquaintance::acquaintance
::aquiantance::acquaintance
::acquiantences::acquaintances
::accquainted::acquainted
::aquainted::acquainted
::aquire::acquire
::aquired::acquired
::aquiring::acquiring
::aquit::acquit
::acquited::acquitted
::aquitted::acquitted
::accross::across
::activly::actively
::activites::activities
::actualy::actually
::actualyl::actually
::adaption::adaptation
::adaptions::adaptations
::addtion::addition
::additinal::additional
::addtional::additional
::additinally::additionally
::addres::address
::adres::address
::adress::address
::addresable::addressable
::adresable::addressable
::adressable::addressable
::addresed::addressed
::adressed::addressed
::addressess::addresses
::addresing::addressing
::adresing::addressing
::adecuate::adequate
::adequit::adequate
::adequite::adequate
::adherance::adherence
::adhearing::adhering
::adminstered::administered
::adminstrate::administrate
::adminstration::administration
::admininistrative::administrative
::adminstrative::administrative
::adminstrator::administrator
::admissability::admissibility
::admissable::admissible
::addmission::admission
::admited::admitted
::admitedly::admittedly
::adolecent::adolescent
::addopt::adopt
::addopted::adopted
::addoptive::adoptive
::adavanced::advanced
::adantage::advantage
::advanage::advantage
::adventrous::adventurous
::advesary::adversary
::advertisment::advertisement
::advertisments::advertisements
::asdvertising::advertising
::adviced::advised
::aeriel::aerial
::aeriels::aerials
::areodynamics::aerodynamics
::asthetic::aesthetic
::asthetical::aesthetic
::asthetically::aesthetically
::afair::affair
::affilate::affiliate
::affilliate::affiliate
::afficionado::aficionado
::afficianados::aficionados
::afficionados::aficionados
::aforememtioned::aforementioned
::affraid::afraid
::afterthe::after the
::agian::again
::agin::again
::againnst::against
::agains::against
::agaisnt::against
::aganist::against
::agianst::against
::aginst::against
::againstt he::against the
::aggaravates::aggravates
::agregate::aggregate
::agregates::aggregates
::agression::aggression
::aggresive::aggressive
::agressive::aggressive
::agressively::aggressively
::agressor::aggressor
::agrieved::aggrieved
::agre::agree
::aggreed::agreed
::agred::agreed
::agreing::agreeing
::aggreement::agreement
::agreeement::agreement
::agreemeent::agreement
::agreemnet::agreement
::agreemnt::agreement
::agreemeents::agreements
::agreemnets::agreements
::agricuture::agriculture
::airbourne::airborne
::aicraft::aircraft
::aircaft::aircraft
::aircrafts::aircraft
::airrcraft::aircraft
::aiport::airport
::airporta::airports
::albiet::albeit
::alchohol::alcohol
::alchol::alcohol
::alcohal::alcohol
::alochol::alcohol
::alchoholic::alcoholic
::alcholic::alcoholic
::alcoholical::alcoholic
::algebraical::algebraic
::algoritm::algorithm
::algorhitms::algorithms
::algoritms::algorithms
::alientating::alienating
::alltime::all-time
::aledge::allege
::alege::allege
::alledge::allege
::aledged::alleged
::aleged::alleged
::alledged::alleged
::alledgedly::allegedly
::allegedely::allegedly
::allegedy::allegedly
::allegely::allegedly
::aledges::alleges
::alledges::alleges
::alegience::allegiance
::allegence::allegiance
::allegience::allegiance
::alliviate::alleviate
::allopone::allophone
::allopones::allophones
::alotted::allotted
::alowed::allowed
::alowing::allowing
::alusion::allusion
::almots::almost
::almsot::almost
::alomst::almost
::alonw::alone
::allready::already
::alraedy::already
::alreayd::already
::alreday::already
::aready::already
::alsation::Alsatian
::alsot::also
::aslo::also
::alternitives::alternatives
::allthough::although
::altho::although
::althought::although
::altough::although
::allwasy::always
::allwyas::always
::alwasy::always
::alwats::always
::alway::always
::alwyas::always
::amalgomated::amalgamated
::amatuer::amateur
::amerliorate::ameliorate
::ammend::amend
::ammended::amended
::admendment::amendment
::amendmant::amendment
::ammendment::amendment
::ammendments::amendments
::amoung::among
::amung::among
::amoungst::amongst
::ammount::amount
::ammused::amused
::analagous::analogous
::analogeous::analogous
::analitic::analytic
::anarchim::anarchism
::anarchistm::anarchism
::ansestors::ancestors
::ancestory::ancestry
::ancilliary::ancillary
::adn::and
::anbd::and
::anmd::and
::andone::and one
::andt he::and the
::andteh::and the
::andthe::and the
::androgenous::androgynous
::androgeny::androgyny
::anihilation::annihilation
::aniversary::anniversary
::annouced::announced
::anounced::announced
::anual::annual
::annualy::annually
::annuled::annulled
::anulled::annulled
::annoint::anoint
::annointed::anointed
::annointing::anointing
::annoints::anoints
::anomolies::anomalies
::anomolous::anomalous
::anomoly::anomaly
::anonimity::anonymity
::anohter::another
::anotehr::another
::anothe::another
::anwsered::answered
::antartic::antarctic
::anthromorphisation::anthropomorphisation
::anthromorphization::anthropomorphization
::anti-semetic::anti-Semitic
::anyother::any other
::anytying::anything
::anyhwere::anywhere
::appart::apart
::aparment::apartment
::appartment::apartment
::appartments::apartments
::apenines::Apennines
::appenines::Apennines
::apolegetics::apologetics
::appologies::apologies
::appology::apology
::aparent::apparent
::apparant::apparent
::apparrent::apparent
::apparantly::apparently
::appealling::appealing
::appeareance::appearance
::appearence::appearance
::apperance::appearance
::apprearance::appearance
::appearences::appearances
::apperances::appearances
::appeares::appears
::aplication::application
::applicaiton::application
::applicaitons::applications
::aplied::applied
::applyed::applied
::appointiment::appointment
::apprieciate::appreciate
::aprehensive::apprehensive
::approachs::approaches
::appropiate::appropriate
::appropraite::appropriate
::appropropiate::appropriate
::approrpiate::appropriate
::approrpriate::appropriate
::apropriate::appropriate
::approproximate::approximate
::aproximate::approximate
::approxamately::approximately
::approxiately::approximately
::approximitely::approximately
::aproximately::approximately
::arbitarily::arbitrarily
::abritrary::arbitrary
::arbitary::arbitrary
::arbouretum::arboretum
::archiac::archaic
::archimedian::Archimedean
::archictect::architect
::archetectural::architectural
::architectual::architectural
::archetecturally::architecturally
::architechturally::architecturally
::archetecture::architecture
::architechture::architecture
::architechtures::architectures
::arn't::aren't
::argubly::arguably
::armamant::armament
::armistace::armistice
::arised::arose
::arond::around
::aroud::around
::arround::around
::arund::around
::aranged::arranged
::arangement::arrangement
::arrangment::arrangement
::arrangments::arrangements
::arival::arrival
::artical::article
::artice::article
::articel::article
::artifical::artificial
::artifically::artificially
::artillary::artillery
::asthe::as the
::aswell::as well
::asetic::ascetic
::aisian::Asian
::asside::aside
::askt he::ask the
::asphyxation::asphyxiation
::assisnate::assassinate
::assassintation::assassination
::assosication::assassination
::asssassans::assassins
::assualt::assault
::assualted::assaulted
::assemple::assemble
::assertation::assertion
::assesment::assessment
::asign::assign
::assit::assist
::assistent::assistant
::assitant::assistant
::assoicate::associate
::assoicated::associated
::assoicates::associates
::assocation::association
::asume::assume
::asteriod::asteroid
::atthe::at the
::athiesm::atheism
::athiest::atheist
::atheistical::atheistic
::athenean::Athenian
::atheneans::Athenians
::atmospher::atmosphere
::attrocities::atrocities
::attatch::attach
::atain::attain
::attemp::attempt
::attemt::attempt
::attemped::attempted
::attemted::attempted
::attemting::attempting
::attemts::attempts
::attendence::attendance
::attendent::attendant
::attendents::attendants
::attened::attended
::atention::attention
::attension::attention
::attentioin::attention
::attitide::attitude
::atorney::attorney
::attributred::attributed
::audeince::audience
::audiance::audience
::austrailia::Australia
::austrailian::Australian
::australian::Australian
::auther::author
::autor::author
::authorative::authoritative
::authoritive::authoritative
::authorites::authorities
::authoritiers::authorities
::authrorities::authorities
::authorithy::authority
::autority::authority
::authobiographic::autobiographic
::authobiography::autobiography
::autochtonous::autochthonous
::autoctonous::autochthonous
::automaticly::automatically
::automibile::automobile
::automonomous::autonomous
::auxillaries::auxiliaries
::auxilliaries::auxiliaries
::auxilary::auxiliary
::auxillary::auxiliary
::auxilliary::auxiliary
::availablility::availability
::availaible::available
::availalbe::available
::availble::available
::availiable::available
::availible::available
::avalable::available
::avaliable::available
::avilable::available
::avalance::avalanche
::averageed::averaged
::avation::aviation
::awared::awarded
::awya::away
::aywa::away
::abck::back
::bakc::back
::bcak::back
::backgorund::background
::backrounds::backgrounds
::balence::balance
::ballance::balance
::banannas::bananas
::bandwith::bandwidth
::bankrupcy::bankruptcy
::banruptcy::bankruptcy
::barbeque::barbecue
::basicaly::basically
::basicly::basically
::cattleship::battleship
::bve::be
::eb::be
::beachead::beachhead
::beatiful::beautiful
::beautyfull::beautiful
::beutiful::beautiful
::becamae::became
::baceause::because
::beacuse::because
::becasue::because
::becaus::because
::beccause::because
::becouse::because
::becuase::because
::becuse::because
::becausea::because a
::becauseof::because of
::becausethe::because the
::becauseyou::because you
::becoe::become
::becomeing::becoming
::becomming::becoming
::bedore::before
::befoer::before
::begginer::beginner
::begginers::beginners
::beggining::beginning
::begining::beginning
::beginining::beginning
::beginnig::beginning
::begginings::beginnings
::beggins::begins
::behavour::behaviour
::beng::being
::beleagured::beleaguered
::beligum::belgium
::beleif::belief
::beleiev::believe
::beleieve::believe
::beleive::believe
::belive::believe
::beleived::believed
::belived::believed
::beleives::believes
::beleiving::believing
::belligerant::belligerent
::bellweather::bellwether
::bemusemnt::bemusement
::benefical::beneficial
::benificial::beneficial
::beneficary::beneficiary
::benifit::benefit
::benifits::benefits
::bergamont::bergamot
::bernouilli::Bernoulli
::beseige::besiege
::beseiged::besieged
::beseiging::besieging
::beastiality::bestiality
::betweeen::between
::betwen::between
::bewteen::between
::inbetween::between
::vetween::between
::bicep::biceps
::bilateraly::bilaterally
::billingualism::bilingualism
::binominal::binomial
::bizzare::bizarre
::blaim::blame
::blaimed::blamed
::blessure::blessing
::blitzkreig::Blitzkrieg
::bodydbuilder::bodybuilder
::bombardement::bombardment
::bombarment::bombardment
::bonnano::Bonanno
::bondary::boundary
::boundry::boundary
::boxs::boxes
::brasillian::Brazilian
::breakthough::breakthrough
::breakthroughts::breakthroughs
::brethen::brethren
::bretheren::brethren
::breif::brief
::breifly::briefly
::briliant::brilliant
::brillant::brilliant
::brimestone::brimstone
::britian::Britain
::brittish::British
::broacasted::broadcast
::brodcast::broadcast
::broadacasting::broadcasting
::broady::broadly
::borke::broke
::buddah::Buddha
::bouy::buoy
::bouyancy::buoyancy
::buoancy::buoyancy
::bouyant::buoyant
::boyant::buoyant
::beaurocracy::bureaucracy
::beaurocratic::bureaucratic
::burried::buried
::buisness::business
::busness::business
::bussiness::business
::busineses::businesses
::buisnessman::businessman
::butthe::but the
::byt he::by the
::ceasar::Caesar
::casion::caisson
::caluclate::calculate
::caluculate::calculate
::calulate::calculate
::calcullated::calculated
::caluclated::calculated
::caluculated::calculated
::calulated::calculated
::calculs::calculus
::calander::calendar
::calenders::calendars
::califronia::California
::califronian::Californian
::caligraphy::calligraphy
::callipigian::callipygian
::cambrige::Cambridge
::camoflage::camouflage
::campain::campaign
::campains::campaigns
::acn::can
::cna::can
::cxan::can
::can't of::can't have
::candadate::candidate
::candiate::candidate
::candidiate::candidate
::candidtae::candidate
::candidtaes::candidates
::cannister::canister
::cannisters::canisters
::cannnot::cannot
::cannonical::canonical
::cantalope::cantaloupe
::caperbility::capability
::capible::capable
::capetown::Cape Town
::captial::capital
::captued::captured
::capturd::captured
::carcas::carcass
::carreer::career
::carrers::careers
::carefull::careful
::carribbean::Caribbean
::carribean::Caribbean
::careing::caring
::carmalite::Carmelite
::carniverous::carnivorous
::carthagian::Carthaginian
::cartilege::cartilage
::cartilidge::cartilage
::carthographer::cartographer
::cartdridge::cartridge
::cartrige::cartridge
::casette::cassette
::cassawory::cassowary
::cassowarry::cassowary
::casulaties::casualties
::causalities::casualties
::casulaty::casualty
::categiory::category
::ctaegory::category
::catterpilar::caterpillar
::catterpilars::caterpillars
::cathlic::catholic
::catholocism::catholicism
::caucasion::Caucasian
::cacuses::caucuses
::cieling::ceiling
::cellpading::cellpadding
::celcius::Celsius
::cemetaries::cemeteries
::cementary::cemetery
::cemetarey::cemetery
::cemetary::cemetery
::sensure::censure
::cencus::census
::cententenial::centennial
::centruies::centuries
::centruy::century
::cerimonial::ceremonial
::cerimonies::ceremonies
::cerimonious::ceremonious
::cerimony::ceremony
::ceromony::ceremony
::certian::certain
::certainity::certainty
::chariman::chairman
::challange::challenge
::challege::challenge
::challanged::challenged
::challanges::challenges
::chalenging::challenging
::champange::champagne
::chaneg::change
::chnage::change
::changable::changeable
::chanegs::changes
::changeing::changing
::changng::changing
::caharcter::character
::carachter::character
::charachter::character
::charactor::character
::charecter::character
::charector::character
::chracter::character
::caracterised::characterised
::charaterised::characterised
::charactersistic::characteristic
::charistics::characteristics
::caracterized::characterized
::charaterized::characterized
::cahracters::characters
::charachters::characters
::charactors::characters
::carismatic::charismatic
::charasmatic::charismatic
::chartiable::charitable
::caht::chat
::chekc::check
::chemcial::chemical
::chemcially::chemically
::chemicaly::chemically
::chemestry::chemistry
::cheif::chief
::childbird::childbirth
::childen::children
::childrens::children's
::chilli::chili
::choosen::chosen
::chuch::church
::curch::church
::churchs::churches
::cincinatti::Cincinnati
::cincinnatti::Cincinnati
::circut::circuit
::ciricuit::circuit
::curcuit::circuit
::circulaton::circulation
::circumsicion::circumcision
::sercumstances::circumstances
::cirtus::citrus
::civillian::civilian
::claimes::claims
::clas::class
::clasic::classic
::clasical::classical
::clasically::classically
::claer::clear
::cleareance::clearance
::claered::cleared
::claerer::clearer
::claerly::clearly
::cliant::client
::clincial::clinical
::clinicaly::clinically
::caost::coast
::coctail::cocktail
::cognizent::cognizant
::co-incided::coincided
::coincedentally::coincidentally
::colaborations::collaborations
::collaberative::collaborative
::colateral::collateral
::collegue::colleague
::collegues::colleagues
::collectable::collectible
::colection::collection
::collecton::collection
::colelctive::collective
::collonies::colonies
::colonisators::colonisers
::colonizators::colonizers
::collonade::colonnade
::collony::colony
::collosal::colossal
::colum::column
::combintation::combination
::combanations::combinations
::combinatins::combinations
::combusion::combustion
::comback::comeback
::commedic::comedic
::confortable::comfortable
::comming::coming
::commadn::command
::comander::commander
::comando::commando
::comandos::commandos
::commandoes::commandos
::comemmorate::commemorate
::commemmorate::commemorate
::commmemorated::commemorated
::comemmorates::commemorates
::commemmorating::commemorating
::comemoretion::commemoration
::commemerative::commemorative
::commerorative::commemorative
::commerical::commercial
::commericial::commercial
::commerically::commercially
::commericially::commercially
::comission::commission
::commision::commission
::comissioned::commissioned
::commisioned::commissioned
::comissioner::commissioner
::commisioner::commissioner
::comissioning::commissioning
::commisioning::commissioning
::comissions::commissions
::commisions::commissions
::comit::commit
::committment::commitment
::committments::commitments
::comited::committed
::comitted::committed
::commited::committed
::comittee::committee
::commitee::committee
::committe::committee
::committy::committee
::comiting::committing
::comitting::committing
::commiting::committing
::commongly::commonly
::commonweath::commonwealth
::comunicate::communicate
::comminication::communication
::communciation::communication
::communiation::communication
::commuications::communications
::commuinications::communications
::communites::communities
::comunity::community
::comanies::companies
::comapnies::companies
::comany::company
::comapany::company
::comapny::company
::company;s::company's
::comparitive::comparative
::comparitively::comparatively
::compair::compare
::comparision::comparison
::comparisions::comparisons
::compability::compatibility
::compatiable::compatible
::compensantion::compensation
::competance::competence
::competant::competent
::compitent::competent
::competitiion::competition
::compeitions::competitions
::competative::competitive
::competive::competitive
::competiveness::competitiveness
::copmetitors::competitors
::complier::compiler
::compleated::completed
::completedthe::completed the
::competely::completely
::compleatly::completely
::completelyl::completely
::completly::completely
::compleatness::completeness
::completness::completeness
::completetion::completion
::componant::component
::composate::composite
::comphrehensive::comprehensive
::comprimise::compromise
::compulsary::compulsory
::compulsery::compulsory
::cmoputer::computer
::coputer::computer
::computarised::computerised
::computarized::computerized
::concieted::conceited
::concieve::conceive
::concieved::conceived
::consentrate::concentrate
::consentrated::concentrated
::consentrates::concentrates
::consept::concept
::consern::concern
::conserned::concerned
::conserning::concerning
::comdemnation::condemnation
::condamned::condemned
::condemmed::condemned
::condidtion::condition
::condidtions::conditions
::conditionsof::conditions of
::condolances::condolences
::conferance::conference
::confidental::confidential
::confidentally::confidentially
::confids::confides
::configureable::configurable
::confirmmation::confirmation
::coform::conform
::congradulations::congratulations
::congresional::congressional
::conjecutre::conjecture
::conjuction::conjunction
::conected::connected
::conneticut::Connecticut
::conection::connection
::conived::connived
::cannotation::connotation
::cannotations::connotations
::conotations::connotations
::conquerd::conquered
::conqured::conquered
::conquerer::conqueror
::conquerers::conquerors
::concious::conscious
::consious::conscious
::conciously::consciously
::conciousness::consciousness
::consciouness::consciousness
::consiciousness::consciousness
::consicousness::consciousness
::consectutive::consecutive
::concensus::consensus
::conesencus::consensus
::conscent::consent
::consequeseces::consequences
::consenquently::consequently
::consequentually::consequently
::conservitive::conservative
::concider::consider
::consdider::consider
::considerit::considerate
::considerite::considerate
::concidered::considered
::consdidered::considered
::consdiered::considered
::considerd::considered
::consideres::considered
::concidering::considering
::conciders::considers
::consistant::consistent
::consistantly::consistently
::consolodate::consolidate
::consolodated::consolidated
::consonent::consonant
::consonents::consonants
::consorcium::consortium
::conspiracys::conspiracies
::conspiricy::conspiracy
::conspiriator::conspirator
::constatn::constant
::constanly::constantly
::constarnation::consternation
::consituencies::constituencies
::consituency::constituency
::constituant::constituent
::constituants::constituents
::consituted::constituted
::consitution::constitution
::constituion::constitution
::costitution::constitution
::consitutional::constitutional
::constituional::constitutional
::constaints::constraints
::consttruction::construction
::constuction::construction
::contruction::construction
::consulant::consultant
::consultent::consultant
::consumber::consumer
::consumate::consummate
::consumated::consummated
::comntain::contain
::comtain::contain
::comntains::contains
::comtains::contains
::containes::contains
::countains::contains
::contaiminate::contaminate
::contemporaneus::contemporaneous
::contamporaries::contemporaries
::contamporary::contemporary
::contempoary::contemporary
::contempory::contemporary
::contendor::contender
::constinually::continually
::contined::continued
::continueing::continuing
::continous::continuous
::continously::continuously
::contritutions::contributions
::contributer::contributor
::contributers::contributors
::controll::control
::controled::controlled
::controling::controlling
::controlls::controls
::contravercial::controversial
::controvercial::controversial
::controversal::controversial
::controvertial::controversial
::controveries::controversies
::contraversy::controversy
::controvercy::controversy
::controvery::controversy
::conveinent::convenient
::convienient::convenient
::convential::conventional
::convertion::conversion
::convertor::converter
::convertors::converters
::convertable::convertible
::convertables::convertibles
::conveyer::conveyor
::conviced::convinced
::cooparate::cooperate
::cooporate::cooperate
::coordiantion::coordination
::cpoy::copy
::copywrite::copyright
::coridal::cordial
::corparate::corporate
::corproation::corporation
::coorperations::corporations
::corperations::corporations
::corproations::corporations
::correcters::correctors
::corrispond::correspond
::corrisponded::corresponded
::correspondant::correspondent
::corrispondant::correspondent
::correspondants::correspondents
::corrispondants::correspondents
::correponding::corresponding
::correposding::corresponding
::corrisponding::corresponding
::corrisponds::corresponds
::corridoors::corridors
::corosion::corrosion
::corruptable::corruptible
::cotten::cotton
::coudl::could
::could of::could have
::couldthe::could the
::coudln't::couldn't
::coudn't::couldn't
::couldnt::couldn't
::coucil::council
::counries::countries
::countires::countries
::ocuntries::countries
::ocuntry::country
::coururier::courier
::convenant::covenant
::creaeted::created
::creedence::credence
::criterias::criteria
::critereon::criterion
::crtical::critical
::critised::criticised
::criticing::criticising
::criticists::critics
::crockodiles::crocodiles
::crucifiction::crucifixion
::crusies::cruises
::crystalisation::crystallisation
::culiminating::culminating
::cumulatative::cumulative
::currenly::currently
::ciriculum::curriculum
::curriculem::curriculum
::cusotmer::customer
::cutsomer::customer
::cusotmers::customers
::cutsomers::customers
::cxan::cyan
::cilinder::cylinder
::cyclinder::cylinder
::dakiri::daiquiri
::dalmation::dalmatian
::danceing::dancing
::dardenelles::Dardanelles
::dael::deal
::debateable::debatable
::decaffinated::decaffeinated
::decathalon::decathlon
::decieved::deceived
::decideable::decidable
::deside::decide
::decidely::decidedly
::ecidious::deciduous
::decison::decision
::descision::decision
::desicion::decision
::desision::decision
::decisons::decisions
::descisions::decisions
::desicions::decisions
::desisions::decisions
::decomissioned::decommissioned
::decomposit::decompose
::decomposited::decomposed
::decomposits::decomposes
::decompositing::decomposing
::decress::decrees
::deafult::default
::defendent::defendant
::defendents::defendants
::defencive::defensive
::deffensively::defensively
::definance::defiance
::deffine::define
::deffined::defined
::definining::defining
::definate::definite
::definit::definite
::definately::definitely
::definatly::definitely
::definetly::definitely
::definitly::definitely
::definiton::definition
::defintion::definition
::degredation::degradation
::degrate::degrade
::dieties::deities
::diety::deity
::delagates::delegates
::deliberatly::deliberately
::delerious::delirious
::delusionally::delusively
::devels::delves
::damenor::demeanor
::demenor::demeanor
::damenor::demeanour
::damenour::demeanour
::demenour::demeanour
::demorcracy::democracy
::demographical::demographic
::demolision::demolition
::demostration::demonstration
::denegrating::denigrating
::densly::densely
::deparment::department
::deptartment::department
::dependance::dependence
::dependancy::dependency
::dependant::dependent
::despict::depict
::derivitive::derivative
::deriviated::derived
::dirived::derived
::derogitory::derogatory
::decendant::descendant
::decendent::descendant
::decendants::descendants
::decendents::descendants
::descendands::descendants
::decribe::describe
::discribe::describe
::decribed::described
::descibed::described
::discribed::described
::decribes::describes
::descriibes::describes
::discribes::describes
::decribing::describing
::discribing::describing
::descriptoin::description
::descripton::description
::descripters::descriptors
::dessicated::desiccated
::disign::design
::desgined::designed
::dessigned::designed
::desigining::designing
::desireable::desirable
::desktiop::desktop
::dispair::despair
::desparate::desperate
::despiration::desperation
::dispicable::despicable
::dispite::despite
::destablised::destabilised
::destablized::destabilized
::desinations::destinations
::desitned::destined
::destory::destroy
::desctruction::destruction
::distruction::destruction
::distructive::destructive
::detatched::detached
::detailled::detailed
::deatils::details
::dectect::detect
::deteriate::deteriorate
::deteoriated::deteriorated
::deterioriating::deteriorating
::determinining::determining
::detremental::detrimental
::devasted::devastated
::devestated::devastated
::devestating::devastating
::devistating::devastating
::devellop::develop
::devellops::develop
::develloped::developed
::developped::developed
::develloper::developer
::developor::developer
::develeoprs::developers
::devellopers::developers
::developors::developers
::develloping::developing
::delevopment::development
::devellopment::development
::develpment::development
::devolopement::development
::devellopments::developments
::divice::device
::diablical::diabolical
::diamons::diamonds
::diarhea::diarrhoea
::dichtomy::dichotomy
::didnot::did not
::didint::didn't
::didnt::didn't
::differance::difference
::diferences::differences
::differances::differences
::difefrent::different
::diferent::different
::diferrent::different
::differant::different
::differemt::different
::differnt::different
::diffrent::different
::differentiatiations::differentiations
::diffcult::difficult
::diffculties::difficulties
::dificulties::difficulties
::diffculty::difficulty
::difficulity::difficulty
::dificulty::difficulty
::delapidated::dilapidated
::dimention::dimension
::dimentional::dimensional
::dimesnional::dimensional
::dimenions::dimensions
::dimentions::dimensions
::diminuitive::diminutive
::diosese::diocese
::diptheria::diphtheria
::diphtong::diphthong
::dipthong::diphthong
::diphtongs::diphthongs
::dipthongs::diphthongs
::diplomancy::diplomacy
::directiosn::direction
::driectly::directly
::directer::director
::directers::directors
::disagreeed::disagreed
::dissagreement::disagreement
::disapear::disappear
::dissapear::disappear
::dissappear::disappear
::dissapearance::disappearance
::disapeared::disappeared
::disappearred::disappeared
::dissapeared::disappeared
::dissapearing::disappearing
::dissapears::disappears
::dissappears::disappears
::dissappointed::disappointed
::disapointing::disappointing
::disaproval::disapproval
::dissarray::disarray
::diaster::disaster
::disasterous::disastrous
::disatrous::disastrous
::diciplin::discipline
::disiplined::disciplined
::unconfortability::discomfort
::diconnects::disconnects
::discontentment::discontent
::dicover::discover
::disover::discover
::dicovered::discovered
::discoverd::discovered
::dicovering::discovering
::dicovers::discovers
::dicovery::discovery
::descuss::discuss
::dicussed::discussed
::desease::disease
::disenchanged::disenchanted
::desintegrated::disintegrated
::desintegration::disintegration
::disobediance::disobedience
::dissobediance::disobedience
::dissobedience::disobedience
::disobediant::disobedient
::dissobediant::disobedient
::dissobedient::disobedient
::desorder::disorder
::desoriented::disoriented
::disparingly::disparagingly
::despatched::dispatched
::dispell::dispel
::dispeled::dispelled
::dispeling::dispelling
::dispells::dispels
::dispence::dispense
::dispenced::dispensed
::dispencing::dispensing
::diaplay::display
::dispaly::display
::unplease::displease
::dispostion::disposition
::disproportiate::disproportionate
::disputandem::disputandum
::disatisfaction::dissatisfaction
::disatisfied::dissatisfied
::disemination::dissemination
::disolved::dissolved
::dissonent::dissonant
::disctinction::distinction
::distiction::distinction
::disctinctive::distinctive
::distingish::distinguish
::distingished::distinguished
::distingquished::distinguished
::distingishes::distinguishes
::distingishing::distinguishing
::ditributed::distributed
::distribusion::distribution
::distrubution::distribution
::disricts::districts
::devide::divide
::devided::divided
::divison::division
::divisons::divisions
::docrines::doctrines
::doctines::doctrines
::doccument::document
::docuemnt::document
::documetn::document
::documnet::document
::documenatry::documentary
::doccumented::documented
::doccuments::documents
::docuement::documents
::documnets::documents
::doens::does
::doese::does
::doe snot::does not ; *could* be legitimate... but very unlikely!
::doens't::doesn't
::doesnt::doesn't
::dosen't::doesn't
::dosn't::doesn't
::doign::doing
::doimg::doing
::doind::doing
::donig::doing
::dollers::dollars
::dominent::dominant
::dominiant::dominant
::dominaton::domination
::do'nt::don't
::dont::don't
::don't no::don't know
::doulbe::double
::dowloads::downloads
::dramtic::dramatic
::draughtman::draughtsman
::dravadian::Dravidian
::deram::dream
::derams::dreams
::dreasm::dreams
::drnik::drink
::driveing::driving
::drummless::drumless
::druming::drumming
::drunkeness::drunkenness
::dukeship::dukedom
::dumbell::dumbbell
::dupicate::duplicate
::durig::during
::durring::during
::duting::during
::dieing::dying
::eahc::each
::eachotehr::eachother
::ealier::earlier
::earlies::earliest
::eearly::early
::earnt::earned
::ecclectic::eclectic
::eclispe::eclipse
::ecomonic::economic
::eceonomy::economy
::esctasy::ecstasy
::eles::eels
::effeciency::efficiency
::efficency::efficiency
::effecient::efficient
::efficent::efficient
::effeciently::efficiently
::efficently::efficiently
::effulence::effluence
::efort::effort
::eforts::efforts
::aggregious::egregious
::eight o::eight o
::eigth::eighth
::eiter::either
::ellected::elected
::electrial::electrical
::electricly::electrically
::electricty::electricity
::eletricity::electricity
::elementay::elementary
::elimentary::elementary
::elphant::elephant
::elicided::elicited
::eligable::eligible
::eleminated::eliminated
::eleminating::eliminating
::alse::else
::esle::else
::eminate::emanate
::eminated::emanated
::embargos::embargoes
::embarras::embarrass
::embarrased::embarrassed
::embarrasing::embarrassing
::embarrasment::embarrassment
::embezelled::embezzled
::emblamatic::emblematic
::emmigrated::emigrated
::emmisaries::emissaries
::emmisarries::emissaries
::emmisarry::emissary
::emmisary::emissary
::emision::emission
::emmision::emission
::emmisions::emissions
::emited::emitted
::emmited::emitted
::emmitted::emitted
::emiting::emitting
::emmiting::emitting
::emmitting::emitting
::emphsis::emphasis
::emphaised::emphasised
::emphysyma::emphysema
::emperical::empirical
::imploys::employs
::enameld::enamelled
::encouraing::encouraging
::encryptiion::encryption
::encylopedia::encyclopedia
::endevors::endeavors
::endevour::endeavour
::endevours::endeavours
::endig::ending
::endolithes::endoliths
::enforceing::enforcing
::engagment::engagement
::engeneer::engineer
::engieneer::engineer
::engeneering::engineering
::engieneers::engineers
::enlish::English
::enchancement::enhancement
::emnity::enmity
::enourmous::enormous
::enourmously::enormously
::enought::enough
::ensconsed::ensconced
::entaglements::entanglements
::intertaining::entertaining
::enteratinment::entertainment
::entitlied::entitled
::entitity::entity
::entrepeneur::entrepreneur
::entrepeneurs::entrepreneurs
::intrusted::entrusted
::enviornment::environment
::enviornmental::environmental
::enviornmentalist::environmentalist
::enviornmentally::environmentally
::enviornments::environments
::envrionments::environments
::epsiode::episode
::epidsodes::episodes
::equitorial::equatorial
::equilibium::equilibrium
::equilibrum::equilibrium
::equippment::equipment
::equiped::equipped
::equialent::equivalent
::equivalant::equivalent
::equivelant::equivalent
::equivelent::equivalent
::equivilant::equivalent
::equivilent::equivalent
::equivlalent::equivalent
::eratic::erratic
::eratically::erratically
::eraticly::erratically
::errupted::erupted
::especally::especially
::especialy::especially
::especialyl::especially
::espesially::especially
::expecially::especially
::expresso::espresso
::essense::essence
::esential::essential
::essencial::essential
::essentail::essential
::essentual::essential
::essesital::essential
::essentialy::essentially
::estabishes::establishes
::establising::establishing
::esitmated::estimated
::ect::etc
::ethnocentricm::ethnocentrism
::europian::European
::eurpean::European
::eurpoean::European
::europians::Europeans
::evenhtually::eventually
::eventally::eventually
::eventially::eventually
::eventualy::eventually
::eveyr::every
::everytime::every time
::everthing::everything
::evidentally::evidently
::efel::evil
::envolutionary::evolutionary
::exerbate::exacerbate
::exerbated::exacerbated
::excact::exact
::exagerate::exaggerate
::exagerrate::exaggerate
::exagerated::exaggerated
::exagerrated::exaggerated
::exagerates::exaggerates
::exagerrates::exaggerates
::exagerating::exaggerating
::exagerrating::exaggerating
::exhalted::exalted
::examinated::examined
::exemple::example
::exmaple::example
::excedded::exceeded
::exeedingly::exceedingly
::excell::excel
::excellance::excellence
::excelent::excellent
::excellant::excellent
::exelent::excellent
::exellent::excellent
::excells::excels
::exept::except
::exeptional::exceptional
::exerpt::excerpt
::exerpts::excerpts
::excange::exchange
::exchagne::exchange
::exhcange::exchange
::exchagnes::exchanges
::exhcanges::exchanges
::exchanching::exchanging
::excitment::excitement
::exicting::exciting
::exludes::excludes
::exculsivly::exclusively
::excecute::execute
::excecuted::executed
::exectued::executed
::excecutes::executes
::excecuting::executing
::excecution::execution
::exection::execution
::exampt::exempt
::excercise::exercise
::exersize::exercise
::exerciese::exercises
::execising::exercising
::extered::exerted
::exhibtion::exhibition
::exibition::exhibition
::exibitions::exhibitions
::exliled::exiled
::excisted::existed
::existance::existence
::existince::existence
::existant::existent
::exisiting::existing
::exonorate::exonerate
::exoskelaton::exoskeleton
::exapansion::expansion
::expeced::expected
::expeditonary::expeditionary
::expiditions::expeditions
::expell::expel
::expells::expels
::experiance::experience
::experienc::experience
::expierence::experience
::exprience::experience
::experianced::experienced
::exprienced::experienced
::expeiments::experiments
::expalin::explain
::explaning::explaining
::explaination::explanation
::explictly::explicitly
::explotation::exploitation
::exploititive::exploitative
::exressed::expressed
::expropiated::expropriated
::expropiation::expropriation
::extention::extension
::extentions::extensions
::exerternal::external
::exinct::extinct
::extradiction::extradition
::extrordinarily::extraordinarily
::extrordinary::extraordinary
::extravagent::extravagant
::extemely::extremely
::extrememly::extremely
::extremly::extremely
::extermist::extremist
::extremeophile::extremophile
::fascitious::facetious
::facillitate::facilitate
::facilites::facilities
::farenheit::Fahrenheit
::familair::familiar
::familar::familiar
::familliar::familiar
::fammiliar::familiar
::familes::families
::fimilies::families
::famoust::famous
::fanatism::fanaticism
::facia::fascia
::fascitis::fasciitis
::facinated::fascinated
::facist::fascist
::favoutrable::favourable
::feasable::feasible
::faeture::feature
::faetures::features
::febuary::February
::fedreally::federally
::efel::feel
::fertily::fertility
::fued::feud
::fwe::few
::ficticious::fictitious
::fictious::fictitious
::feild::field
::feilds::fields
::fiercly::fiercely
::firey::fiery
::fightings::fighting
::filiament::filament
::fiel::file
::fiels::files
::fianlly::finally
::finaly::finally
::finalyl::finally
::finacial::financial
::financialy::financially
::fidn::find
::fianite::finite
::firts::first
::fisionable::fissionable
::ficed::fixed
::flamable::flammable
::flawess::flawless
::flemmish::Flemish
::glight::flight
::fluorish::flourish
::florescent::fluorescent
::flourescent::fluorescent
::flouride::fluoride
::foucs::focus
::focussed::focused
::focusses::focuses
::focussing::focusing
::follwo::follow
::follwoing::following
::folowing::following
::formalhaut::Fomalhaut
::foootball::football
::fora::for a
::forthe::for the
::forbad::forbade
::forbiden::forbidden
::forhead::forehead
::foriegn::foreign
::formost::foremost
::forunner::forerunner
::forsaw::foresaw
::forseeable::foreseeable
::fortelling::foretelling
::foreward::foreword
::forfiet::forfeit
::formallise::formalise
::formallised::formalised
::formallize::formalize
::formallized::formalized
::formaly::formally
::fomed::formed
::fromed::formed
::formelly::formerly
::fourties::forties
::fourty::forty
::forwrd::forward
::foward::forward
::forwrds::forwards
::fowards::forwards
::faught::fought
::fougth::fought
::foudn::found
::foundaries::foundries
::foundary::foundry
::fouth::fourth
::fransiscan::Franciscan
::fransiscans::Franciscans
::frequentily::frequently
::freind::friend
::freindly::friendly
::firends::friends
::freinds::friends
::frmo::from
::frome::from
::fromt he::from the
::fromthe::from the
::froniter::frontier
::fufill::fulfill
::fufilled::fulfilled
::fulfiled::fulfilled
::funtion::function
::fundametal::fundamental
::fundametals::fundamentals
::furneral::funeral
::funguses::fungi
::firc::furc
::furuther::further
::futher::further
::futhermore::furthermore
::galatic::galactic
::galations::Galatians
::gallaxies::galaxies
::galvinised::galvanised
::galvinized::galvanized
::gameboy::Game Boy
::ganes::games
::ghandi::Gandhi
::ganster::gangster
::garnison::garrison
::guage::gauge
::geneological::genealogical
::geneologies::genealogies
::geneology::genealogy
::gemeral::general
::generaly::generally
::generatting::generating
::genialia::genitalia
::gentlemens::gentlemen's
::geographicial::geographical
::geometrician::geometer
::geometricians::geometers
::geting::getting
::gettin::getting
::guilia::Giulia
::guiliani::Giuliani
::guilio::Giulio
::guiseppe::Giuseppe
::gievn::given
::giveing::giving
::glace::glance
::gloabl::global
::gnawwed::gnawed
::godess::goddess
::godesses::goddesses
::godounov::Godunov
::goign::going
::gonig::going
::oging::going
::giid::good
::gothenberg::Gothenburg
::gottleib::Gottlieb
::goverance::governance
::govement::government
::govenment::government
::govenrment::government
::goverment::government
::governmnet::government
::govorment::government
::govornment::government
::govermental::governmental
::govormental::governmental
::gouvener::governor
::governer::governor
::gracefull::graceful
::graffitti::graffiti
::grafitti::graffiti
::grammer::grammar
::gramatically::grammatically
::grammaticaly::grammatically
::greatful::grateful
::greatfully::gratefully
::gratuitious::gratuitous
::gerat::great
::graet::great
::grat::great
::gridles::griddles
::greif::grief
::gropu::group
::gruop::group
::gruops::groups
::grwo::grow
::guadulupe::Guadalupe
::gunanine::guanine
::gauarana::guarana
::gaurantee::guarantee
::gaurentee::guarantee
::guarentee::guarantee
::gurantee::guarantee
::gauranteed::guaranteed
::gaurenteed::guaranteed
::guarenteed::guaranteed
::guranteed::guaranteed
::gaurantees::guarantees
::gaurentees::guarantees
::guarentees::guarantees
::gurantees::guarantees
::gaurd::guard
::guatamala::Guatemala
::guatamalan::Guatemalan
::guidence::guidance
::guiness::Guinness
::guttaral::guttural
::gutteral::guttural
::gusy::guys
::habaeus::habeas
::habeus::habeas
::habsbourg::Habsburg
::hda::had
::hadbeen::had been
::haemorrage::haemorrhage
::hallowean::Halloween
::ahppen::happen
::hapen::happen
::hapened::happened
::happend::happened
::happended::happened
::happenned::happened
::hapening::happening
::hapens::happens
::harras::harass
::harased::harassed
::harrased::harassed
::harrassed::harassed
::harrasses::harassed
::harases::harasses
::harrases::harasses
::harrasing::harassing
::harrassing::harassing
::harassement::harassment
::harrasment::harassment
::harrassment::harassment
::harrasments::harassments
::harrassments::harassments
::hace::hare
::hsa::has
::hasbeen::has been
::hasnt::hasn't
::ahev::have
::ahve::have
::haev::have
::hvae::have
::havebeen::have been
::haveing::having
::hvaing::having
::hge::he
::hesaid::he said
::hewas::he was
::headquater::headquarter
::headquatered::headquartered
::headquaters::headquarters
::healthercare::healthcare
::heathy::healthy
::heared::heard
::hearign::hearing
::herat::heart
::haviest::heaviest
::heidelburg::Heidelberg
::hieght::height
::hier::heir
::heirarchy::heirarchy
::helment::helmet
::halp::help
::hlep::help
::helpped::helped
::helpfull::helpful
::hemmorhage::hemorrhage
::ehr::her
::ehre::here
::here;s::here's
::heridity::heredity
::heroe::hero
::heros::heroes
::hertzs::hertz
::hesistant::hesitant
::heterogenous::heterogeneous
::heirarchical::hierarchical
::hierachical::hierarchical
::hierarcical::hierarchical
::heirarchies::hierarchies
::hierachies::hierarchies
::heirarchy::hierarchy
::hierachy::hierarchy
::hierarcy::hierarchy
::hieroglph::hieroglyph
::heiroglyphics::hieroglyphics
::hieroglphs::hieroglyphs
::heigher::higher
::higer::higher
::higest::highest
::higway::highway
::hillarious::hilarious
::himselv::himself
::hismelf::himself
::hinderance::hindrance
::hinderence::hindrance
::hindrence::hindrance
::hipopotamus::hippopotamus
::hersuit::hirsute
::hsi::his
::ihs::his
::historicians::historians
::hsitorians::historians
::hstory::history
::hitsingles::hit singles
::hosited::hoisted
::holliday::holiday
::homestate::home state
::homogeneize::homogenize
::homogeneized::homogenized
::honourarium::honorarium
::honory::honorary
::honourific::honorific
::hounour::honour
::horrifing::horrifying
::hospitible::hospitable
::housr::hours
::howver::however
::huminoid::humanoid
::humoural::humoral
::humer::humour
::humerous::humourous
::humurous::humourous
::husban::husband
::hydogen::hydrogen
::hydropile::hydrophile
::hydropilic::hydrophilic
::hydropobe::hydrophobe
::hydropobic::hydrophobic
::hygeine::hygiene
::hypocracy::hypocrisy
::hypocrasy::hypocrisy
::hypocricy::hypocrisy
::hypocrit::hypocrite
::hypocrits::hypocrites
::i;d::I'd
::i"m::I'm
::iconclastic::iconoclastic
::idae::idea
::idaeidae::idea
::idaes::ideas
::identicial::identical
::identifers::identifiers
::identofy::identify
::idealogies::ideologies
::idealogy::ideology
::idiosyncracy::idiosyncrasy
::ideosyncratic::idiosyncratic
::ignorence::ignorance
::illiegal::illegal
::illegimacy::illegitimacy
::illegitmate::illegitimate
::illess::illness
::ilness::illness
::ilogical::illogical
::ilumination::illumination
::illution::illusion
::imagenary::imaginary
::imagin::imagine
::inbalance::imbalance
::inbalanced::imbalanced
::imediate::immediate
::emmediately::immediately
::imediately::immediately
::imediatly::immediately
::immediatley::immediately
::immediatly::immediately
::immidately::immediately
::immidiately::immediately
::imense::immense
::inmigrant::immigrant
::inmigrants::immigrants
::imanent::imminent
::immunosupressant::immunosuppressant
::inpeach::impeach
::impecabbly::impeccably
::impedence::impedance
::implamenting::implementing
::inpolite::impolite
::importamt::important
::importent::important
::importnat::important
::impossable::impossible
::emprisoned::imprisoned
::imprioned::imprisoned
::imprisonned::imprisoned
::inprisonment::imprisonment
::improvemnt::improvement
::improvment::improvement
::improvments::improvements
::inproving::improving
::improvision::improvisation
::int he::in the
::inteh::in the
::inthe::in the
::inwhich::in which
::inablility::inability
::inaccessable::inaccessible
::inadiquate::inadequate
::inadquate::inadequate
::inadvertant::inadvertent
::inadvertantly::inadvertently
::inappropiate::inappropriate
::inagurated::inaugurated
::inaugures::inaugurates
::inaguration::inauguration
::incarcirated::incarcerated
::incidentially::incidentally
::incidently::incidentally
::includ::include
::includng::including
::incuding::including
::incomptable::incompatible
::incompetance::incompetence
::incompetant::incompetent
::incomptetent::incompetent
::imcomplete::incomplete
::inconsistant::inconsistent
::incorportaed::incorporated
::incorprates::incorporates
::incorperation::incorporation
::incorruptable::incorruptible
::inclreased::increased
::increadible::incredible
::incredable::incredible
::incramentally::incrementally
::incunabla::incunabula
::indefinately::indefinitely
::indefinitly::indefinitely
::indepedence::independence
::independance::independence
::independece::independence
::indipendence::independence
::indepedent::independent
::independant::independent
::independendet::independent
::indipendent::independent
::indpendent::independent
::indepedantly::independently
::independantly::independently
::indipendently::independently
::indpendently::independently
::indecate::indicate
::indite::indict
::indictement::indictment
::indigineous::indigenous
::indispensible::indispensable
::individualy::individually
::indviduals::individuals
::enduce::induce
::indulgue::indulge
::indutrial::industrial
::inudstry::industry
::inefficienty::inefficiently
::unequalities::inequalities
::inevatible::inevitable
::inevitible::inevitable
::inevititably::inevitably
::infalability::infallibility
::infallable::infallible
::infrantryman::infantryman
::infectuous::infectious
::infered::inferred
::infilitrate::infiltrate
::infilitrated::infiltrated
::infilitration::infiltration
::infinit::infinite
::infinitly::infinitely
::enflamed::inflamed
::inflamation::inflammation
::influance::influence
::influented::influenced
::influencial::influential
::infomation::information
::informatoin::information
::informtion::information
::infrigement::infringement
::ingenius::ingenious
::ingreediants::ingredients
::inhabitans::inhabitants
::inherantly::inherently
::inheritence::inheritance
::inital::initial
::intial::initial
::ititial::initial
::initally::initially
::intially::initially
::initation::initiation
::initiaitive::initiative
::inate::innate
::inocence::innocence
::inumerable::innumerable
::innoculate::inoculate
::innoculated::inoculated
::insectiverous::insectivorous
::insensative::insensitive
::inseperable::inseparable
::insistance::insistence
::instaleld::installed
::instatance::instance
::instade::instead
::insted::instead
::institue::institute
::instutionalized::institutionalized
::instuction::instruction
::instuments::instruments
::insufficent::insufficient
::insufficently::insufficiently
::insurence::insurance
::intergrated::integrated
::intergration::integration
::intelectual::intellectual
::inteligence::intelligence
::inteligent::intelligent
::interchangable::interchangeable
::interchangably::interchangeably
::intercontinetal::intercontinental
::intrest::interest
::itnerest::interest
::itnerested::interested
::itneresting::interesting
::itnerests::interests
::interferance::interference
::interfereing::interfering
::interm::interim
::interrim::interim
::interum::interim
::intenational::international
::interational::international
::internation::international
::interpet::interpret
::intepretation::interpretation
::intepretator::interpretor
::interrugum::interregnum
::interelated::interrelated
::interupt::interrupt
::intevene::intervene
::intervines::intervenes
::inot::into
::inctroduce::introduce
::inctroduced::introduced
::intrduced::introduced
::introdued::introduced
::intruduced::introduced
::itnroduced::introduced
::instutions::intuitions
::intutive::intuitive
::intutively::intuitively
::inventer::inventor
::invertibrates::invertebrates
::investingate::investigate
::involvment::involvement
::ironicly::ironically
::irelevent::irrelevant
::irrelevent::irrelevant
::irreplacable::irreplaceable
::iresistable::irresistible
::iresistible::irresistible
::irresistable::irresistible
::iresistably::irresistibly
::iresistibly::irresistibly
::irresistably::irresistibly
::irregulairty::irregularity
::iritable::irritable
::iritated::irritated
::i snot::is not
::isthe::is the
::isnt::isn't
::issueing::issuing
::itis::it is
::itwas::it was
::it;s::it's
::its a::it's a
::it snot::it's not
::it' snot::it's not
::iits the::it's the
::its the::it's the
::ihaca::Ithaca
::jaques::jacques
::japanes::Japanese
::jeapardy::jeopardy
::jewelery::jewellery
::jewllery::jewellery
::johanine::Johannine
::jospeh::Joseph
::jouney::journey
::journied::journeyed
::journies::journeys
::juadaism::Judaism
::juadism::Judaism
::judgment::judgement
::jugment::judgment
::judical::judicial
::juducial::judicial
::judisuary::judiciary
::iunior::junior
::juristiction::jurisdiction
::juristictions::jurisdictions
::jstu::just
::jsut::just
::kindergarden::kindergarten
::klenex::kleenex
::knive::knife
::knifes::knives
::konw::know
::kwno::know
::nkow::know
::nkwo::know
::knowldge::knowledge
::knowlege::knowledge
::knowlegeable::knowledgeable
::knwon::known
::konws::knows
::labled::labelled
::labratory::laboratory
::labourious::laborious
::layed::laid
::laguage::language
::laguages::languages
::larg::large
::largst::largest
::larrry::larry
::lavae::larvae
::lazer::laser
::lasoo::lasso
::lastr::last
::lsat::last
::lastyear::last year
::lastest::latest
::lattitude::latitude
::launchs::launch
::launhed::launched
::lazyness::laziness
::leage::league
::leran::learn
::learnign::learning
::lerans::learns
::elast::least
::leaded::led
::lefted::left
::legitamate::legitimate
::legitmate::legitimate
::leibnitz::leibniz
::liesure::leisure
::lenght::length
::let;s::let's
::leathal::lethal
::let's him::lets him
::let's it::lets it
::levle::level
::levetate::levitate
::levetated::levitated
::levetates::levitates
::levetating::levitating
::liasion::liaison
::liason::liaison
::liasons::liaisons
::libell::libel
::libitarianisn::libertarianism
::libary::library
::librarry::library
::librery::library
::lybia::Libya
::lisense::license
::leutenant::lieutenant
::lieutenent::lieutenant
::liftime::lifetime
::lightyear::light year
::lightyears::light years
::lightening::lightning
::liek::like
::liuke::like
::liekd::liked
::likelyhood::likelihood
::likly::likely
::lukid::likud
::lmits::limits
::libguistic::linguistic
::libguistics::linguistics
::linnaena::linnaean
::lippizaner::lipizzaner
::liquify::liquefy
::listners::listeners
::litterally::literally
::litature::literature
::literture::literature
::littel::little
::litttle::little
::liev::live
::lieved::lived
::livley::lively
::liveing::living
::lonelyness::loneliness
::lonley::lonely
::lonly::lonely
::longitudonal::longitudinal
::lookign::looking
::loosing::losing
::lotharingen::lothringen
::loev::love
::lveo::love
::lvoe::love
::lieing::lying
::mackeral::mackerel
::amde::made
::magasine::magazine
::magincian::magician
::magnificient::magnificent
::magolia::magnolia
::mailny::mainly
::mantain::maintain
::mantained::maintained
::maintinaing::maintaining
::maintainance::maintenance
::maintainence::maintenance
::maintance::maintenance
::maintenence::maintenance
::majoroty::majority
::marjority::majority
::amke::make
::mkae::make
::mkea::make
::amkes::makes
::makse::makes
::mkaes::makes
::amking::making
::makeing::making
::mkaing::making
::malcom::Malcolm
::maltesian::Maltese
::mamal::mammal
::mamalian::mammalian
::managable::manageable
::managment::management
::manuver::maneuver
::manoeuverability::maneuverability
::manifestion::manifestation
::manisfestations::manifestations
::manufature::manufacture
::manufacturedd::manufactured
::manufatured::manufactured
::manufaturing::manufacturing
::mrak::mark
::maked::marked
::marketting::marketing
::markes::marks
::marmelade::marmalade
::mariage::marriage
::marrage::marriage
::marraige::marriage
::marryied::married
::marrtyred::martyred
::massmedia::mass media
::massachussets::Massachusetts
::massachussetts::Massachusetts
::masterbation::masturbation
::materalists::materialist
::mathmatically::mathematically
::mathematican::mathematician
::mathmatician::mathematician
::matheticians::mathematicians
::mathmaticians::mathematicians
::mathamatics::mathematics
::mathematicas::mathematics
::may of::may have
::mccarthyst::mccarthyist
::meaninng::meaning
::menat::meant
::mchanics::mechanics
::medieval::mediaeval
::medacine::medicine
::mediciney::mediciny
::medeival::medieval
::medevial::medieval
::medievel::medieval
::mediterainnean::mediterranean
::mediteranean::Mediterranean
::meerkrat::meerkat
::memeber::member
::membranaphone::membranophone
::momento::memento
::rememberable::memorable
::menally::mentally
::maintioned::mentioned
::mercentile::mercantile
::mechandise::merchandise
::merchent::merchant
::mesage::message
::mesages::messages
::messenging::messaging
::messanger::messenger
::metalic::metallic
::metalurgic::metallurgic
::metalurgical::metallurgical
::metalurgy::metallurgy
::metamorphysis::metamorphosis
::methaphor::metaphor
::metaphoricial::metaphorical
::methaphors::metaphors
::mataphysical::metaphysical
::meterologist::meteorologist
::meterology::meteorology
::micheal::Michael
::michagan::Michigan
::micoscopy::microscopy
::midwifes::midwives
::might of::might have
::mileau::milieu
::mileu::milieu
::melieux::milieux
::miliary::military
::miliraty::military
::millitary::military
::miltary::military
::milennia::millennia
::millenia::millennia
::millenial::millennial
::millenialism::millennialism
::milennium::millennium
::millenium::millennium
::milion::million
::millon::million
::millioniare::millionaire
::millepede::millipede
::minerial::mineral
::minature::miniature
::minumum::minimum
::minstries::ministries
::ministery::ministry
::minstry::ministry
::miniscule::minuscule
::mirrorred::mirrored
::miscelaneous::miscellaneous
::miscellanious::miscellaneous
::miscellanous::miscellaneous
::mischeivous::mischievous
::mischevious::mischievous
::mischievious::mischievous
::misdameanor::misdemeanor
::misdemenor::misdemeanor
::misdameanors::misdemeanors
::misdemenors::misdemeanors
::misfourtunes::misfortunes
::mysogynist::misogynist
::mysogyny::misogyny
::misile::missile
::missle::missile
::missonary::missionary
::missisipi::Mississippi
::missisippi::Mississippi
::misouri::Missouri
::mispell::misspell
::mispelled::misspelled
::mispelling::misspelling
::mispellings::misspellings
::mythraic::Mithraic
::missen::mizzen
::modle::model
::moderm::modem
::moil::mohel
::mosture::moisture
::moleclues::molecules
::moent::moment
::monestaries::monasteries
::monestary::monastery
::moeny::money
::monickers::monikers
::monkies::monkeys
::monolite::monolithic
::montypic::monotypic
::mounth::month
::monts::months
::monserrat::Montserrat
::mroe::more
::omre::more
::moreso::more so
::morisette::Morissette
::morrisette::Morissette
::morroccan::moroccan
::morrocco::morocco
::morroco::morocco
::morgage::mortgage
::motiviated::motivated
::mottos::mottoes
::montanous::mountainous
::montains::mountains
::movment::movement
::movei::movie
::mucuous::mucous
::multicultralism::multiculturalism
::multipled::multiplied
::multiplers::multipliers
::muncipalities::municipalities
::muncipality::municipality
::munnicipality::municipality
::muder::murder
::mudering::murdering
::muscial::musical
::muscician::musician
::muscicians::musicians
::muhammadan::muslim
::mohammedans::muslims
::must of::must have
::mutiliated::mutilated
::myu::my
::myraid::myriad
::mysef::myself
::mysefl::myself
::misterious::mysterious
::misteryous::mysterious
::mysterous::mysterious
::mistery::mystery
::naieve::naive
::napoleonian::Napoleonic
::ansalisation::nasalisation
::ansalization::nasalization
::naturual::natural
::naturaly::naturally
::naturely::naturally
::naturually::naturally
::nazereth::Nazareth
::neccesarily::necessarily
::neccessarily::necessarily
::necesarily::necessarily
::nessasarily::necessarily
::neccesary::necessary
::neccessary::necessary
::necesary::necessary
::nessecary::necessary
::necessiate::necessitate
::neccessities::necessities
::ened::need
::neglible::negligible
::negligable::negligible
::negociable::negotiable
::negotiaing::negotiating
::negotation::negotiation
::neigbourhood::neighbourhood
::neolitic::neolithic
::nestin::nesting
::nver::never
::neverthless::nevertheless
::nwe::new
::newyorker::New Yorker
::foundland::Newfoundland
::newletters::newsletters
::enxt::next
::nickle::nickel
::neice::niece
::nightime::nighttime
::ninteenth::nineteenth
::ninties::nineties ; fixed from "1990s": could refer to temperatures too.
::ninty::ninety
::nineth::ninth
::noone::no one
::noncombatents::noncombatants
::nontheless::nonetheless
::unoperational::nonoperational
::nonsence::nonsense
::noth::north
::northereastern::northeastern
::norhern::northern
::northen::northern
::nothern::northern
:C:Nto::Not
:C:nto::not
::noteable::notable
::notabley::notably
::noteably::notably
::nothign::nothing
::notive::notice
::noticable::noticeable
::noticably::noticeably
::noticeing::noticing
::noteriety::notoriety
::notwhithstanding::notwithstanding
::noveau::nouveau
::nowe::now
::nwo::now
::nowdays::nowadays
::nucular::nuclear
::nuculear::nuclear
::nuisanse::nuisance
::nusance::nuisance
::nullabour::Nullarbor
::munbers::numbers
::numberous::numerous
::nuptual::nuptial
::nuremburg::Nuremberg
::nuturing::nurturing
::nutritent::nutrient
::nutritents::nutrients
::obediance::obedience
::obediant::obedient
::obssessed::obsessed
::obession::obsession
::obsolecence::obsolescence
::obstacal::obstacle
::obstancles::obstacles
::obstruced::obstructed
::ocassion::occasion
::occaison::occasion
::occassion::occasion
::ocassional::occasional
::occassional::occasional
::ocassionally::occasionally
::ocassionaly::occasionally
::occassionally::occasionally
::occassionaly::occasionally
::occationally::occasionally
::ocassioned::occasioned
::occassioned::occasioned
::ocassions::occasions
::occassions::occasions
::occour::occur
::occurr::occur
::ocur::occur
::ocurr::occur
::occured::occurred
::ocurred::occurred
::occurence::occurrence
::occurrance::occurrence
::ocurrance::occurrence
::ocurrence::occurrence
::occurences::occurrences
::occurrances::occurrences
::occuring::occurring
::octohedra::octahedra
::octohedral::octahedral
::octohedron::octahedron
::odouriferous::odoriferous
::odourous::odorous
::ouevre::oeuvre
::fo::of
:C:fo::of
:C:od::of
::ofits::of its
::ofthe::of the
::oft he::of the ; Could be legitimate in poetry, but more usually a typo.
::offereings::offerings
::offcers::officers
::offical::official
::offcially::officially
::offically::officially
::officaly::officially
::officialy::officially
::oftenly::often
::omlette::omelette
::omnious::ominous
::omision::omission
::ommision::omission
::omited::omitted
::ommited::omitted
::ommitted::omitted
::omiting::omitting
::ommiting::omitting
::ommitting::omitting
::omniverous::omnivorous
::omniverously::omnivorously
::ont he::on the
::onthe::on the
::oneof::one of
::onepoint::one point
::onyl::only
::onomatopeia::onomatopoeia
::oppenly::openly
::openess::openness
::opperation::operation
::oeprator::operator
::opthalmic::ophthalmic
::opthalmologist::ophthalmologist
::opthamologist::ophthalmologist
::opthalmology::ophthalmology
::oppinion::opinion
::oponent::opponent
::opponant::opponent
::oppononent::opponent
::oppotunities::opportunities
::oportunity::opportunity
::oppertunity::opportunity
::oppotunity::opportunity
::opprotunity::opportunity
::opposible::opposable
::opose::oppose
::oppossed::opposed
::oposite::opposite
::oppasite::opposite
::opposate::opposite
::opposit::opposite
::oposition::opposition
::oppositition::opposition
::opression::oppression
::opressive::oppressive
::optomism::optimism
::optmizations::optimizations
::orded::ordered
::oridinarily::ordinarily
::orginize::organise
::organim::organism
::organiztion::organization
::orginization::organization
::orginized::organized
::orgin::origin
::orginal::original
::origional::original
::orginally::originally
::origanaly::originally
::originall::originally
::originaly::originally
::originially::originally
::originnally::originally
::orignally::originally
::orignially::originally
::orthagonal::orthogonal
::orthagonally::orthogonally
::ohter::other
::otehr::other
::otherw::others
::otu::out
::outof::out of
::overthe::over the
::overthere::over there
::overshaddowed::overshadowed
::overwelming::overwhelming
::overwheliming::overwhelming
::pwn::own
::oxident::oxidant
::oxigen::oxygen
::oximoron::oxymoron
::peageant::pageant
::paide::paid
::payed::paid
::paleolitic::paleolithic
::palistian::Palestinian
::palistinian::Palestinian
::palistinians::Palestinians
::pallete::palette
::pamflet::pamphlet
::pamplet::pamphlet
::pantomine::pantomime
::papanicalou::Papanicolaou
::papaer::paper
::perade::parade
::parrakeets::parakeets
::paralel::parallel
::paralell::parallel
::parralel::parallel
::parrallel::parallel
::parrallell::parallel
::paralelly::parallelly
::paralely::parallelly
::parallely::parallelly
::parrallelly::parallelly
::parrallely::parallelly
::parellels::parallels
::paraphenalia::paraphernalia
::paranthesis::parenthesis
::parliment::parliament
::paliamentarian::parliamentarian
::partof::part of
::partialy::partially
::parituclar::particular
::particualr::particular
::paticular::particular
::particuarly::particularly
::particularily::particularly
::particulary::particularly
::pary::party
::pased::passed
::pasengers::passengers
::passerbys::passersby
::pasttime::pastime
::pastural::pastoral
::pattented::patented
::paitience::patience
::pavillion::pavilion
::paymetn::payment
::paymetns::payments
::peacefuland::peaceful and
::peculure::peculiar
::pedestrain::pedestrian
::perjorative::pejorative
::peloponnes::Peloponnesus
::peleton::peloton
::penatly::penalty
::penerator::penetrator
::penisula::peninsula
::penninsula::peninsula
::pennisula::peninsula
::pensinula::peninsula
::penisular::peninsular
::penninsular::peninsular
::peolpe::people
::peopel::people
::poeple::people
::poeoples::peoples
::percieve::perceive
::percepted::perceived
::percieved::perceived
::percentof::percent of
::percentto::percent to
::precentage::percentage
::perenially::perennially
::performence::performance
::perfomers::performers
::performes::performs
::perhasp::perhaps
::perheaps::perhaps
::perhpas::perhaps
::perphas::perhaps
::preiod::period
::preriod::period
::peripathetic::peripatetic
::perjery::perjury
::permanant::permanent
::permenant::permanent
::perminent::permanent
::permenantly::permanently
::permissable::permissible
::premission::permission
::perpindicular::perpendicular
::perseverence::perseverance
::persistance::persistence
::peristent::persistent
::persistant::persistent
::peronal::personal
::perosnality::personality
::personalyl::personally
::personell::personnel
::personnell::personnel
::prespective::perspective
::pursuade::persuade
::persuded::persuaded
::pursuaded::persuaded
::pursuades::persuades
::pususading::persuading
::pertubation::perturbation
::pertubations::perturbations
::preverse::perverse
::pessiary::pessary
::petetion::petition
::pharoah::Pharaoh
::phenonmena::phenomena
::phenomenonal::phenomenal
::phenomenonly::phenomenally
::phenomenom::phenomenon
::phenomonenon::phenomenon
::phenomonon::phenomenon
::feromone::pheromone
::phillipine::Philippine
::philipines::Philippines
::phillipines::Philippines
::phillippines::Philippines
::philisopher::philosopher
::philospher::philosopher
::philisophical::philosophical
::phylosophical::philosophical
::phillosophically::philosophically
::philosphies::philosophies
::philisophy::philosophy
::philosphy::philosophy
::phonecian::Phoenecian
::pheonix::phoenix ; Not forcing caps, as it could be the bird
::fonetic::phonetic
::phongraph::phonograph
::physicaly::physically
::pciture::picture
::peice::piece
::peices::pieces
::pilgrimmage::pilgrimage
::pilgrimmages::pilgrimages
::pinapple::pineapple
::pinnaple::pineapple
::pinoneered::pioneered
::pich::pitch
::palce::place
::plagarism::plagiarism
::plantiff::plaintiff
::planed::planned
::planation::plantation
::plateu::plateau
::plausable::plausible
::playright::playwright
::playwrite::playwright
::playwrites::playwrights
::pleasent::pleasant
::plesant::pleasant
::plebicite::plebiscite
::peom::poem
::peoms::poems
::peotry::poetry
::poety::poetry
::poisin::poison
::posion::poison
::polical::political
::poltical::political
::politican::politician
::politicans::politicians
::polinator::pollinator
::polinators::pollinators
::polute::pollute
::poluted::polluted
::polutes::pollutes
::poluting::polluting
::polution::pollution
::polyphonyic::polyphonic
::polysaccaride::polysaccharide
::polysaccharid::polysaccharide
::pomegranite::pomegranate
::populare::popular
::popularaty::popularity
::popoulation::population
::poulations::populations
::portayed::portrayed
::potrayed::portrayed
::protrayed::portrayed
::portraing::portraying
::portugese::Portuguese
::portuguease::portuguese
::possition::position
::postion::position
::postition::position
::psoition::position
::postive::positive
::posess::possess
::posessed::possessed
::posesses::possesses
::posseses::possesses
::possessess::possesses
::posessing::possessing
::possesing::possessing
::posession::possession
::possesion::possession
::posessions::possessions
::possiblility::possibility
::possiblilty::possibility
::possable::possible
::possibile::possible
::possably::possibly
::posthomous::posthumous
::potatoe::potato
::potatos::potatoes
::potentialy::potentially
::postdam::Potsdam
::pwoer::power
::poverful::powerful
::poweful::powerful
::powerfull::powerful
::practial::practical
::practially::practically
::practicaly::practically
::practicly::practically
::pratice::practice
::practicioner::practitioner
::practioner::practitioner
::practicioners::practitioners
::practioners::practitioners
::prairy::prairie
::prarie::prairie
::praries::prairies
::pre-Colombian::pre-Columbian
::preample::preamble
::preceed::precede
::preceeded::preceded
::preceeds::precedes
::preceeding::preceding
::precice::precise
::precisly::precisely
::precurser::precursor
::precedessor::predecessor
::predecesors::predecessors
::predicatble::predictable
::predicitons::predictions
::predomiantly::predominately
::preminence::preeminence
::preferrably::preferably
::prefernece::preference
::preferneces::preferences
::prefered::preferred
::prefering::preferring
::pregancies::pregnancies
::pregnent::pregnant
::premeire::premiere
::premeired::premiered
::premillenial::premillennial
::premonasterians::Premonstratensians
::preocupation::preoccupation
::prepartion::preparation
::preperation::preparation
::preperations::preparations
::prepatory::preparatory
::prepair::prepare
::perogative::prerogative
::presance::presence
::presense::presence
::presedential::presidential
::presidenital::presidential
::presidental::presidential
::presitgious::prestigious
::prestigeous::prestigious
::prestigous::prestigious
::presumabely::presumably
::presumibly::presumably
::prevelant::prevalent
::previvous::previous
::priestood::priesthood
::primarly::primarily
::primative::primitive
::primatively::primitively
::primatives::primitives
::primordal::primordial
::pricipal::principal
::priciple::principle
::privte::private
::privelege::privilege
::privelige::privilege
::privilage::privilege
::priviledge::privilege
::privledge::privilege
::priveleged::privileged
::priveliged::privileged
::priveleges::privileges
::priveliges::privileges
::privelleges::privileges
::priviledges::privileges
::protem::pro tem
::probablistic::probabilistic
::probabilaty::probability
::probalibity::probability
::probablly::probably
::probaly::probably
::porblem::problem
::probelm::problem
::porblems::problems
::probelms::problems
::procedger::procedure
::proceedure::procedure
::procede::proceed
::proceded::proceeded
::proceding::proceeding
::procedings::proceedings
::procedes::proceeds
::proccess::process
::proces::process
::proccessing::processing
::processer::processor
::proclamed::proclaimed
::proclaming::proclaiming
::proclaimation::proclamation
::proclomation::proclamation
::proffesed::professed
::profesion::profession
::proffesion::profession
::proffesional::professional
::profesor::professor
::professer::professor
::proffesor::professor
::programable::programmable
::ptogress::progress
::progessed::progressed
::prohabition::prohibition
::prologomena::prolegomena
::preliferation::proliferation
::profilic::prolific
::prominance::prominence
::prominant::prominent
::prominantly::prominently
::promiscous::promiscuous
::promotted::promoted
::pomotion::promotion
::propmted::prompted
::pronomial::pronominal
::pronouced::pronounced
::pronounched::pronounced
::prouncements::pronouncements
::pronounciation::pronunciation
::propoganda::propaganda
::propogate::propagate
::propogates::propagates
::propogation::propagation
::propper::proper
::propperly::properly
::prophacy::prophecy
::poportional::proportional
::propotions::proportions
::propostion::proposition
::propietary::proprietary
::proprietory::proprietary
::proseletyzing::proselytizing
::protaganist::protagonist
::protoganist::protagonist
::protaganists::protagonists
::pretection::protection
::protien::protein
::protocal::protocol
::protruberance::protuberance
::protruberances::protuberances
::proove::prove
::prooved::proved
::porvide::provide
::provded::provided
::provicial::provincial
::provinicial::provincial
::provisonal::provisional
::provacative::provocative
::proximty::proximity
::psuedo::pseudo
::pseudonyn::pseudonym
::pseudononymous::pseudonymous
::psyhic::psychic
::pyscic::psychic
::psycology::psychology
::publically::publicly
::publicaly::publicly
::pucini::Puccini
::puertorrican::Puerto Rican
::puertorricans::Puerto Ricans
::pumkin::pumpkin
::puchasing::purchasing
::puritannical::puritanical
::purpotedly::purportedly
::purposedly::purposely
::persue::pursue
::persued::pursued
::persuing::pursuing
::persuit::pursuit
::persuits::pursuits
::puting::putting
::quantaty::quantity
::quantitiy::quantity
::quarantaine::quarantine
::quater::quarter
::quaters::quarters
::quesion::question
::questoin::question
::quetion::question
::questonable::questionable
::questionnair::questionnaire
::quesions::questions
::questioms::questions
::questiosn::questions
::quetions::questions
::quicklyu::quickly
::quinessential::quintessential
::quitted::quit
::quizes::quizzes
::rabinnical::rabbinical
::radiactive::radioactive
::rancourous::rancorous
::repid::rapid
::rarified::rarefied
::rasberry::raspberry
::ratehr::rather
::radify::ratify
::racaus::raucous
::reched::reached
::reacing::reaching
::readmition::readmission
::rela::real
::relized::realised
::realsitic::realistic
::erally::really
::raelly::really
::realy::really
::realyl::really
::relaly::really
::rebllions::rebellions
::rebounce::rebound
::rebiulding::rebuilding
::reacll::recall
::receeded::receded
::receeding::receding
::receieve::receive
::receivedfrom::received from
::receving::receiving
::rechargable::rechargeable
::recipiant::recipient
::reciepents::recipients
::recipiants::recipients
::recogise::recognise
::recogize::recognize
::reconize::recognize
::reconized::recognized
::reccommend::recommend
::recomend::recommend
::reommend::recommend
::recomendation::recommendation
::recomendations::recommendations
::recommedations::recommendations
::reccommended::recommended
::recomended::recommended
::reccommending::recommending
::recomending::recommending
::recomends::recommends
::reconcilation::reconciliation
::reconaissance::reconnaissance
::reconnaissence::reconnaissance
::recontructed::reconstructed
::recrod::record
::rocord::record
::recordproducer::record producer
::recrational::recreational
::recuiting::recruiting
::rucuperate::recuperate
::recurrance::recurrence
::reoccurrence::recurrence
::reaccurring::recurring
::reccuring::recurring
::recuring::recurring
::recyling::recycling
::reedeming::redeeming
::relected::reelected
::revaluated::reevaluated
::referrence::reference
::refference::reference
::refrence::reference
::refernces::references
::refrences::references
::refedendum::referendum
::referal::referral
::refered::referred
::reffered::referred
::referiang::referring
::refering::referring
::referrs::refers
::refrers::refers
::refect::reflect
::refromist::reformist
::refridgeration::refrigeration
::refridgerator::refrigerator
::refusla::refusal
::irregardless::regardless
::regardes::regards
::regluar::regular
::reguarly::regularly
::regularily::regularly
::regulaion::regulation
::regulaotrs::regulators
::rehersal::rehearsal
::reigining::reigning
::reicarnation::reincarnation
::reenforced::reinforced
::realtions::relations
::relatiopnship::relationship
::realitvely::relatively
::relativly::relatively
::relitavely::relatively
::releses::releases
::relevence::relevance
::relevent::relevant
::relient::reliant
::releive::relieve
::releived::relieved
::releiver::reliever
::religeous::religious
::religous::religious
::religously::religiously
::relinqushment::relinquishment
::reluctent::reluctant
::remaing::remaining
::remeber::remember
::rememberance::remembrance
::remembrence::remembrance
::remenicent::reminiscent
::reminescent::reminiscent
::reminscent::reminiscent
::reminsicent::reminiscent
::remenant::remnant
::reminent::remnant
::renedered::rende
::rendevous::rendezvous
::rendezous::rendezvous
::renewl::renewal
::reknown::renown
::reknowned::renowned
::rentors::renters
::reorganision::reorganisation
::repeteadly::repeatedly
::repentence::repentance
::repentent::repentant
::reprtoire::repertoire
::repetion::repetition
::reptition::repetition
::relpacement::replacement
::reportadly::reportedly
::represnt::represent
::represantative::representative
::representive::representative
::representativs::representatives
::representives::representatives
::represetned::represented
::reproducable::reproducible
::requred::required
::reasearch::research
::reserach::research
::resembelance::resemblance
::resemblence::resemblance
::ressemblance::resemblance
::ressemblence::resemblance
::ressemble::resemble
::ressembled::resembled
::resembes::resembles
::ressembling::resembling
::resevoir::reservoir
::recide::reside
::recided::resided
::recident::resident
::recidents::residents
::reciding::residing
::resignement::resignment
::resistence::resistance
::resistent::resistant
::resistable::resistible
::resollution::resolution
::resorces::resources
::repsectively::respectively
::respectivly::respectively
::respomse::response
::responce::response
::responibilities::responsibilities
::responsability::responsibility
::responisble::responsible
::responsable::responsible
::responsibile::responsible
::resaurant::restaurant
::restaraunt::restaurant
::restauraunt::restaurant
::resteraunt::restaurant
::restuarant::restaurant
::resturant::restaurant
::resturaunt::restaurant
::restaraunts::restaurants
::resteraunts::restaurants
::restaraunteur::restaurateur
::restaraunteurs::restaurateurs
::restauranteurs::restaurateurs
::restauration::restoration
::resticted::restricted
::reult::result
::resurgance::resurgence
::resssurecting::resurrecting
::resurecting::resurrecting
::ressurrection::resurrection
::retalitated::retaliated
::retalitation::retaliation
::retreive::retrieve
::returnd::returned
::reveral::reversal
::reversable::reversible
::reveiw::review
::reveiwing::reviewing
::revolutionar::revolutionary
::rewriet::rewrite
::rewitten::rewritten
::rhymme::rhyme
::rhythem::rhythm
::rhythim::rhythm
::rythem::rhythm
::rythim::rhythm
::rythm::rhythm
::rhytmic::rhythmic
::rythmic::rhythmic
::rythyms::rhythms
::rediculous::ridiculous
::rigourous::rigorous
::rigeur::rigueur
::rininging::ringing
::rockerfeller::Rockefeller
::rococco::rococo
::roomate::roommate
::rised::rose
::rougly::roughly
::rudimentatry::rudimentary
::rulle::rule
::rumers::rumors
::runing::running
::runnung::running
::russina::Russian
::russion::Russian
::sacrafice::sacrifice
::sacrifical::sacrificial
::sacreligious::sacrilegious
::sandess::sadness
::saftey::safety
::safty::safety
::saidhe::said he
::saidit::said it
::saidthat::said that
::saidt he::said the
::saidthe::said the
::salery::salary
::smae::same
::santioned::sanctioned
::sanctionning::sanctioning
::sandwhich::sandwich
::sanhedrim::Sanhedrin
::satelite::satellite
::sattelite::satellite
::satelites::satellites
::sattelites::satellites
::satric::satiric
::satrical::satirical
::satrically::satirically
::satisfactority::satisfactorily
::saterday::Saturday
::saterdays::Saturdays
::svae::save
::svaes::saves
::saxaphone::saxophone
::sasy::says
::syas::says
::scaleable::scalable
::scandanavia::Scandinavia
::scaricity::scarcity
::scavanged::scavenged
::senarios::scenarios
::scedule::schedule
::schedual::schedule
::sceduled::scheduled
::scholarhip::scholarship
::scholarstic::scholastic
::shcool::school
::scince::science
::scinece::science
::scientfic::scientific
::scientifc::scientific
::screenwrighter::screenwriter
::scirpt::script
::scoll::scroll
::scrutinity::scrutiny
::scuptures::sculptures
::seach::search
::seached::searched
::seaches::searches
::secratary::secretary
::secretery::secretary
::sectino::section
::seing::seeing
::segementation::segmentation
::seguoys::segues
::sieze::seize
::siezed::seized
::siezing::seizing
::siezure::seizure
::siezures::seizures
::seldomly::seldom
::selectoin::selection
::seinor::senior
::sence::sense
::senstive::sensitive
::sentance::sentence
::separeate::separate
::sepulchure::sepulchre
::sargant::sergeant
::sargeant::sergeant
::sergent::sergeant
::settelement::settlement
::settlment::settlement
::severeal::several
::severley::severely
::severly::severely
::shaddow::shadow
::seh::she
::shesaid::she said
::sherif::sheriff
::sheild::shield
::shineing::shining
::shiped::shipped
::shiping::shipping
::shopkeeepers::shopkeepers
::shortwhile::short while
::shorly::shortly
::shoudl::should
::should of::should have
::shoudln't::shouldn't
::shouldent::shouldn't
::shouldnt::shouldn't
::sohw::show
::showinf::showing
::shreak::shriek
::shrinked::shrunk
::sedereal::sidereal
::sideral::sidereal
::seige::siege
::signitories::signatories
::signitory::signatory
::siginificant::significant
::signficant::significant
::signficiant::significant
::signifacnt::significant
::signifigant::significant
::signifantly::significantly
::significently::significantly
::signifigantly::significantly
::signfies::signifies
::silicone chip::silicon chip
::simalar::similar
::similiar::similar
::simmilar::similar
::similiarity::similarity
::similarily::similarly
::similiarly::similarly
::simplier::simpler
::simpley::simply
::simpyl::simply
::simultanous::simultaneous
::simultanously::simultaneously
::sicne::since
::sincerley::sincerely
::sincerly::sincerely
::singsog::singsong
::sixtin::Sistine
::skagerak::Skagerrak
::skateing::skating
::slaugterhouses::slaughterhouses
::slowy::slowly
::smoothe::smooth
::smoothes::smooths
::sneeks::sneaks
::snese::sneeze
::sot hat::so that
::soical::social
::socalism::socialism
::socities::societies
::sofware::software
::soilders::soldiers
::soliders::soldiers
::soley::solely
::soliliquy::soliloquy
::solatary::solitary
::soluable::soluble
::soem::some
::somene::someone
::somethign::something
::someting::something
::somthing::something
::somtimes::sometimes
::somewaht::somewhat
::somwhere::somewhere
::sophicated::sophisticated
::suphisticated::sophisticated
::sophmore::sophomore
::sorceror::sorcerer
::saught::sought
::seeked::sought
::soudn::sound
::soudns::sounds
::sountrack::soundtrack
::suop::soup
::sourth::south
::sourthern::southern
::souvenier::souvenir
::souveniers::souvenirs
::soverign::sovereign
::sovereignity::sovereignty
::soverignity::sovereignty
::soverignty::sovereignty
::soveits::soviets
::soveits::soviets(x
::spoace::space
::spainish::Spanish
::speciallized::specialised
::speices::species
::specfic::specific
::specificaly::specifically
::specificalyl::specifically
::specifiying::specifying
::speciman::specimen
::spectauclar::spectacular
::spectaulars::spectaculars
::spectum::spectrum
::speach::speech
::sprech::speech
::sppeches::speeches
::spermatozoan::spermatozoon
::spriritual::spiritual
::spritual::spiritual
::spendour::splendour
::sponser::sponsor
::sponsered::sponsored
::sponzored::sponsored
::spontanous::spontaneous
::spoonfulls::spoonfuls
::sportscar::sports car
::spreaded::spread
::spred::spread
::sqaure::square
::stablility::stability
::stainlees::stainless
::stnad::stand
::standars::standards
::strat::start
::statment::statement
::statememts::statements
::statments::statements
::stateman::statesman
::staion::station
::sterotypes::stereotypes
::steriods::steroids
::sitll::still
::stiring::stirring
::stirrs::stirs
::stpo::stop
::storeis::stories
::storise::stories
::sotry::story
::stopry::story
::stoyr::story
::stroy::story
::strnad::strand
::stange::strange
::startegic::strategic
::stratagically::strategically
::startegies::strategies
::stradegies::strategies
::startegy::strategy
::stradegy::strategy
::streemlining::streamlining
::stregth::strength
::strenght::strength
::strentgh::strength
::strenghen::strengthen
::strenghten::strengthen
::strenghened::strengthened
::strenghtened::strengthened
::strengtened::strengthened
::strenghening::strengthening
::strenghtening::strengthening
::strenous::strenuous
::strictist::strictest
::strikely::strikingly
::stingent::stringent
::stong::strong
::stornegst::strongest
::stucture::structure
::sturcture::structure
::stuctured::structured
::struggel::struggle
::strugle::struggle
::stuggling::struggling
::stubborness::stubbornness
::studnet::student
::studdy::study
::studing::studying
::stlye::style
::sytle::style
::stilus::stylus
::subconsiously::subconsciously
::subjudgation::subjugation
::submachne::submachine
::sepina::subpoena
::subsquent::subsequent
::subsquently::subsequently
::subsidary::subsidiary
::subsiduary::subsidiary
::subpecies::subspecies
::substace::substance
::subtances::substances
::substancial::substantial
::substatial::substantial
::substituded::substituted
::subterranian::subterranean
::substract::subtract
::substracted::subtracted
::substracting::subtracting
::substraction::subtraction
::substracts::subtracts
::suburburban::suburban
::suceed::succeed
::succceeded::succeeded
::succedded::succeeded
::succeded::succeeded
::suceeded::succeeded
::suceeding::succeeding
::succeds::succeeds
::suceeds::succeeds
::succsess::success
::sucess::success
::succcesses::successes
::sucesses::successes
::succesful::successful
::successfull::successful
::succsessfull::successful
::sucesful::successful
::sucessful::successful
::sucessfull::successful
::succesfully::successfully
::succesfuly::successfully
::successfuly::successfully
::successfulyl::successfully
::successully::successfully
::sucesfully::successfully
::sucesfuly::successfully
::sucessfully::successfully
::sucessfuly::successfully
::succesion::succession
::sucesion::succession
::sucession::succession
::succesive::successive
::sucessive::successive
::sucessor::successor
::sucessot::successor
::sufferred::suffered
::sufferring::suffering
::suffcient::sufficient
::sufficent::sufficient
::sufficiant::sufficient
::suffciently::sufficiently
::sufficently::sufficiently
::sufferage::suffrage
::suggestable::suggestible
::sucidial::suicidal
::sucide::suicide
::sumary::summary
::sunglases::sunglasses
::superintendant::superintendent
::surplanted::supplanted
::suplimented::supplemented
::supplamented::supplemented
::suppliementing::supplementing
::suppy::supply
::wupport::support
::supose::suppose
::suposed::supposed
::suppoed::supposed
::suppossed::supposed
::suposedly::supposedly
::supposingly::supposedly
::suposes::supposes
::suposing::supposing
::supress::suppress
::surpress::suppress
::supressed::suppressed
::surpressed::suppressed
::supresses::suppresses
::supressing::suppressing
::surley::surely
::surfce::surface
::suprise::surprise
::suprize::surprise
::surprize::surprise
::suprised::surprised
::suprized::surprised
::surprized::surprised
::suprising::surprising
::suprizing::surprising
::surprizing::surprising
::suprisingly::surprisingly
::suprizingly::surprisingly
::surprizingly::surprisingly
::surrended::surrendered
::surrundering::surrendering
::surrepetitious::surreptitious
::surreptious::surreptitious
::surrepetitiously::surreptitiously
::surreptiously::surreptitiously
::suround::surround
::surounded::surrounded
::surronded::surrounded
::surrouded::surrounded
::sorrounding::surrounding
::surounding::surrounding
::surrouding::surrounding
::suroundings::surroundings
::surounds::surrounds
::surveill::surveil
::surveilence::surveillance
::surveyer::surveyor
::survivied::survived
::surviver::survivor
::survivers::survivors
::suseptable::susceptible
::suseptible::susceptible
::suspention::suspension
::swaer::swear
::swaers::swears
::swepth::swept
::swiming::swimming
::symettric::symmetric
::symmetral::symmetric
::symetrical::symmetrical
::symetrically::symmetrically
::symmetricaly::symmetrically
::symetry::symmetry
::synphony::symphony
::sypmtoms::symptoms
::synagouge::synagogue
::syncronization::synchronization
::synonomous::synonymous
::synonymns::synonyms
::syphyllis::syphilis
::syrap::syrup
::sytem::system
::sysmatically::systematically
::tkae::take
::tkaes::takes
::tkaing::taking
::talekd::talked
::talkign::talking
::tlaking::talking
::targetted::targeted
::targetting::targeting
::tast::taste
::tatoo::tattoo
::tattooes::tattoos
::teached::taught
::taxanomic::taxonomic
::taxanomy::taxonomy
::tecnical::technical
::techician::technician
::technitian::technician
::techicians::technicians
::techiniques::techniques
::technnology::technology
::technolgy::technology
::telphony::telephony
::televize::televise
::telelevision::television
::televsion::television
::tellt he::tell the
::temperment::temperament
::tempermental::temperamental
::temparate::temperate
::temerature::temperature
::tempertaure::temperature
::temperture::temperature
::temperarily::temporarily
::tepmorarily::temporarily
::temprary::temporary
::tendancies::tendencies
::tendacy::tendency
::tendancy::tendency
::tendonitis::tendinitis
::tennisplayer::tennis player
::tenacle::tentacle
::tenacles::tentacles
::terrestial::terrestrial
::terriories::territories
::terriory::territory
::territoy::territory
::territorist::terrorist
::terroist::terrorist
::testiclular::testicular
::tahn::than
::thna::than
::thansk::thanks
::taht::that
::tath::that
::thgat::that
::thta::that
::thyat::that
::tyhat::that
::thatt he::that the
::thatthe::that the
::thast::that's
::thats::that's
::hte::the
::teh::the
::tehw::the
::tghe::the
::theh::the
::thge::the
::thw::the
::tje::the
::tjhe::the
::tthe::the
::tyhe::the
::thecompany::the company
::thefirst::the first
::thegovernment::the government
::thenew::the new
::thesame::the same
::thetwo::the two
::theather::theatre
::theri::their
::thier::their
::there's is::theirs is
::htem::them
::themself::themselves
::themselfs::themselves
::themslves::themselves
::hten::then
::thn::then
::thne::then
::htere::there
::their are::there are
::they're are::there are
::their is::there is
::they're is::there is
::therafter::thereafter
::therby::thereby
::htese::these
::theese::these
::htey::they
::tehy::they
::tyhe::they
::they;l::they'll
::theyll::they'll
::they;r::they're
::they;v::they've
::theyve::they've
::theif::thief
::theives::thieves
::hting::thing
::thign::thing
::thnig::thing
::thigns::things
::thigsn::things
::thnigs::things
::htikn::think
::htink::think
::thikn::think
::thiunk::think
::tihkn::think
::thikning::thinking
::thikns::thinks
::thrid::third
::htis::this
::tghis::this
::thsi::this
::tihs::this
::thisyear::this year
::throrough::thorough
::throughly::thoroughly
::thsoe::those
::threatend::threatened
::threatning::threatening
::threee::three
::threshhold::threshold
::throuhg::through
::thru::through
::thoughout::throughout
::througout::throughout
::tiget::tiger
::tiem::time
::timne::time
::tot he::to the
::tothe::to the
::tabacco::tobacco
::tobbaco::tobacco
::todya::today
::todays::today's
::tiogether::together
::togehter::together
::toghether::together
::toldt he::told the
::tolerence::tolerance
::tolkein::Tolkien
::tomatos::tomatoes
::tommorow::tomorrow
::tommorrow::tomorrow
::tomorow::tomorrow
::tounge::tongue
::tongiht::tonight
::tonihgt::tonight
::tormenters::tormentors
::toriodal::toroidal
::torpeados::torpedoes
::torpedos::torpedoes
::totaly::totally
::totalyl::totally
::towrad::toward
::towords::towards
::twon::town
::traditition::tradition
::traditionnal::traditional
::tradionally::traditionally
::traditionaly::traditionally
::traditionalyl::traditionally
::tradtionally::traditionally
::trafic::traffic
::trafficed::trafficked
::trafficing::trafficking
::transcendance::transcendence
::trancendent::transcendent
::transcendant::transcendent
::transcendentational::transcendental
::trancending::transcending
::transending::transcending
::transcripting::transcribing
::transfered::transferred
::transfering::transferring
::tranform::transform
::transformaton::transformation
::tranformed::transformed
::transistion::transition
::translater::translator
::translaters::translators
::transmissable::transmissible
::transporation::transportation
::transesxuals::transsexuals
::tremelo::tremolo
::tremelos::tremolos
::triathalon::triathlon
::tryed::tried
::triguered::triggered
::triology::trilogy
::troling::trolling
::toubles::troubles
::troup::troupe
::truely::truly
::truley::truly
::turnk::trunk
::tust::trust
::trustworthyness::trustworthiness
::tuscon::Tucson
::termoil::turmoil
::twpo::two
::typcial::typical
::typicaly::typically
::tyranies::tyrannies
::tyrranies::tyrannies
::tyrany::tyranny
::tyrrany::tyranny
::ubiquitious::ubiquitous
::ukranian::Ukrainian
::ukelele::ukulele
::alterior::ulterior
::ultimely::ultimately
::unacompanied::unaccompanied
::unanymous::unanimous
::unathorised::unauthorised
::unavailible::unavailable
::unballance::unbalance
::unbeleivable::unbelievable
::uncertainity::uncertainty
::unchallengable::unchallengeable
::unchangable::unchangeable
::uncompetive::uncompetitive
::unconcious::unconscious
::unconciousness::unconsciousness
::uncontitutional::unconstitutional
::unconvential::unconventional
::undecideable::undecidable
::indefineable::undefinable
::undert he::under the
::undreground::underground
::udnerstand::understand
::understnad::understand
::understoon::understood
::undesireable::undesirable
::undetecable::undetectable
::undoubtely::undoubtedly
::unforgetable::unforgettable
::unforgiveable::unforgivable
::unforetunately::unfortunately
::unfortunatley::unfortunately
::unfortunatly::unfortunately
::unfourtunately::unfortunately
::unahppy::unhappy
::unilatreal::unilateral
::unilateraly::unilaterally
::unilatreally::unilaterally
::unihabited::uninhabited
::uninterruped::uninterrupted
::uninterupted::uninterrupted
::unitedstates::United States
::unitesstates::United States
::univeral::universal
::univeristies::universities
::univesities::universities
::univeristy::university
::universtiy::university
::univesity::university
::unviersity::university
::unkown::unknown
::unliek::unlike
::unlikey::unlikely
::unmanouverable::unmanoeuvrable
::unmistakeably::unmistakably
::unneccesarily::unnecessarily
::unneccessarily::unnecessarily
::unnecesarily::unnecessarily
::uneccesary::unnecessary
::unecessary::unnecessary
::unneccesary::unnecessary
::unneccessary::unnecessary
::unnecesary::unnecessary
::unoticeable::unnoticeable
::inofficial::unofficial
::unoffical::unofficial
::unplesant::unpleasant
::unpleasently::unpleasantly
::unprecendented::unprecedented
::unprecidented::unprecedented
::unrepentent::unrepentant
::unrepetant::unrepentant
::unrepetent::unrepentant
::unsubstanciated::unsubstantiated
::unsuccesful::unsuccessful
::unsuccessfull::unsuccessful
::unsucesful::unsuccessful
::unsucessful::unsuccessful
::unsucessfull::unsuccessful
::unsuccesfully::unsuccessfully
::unsucesfuly::unsuccessfully
::unsucessfully::unsuccessfully
::unsuprised::unsurprised
::unsuprized::unsurprised
::unsurprized::unsurprised
::unsuprising::unsurprising
::unsuprizing::unsurprising
::unsurprizing::unsurprising
::unsuprisingly::unsurprisingly
::unsuprizingly::unsurprisingly
::unsurprizingly::unsurprisingly
::untill::until
::untranslateable::untranslatable
::unuseable::unusable
::unusuable::unusable
::unwarrented::unwarranted
::unweildly::unwieldy
::unwieldly::unwieldy
::tjpanishad::upanishad
::upcomming::upcoming
::upgradded::upgraded
::useage::usage
::uise::use
::usefull::useful
::usefuly::usefully
::useing::using
::usally::usually
::usualy::usually
::usualyl::usually
::ususally::usually
::vaccum::vacuum
::vaccume::vacuum
::vaguaries::vagaries
::vailidty::validity
::valetta::valletta
::valuble::valuable
::valueable::valuable
::varient::variant
::varations::variations
::vaieties::varieties
::varities::varieties
::variey::variety
::varity::variety
::vreity::variety
::vriety::variety
::varous::various
::varing::varying
::vasall::vassal
::vasalls::vassals
::vegitable::vegetable
::vegtable::vegetable
::vegitables::vegetables
::vegatarian::vegetarian
::vehicule::vehicle
::vengance::vengeance
::vengence::vengeance
::venemous::venomous
::verfication::verification
::vermillion::vermilion
::versitilaty::versatility
::versitlity::versatility
::verison::version
::verisons::versions
::veyr::very
::vrey::very
::vyer::very
::vyre::very
::vacinity::vicinity
::vincinity::vicinity
::vitories::victories
::wiew::view
::vigilence::vigilance
::vigourous::vigorous
::villification::vilification
::villify::vilify
::villian::villain
::violentce::violence
::virgina::Virginia
::virutal::virtual
::virtualyl::virtually
::visable::visible
::visably::visibly
::visting::visiting
::vistors::visitors
::volcanoe::volcano
::volkswagon::Volkswagen
::voleyball::volleyball
::volontary::voluntary
::volonteer::volunteer
::volounteer::volunteer
::volonteered::volunteered
::volounteered::volunteered
::volonteering::volunteering
::volounteering::volunteering
::volonteers::volunteers
::volounteers::volunteers
::vulnerablility::vulnerability
::vulnerible::vulnerable
::watn::want
::whant::want
::wnat::want
::wan tit::want it
::wanna::want to
::wnated::wanted
::whants::wants
::wnats::wants
::wardobe::wardrobe
::warrent::warrant
::warantee::warranty
::warrriors::warriors
::wass::was
::weas::was
::ws::was
::wa snot::was not
::wasnt::wasn't
::wya::way
::wayword::wayward
::we;d::we'd
::weaponary::weaponry
::wether::weather
::wendsay::Wednesday
::wensday::Wednesday
::wiegh::weigh
::wierd::weird
::vell::well
::werre::were
::wern't::weren't
::waht::what
::whta::what
::what;s::what's
::wehn::when
::whn::when
::whent he::when the
::wehre::where
::wherre::where
::where;s::where's
::wereabouts::whereabouts
::wheras::whereas
::wherease::whereas
::whereever::wherever
::whther::whether
::hwich::which
::hwihc::which
::whcih::which
::whic::which
::whihc::which
::whlch::which
::wihch::which
::whicht he::which the
::hwile::while
::woh::who
::who;s::who's
::hwole::whole
::wohle::whole
::wholey::wholly
::widesread::widespread
::weilded::wielded
::wief::wife
::iwll::will
::wille::will
::wiull::will
::willbe::will be
::will of::will have
::willingless::willingness
::windoes::windows
::wintery::wintry
::iwth::with
::whith::with
::wih::with
::wiht::with
::withe::with
::witht::with
::witn::with
::wtih::with
::witha::with a
::witht he::with the
::withthe::with the
::withdrawl::withdrawal
::witheld::withheld
::withold::withhold
::withing::within
::womens::women's
::wo'nt::won't
::wonderfull::wonderful
::wrod::word
::owrk::work
::wokr::work
::wrok::work
::wokring::working
::wroking::working
::workststion::workstation
::worls::world
::worstened::worsened
::owudl::would
::owuld::would
::woudl::would
::wuould::would
::wouldbe::would be
::would of::would have
::woudln't::wouldn't
::wouldnt::wouldn't
::wresters::wrestlers
::rwite::write
::wriet::write
::wirting::writing
::writting::writing
::writen::written
::wroet::wrote
::x-Box::Xbox
::xenophoby::xenophobia
::yatch::yacht
::yaching::yachting
::eyar::year
::yera::year
::eyars::years
::yeasr::years
::yeras::years
::yersa::years
::yelow::yellow
::eyt::yet
::yeild::yield
::yeilding::yielding
::yoiu::you
::ytou::you
::yuo::you
::youare::you are
::you;d::you'd
::your a::you're a
::your an::you're an
::your her::you're her
::your here::you're here
::your his::you're his
::your my::you're my
::your the::you're the
::your their::you're their
::your your::you're your
::youve::you've
::yoru::your
::yuor::your
::you're own::your own
::youself::yourself
::youseff::yousef
::zeebra::zebra
::sionist::Zionist
::sionists::Zionists

;------------------------------------------------------------------------------
; Ambiguous entries.  Where desired, pick the one that's best for you, edit,
; and move into the above list or, preferably, the autocorrect user file.
;------------------------------------------------------------------------------
/*
:*:cooperat::coöperat
::(c)::©
::(r)::®
::(tm)::™
::a gogo::à gogo
::abbe::abbé
::accension::accession, ascension
::achive::achieve, archive
::achived::achieved, archived
::ackward::awkward, backward
::addres::address, adders
::adress::address, A dress
::adressing::addressing, dressing
::afair::affair, afar, Afar (African place), a fair, acronym "as far as I recall"
::affort::afford, effort
::agin::again, a gin, aging
::agina::again, angina
::ago-go::àgo-go
::aledge::allege, a ledge
::alot::a lot, allot
::alusion::allusion, illusion
::amature::armature, amateur
::anu::añu
::anual::annual, anal
::anual::annual, manual
::aparent::apparent, a parent
::apon::upon, apron
::appealling::appealing, appalling
::archaoelogy::archeology, archaeology
::archaology::archeology, archaeology
::archeaologist::archeologist, archaeologist
::archeaologists::archeologists, archaeologists
::assosication::assassination, association
::attaindre::attainder, attained
::attened::attended or attend
::baout::about, bout
::beggin::begin, begging
::behavour::behavior, behaviour
::belives::believes, beliefs
::boaut::bout, boat, about
::Bon::Bön

::assasined::assassinated Broken by ":*:assasin::", but no great loss.
::Bootes::Boötes
::bric-a-brac::bric-à-brac
::buring::burying, burning, burin, during
::busineses::business, businesses
::cafe::café
::calaber::caliber, calibre
::calander::calendar, calender, colander
::cancelled::canceled  ; commonwealth vs US
::cancelling::canceling  ; commonwealth vs US
::canon::cañon
::cant::cannot, can not, can't
::carcas::carcass, Caracas
::carmel::caramel, carmel-by-the-sea
::Cataline::Catiline, Catalina
::censur::censor, censure
::ceratin::certain, keratin
::cervial::cervical, servile, serval
::chasr::chaser, chase
::clera::clear, sclera
::comander::commander, commandeer
::competion::competition, completion
::continuum::continuüm
::coopt::coöpt
::coordinat::coördinat
::coorperation::cooperation, corporation
::coudl::could, cloud
::councellor::councillor, counselor, councilor
::councellors::councillors, counselors, councilors
::coururier::courier, couturier
::coverted::converted, covered, coveted
::cpoy::coy, copy
::creme::crème
::dael::deal, dial, dahl
::deram::dram, dream
::desparate::desperate, disparate
::diea::idea, die
::dieing::dying, dyeing
::diversed::diverse, diverged
::divorce::divorcé
::Dona::Doña
::doub::doubt, daub
::dyas::dryas, Dyas (Robert Dyas is a hardware chain), dais
::efford::effort, afford
::effords::efforts, affords
::eigth::eighth, eight
::electic::eclectic, electric
::electon::election, electron
::elite::élite
::emition::emission, emotion
::emminent::eminent, imminent
::empirial::empirical, imperial
::Enlish::English, enlist
::erally::orally, really
::erested::arrested, erected
::ethose::those, ethos
::etude::étude
::expose::exposé
::extint::extinct, extant
::eyar::year, eyas
::eyars::years, eyas
::eyasr::years, eyas
::fiel::feel, field, file, phial
::fiels::feels, fields, files, phials
::firts::flirts, first
::fleed::fled, freed
::fo::for, of
::fomr::from, form
::fontrier::fontier, frontier
::fro::for, to and fro, (a)fro
::futhroc::futhark, futhorc
::gae::game, Gael, gale
::gaurd::guard, gourd
::gogin::going, Gauguin
::Guaduloupe::Guadalupe, Guadeloupe
::Guadulupe::Guadalupe, Guadeloupe
::guerrila::guerilla, guerrilla
::guerrilas::guerillas, guerrillas
::haev::have, heave
::Hallowean::Hallowe'en, Halloween
::herad::heard, Hera
::housr::hours, house
::hten::then, hen, the
::htere::there, here
::humer::humor, humour
::humerous::humorous, humourous, humerus
::hvea::have, heave
::idesa::ideas, ides
::imaginery::imaginary, imagery
::imanent::eminent, imminent
::iminent::eminent, imminent, immanent
::indispensable::indispensible ; commonwealth vs US?
::indispensible::indispensable ; commonwealth vs US?
::inheritage::heritage, inheritance
::inspite::in spite, inspire
::interbread::interbreed, interbred
::intered::interred, interned
::inumerable::enumerable, innumerable
::israelies::Israelis, Israelites
::labatory::lavatory, laboratory
::labled::labelled, labeled
::lame::lamé
::leanr::lean, learn, leaner
::lible::libel, liable
::liscense::license, licence
::lisence::license, licence
::lisense::license, licence
::lonly::lonely, only
::maked::marked, made
::managable::manageable, manageably
::manoeuver::maneuver ; Commonwealth vs US?
::manouver::maneuver, manoeuvre
::manouver::manoeuvre ; Commonwealth vs US?
::manouverability::maneuverability, manoeuvrability, manoeuverability
::manouverable::maneuverable, manoeuvrable
::manouvers::maneuvers, manoeuvres
::manuever::maneuver, manoeuvre
::manuevers::maneuvers, manoeuvres
::mear::wear, mere, mare
::meranda::veranda, Miranda
::Metis::Métis
::mit::mitt, M.I.T., German "with"
::monestary::monastery, monetary
::moreso::more, more so
::muscels::mussels, muscles
::ne::né
::neice::niece, nice
::neigbour::neighbour, neighbor
::neigbouring::neighbouring, neighboring
::neigbours::neighbours, neighbors
::nto:: not ; Replaced with case sensitive for NTO acronym.
::oging::going, ogling
::ole::olé
::onot::note, not
::opium::opïum
::ore::öre
::ore::øre
::orgin::origin, organ
::palce::place, palace
::pate::pâte
::pate::pâté
::performes::performed, performs
::personel::personnel, personal
::positon::position, positron
::preëmpt
::premiere::première
::premiered::premièred
::premieres::premières
::premiering::premièring
::procede::proceed, precede
::proceded::proceeded, preceded
::procedes::proceeds, precedes
::proceding::proceeding, preceding
::profesion::profusion, profession
::progrom::pogrom, program
::progroms::pogroms, programs
::prominately::prominently, predominately
::qtuie::quite, quiet
::qutie::quite, quiet
::reenter::reënter
::relized::realised, realized
::repatition::repetition, repartition
::residuum::residuüm
::restraunt::restraint, restaurant
::resume::résumé
::rigeur::rigueur, rigour, rigor
::role::rôle
::rose::rosé
::sasy::says, sassy
::scholarstic::scholastic, scholarly
::secceeded::seceded, succeeded
::seceed::succeed, secede
::seceeded::succeeded, seceded
::sepulchure::sepulchre, sepulcher
::sepulcre::sepulchre, sepulcher
::shamen::shaman, shamans
::sheat::sheath, sheet, cheat
::shoudln::should, shouldn't
::sieze::seize, size
::siezed::seized, sized
::siezing::seizing, sizing
::sinse::sines, since
::snese::sneeze, sense
::sotyr::satyr, story
::sould::could, should, sold
::speciallized::specialised, specialized
::specif::specific, specify
::spects::aspects, expects
::strat::start, strata
::stroy::story, destroy
::surley::surly, surely
::surrended::surrounded, surrendered
::thast::that, that's
::theather::theater, theatre
::ther::there, their, the
::thikning::thinking, thickening
::throught::thought, through, throughout
::tiem::time, Tim
::tiome::time, tome
::tourch::torch, touch
::transcripting::transcribing, transcription
::travelling::traveling   ; commonwealth vs US
::troups::troupes, troops
::turnk::turnkey, trunk
::uber::über
::unmanouverable::unmaneuverable, unmanoeuvrable
::unsed::used, unused, unsaid
::vigeur::vigueur, vigour, vigor
::villin::villi, villain, villein
::vistors::visitors, vistas
::wanna::want to - often deliberate
::weild::wield, wild
::wholy::wholly, holy
::wich::which, witch
::withdrawl::withdrawal, withdraw
::woulf::would, wolf
::ws::was, www.example.ws
::Yementite::Yemenite, Yemeni
:?:oology::oölogy
:?:t he:: the  ; Can't use this. Needs to be cleverer.
*/

;-------------------------------------------------------------------------------
;  Capitalise dates
;-------------------------------------------------------------------------------
::monday::Monday
::tuesday::Tuesday
::wednesday::Wednesday
::thursday::Thursday
::friday::Friday
::saturday::Saturday
::sunday::Sunday

::january::January
::february::February
; ::march::March  ; Commented out because it matches the common word "march".
::april::April
; ::may::May  ; Commented out because it matches the common word "may".
::june::June
::july::July
::august::August
::september::September
::october::October
::november::November
::december::December


;-------------------------------------------------------------------------------
; Anything below this point was added to the script by the user via the Win+H hotkey.
;-------------------------------------------------------------------------------

;Code from KG

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;            Alt+Shift+F Open Form and results

!+F::
;sleep, 300
IfWinActive, Medical Validation :   (Authorise By Episode)
{
episode:=EpMedVal()
;txt := ClipBoard
;sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . episode . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
sleep, 1000
send, !{tab}
sleep, 1000
}
IfWinActive, Medical Validation :   (Authorise By Episode)
{
mrn := MRNMedVal()
;txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . mrn
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
sleep, 300
Remove_IE_notice()
return
}
else
IfWinActive, Result Verify - Single -
{
episode := EpResultVerSingle()
txt := ClipBoard
url := "http://172.22.4.40/multipdfsearch.php?file=" . episode . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
sleep, 1000
send, !{tab}
sleep, 1000
}
IfWinActive, Result Verify - Single -
{
MRNResultVerSingle()
txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
sleep, 300
Remove_IE_notice()
return
}
else
IfWinActive,  Result Entry - Single -
{
episode := EpResultSingle()
txt := ClipBoard
;sleep, 100
url := "http://172.22.4.40/multipdfsearch.php?file=" . episode . ""
run, %url%
sleep, 1500
WinMove, Internet Explorer ahk_class IEFrame,, A_ScreenWidth-537, 0, 545, A_ScreenHeight-25
sleep, 1000
send, !{tab}
sleep, 1000
}
IfWinActive, Result Entry - Single -
{
MRNSingle()
txt := Clipboard
url := "https://trakdb-prod.nhls.ac.za/csp/reporting/epr.csp?PAGE=4&vstRID=*&MRN=" . txt
run, %url%
sleep, 800
WinMove, Internet Explorer ahk_class IEFrame,, 0, 0, A_ScreenWidth-0, A_ScreenHeight-25
sleep, 300
Remove_IE_notice()
return
}
