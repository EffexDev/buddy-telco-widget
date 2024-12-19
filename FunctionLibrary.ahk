;------------------------ Make it Draggable --------------------------

LShift & LButton::
EWD_MoveWindow(*)
{
    CoordMode "Mouse"  ; Switch to screen/absolute coordinates.
    MouseGetPos &EWD_MouseStartX, &EWD_MouseStartY, &EWD_MouseWin
    WinGetPos &EWD_OriginalPosX, &EWD_OriginalPosY,,, EWD_MouseWin
    if !WinGetMinMax(EWD_MouseWin)  ; Only if the window isn't maximized 
        SetTimer EWD_WatchMouse, 10 ; Track the mouse as the user drags it.

    EWD_WatchMouse()
    {
        if !GetKeyState("LButton", "P")  ; Button has been released, so drag is complete.
        {
            SetTimer , 0
            return
        }
        if GetKeyState("Escape", "P")  ; Escape has been pressed, so drag is cancelled.
        {
            SetTimer , 0
            WinMove EWD_OriginalPosX, EWD_OriginalPosY,,, EWD_MouseWin
            return
        }
        ; Otherwise, reposition the window to match the change in mouse coordinates
        ; caused by the user having dragged the mouse:
        CoordMode "Mouse"
        MouseGetPos &EWD_MouseX, &EWD_MouseY
        WinGetPos &EWD_WinX, &EWD_WinY,,, EWD_MouseWin
        SetWinDelay -1   ; Makes the below move faster/smoother.
        WinMove EWD_WinX + EWD_MouseX - EWD_MouseStartX, EWD_WinY + EWD_MouseY - EWD_MouseStartY,,, EWD_MouseWin
        EWD_MouseStartX := EWD_MouseX  ; Update for the next timer-call to this subroutine.
        EWD_MouseStartY := EWD_MouseY
    }
}

;------------------------- Tools -------------------------------------

; Runs a ping test using an IP address that is saved to the clipboard. Will open CMD if it is not open already
PingTest(*)
{
    if WinExist("C:\Windows\SYSTEM32\cmd.exe")
    WinActivate
else
    Run "cmd.exe"
    Sleep 300
    Send "ping " A_Clipboard "{Enter}"
    return
}

; Runs a traceroute using an IP address that is saved to the clipboard. Will open CMD if it is not open already
Traceroute(*)
{
    if WinExist("C:\Windows\SYSTEM32\cmd.exe")
        WinActivate
    else
        Run "cmd.exe"
        Sleep 300
        Send "tracert " A_Clipboard "{Enter}"
        return
}

; Runs a nslookup using an IP address that is saved to the clipboard. Will open CMD if it is not open already
NSLookup(*)
{
    if WinExist("C:\Windows\SYSTEM32\cmd.exe")
        WinActivate
    else
        Run "cmd.exe"
        Sleep 300
        Send "nslookup " A_Clipboard "{Enter}"
        return
}

; This is a function to open a GUI containing two calendars, 2 edit fields and a submit button. It's for calculating prorata credits
ProRataCalc(*) {
    ProrataGui := Gui(,"Buddy Tool Kit")
    ProrataGui.BackColor := "c007ba8"
    ProrataGui.SetFont("s10")
    ProrataGui.Add("Text", "cFFFFFF", "Start of billing cycle")
    ProrataGui.Add("MonthCal", "yp+20 vBillingStart")
    ProrataGui.Add("Text", "yp-20 xp+240 cFFFFFF", "Service Closure Date")
    ProrataGui.Add("MonthCal", "yp+20 vServiceEnd")
    ProrataGui.Show("w490 h300")
    ProrataGui.Add("Edit","xm w100 vMonthlyCost", "")
    ProrataGui.Add("Text","yp cFFFFFF", "Enter the monthly billing amount")
    ProrataGui.Add("Edit","xm w100 vdaysInMonth", "")
    ProrataGui.Add("Text","yp cFFFFFF", "How many days this month?")
    ProrataGui.Add("Button","xm", "Calculate").OnEvent("Click", PRCalcBox)

        ; Generates a third GUI with the correct prorata amount
        PRCalcBox(*) {
            Saved:= ProrataGui.Submit(False)
            DaysPassed := DateDiff(Saved.ServiceEnd, Saved.BillingStart, "days")
            Month := 
            DailyCost := Saved.MonthlyCost / Saved.daysInMonth
            DaysUsed := DailyCost * DaysPassed
            ProrataAmount := Saved.MonthlyCost - DaysUsed
            MsgBox "The prorata amount is " ProrataAmount
        }
}

; Generates a GUI with a large blank input field for typing notes into. There is no check to see if it is already open as having more than one open could be intended. THIS IS NOT PERSISTENMT AND CANNOT SAVE FILES
Notepad(DarkmodeButton, *) {
    Global NotesGui, Notes    

    if !NotesGui {  ; Only create a new GUI if it doesn't exist    
        NotesGui := Gui(,"Notepad")
        NotesGui.SetFont("s10","Nunito")
        NotesGui.BackColor := "c007ba8"
        Notes := NotesGui.Add("Edit", "h600 w685", "")
        
        NotesGui.OnEvent("Close", (*) => (NotesGui := "", Notes := ""))
    }
    NotesGui.Show("w710 h620")
}

; Generates a GUI with an input field. Will only open if it is not already open to prevent duplicates. 
TemplatesPad(*) {
    Global TemplatesGui, Templates

    if !TemplatesGui {  ; Only create a new GUI if it doesn't exist
        TemplatesGui := Gui(,"Templates")
        TemplatesGui.SetFont("s10","Nunito")
        TemplatesGui.BackColor := "c007ba8"
        Templates := TemplatesGui.Add("Edit", "h600 w685", "")
        
        ; Add handler for GUI close to reset the variables
        TemplatesGui.OnEvent("Close", (*) => (TemplatesGui := "", Templates := ""))
    }
    TemplatesGui.Show("w710 h620")
}

; Generates a static GUI with a list of hotkeys that do not have GUI buttons
HotkeysPad(*) {
    Global HotkeysGui, Hotkeys
    if !HotkeysGui {
        HotkeysGui := Gui(,"Available Hotkeys")
        HotkeysGui.SetFont("s10","Nunito")
        HotkeysGui.BackColor := "c007ba8"
        Hotkeys := HotkeysGui.Add("Text", "+Wrap cFFFFFF", "Available Hotkeys:`n`n@@ - Your email`n`n~~ - Date 7 calendar days from now. For abandoment comms`n`nCTRL+S - Randomized signature for app faults. Also checks the publish to app checkbox`n`nCTRL+DEL - Content aware search ie. Superlookup")
        
        HotkeysGui.OnEvent("Close", (*) => (HotkeyssGui := "", Hotkeys := ""))
    }
    HotkeysGui.Show("w710 h250")
}

; Runs when you press the lock terminal button. Locks the Terminal.
LockTerminal(*)
{
    Run "rundll32 user32.dll`,LockWorkStation"
}

; Hotkey to run the function below
^Del::ProcessSuperlookup()

; Content-aware search. Can be run with the hotkey above. Just a bunch of regex to search the clipboard for certain strings and then append them to URLs and then open said URL
ProcessSuperlookup(*)
{
    ;NOC Jira Tickets
    if (RegExMatch(A_Clipboard, "(NOC)\-\d*", &Match)) 
        Run "https://aussiebb.atlassian.net/servicedesk/customer/portal/18/" MATCH[0]
    ;NBN AVC/INC/ORD/PRI/WRI/HRI/CVCs
    else if (RegExMatch(A_Clipboard, "(AVC|INC|ORD|PRI|WRI|HRI|CVC)\d*", &Match))
        Run "https://nbnportals.nbnco.net.au/online_customers/page/home?headerSearch=" MATCH[0]
    ;NBN CRQs
    else if (RegExMatch(A_Clipboard, "^CRQ\d*", &Match))
        Run "https://nbnportals.nbnco.net.au/online_customers/page/change_activity/list?criteriaType=CHANGE_REF_NO&criteria=" MATCH[0]
    ;NBN APTs
    else if (RegExMatch(A_Clipboard, "APT\d*", &Match))
        Run "https://nbnportals.nbnco.net.au/online_customers/page/appointment/view/ASI000000001104/" MATCH[0]
	;NBN LOC
    else if (RegExMatch(A_Clipboard, "LOC\d+", &Match))
        Run "https://nbnportals.nbnco.net.au/online_customers/page/manageaddress/site_qualification/search?address.nbnLocationId=" MATCH[0]
    ;Australia Post T/Ns
    else if (RegExMatch(A_Clipboard, "(?:R|92A|34ECK|I8|2KUZ|34TDC|36AAC|36BRU|36LFM|36LDQ|030|0207|47U4|4GPZ|4HBZ)\d+", &Match))
        Run "https://auspost.com.au/mypost/track/#/details/" MATCH [0]
    ;Sites
    else if (RegExMatch(A_Clipboard, "^http(s)?:\/\/|www\.", &Match))
        Run A_Clipboard
    ;Mobile Numbers
    else if (RegExMatch(A_Clipboard, "^04[\d\s]+\s?$", &Match))
        Run "https://cms.aussiebroadband.com.au/?type=billing_name&search=" MATCH[0]
    ;Emails
    else if (RegExMatch(A_Clipboard, "^[^@]+@[^@]+\.[^@]+$", &Match))
        Run "https://cms.aussiebroadband.com.au/?type=email&search=" MATCH[0]

    else if (RegExMatch(A_Clipboard, "^16937\d{6}$", &Match))
        Run "https://www.speedtest.net/result/" MATCH[0]
    
    else if (RegExMatch(A_Clipboard, "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$", &Match))
        Run "https://ipinfo.io/" MATCH[0]
    ;Google search
    else {
        Run "https://www.google.com/search?q=" A_Clipboard
	}
}

; Runs the stuff I need to work for the day. Only opens a single instance of them
Startup(*)
{
        if WinExist("C:\Program Files\Slack\slack.exe")
        {
            WinActivate
        } 
    else 
        {
            Run "C:\Program Files\Slack\Slack.exe"
        }

    if WinExist("ahk_exe C:\Program Files\Google\Chrome\Application\msedge.exe")
        {
            WinActivate
        }
    else 
        {
            Run "msedge.exe"
        }
}

WhatTheDogDoin(*) {
    ToolTip("I don't talk")
    SetTimer () => ToolTip(), -2000
}

SaveCustomerName(ctrl, *)
{
    Global ShowNotesButton
    ; Takes the options submitted from the dropdowns and saves them to an array called "Saved"
    Saved:= BuddyGui.Submit(False)

    Global Customer := Saved.CustomerName
}

; Sets the variables for the logged in users full name, and first name.
csFullName := A_UserName
csFirstName := ""
Match:=""
;  Declares a Dice variable at 0
Dice:= 0
; Declares an array of sign offs for CSP faults
sign := ["Regards,", "Cheers,", "Have a great day{!}", "All the best,"]

; Regex matches the full name, removes anything after the full stop and converts to Title Case to get the first name in a variable.
if (RegExMatch(csFullName, "^[^.]*",&csFirstName)) {
csTitle:=StrTitle(csFirstName[0])
}

:*:@@::  ; Double @ for sending email address to input field
{    
    SendInput csFullName
    SendInput "@team.aussiebroadband.com.au"
}

:*:~~:: ; Gives an address 7 calendar days in the future for abandonment comms
{
    CurrentDate := FormatTime(, "yyyyMMdd")
    NewDate := FormatTime(DateAdd(CurrentDate, 7, "days"), "dd/MM/yyyy")
    Send(NewDate)
}

^+s::{ ; Randomized signature. Also auto-ticks the publish to app button and focuses the submit button so you can just hit enter.
    Dice:= Random(1,4)
    Selected := sign.Get(Dice)
    Send (Selected)
    Send "{Enter}"
    Send (csTitle)
    SendInput "{Tab 1}{Space}"
	SendInput "{Tab 3}"
    Dice:= Random(1,4)   
}

RunCMS(*) {
    Run "https://cms.aussiebroadband.com.au/"
}

RunOrderSupport(*) {
    Run "https://cms.aussiebroadband.com.au/nbnapp.php?bc=buddytelco"
}

RunNBNSQ(*) {
    Run "https://cms.aussiebroadband.com.au/nbnsq2.php"
}

RunComplaints(*) {
    Run "https://cms.aussiebroadband.com.au/complaints.php"
}

RunBuddy(*) {
    Run "https://www.buddytelco.com.au/"
}

RunOutages(*) {
    Run "https://www.nbnco.com.au/support/network-status"
    Run "https://www.buddytelco.com.au/network/"
}

RunGPT(*) {
    Run "https://chatgpt.com/"
}

^!z::  ; Control+Alt+Z hotkey.
{
    MouseGetPos &MouseX, &MouseY
    MsgBox "The color at the current cursor position is " PixelGetColor(MouseX, MouseY)
}

VersionNumber := "4.0"

Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/version.ini", A_WorkingDir . "\version.ini")
global VersionNumberCheck := IniRead("version.ini", "Version", "VersionNumber")

if VersionNumberCheck > VersionNumber {
    AutoUpdateGui := Gui("-Caption +AlwaysOnTop","Buddy Tool Kit")
    AutoUpdateGui.BackColor := "c007ba8"
    AutoUpdateGui.SetFont("s10")
    AutoUpdateGui.Show("w150 h70")
    AutoUpdateGui.Add("Text","x+13 y+5 cFFFFFF", "Update Available")
    AutoUpdateGui.Add("Button","xp+20 y+10","Update").OnEvent("Click", UpdateWidget)
}

UpdateWidgetCheck(*) {
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/version.ini", A_WorkingDir . "\version.ini")
    global VersionNumberCheck := IniRead("version.ini", "Version", "VersionNumber")
    
    if VersionNumberCheck > VersionNumber {
        Global CheckUpdateGui := Gui("-Caption +AlwaysOnTop","Buddy Tool Kit")
        CheckUpdateGui.BackColor := "c007ba8"
        CheckUpdateGui.SetFont("s10")
        CheckUpdateGui.Show("w150 h70")
        CheckUpdateGui.Add("Text","x+13 y+5 cFFFFFF", "Update Available")
        CheckUpdateButton := CheckUpdateGui.Add("Button","xp+20 y+10","Update")
        CheckUpdateButton.OnEvent("Click", UpdateWidget)
        CheckUpdateButton.OnEvent("Click", CloseGui)
    }
    else {
        CheckUpdateGui := Gui("-Caption +AlwaysOnTop","Buddy Tool Kit")
        CheckUpdateGui.BackColor := "c007ba8"
        CheckUpdateGui.SetFont("s10")
        CheckUpdateGui.Show("w150 h70")
        CheckUpdateGui.Add("Text","x+13 y+5 cFFFFFF", "You're up to date!")
        CheckUpdateGui.Add("Button","xp+28 y+10","Close").OnEvent("Click", CloseGui)
    }

    CloseGui(*) {
        CheckUpdateGui.Destroy
    }
}

UpdateWidget(*) {
    AutoUpdateGui.Destroy
    CheckUpdateGui.Destroy
    LoadingGui := Gui("-Caption","Buddy Tool Kit")
    LoadingGui.BackColor := "c007ba8"
    LoadingGui.SetFont("s10")
    LoadingGui.Show("w200 h60")
    LoadingGui.Add("Text", "cFFFFFF", "Updating")
    LoadingGui.Add("Progress", " w180 h20 cGreen vMyProgress", 0)
    LoadingGui["MyProgress"].Value := 14
    Sleep "500"
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/Changelog.txt", A_WorkingDir . "\Changelog.txt")
    LoadingGui["MyProgress"].Value := 28
    Sleep "500"
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/BuddyToolKit.ahk", A_WorkingDir . "\BuddyToolKit.ahk")
    LoadingGui["MyProgress"].Value := 42
    Sleep "500"
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/FunctionLibrary.ahk", A_WorkingDir . "\FunctionLibrary.ahk")
    LoadingGui["MyProgress"].Value := 56
    Sleep "500"
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/Generate.ahk", A_WorkingDir . "\Generate.ahk")
    LoadingGui["MyProgress"].Value := 70
    Sleep "1000"
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/Settings.ahk", A_WorkingDir . "\Settings.ahk")
    LoadingGui["MyProgress"].Value := 84
    Sleep "500"
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/Templates.ahk", A_WorkingDir . "\Templates.ahk")
    LoadingGui["MyProgress"].Value := 100
    Sleep "500"
    Download("https://raw.githubusercontent.com/EffexDev/Buddy-Telco-Widget/refs/heads/main/config.ini", A_WorkingDir . "\config.ini")
    LoadingGui.Destroy
    MsgBox "Update Complete"
    Reload
}
