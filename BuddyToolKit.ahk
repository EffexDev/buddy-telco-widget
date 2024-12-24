#Requires AutoHotkey v2.0
#Include Templates.ahk

; Hotkey to reload the script. I mostly use this for debugging and testing
!F1::
{
    Reload
}

;Below is just the variable declarations for the template maps to be global, as well as setting the customername field. Ig it is not set you might accidentally pout the wrong customers name in
IniWrite("xxx", "config.ini", "Customer", "CustomerName")

Global LiveChatMap := Map()
Global GeneralNBNMap := Map()
Global GeneralHardwareMap := Map()
Global GeneralOTRSMap := Map()
Global RedmineJiraMap := Map()
Global PPMap := Map()
Global FHMap := Map()
Global ReconnectionMap := Map()
Global CallbackMap := Map()
Global FaultTemplatesMap := Map()
Global DiscoveryMap := Map()
Global SpeedsMap := Map()
Global DropoutsMap := Map()
Global ConnectionMap := Map()
Global SetupMap := Map()
Global LinkMap := Map()
Global HardwareMap := Map()
Global ServiceRequestMap := Map()
Global DelaysMap := Map()
Global HFCMap := Map()
Global FTTCMap := Map()
Global FTTPMap := Map()
Global ValidationMap := Map()
Global BanlistingMap := Map()
Global FailedPaymentMap := Map()
Global PaymentMap := Map()
Global NBNCompMap := Map()
Global RaisingMap := Map()
Global ClarificationMap := Map()
Global ResolutionsMap := Map()
Global ChangesMap := Map()
Global TIOMap := Map()
Global ContactsMap := Map()
Global TCSGeneralMap := Map()
Global TCSBillingMap := Map()
Global TCSSuspensionMap := Map()
Global TCSChangesMap := Map()

Global NotesGui := Gui(,"Notepad"), Notes := NotesGui.Add("Edit", "h600 w685", "")
Global TemplatesGui := Gui(,"Templates"), Templates := TemplatesGui.Add("Edit", "h600 w685", "")
Global HotkeysGui := Gui(,"Useful Links"), Hotkeys := HotkeysGui.Add("Text", "+Wrap h230 cFFFFFF", "Available Hotkeys:`n`n@@ - Your email`n`n~~ - Date 7 calendar days from now. For abandoment comms`n`nCTRL+Shift+S - Randomized signature for app faults. Also checks the publish to app checkbox`n`nCTRL+DEL - Content aware search ie. Superlookup")

; Making all of the popout GUIs consistent
Hotkeys.SetFont("s10","Nunito")
Notes.SetFont("s10","Nunito")
Templates.SetFont("s10","Nunito")

NotesGui.BackColor := "c007ba8"
TemplatesGui.BackColor := "c007ba8"
HotkeysGui.BackColor := "c007ba8"

; The main body of the GUI itself. Dimensions and tabs etc
Global BuddyGui := Gui("-Caption +Border","Buddy Tool Kit V2.0")
BuddyGui.BackColor := "c007ba8"
BuddyGui.Add("Picture", "ym+10 x+20 w180 h-1","BuddyLogo.png")
DogImage := BuddyGui.Add("Picture", "ym xm+480 ym w-1 h90","BuddyPC.png")
DogImage.OnEvent("Click", WhatTheDogDoin)
BuddyGui.SetFont("s10 c000000","Nunito")
BuddyGui.Add("Text", " xm cFFFFFF" , "Customer Name:")
Global CustomerNameField := BuddyGui.Add("Edit", "yp-3 xm+105 w150 vCustomerNameValue", "").OnEvent("Change", CustomerNameEdit)
TemplateTab := BuddyGui.Add("Tab3","xm h70 w610 BackgroundWhite", ["General", "Accounts", "Faults","Order Support","Complaints","T and Cs"])
ToolsTab := BuddyGui.Add("Tab3", "WP h80 BackgroundWhite c222222 vToolsTab", ["QOL", "Automations", "Useful Links", "Options"])

iconsize := 32
hIcon := LoadPicture("TaskBarIcon.ico", "Icon1 w" iconsize " h" iconsize, &imgtype)
SendMessage(0x0080, 0, hIcon, BuddyGui)
SendMessage(0x0080, 1, hIcon, BuddyGui) 
BuddyGui.Show("x1920 y0 w630")

TraySetIcon(A_ScriptDir "\TaskBarIcon.ico") 

;DO NOT REMOVE it will break stuff
Send "xxx"

;First set of tabs, for department selection to segregate templates and keep things organised. This grabs the options selected in both dropdowns and saves them into a variable to be used later.
TemplateTab.UseTab(1)
SelGeneralReason := BuddyGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedGeneralReason Choose1", GeneralReasons)
SelGeneralReason.OnEvent('Change', SelGeneralReasonSelected)
SelGeneralTemplate := BuddyGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedGeneral", GeneralTemplates
[SelGeneralReason.Value])
GenerateFault := BuddyGui.Add("Button", "yp", "Generate").OnEvent("Click", RunGeneral)

TemplateTab.UseTab(2)
SelAccountReason := BuddyGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedAccountReason Choose1", AccountReasons)
SelAccountReason.OnEvent('Change', SelAccountReasonSelected)
SelAccountTemplate := BuddyGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedAccount", AccountTemplates
[SelAccountReason.Value])
GenerateFault := BuddyGui.Add("Button", "yp", "Generate").OnEvent("Click", RunAccount)

TemplateTab.UseTab(3)
SelFaultReason := BuddyGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedFaultReason Choose1", FaultReasons)
SelFaultReason.OnEvent('Change', SelFaultReasonSelected)
SelFaultTemplate := BuddyGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedFault", FaultTemplates[SelFaultReason.Value])
GenerateFault := BuddyGui.Add("Button", "yp", "Generate").OnEvent("Click", RunFault)

TemplateTab.UseTab(4)
SelDeliveryReason := BuddyGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedDeliveryReason Choose1", DeliveryReasons)
SelDeliveryReason.OnEvent('Change', SelDeliveryReasonSelected)
SelDeliveryTemplate := BuddyGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedDelivery", DeliveryTemplates[SelDeliveryReason.Value])
GenerateFault := BuddyGui.Add("Button", "yp", "Generate").OnEvent("Click", RunDelivery)

TemplateTab.UseTab(5)
SelComplaintReason := BuddyGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedComplaintReason Choose1", ComplaintReasons)
SelComplaintReason.OnEvent('Change', SelComplaintReasonSelected)
SelComplaintTemplate := BuddyGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedComplaint", ComplaintTemplates[SelComplaintReason.Value])
GenerateFault := BuddyGui.Add("Button", "yp", "Generate").OnEvent("Click", RunComplaint)

TemplateTab.UseTab(6)
SelTCSReason := BuddyGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedTCSReason Choose1", TCSReasons)
SelTCSReason.OnEvent('Change', SelTCSReasonSelected)
SelTCSTemplate := BuddyGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedTCS", TCSTemplates[SelTCSReason.Value])
GenerateFault := BuddyGui.Add("Button", "yp", "Generate").OnEvent("Click", RunTCS)

; The tools tab. Controls the bottom set of tabs and the content of them. Functions are in the FunctionsLibrary file
ToolsTab.UseTab(1)
Superlookup := BuddyGui.Add("Button","xm+15 y245 vSuperlookup", "Superlookup").OnEvent("Click", ProcessSuperlookup)
BuddyGui.Add("Button", "yp w90", "Notes").OnEvent("Click", NotePad)
BuddyGui.Add("Button","yp w90", "Hotkeys").OnEvent("Click", HotkeysPad)
BuddyGui.Add("Button","yp w90","Startup").OnEvent("Click", Startup)
BuddyGui.Add("Button","yp w100", "Lock Terminal").OnEvent("Click", LockTerminal)
NotePadEmbedded := BuddyGui.Add("Edit", "yp+40 xm+10 h510 w585 vNotePadEmbedded", "")

ToolsTab.UseTab(2)
BuddyGui.Add("Button", "xm+15 y245 Section", "Ping Test").OnEvent("Click", PingTest)
BuddyGui.Add("Button", "yp", "Traceroute").OnEvent("Click", Traceroute)
BuddyGui.Add("Button", "yp", "NSLookup").OnEvent("Click", NSLookup)
BuddyGui.Add("Button", "yp", "Prorata Calc").OnEvent("Click", ProRataCalc)

ToolsTab.UseTab(3)
BuddyGui.Add("Button", "xm+15 y245 Section", "CMS").OnEvent("Click", RunCMS)
BuddyGui.Add("Button", "yp", "Order Support").OnEvent("Click", RunOrderSupport)
BuddyGui.Add("Button", "yp", "NBN SQ").OnEvent("Click", RunNBNSQ)
BuddyGui.Add("Button", "yp", "Complaints").OnEvent("Click", RunComplaints)
BuddyGui.Add("Button", "yp", "Buddy Website").OnEvent("Click", RunBuddy)
BuddyGui.Add("Button", "yp", "Outages").OnEvent("Click", RunOutages)
BuddyGui.Add("Button", "yp", "ChatGPT").OnEvent("Click", RunGPT)

ToolsTab.UseTab(4)
Global AlwaysOnTopButton := BuddyGui.Add("Checkbox", "xm+15 y245 Section vAlwaysOnTop ").OnEvent("Click", AlwaysOnTopToggle)
AlwaysOnTopCheckBoxText := BuddyGui.Add("Text", "yp xp+20 c000000", "Always on Top")
Global ShowNotesButton := BuddyGui.Add("Checkbox", "yp x+20 vShowNotesButton").OnEvent("Click", ShowNotes)
ShowNotesButtonText := BuddyGui.Add("Text", "yp xp+20 c000000", "Show Notepad")
Global DarkmodeButton := BuddyGui.Add("Checkbox", "yp x+20 vDarkModeButton ").OnEvent("Click", Darkmode)
DarkmodeButtonText := BuddyGui.Add("Text", "yp xp+20 c000000", "Darkmode")
Global UpdateButton := BuddyGui.Add("Button", "yp-5 x+120", "Check for Updates")
UpdateButton.OnEvent("Click", UpdateWidgetCheck)
BuddyGui["NotePadEmbedded"].Visible := 0

; Customer name edit field. This is sanitised because tabs and symbols can cause output generation errors
global CustomerName := ""  
CustomerNameEdit(CustomerNameValue, *) {
    global CustomerName
    CustomerName := BuddyGui["CustomerNameValue"].Value
    global CustomerNameSanitised := RegExReplace(CustomerName, "[^a-zA-Z]", "")
    IniWrite(CustomerNameSanitised, "config.ini", "Customer", "CustomerName")
    UpdateTemplates()
}

; Logic for cascading dropdowns
SelGeneralReasonSelected(*) 
{
    SelGeneralTemplate.Delete()
    SelGeneralTemplate.Add(GeneralTemplates[SelGeneralReason.value])
    SelGeneralTemplate.Choose(1)
}

SelAccountReasonSelected(*) 
{
    SelAccountTemplate.Delete()
    SelAccountTemplate.Add(AccountTemplates[SelAccountReason.value])
    SelAccountTemplate.Choose(1)
}

SelFaultReasonSelected(*) 
{
    SelFaultTemplate.Delete()
    SelFaultTemplate.Add(FaultTemplates[SelFaultReason.value])
    SelFaultTemplate.Choose(1)
}

SelDeliveryReasonSelected(*) 
{
    SelDeliveryTemplate.Delete()
    SelDeliveryTemplate.Add(DeliveryTemplates[SelDeliveryReason.value])
    SelDeliveryTemplate.Choose(1)
}

SelComplaintReasonSelected(*) 
{
    SelComplaintTemplate.Delete()
    SelComplaintTemplate.Add(ComplaintTemplates[SelComplaintReason.value])
    SelComplaintTemplate.Choose(1)
}

SelTCSReasonSelected(*) 
{
    SelTCSTemplate.Delete()
    SelTCSTemplate.Add(TCSTemplates[SelTCSReason.value])
    SelTCSTemplate.Choose(1)
}

#Include Settings.ahk
#Include Generate.ahk
#Include FunctionLibrary.ahk
