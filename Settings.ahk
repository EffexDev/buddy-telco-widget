AlwaysOnTopSetting := IniRead("config.ini", "Settings", "AlwaysOnTop")
InitialOntopState := AlwaysOnTopSetting == "True" ? 1 : 0

BuddyGui["AlwaysOnTop"].Value := InitialOntopState

ApplyOntop(isOntop) {
    if isOntop {
        BuddyGui.Opt("+AlwaysOnTop")  ; Enable Always on Top
    }
    else {
        BuddyGui.Opt("-AlwaysOnTop")  ; Disable Always on Top
    }
}

AlwaysOnTopToggle(ctrl, *) {
    if ctrl.Value {
        IniWrite("True", "config.ini", "Settings", "AlwaysOnTop")
    } else {
        IniWrite("False", "config.ini", "Settings", "AlwaysOnTop")
    }
    ApplyOntop(ctrl.Value)
}

ApplyOntop(initialOntopState)

NotepadSetting := IniRead("config.ini", "Settings", "ShowNotepad")
InitialNotepadState := NotepadSetting == "True" ? 1 : 0

BuddyGui["ShowNotesButton"].Value := InitialNotepadState

ShowNotepad(isExpanded) {

    Global NotePadEmbedded
    Global ToolsTab

    if isExpanded {
        BuddyGui["NotePadEmbedded"].Visible := 1
        BuddyGui.Show("h815")
        ToolsTab.Move(,, , 600)
        ToolsTab.Choose(1)
    }
    else {
        BuddyGui["NotePadEmbedded"].Visible := 0
        BuddyGui.Show("h295 w630")
        ToolsTab.Move(,, , 80)
    }
}

ShowNotes(ctrl, *) {
    if ctrl.Value {
        IniWrite("True", "config.ini", "Settings", "ShowNotepad")
    } else {
        IniWrite("False", "config.ini", "Settings", "ShowNotepad")
    }
    ShowNotepad(ctrl.Value)
}

ShowNotepad(initialNotepadState)

DarkmodeSetting := IniRead("config.ini", "Settings", "Darkmode")
initialDarkmodeState := DarkmodeSetting == "True" ? 1 : 0

BuddyGui["DarkModeButton"].Value := initialDarkmodeState

; Function to apply darkmode settings
ApplyDarkmode(isDark) {
    if isDark {
        BuddyGui.SetFont("cFFFFFF")
        BuddyGui.BackColor := "004069"
        NotesGui.BackColor := "004069"
        TemplatesGui.BackColor := "004069"
        HotkeysGui.BackColor := "004069"
        TemplateTab.Opt("xm h80 w610 Background00072D cFFFFFF")
        ToolsTab.Opt("WP h80 Background00072D cFFFFFF vToolsTab")
        AlwaysOnTopCheckBoxText.Opt("+BackgroundTrans cFFFFFF")
        ShowNotesButtonText.Opt("cFFFFFF")
        DarkmodeButtonText.Opt("cFFFFFF")
        NotePadEmbedded.Opt("Backgroundc0f0f13 cFFFFFF")
        Notes.Opt("Backgroundc000000 cFFFFFF")
        Notes.SetFont("s10", "Nunito")
        Templates.Opt("Backgroundc0f0f13 cFFFFFF")
        Templates.SetFont("s10", "Nunito")
        Hotkeys.Opt("cFFFFFF")
        Hotkeys.SetFont("s10", "Nunito")
    } else {
        BuddyGui.BackColor := "c007ba8"
        BuddyGui.Color := "c000000"
        NotesGui.BackColor := "c007ba8"
        TemplatesGui.BackColor := "c007ba8"
        HotkeysGui.BackColor := "c007ba8"
        TemplateTab.Opt("xm h80 w610 BackgroundWhite c000000")
        ToolsTab.Opt("WP h80 BackgroundWhite vToolsTab c000000")
        AlwaysOnTopCheckBoxText.Opt("c000000")
        ShowNotesButtonText.Opt("c000000")
        DarkmodeButtonText.Opt("c000000")
        NotePadEmbedded.Opt("BackgroundcFFFFFF c000000")
        Notes.Opt("BackgroundcFFFFFF c000000")
        Notes.SetFont("s10", "Nunito")
        Templates.Opt("BackgroundcFFFFFF c000000")
        Templates.SetFont("s10", "Nunito")
        Hotkeys.Opt("cFFFFFF")
        Hotkeys.SetFont("s10", "Nunito")
    }
}

Darkmode(ctrl, *) {
    if ctrl.Value {
        IniWrite("True", "config.ini", "Settings", "Darkmode")
    } else {
        IniWrite("False", "config.ini", "Settings", "Darkmode")
    }
    ApplyDarkmode(ctrl.Value)
}

ApplyDarkmode(initialDarkmodeState)