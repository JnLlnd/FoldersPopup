InitLanguageVariables:

global lAboutText1 := "~1~ ~2~ (~3~ bits)"
global lAboutText2 := "~1~ is written by Jean Lalonde using the`n<a href=""http://ahkscript.org/"">AutoHotkey</a> programming language.`n`nGerman translation: Edgar ""Fast Edi"" Hoffmann`nFrench translation: Jean Lalonde`nOther languages translation: (help welcomed)`nEnglish proof checking: (help welcomed)`n`nIcons by: <a href=""http://www.visualpharm.com"">Visual Pharm</a>`nAutoHotkey_L v1.1 sources: <a href=""https://github.com/JnLlnd/FoldersPopup"">GitHub</a>"
global lAboutText3 := "~1~ Jean Lalonde 2013-2014. Freeware."
global lAboutText4 := "Support on <a href=""http://code.jeanlalonde.ca/folderspopup"">www.code.jeanlalonde.ca</a>"
global lAboutTitle := "About - ~1~ ~2~"
global lAppTagline := "Jump instantly from one folder to another"
global lButtonRemind := "Remind me"
global lDiagModeCaution = "~1~ is running in diagnostic mode.`n`nInformation about the app's execution will be collected in the file:`n~2~`n`nNothing will be sent without your consent.`n`nDo you want to keep diagnostic mode ON?"
global lDiagModeExit = "~1~ colleted diagnostic information in the file ~2~."
global lDiagModeIntro = "Send this file to ahk@jeanlalonde.ca with a description of the situation requiring diagnostic."
global lDiagModeSee = "Do you want to see the diagnostic file?"
global lDialogAdd := "Add"
global lDialogAddDialogAlready := "This dialog box type is already supported."
global lDialogAddDialogPrompt := "Enter the new dialog box name`n(or part of the name):"
global lDialogAddDialogTitle := "Add Dialog Box - ~1~ ~2~"
global lDialogAddEditFolderTitle = "~1~ folder - ~2~ ~3~"
global lDialogAddFolderManuallyPrompt := "Sorry, we can't detect the current folder in this type of window.`n`nDo you want to add it manually now?"
global lDialogAddFolderManuallyTitle := "Add This Folder - ~1~ ~2~"
global lDialogAddFolderSelect := "Choose or create the new folder:"
global lDialogBrowseButton := "Browse"
global lDialogCancelButton := "Cancel"
global lDialogCancelPrompt := "Discard changes?"
global lDialogCancelTitle := "Cancel - ~1~ ~2~"
global lDialogEditDialogPrompt := "Enter the new dialog box name`n(or part of the name):"
global lDialogEditDialogTitle := "Edit Dialog box - ~1~ ~2~"
global lDialogEditFolderPrompt := "Enter the name for this foler:"
global lDialogEditFolderSelect := "Choose or create the new folder:"
global lDialogEditFolderTitle := "Edit Folder - ~1~ ~2~"
global lDialogFolderLabel := "Folder"
global lDialogFolderLocationEmpty := "The location is empty. Please, choose a location."
global lDialogFolderNameEmpty := "The folder name is empty. Please, choose a short name."
global lDialogFolderNameNoSeparator := "The submenu name cannot include the reserved character ""~1~""."
global lDialogFolderNameNotNew := "The name ""~1~"" is already used in this menu. Please, choose a new name."
global lDialogFolderParentMenu := "Folder parent menu"
global lDialogFolderRemovePrompt := "Remove the folder ""~1~""`nand ALL its content (folders and submenus)?"
global lDialogFolderRemoveTitle := "Remove Folder - ~1~"
global lDialogFolderShortName := "Folder short name for menu"
global lDialogInvalidHotkey := "With your current system keyboard layout, the hotkey ""~1~"" could not be used as a trigger for the popup menu (not a valid key namse error).`n`nPlease, open the ""~2~ Settings"" window from the Tray menu and click ""Options"". In this dialog box, choose another shortcut for ""~3~""."
global lDialogSave := "Save"
global lDialogSelectItemToEdit := "Please, select the item to edit."
global lDialogSelectItemToRemove := "Please, select the item to remove."
global lDialogSubmenuLabel := "Submenu"
global lDialogSubmenuShortName := "Submenu short name"
global lDonateButton := "Support Freeware!"
global lDonatePrompt := "Are you happy with ~1~?`n`n~1~ is not only FREE of charge but also FREE of nasty advertising or adware that you never know if they carry spyware or malware.`n`n~2~ times so far, ~1~ has helped you in your daily work. How about supporting freeware now?"
global lDonateThankyou := "Thank you for being fair and support freeware!"
global lDonateTitle := "~1~ times! - ~2~"
global lGui2Close := "Close"
global lGuiAbout := "About"
global lGuiAddDialog := "Add"
global lGuiAddFolder := "Add"
global lGuiCancel := "&Cancel"
global lGuiClose := "&Close"
global lGuiEditFolder := "Edit"
global lGuiHelp := "Help"
global lGuiLvDialogsHeader := "Supported Dialog Boxes`n(part of the names is enough)" 
global lGuiLvFoldersHeader := "Name|Path"
global lGuiMoveFolderDown := "Move D&own"
global lGuiMoveFolderUp := "Move &Up"
global lGuiOptions := "Options"
global lGuiRemoveFolder := "Remove"
global lGuiSave := "&Save"
global lGuiSeparator := "Se&parator"
global lGuiSubmenuDropdownLabel := "Menu to edit:"
global lGuiSubmenuLocation := ">"
global lGuiSubmenuSeparator := ">"
global lGuiTitle := "Settings - ~1~ ~2~"
global lHelpText1 := "At its launch, FoldersPopup add an icon in the System Tray and await your orders. When you need to change folder in Windows Explorer or in a file dialog box, just click the Middle mouse button or press Windows-K and, in the popup menu, select the desired folder. FoldersPopup will take you there this instantly!"
global lHelpText2 := "Choose ""Settings"" to open the FoldersPopup settings window where you can add folders to your menu, delete, move or rename them. You can add sub-menus and delete, move or rename them as well. Choose the ""Menu to edit:"" in the  dropdown list. Click ""Save"" to save your changes."
global lHelpText3 := "You can quickly add new folders to the popup menu:`n1) Go to a frequently used folder.`n2) Click the Middle mouse button (or press Windows-K) and choose ""Add This Folder"".`n3) Give the folder a short name, click ""OK"" and ""Save"" in the settings window.`nThis new folder is saved in the main menu. Edit the folder to move it to another sub-menu."
global lHelpText4 := "By default, FoldersPopup supports regular dialog boxes (Open, Save, etc.). You can easily teach FoldersPopup to recognize other types of dialog box:`n1) When you are in an unsupported dialog box, click the Middle mouse button (or press Windows-K) and choose ""Add this dialog to the supported list"".`n2 ) Click ""OK"" and ""Save"" in the settings window."
global lHelpText5 := "TIPS AND TRICKS`n- In the Tray menu, check the ""Run at Startup"" option to launch FoldersPopup automatically at startup.`n- Need to open a favorite folder in a New Explorer window? Just press Shift-Middle mouse button (or Shift-Windows-K). This also allow you to open a favorite folder from any application.`n- Configure your popup menu triggers: choose your preferred mouse buttons or keyboard shortcuts in ""Settings, Options"".`n- Also in ""Options"": choose your preferred language, select to display recent folders, add numeric keyboard shortcuts to the folders menu or pin the popup menu at a fix position."
global lHelpText6 := "Support on FoldersPopup is available at <a href=""http://www.code.jeanlalonde.ca/folderspopup"">www.code.jeanlalonde.ca/folderspopup</a>."
global lHelpText7 := "~1~ Jean Lalonde 2013-2014. Freeware."
global lHelpTextLead := "FoldersPopup lets you move like a breeze between your frequently used folders!"
global lHelpTitle := "Help - ~1~ ~2~"
global lMainMenuName := "Main"
global lMenuAbout := "A&bout"
global lMenuAddThisDialog := "&Add This Dialog Box to the supported list"
global lMenuAddThisFolder := "&Add This Folder"
global lMenuControlPanel := "&Control Panel"
global lMenuDesktop := "&Desktop"
global lMenuDialogNotSupported := "This dialog box type is not supported yet"
global lMenuDocuments := "D&ocuments"
global lMenuHelp := "Help"
global lMenuMyComputer := "&My Computer"
global lMenuNetworkNeighborhood := "&Network Neighborhood"
global lMenuPictures := "&Pictures"
global lMenuRecentFolders := "&Recent folders"
global lMenuRecycleBin := "&Recycle Bin"
global lMenuReservedShortcuts := "FP"
global lMenuReservedShortcutsRecent := "R"
global lMenuReservedShortcutsSpecial := "S"
global lMenuReservedShortcutsSwitch := "DE"
global lMenuRunAtStartup := "&Run at Startup"
global lMenuSeparator := "----------------"
global lMenuSettings := "&~1~ Settings"
global lMenuSpecialFolders := "&Special Folders"
global lMenuSwitchDialog := "Switch in &dialog box"
global lMenuSwitchExplorer := "Switch to &Explorer"
global lMenuUpdate := "Check for &update"
global lNavigateFileError := "An error occured while opening the folder:`n~1~."
global lNavigateSpecialError := "An error occured while opening the special folder #~1~."
global lOptionsAlt := "Alt"
global lOptionsArrDescriptions1 := "Choose the MOUSE button and modifiers combination that will open the folders popup menu in Windows Explorer or file dialog boxes. By default, this is the middle mouse button (MButton) without any modifier key."
global lOptionsArrDescriptions2 := "Choose the MOUSE button and modifiers combination that will open the folders popup menu in ANY window and navigate to the selected folder in a NEW Windows Explorer window. By default, this is the middle mouse button (MButton) while the Shift key is pressed."
global lOptionsArrDescriptions3 := "Choose the KEYBOARD hotkey combination that will open the folders popup menu in the active Windows Explorer or file dialog box. By default, this is Win+K (the ""K"" letter while the Windows key is pressed)."
global lOptionsArrDescriptions4 := "Choose the KEYBOARD hotkey combination that will open the folders popup menu in ANY window and navigate to the selected folder in a NEW Windows Explorer window. By default, this is Shift+Win+K (the ""K"" letter while the Shift and Windows keys are pressed)."
global lOptionsArrDescriptions5 := "Choose the hotkey or mouse button combination that will open the FoldersPopup setting dialog box. By default, this is Shift+Windows+F."
global lOptionsChangeHotkey := "Change"
global lOptionsChangeHotkeyTitle := "Change hotkey - ~1~ ~2~"
global lOptionsCtrl := "Ctrl"
global lOptionsDisplayMenuShortcuts := "Display &Numeric Menu Shortcuts"
global lOptionsDisplayRecentFolders := "Display Recent Folders"
global lOptionsDisplaySpecialFolders := "Display S&pecial Folders"
global lOptionsDisplaySwitchMenu := "Display Switch Menu"
global lOptionsGuiIntro := "Define the mouse buttons and keyboard hotkeys that will trigger ~1~."
global lOptionsGuiTitle := "Options - ~1~ ~2~"
global lOptionsKeyboard := "Keyboard"
global lOptionsLanguage := "Language"
global lOptionsLanguageLabels := "English|French|German"
global lOptionsMouse := "Mouse"
global lOptionsNoKeyOrMouseSpecified := "No Key or Mouse Button specified. The existing ""~1~"" hotkey is unchanged."
global lOptionsOtherOptions := "Other Options"
global lOptionsPopupFix := "Popup &Menu at Fix Position"
global lOptionsPopupFixPositionX := "Position X:"
global lOptionsPopupFixPositionY := "Y:"
global lOptionsRunAtStartup := "&Run at Startup"
global lOptionsShift := "Shift"
global lOptionsTitles := "Mouse - Folders Popup|Mouse - New Explorer|Keyboard - Folders Popup|Keyboard - New Explorer|Settings window"
global lOptionsTrayTip := "&Display Startup Tray Tip"
global lOptionsTriggerFor := "Trigger for:"
global lOptionsWin := "Win"
global lReloadPrompt := "Language changed to ~1~. Do you want to reload ~2~ in ~1~ now? Unsaved changes to the folders menu will be lost."
global lTrayTipInstalledDetail := "To show ~1~ menu, press:`n`n~2~ / ~3~`nin Windows Explorer or in a Dialog box`n`n~4~ / ~5~`nto open a New Explorer window"
global lTrayTipinstalledTitle := "~1~ ~2~ ready!"
global lUpdatePrompt := "Update ~1~ from v~2~ to v~3~?"
global lUpdateTitle := "Update ~1~?"
global lUpdateYouHaveLatest := "You have the latest version: ~1~.`n`nVisit the ~2~ web page anyway?"

return
