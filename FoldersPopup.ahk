;===============================================
/*
	FoldersPopup
	Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
	By Jean Lalonde (JnLlnd on AHKScript.org forum), based on DirMenu v2 by Robert Ryan (rbrtryn on AutoHotkey.com forum)

	Version: FoldersPopup v0.9 BETA
	- implemented startup option in tray and check4update
	- removed debugging code, prepare for compiler, removed external pictures
	- standardize dialog box titles, various text fixes
	- renamed the app FoldersPopup

	Version: PopupFolders v0.5 ALPHA (last alpha version)
	- implemented GuiAbout and GuiHelp, added About and Help to tray menu, tray tip displayed only 5 times
	- removed file:/// protocol prefix, added support for ExploreWClass, implemented try/catch to Explore shell method, offer to add manually when add folder failed

	Version: PopupFolders v0.4 ALPHA
	- add settings hotkey to ini file (default Crtl-Windows-F), enable AddThisFolder in all version Explorer and only in WIN_7/Win_8 dialog boxes (not working in WIN_XP)
	- add GuiSave, GuiCancel, RemoveFolder, EditFolder, AddSeparator, MoveFolderUp/Down, RemoveDialog, EditDialog, fix bug in GuiShow, add tray icon

	Version: PopupFolders v0.3 ALPHA
	- add NavigateConsole for console support (command prompt CMD)
	- change .ini filename to new app name

	Version: PopupFolders v0.2 ALPHA
	- renamed app PopupFolders, isolate text into language variables

	Version: DirMenu3 v0.1 ALPHA
	- init skeleton, read ini file and create arrays for folders menu and supported dialog boxes
	- create language file, build gui, tray menu and folder menu, skeleton for front end buttons and commands
	- create AddThisDialog menu, MButton condition, CanOpenFavorite improvements with WindowIsAnExplorer, WindowIsDesktop and DialogIsSupported
	- add SpecialFolders menu, OpenFavorite for Explorer and Desktop, NavigateExplorer
	- support MS Office dialog boxes on WinXP (bosa_sdm_), open special folders in explorers
	- NavigateDialog, add Desktop, Document and Pictures special folders, open these special menus in dialog boxes, enabling/disabling the appropriate menus in dialog boxes or explorers

	Version: DirMenu v2.2 (never released / not stable - base of a total rewrite to DirMenu3)
	- manage (add, modify or delete) supported dialog box titles in the Gui
	- suggest current dialog box when adding a name
	- save the supported dialog box names on the first line of the settings file (dirmenu.txt)
	- add "Add This Dialog" to the MButton menu to add the current dialog box name (need to desactivate when in an already supported dialog box)
	- added Win8 to the list of supported versions (assumed as equal to Win7 - could not test myself)
	- removed the "Menu File" button because not needed anymore
	- fixed an issue when 2 folders had the same name (now preventing the use of an existing name)
	- change default setting filename to "DirMenu2.txt" to avoid upgrade errors
	- upgrade previous versions settings files to v2.2
	- ask confirmation before discarding changes with Revert or Cancel buttons
	- replaces RegEx on strDialogNames with DialogIsSupported() function on the ListView
	
	Version: DirMenu v2.1
	- make it work with any locale (still working with English)
	- put supported dialog box titles in a variable (strDialogNames) at the top of the script for easy editing
	- put DirMenu data file name in a variable (strDirMenuFile) at the top of the script for easy editing
	- add "Add This Folder" to the MButton menu to add the current folder
	- add "Menu File" button to open de DirMenu.txt file for edition in Notepad
	- propose the deepest folder name as default name for a new folder

*/ 
;===============================================

; --- COMPILER DIRECTIVES ---

; Doc: http://fincs.ahk4.net/Ahk2ExeDirectives.htm
; Note: prefix comma with `

;@Ahk2Exe-SetName FoldersPopup
;@Ahk2Exe-SetDescription Popup menu to jump instantly from one folder to another. Freeware.
;@Ahk2Exe-SetVersion 0.9
;@Ahk2Exe-SetOrigFilename FoldersPopup.exe


;============================================================
; INITIALIZATION
;============================================================

#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off

global strCurrentVersion := "1.0 beta"
#Include %A_ScriptDir%\FoldersPopup_LANG.ahk
SetWorkingDir, %A_ScriptDir%

global strIniFile := A_ScriptDir . "\" . lAppName . ".ini"

strMouseButtons := " |LButton|MButton|RButton|XButton1|XButton2|WheelUp|WheelDown|WheelLeft|WheelRight|"
; leave last | to enable default value on the last item

;@Ahk2Exe-IgnoreBegin
; Piece of code for developement phase only - won't be compiled
if (A_ComputerName = "JEAN-PC") ; for my home PC
	strIniFile := A_ScriptDir . "\" . lAppName . "-HOME.ini"
else if InStr(A_ComputerName, "STIC") ; for my work hotkeys
	strIniFile := A_ScriptDir . "\" . lAppName . "-WORK.ini"
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd

Gosub, BuildSpecialFoldersMenu
Gosub, LoadIniFile
Gosub, BuildFoldersMenu
Gosub, BuildGUI
Gosub, BuildAddDialogMenu
Gosub, Check4Update
Gosub, BuildTrayMenu

IfExist, %A_Startup%\%lAppName%.lnk
{
	FileDelete, %A_Startup%\%lAppName%.lnk
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%lAppName%.lnk
	Menu, Tray, Check, %lMenuRunAtStartup%
}

if (intDisplayTrayTip <> 0)
	TrayTip, % L(lTrayTipInstalledTitle, lAppName, lAppVersion)
		, % L(lTrayTipInstalledDetail, lAppName), , 1
	; ### adapt to current hotkeys
return



;============================================================
; BACK END FUNCTIONS AND COMMANDS
;============================================================


;-----------------------------------------------------------
LoadIniFile:
;-----------------------------------------------------------

global arrGlobalFolders := Object()
global arrGlogalDialogs := Object()

strPopupHotkeyMouseDefault := "MButton"
strPopupHotkeyMouseNewDefault := "+MButton"
strPopupHotkeyKeyboardDefault := "#k"
strPopupHotkeyKeyboardNewDefault := "+#k"
strSettingsHotkeyDefault := "+#f"

IfNotExist, %strIniFile%
	FileAppend,
		(LTrim Join`r`n
			[Global]
			PopupHotkeyMouse=%strPopupHotkeyMouseDefault%
			PopupHotkeyNewMouse=%strPopupHotkeyMouseNewDefault%
			PopupHotkeyKeyboard=%strPopupHotkeyKeyboardDefault%
			PopupHotkeyNewKeyboard=%strPopupHotkeyKeyboardNewDefault%
			SettingsHotkey=%strSettingsHotkeyDefault%
			DisplayTrayTip=5
			[Folders]
			Folder1=C:\|C:\
			Folder2=Windows|%A_WinDir%
			Folder3=Program Files|%A_ProgramFiles%
			[Dialogs]
			Dialog1=Export
			Dialog2=Import
			Dialog3=Insert
			Dialog4=Open
			Dialog5=Save
			Dialog6=Select
			Dialog7=Upload

)
		, %strIniFile%
	
; Hotkeys: ini names, hotkey variables name, default values and gosub label
strIniKeyNames := "PopupHotkeyMouse|PopupHotkeyNewMouse|PopupHotkeyKeyboard|PopupHotkeyNewKeyboard|SettingsHotkey"
StringSplit, arrIniKeyNames, strIniKeyNames, |
strHotkeyVarNames := "strPopupHotkeyMouse|strPopupHotkeyMouseNew|strPopupHotkeyKeyboard|strPopupHotkeyKeyboardNew|strSettingsHotkey"
StringSplit, arrHotkeyVarNames, strHotkeyVarNames, |
strHotkeyDefaults := "MButton|+MButton|#k|+#k|+#f"
StringSplit, arrHotkeyDefaults, strHotkeyDefaults, |
strHotkeyLabels := "PopupMenuMouse|PopupMenuNewWindowMouse|PopupMenuKeyboard|PopupMenuNewWindowKeyboard|GuiShow"
StringSplit, arrHotkeyLabels, strHotkeyLabels, |

Gosub, LoadIniHotkeys

IniRead, intDisplayTrayTip, %strIniFile%, Global, DisplayTrayTip
if (intDisplayTrayTip > 0)
	IniWrite, % (intDisplayTrayTip - 1), %strIniFile%, Global, DisplayTrayTip

IniRead, strLatestSkipped, %strIniFile%, Global, LatestVersionSkipped, 0.0

Loop
{
	IniRead, strIniLine, %strIniFile%, Folders, Folder%A_Index%
	if (strIniLine = "ERROR")
		Break
	StringSplit, arrThisObject, strIniLine, |
	objFolder := Object()
	objFolder.Name := arrThisObject1
	objFolder.Path := arrThisObject2
	arrGlobalFolders.Insert(objFolder)
}
Loop
{
	IniRead, strIniLine, %strIniFile%, Dialogs, Dialog%A_Index%
	if (strIniLine = "ERROR")
		Break
	arrGlogalDialogs.Insert(strIniLine)
}

return
;------------------------------------------------------------


;-----------------------------------------------------------
LoadIniHotkeys:
;-----------------------------------------------------------
; Read the values and set hotkey shortcuts
loop, % arrIniKeyNames%0%
{
	IniRead, arrHotkeyVarNames%A_Index%, %strIniFile%, Global, % arrIniKeyNames%A_Index%, % arrHotkeyDefaults%A_Index%
	; example: Hotkey, $MButton, PopupMenuMouse
	Hotkey, % "$" . arrHotkeyVarNames%A_Index%, % arrHotkeyLabels%A_Index%, On
	; Prepare global arrays used by GuiHotkey function
	SplitHotkey(arrHotkeyVarNames%A_Index%, strMouseButtons
		, strModifiers%A_Index%, strOptionsKey%A_Index%, strMouseButton%A_Index%, strMouseButtonsWithDefault%A_Index%)
}

return
;------------------------------------------------------------



;============================================================
; FRONT END FUNCTIONS AND COMMANDS
;============================================================


;------------------------------------------------------------
PopupMenuMouse: ; default MButton
PopupMenuKeyboard: ; default #k
;------------------------------------------------------------

If !CanOpenFavorite(A_ThisLabel, strGlobalWinId, strGlobalClass)
{
	if (A_ThisLabel = "PopupMenuMouse")
		strThisHotkey := A_ThisHotkey
	else
		StringReplace, strThisHotkey, A_ThisHotkey, $ ; remove $ from hotkey
	Send, {%strThisHotkey%} ; for example {MButton} or #k
	###_T(A_ThisLabel, strThisHotkey)
	return
}

if (A_ThisLabel = "PopupMenuMouse")
{
	WinActivate, % "ahk_id " . strGlobalWinId
	; display menu at mouse pointer location
	intMenuPosX :=
	intMenuPosY :=
}
else
{
	; display menu at an offset of 20x20 pixel from top-left client area
	intMenuPosX := 20
	intMenuPosY := 20
}
###_T(A_ThisLabel , "strGlobalWinId: " . strGlobalWinId . "`nstrGlobalClass: " . strGlobalClass)

; Can't find how to navigate a dialog box to My Computer or Network Neighborhood... need help ???
Menu, menuSpecialFolders
	, % WindowIsConsole(strGlobalClass) or WindowIsDialog(strGlobalClass) ? "Disable" : "Enable"
	, %lMenuMyComputer%
Menu, menuSpecialFolders
	, % WindowIsConsole(strGlobalClass) or WindowIsDialog(strGlobalClass) ? "Disable" : "Enable"
	, %lMenuNetworkNeighborhood%

; There is no point to navigate a dialog box or console to Control Panel or Recycle Bin
Menu, menuSpecialFolders
	, % WindowIsConsole(strGlobalClass) or WindowIsDialog(strGlobalClass) ? "Disable" : "Enable"
	, %lMenuControlPanel%
Menu, menuSpecialFolders
	, % WindowIsConsole(strGlobalClass) or WindowIsDialog(strGlobalClass) ? "Disable" : "Enable"
	, %lMenuRecycleBin%

; ONLY IF MOUSE HOTKEY -> moved to PopupMenuMouse
; WinActivate, % "ahk_id " . strGlobalWinId

if (WindowIsAnExplorer(strGlobalClass) or WindowIsDesktop(strGlobalClass) or WindowIsConsole(strGlobalClass) or DialogIsSupported(strGlobalWinId))
{
	; Enable Add This Folder only if the mouse is over an Explorer (tested on WIN_XP and WIN_7) or a dialog box (works on WIN_7, not on WIN_XP)
	; Other tests shown that WIN_8 behaves like WIN_7. So, I assume WIN_8 to work. If someone could confirm (until I can test it myself)?
	Menu, menuFolders
		, % WindowIsAnExplorer(strGlobalClass) or (WindowIsDialog(strGlobalClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
		, %lMenuAddThisFolder%
	Menu, menuFolders, Show, %intMenuPosX%, %intMenuPosY% ; mouse pointer if mouse button, 20x20 offset of active window if keyboard shortcut

}
else
	Menu, menuAddDialog, Show

return
;------------------------------------------------------------


;------------------------------------------------------------
PopupMenuNewWindowMouse: ; default +MButton::
PopupMenuNewWindowKeyboard: ; default +#k
;------------------------------------------------------------
if (A_ThisLabel = "PopupMenuNewWindowMouse")
{
	MouseGetPos, , , strGlobalWinId
	; display menu at mouse pointer location
	intMenuPosX :=
	intMenuPosY :=
}
else
{
	strGlobalWinId := WinExist("A")
	; display menu at an offset of 20x20 pixel from top-left client area
	intMenuPosX := 20
	intMenuPosY := 20
}
WinGetClass strGlobalClass, % "ahk_id " . strGlobalWinId
###_T(A_ThisLabel , "strGlobalWinId: " . strGlobalWinId . "`nstrGlobalClass: " . strGlobalClass)

; In case it was disabled while in a dialog box
Menu, menuSpecialFolders, Enable, %lMenuMyComputer%
Menu, menuSpecialFolders, Enable, %lMenuNetworkNeighborhood%
Menu, menuSpecialFolders, Enable, %lMenuControlPanel%
Menu, menuSpecialFolders, Enable, %lMenuRecycleBin%

; Enable "Add This Folder" only if the target window is an Explorer (tested on WIN_XP and WIN_7)
; or a dialog box under WIN_7 (does not work under WIN_XP).
; Other tests shown that WIN_8 behaves like WIN_7.
Menu, menuFolders
	, % WindowIsAnExplorer(strGlobalClass) or (WindowIsDialog(strGlobalClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
	, %lMenuAddThisFolder%
Menu, menuFolders, Show, %intMenuPosX%, %intMenuPosY% ; mouse pointer if mouse button, 20x20 offset of active window if keyboard shortcut

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildTrayMenu:
;------------------------------------------------------------

;@Ahk2Exe-IgnoreBegin
; Piece of code for developement phase only - won't be compiled
Menu, Tray, Icon, %A_ScriptDir%\Folders-Likes-icon-256-light-center.ico, 1
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd
Menu, Tray, Add
Menu, Tray, Add, %lMenuSettings%, GuiShow
Menu, Tray, Add
Menu, Tray, Add, %lMenuRunAtStartup%, RunAtStartup
Menu, Tray, Add
Menu, Tray, Add, %lMenuUpdate%, Check4Update
Menu, Tray, Add, %lMenuHelp%, GuiHelp
Menu, Tray, Add, %lMenuAbout%, GuiAbout
Menu, Tray, Default, %lMenuSettings%

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildSpecialFoldersMenu:
;------------------------------------------------------------

Menu, menuSpecialFolders, Add, %lMenuDesktop%, OpenSpecialFolder
Menu, menuSpecialFolders, Add, %lMenuDocuments%, OpenSpecialFolder
Menu, menuSpecialFolders, Add, %lMenuPictures%, OpenSpecialFolder
Menu, menuSpecialFolders, Add
Menu, menuSpecialFolders, Add, %lMenuMyComputer%, OpenSpecialFolder
Menu, menuSpecialFolders, Add, %lMenuNetworkNeighborhood%, OpenSpecialFolder
Menu, menuSpecialFolders, Add
Menu, menuSpecialFolders, Add, %lMenuControlPanel%, OpenSpecialFolder
Menu, menuSpecialFolders, Add, %lMenuRecycleBin%, OpenSpecialFolder

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildFoldersMenu:
;------------------------------------------------------------

Menu, menuFolders, Add
Menu, menuFolders, DeleteAll
Loop, % arrGlobalFolders.MaxIndex()
{
	if (arrGlobalFolders[A_Index].Name = lMenuSeparator)
		Menu, menuFolders, Add
	else
		Menu, menuFolders, Add, % arrGlobalFolders[A_Index].Name, OpenFavorite
}
Menu, menuFolders, Add
Menu, menuFolders, Add, %lMenuSpecialFolders%, :menuSpecialFolders
Menu, menuFolders, Add
Menu, menuFolders, Add, %lMenuSettings%, GuiShow
Menu, menuFolders, Default, %lMenuSettings%
Menu, menuFolders, Add, %lMenuAddThisFolder%, AddThisFolder

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildAddDialogMenu:
;------------------------------------------------------------

Menu, menuAddDialog, Add, %lMenuDialogNotSupported%, AddThisDialog
Menu, menuAddDialog, Disable, %lMenuDialogNotSupported%
Menu, menuAddDialog, Add, %lMenuAddThisDialog%, AddThisDialog
Menu, menuAddDialog, Add
Menu, menuAddDialog, Add, %lMenuSettings%, GuiShow
Menu, menuAddDialog, Default, %lMenuAddThisDialog%

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildGui:
;------------------------------------------------------------

Gui, 1:New, , % L(lGuiTitle, lAppName, lAppVersion)
Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Text, x10 y10 w490 h25, %lAppName%
Gui, 1:Font, s8 w400, Arial
Gui, 1:Add, Button, y10 x315 w45 h22 gGuiAbout, % L(lGuiAbout)
Gui, 1:Font, s8 w400, Verdana
Gui, 1:Add, Button, x+10 w75 gGuiHelp, %lGuiHelp%
Gui, 1:Add, Text, x10 y30, %lAppTagline%

Gui, 1:Font, s8 w400, Verdana
Gui, 1:Add, ListView, x10 w350 h220 Count32 -Multi NoSortHdr LV0x10 vlvFoldersList, %lGuiLvFoldersHeader%
Gui, 1:Add, Button, x+10 w75 gGuiAddFolder, %lGuiAddFolder%
Gui, 1:Add, Button, w75 gGuiRemoveFolder, %lGuiRemoveFolder%
Gui, 1:Add, Button, w75 gGuiEditFolder, %lGuiEditFolder%
Gui, 1:Add, Button, w75 gGuiAddSeparator, %lGuiSeparator%
Gui, 1:Add, Button, w75 gGuiMoveFolderUp, %lGuiMoveFolderUp%
Gui, 1:Add, Button, w75 gGuiMoveFolderDown, %lGuiMoveFolderDown%

Gui, 1:Add, ListView
	, x10 w350 h120 Count16 -Multi NoSortHdr +0x10 LV0x10 vlvDialogsList, %lGuiLvDialogsHeader%
Gui, 1:Add, Button, x+10 w75 gGuiAddDialog, %lGuiAddDialog%
Gui, 1:Add, Button, w75 gGuiRemoveDialog, %lGuiRemoveDialog%
Gui, 1:Add, Button, w75 gGuiEditDialog, %lGuiEditDialog%

Gui, 1:Add, Button, x100 w75 r1 Disabled Default gGuiSave, %lGuiSave%
Gui, 1:Add, Button, x+40 w75 r1 gGuiCancel, %lGuiCancel%
Gui, 1:Add, Button, x+80 w75 gGuiOptions, %lGuiOptions%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFolder:
;------------------------------------------------------------

FileSelectFolder, strNewPath, *C:\, 3, %lDialogAddFolderSelect%
if (strNewPath = "")
	return
AddFolder(strNewPath)

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveFolder:
;------------------------------------------------------------

GuiControl, Focus, lvFoldersList
Gui, 1:ListView, lvFoldersList
intItemToRemove := LV_GetNext()
if !(intItemToRemove)
{
	Oops(lDialogSelectItemToRemove)
	return
}
LV_Delete(intItemToRemove)
LV_Modify(intItemToRemove, "Select Focus")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")
GuiControl, Enable, %lGuiSave%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditFolder:
;------------------------------------------------------------

Gui, 1:+OwnDialogs
GuiControl, Focus, lvFoldersList

Gui, 1:ListView, lvFoldersList
intRowToEdit := LV_GetNext()
LV_GetText(strCurrentName, intRowToEdit, 1)
if !StrLen(strCurrentName)
{
	Oops(lDialogSelectItemToEdit)
	return
}
LV_GetText(strCurrentPath, intRowToEdit, 2)

FileSelectFolder, strNewPath, *%strCurrentPath%, 3, %lDialogEditFolderSelect%
if (strNewPath = "")
	return

Loop
{
	InputBox strNewName, % L(lDialogEditFolderTitle, lAppName, lAppVersion)
		, %lDialogEditFolderPrompt%, , 250, 120, , , , , %strCurrentName%
	if (ErrorLevel)
		return
} until (strNewName = strCurrentName) or FolderNameIsNew(strNewName)

LV_Modify(intRowToEdit, "Select Focus", strNewName, strNewPath)
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")
GuiControl, Enable, %lGuiSave%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddSeparator:
;------------------------------------------------------------

GuiControl, Focus, lvFoldersList
Gui, 1:ListView, lvFoldersList
LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1, "Select Focus", lMenuSeparator, lMenuSeparator . lMenuSeparator)
GuiControl, Enable, %lGuiSave%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFolderUp:
;------------------------------------------------------------

GuiControl, Focus, lvFoldersList
Gui, 1:ListView, lvFoldersList
intSelectedRow := LV_GetNext()
if (intSelectedRow = 1)
	return

LV_GetText(strThisName, intSelectedRow, 1)
LV_GetText(strThisPath, intSelectedRow, 2)

LV_GetText(PriorName, intSelectedRow - 1, 1)
LV_GetText(PriorPath, intSelectedRow - 1, 2)

LV_Modify(intSelectedRow, "", PriorName, PriorPath)
LV_Modify(intSelectedRow - 1, "Select Focus Vis", strThisName, strThisPath)

GuiControl, Enable, %lGuiSave%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFolderDown:
;------------------------------------------------------------

GuiControl, Focus, lvFoldersList
Gui, 1:ListView, lvFoldersList
intSelectedRow := LV_GetNext()
if (intSelectedRow = LV_GetCount())
	return

LV_GetText(strThisName, intSelectedRow, 1)
LV_GetText(strThisPath, intSelectedRow, 2)

LV_GetText(NextName, intSelectedRow + 1, 1)
LV_GetText(NextPath, intSelectedRow + 1, 2)
	
LV_Modify(intSelectedRow, "", NextName, NextPath)
LV_Modify(intSelectedRow + 1, "Select Focus Vis", strThisName, strThisPath)

GuiControl, Enable, %lGuiSave%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddDialog:
;------------------------------------------------------------

AddDialog("")

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveDialog:
;------------------------------------------------------------

GuiControl, Focus, lvDialogsList
Gui, 1:ListView, lvDialogssList
intItemToRemove := LV_GetNext()
if !(intItemToRemove)
{
	Oops(lDialogSelectItemToRemove)
	return
}
LV_Delete(intItemToRemove)
LV_Modify(intItemToRemove, "Select Focus")
GuiControl, Enable, %lGuiSave%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditDialog:
;------------------------------------------------------------

Gui, 1:+OwnDialogs
GuiControl, Focus, lvDialogsList
Gui, 1:ListView, lvDialogsList
intRowToEdit := LV_GetNext()
LV_GetText(strCurrentDialog, intRowToEdit, 1)

InputBox strNewDialog, % L(lDialogEditDialogTitle, lAppName, lAppVersion)
	, %lDialogEditDialogPrompt%, , 250, 120, , , , , %strCurrentDialog%
if (ErrorLevel) or !StrLen(strNewDialog) or (strNewDialog = strCurrentDialog)
	return

Gui, 1:ListView, lvDialogsList
Loop, % LV_GetCount()
{
	LV_GetText(strThisDialog, A_Index, 1)
	if (strNewDialog = strThisDialog)
	{
		Oops(lDialogAddDialogAlready)
		return
	}
}

LV_Modify(intRowToEdit, "Select Focus", strNewDialog)
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(1, "Sort")
LV_Modify(LV_GetNext(), "Vis")
GuiControl, Enable, %lGuiSave%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiSave:
;------------------------------------------------------------

IniDelete, %strIniFile%, Folders
Gui, 1:ListView, lvFoldersList
Loop % LV_GetCount()
{
	LV_GetText(strName, A_Index, 1)
	LV_GetText(strPath, A_Index, 2)
	IniWrite, %strName%|%strPath%, %strIniFile%, Folders, Folder%A_Index%
}

IniDelete, %strIniFile%, Dialogs
Gui, 1:ListView, lvDialogsList
Loop % LV_GetCount()
{
	LV_GetText(strDialog, A_Index, 1)
	IniWrite, %strDialog%, %strIniFile%, Dialogs, Dialog%A_Index%
}

Gosub, LoadIniFile
Gosub, BuildFoldersMenu
GuiControl, Disable, %lGuiSave%
Gosub, GuiCancel

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiShow:
;------------------------------------------------------------

Gosub, LoadSettingsToGui
Gui, 1:Show, w455 h455

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiCancel:
;------------------------------------------------------------

GuiControlGet, blnSaveEnabled, Enabled, %lGuiSave%
if (blnSaveEnabled)
{
	Gui, 1:+OwnDialogs
	MsgBox, 36, % L(lDialogCancelTitle, lAppName, lAppVersion), %lDialogCancelPrompt%
	IfMsgBox, Yes
		GuiControl, Disable, %lGuiSave%
	IfMsgBox, No
		return
}
Gui, 1:Cancel

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiClose:
;------------------------------------------------------------

GoSub, GuiCancel

return
;------------------------------------------------------------


;------------------------------------------------------------
AddThisFolder:
;------------------------------------------------------------

objPrevClipboard := ClipboardAll ; Save the entire clipboard
ClipBoard := ""

; Add This folder menu is active only if we are in Explorer (WIN_XP, WIN_7 or WIN_8) or in a Dialog box (WIN_7 or WIN_8).
; In all these OS, the key sequence {F4}{Esc} selects the current location of the window.
intWaitTimeIncrement := 150 ; time required on an XP average machine
intTries := 3
Loop, %intTries%
{
	Sleep, intWaitTimeIncrement * A_Index
	Send {F4}{Esc} ; F4 move the caret the "Go To A Different Folder box" and {Esc} select it content ({Esc} could be replaced by ^a to Select All)
	Sleep, intWaitTimeIncrement * A_Index
	Send ^c ; Copy
	Sleep, intWaitTimeIncrement * A_Index
} Until (StrLen(ClipBoard))

strCurrentFolder := ClipBoard

If StrLen(strCurrentFolder)
{
	Gosub, GuiShow
	AddFolder(strCurrentFolder)
}
else
{
	Gui, 1:+OwnDialogs 
	MsgBox, 52, % L(lDialogAddFolderManuallyTitle, lAppName, lAppVersion), %lDialogAddFolderManuallyPrompt%
	IfMsgBox, Yes
	{
		Gosub, GuiShow
		Gosub, GuiAddFolder
	}
}

Clipboard := objPrevClipboard ; Restore the original clipboard
objPrevClipboard := "" ; Free the memory in case the clipboard was very large

return
;------------------------------------------------------------


;------------------------------------------------------------
AddFolder(strPath)
;------------------------------------------------------------
{
	GuiControl, Focus, lvFoldersList
	Gui, 1:+OwnDialogs

	; suggest the deepest folder's name as default name for the added folder
	SplitPath, strPath, strDefaultName, , , , strDrive
	if !StrLen(strDefaultName) ; we are probably at the root of a drive
		strDefaultName := strDrive

	Loop
	{
		InputBox strName, % L(lDialogFolderNameTitle, lAppName, lAppVersion), %lDialogFolderNamePrompt%, , 250, 120, , , , , %strDefaultName%
		if (ErrorLevel) or !StrLen(strName)
			return
	} until FolderNameIsNew(strName)
	
	Gui, 1:ListView, lvFoldersList
	LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1, "Select Focus", strName, strPath)
	LV_Modify(LV_GetNext(), "Vis")
	LV_ModifyCol(1, "AutoHdr")
	LV_ModifyCol(2, "AutoHdr")
	GuiControl, Enable, %lGuiSave%
}
;------------------------------------------------------------


;------------------------------------------------------------
FolderNameIsNew(strCandidateName)
;------------------------------------------------------------
{
	Gui, 1:ListView, lvFoldersList
	Loop, % LV_GetCount()
	{
		LV_GetText(strThisName, A_Index, 1)
		if (strCandidateName = strThisName)
		{
			Oops(lDialogFolderNameNotNew, strCandidateName)
			return False
		}
	}
	return True
}
;------------------------------------------------------------


;------------------------------------------------------------
AddThisDialog:
;------------------------------------------------------------

WinGetTitle, strDialogTitle, ahk_id %strGlobalWinId%
Gosub, GuiShow
AddDialog(strDialogTitle)

return
;------------------------------------------------------------


;------------------------------------------------------------
AddDialog(strCurrentDialogTitle)
;------------------------------------------------------------
{
	Gui, 1:+OwnDialogs
	GuiControl, Focus, lvDialogsList

	InputBox, strNewDialog, % L(lDialogAddDialogTitle, lAppName, lAppVersion), %lDialogAddDialogPrompt%, , 250, 150, , , , , %strCurrentDialogTitle%
	if (ErrorLevel) or !StrLen(strNewDialog)
		return
	
	Gui, 1:ListView, lvDialogsList
	Loop, % LV_GetCount()
	{
		LV_GetText(strThisDialog, A_Index, 1)
		if (strNewDialog = strThisDialog)
		{
			Oops(lDialogAddDialogAlready)
			return
		}
	}

	LV_Add("Select Focus", strNewDialog)
	LV_Modify(LV_GetNext(), "Vis")
	LV_ModifyCol(1, "AutoHdr")
	GuiControl, Enable, %lGuiSave%
}
;------------------------------------------------------------


;------------------------------------------------------------
RunAtStartup:
;------------------------------------------------------------
; Startup code adapted from Avi Aryan Ryan in Clipjump

Menu, Tray, Togglecheck, %lMenuRunAtStartup%
IfExist, %A_Startup%\%lAppName%.lnk
	FileDelete, %A_Startup%\%lAppName%.lnk
else
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%lAppName%.lnk

return
;------------------------------------------------------------


;------------------------------------------------------------
Check4Update:
;------------------------------------------------------------

strLatestVersion := Url2Var("https://raw.github.com/JnLlnd/FoldersPopup/master/latest-version.txt")

if RegExMatch(strCurrentVersion, "(alpha|beta)")
	or (FirstVsSecondIs(strLatestSkipped, strLatestVersion) >= 0 and (A_ThisMenuItem <> lMenuUpdate))
	return

if FirstVsSecondIs(strLatestVersion, strCurrentVersion) = 1
{
	Gui, 1:+OwnDialogs
	SetTimer, ChangeButtonNames, 50

	MsgBox, 3, % l(lUpdateTitle, lAppName), % l(lUpdatePrompt, lAppName, strCurrentVersion, strLatestVersion), 30
	IfMsgBox, Yes
		Run, http://code.jeanlalonde.ca/folderspopup/
	IfMsgBox, No
		IniWrite, %strLatestVersion%, %strIniFile%, Global, LatestVersionSkipped
	IfMsgBox, Cancel ; Remind me
		IniWrite, 0.0, %strIniFile%, Global, LatestVersionSkipped
	IfMsgBox, TIMEOUT ; Remind me
		IniWrite, 0.0, %strIniFile%, Global, LatestVersionSkipped
}
else if (A_ThisMenuItem = lMenuUpdate)
{
	MsgBox, 4, % l(lUpdateTitle, lAppName), % l(lUpdateYouHaveLatest, lAppVersion, lAppName), 30
	IfMsgBox, Yes
		Run, http://code.jeanlalonde.ca/folderspopup/
}

return 
;------------------------------------------------------------


;------------------------------------------------------------
FirstVsSecondIs(strFirstVersion, strSecondVersion)
;------------------------------------------------------------
{
	StringSplit, arrFirstVersion, strFirstVersion, `.
	StringSplit, arrSecondVersion, strSecondVersion, `.
	if (arrFirstVersion0 > arrSecondVersion0)
		intLoop := arrFirstVersion0
	else
		intLoop := arrSecondVersion0

	Loop %intLoop%
		if (arrFirstVersion%A_index% > arrSecondVersion%A_index%)
			return 1 ; greater
		else if (arrFirstVersion%A_index% < arrSecondVersion%A_index%)
			return -1 ; smaller
		
	return 0 ; equal
}
;------------------------------------------------------------

/*
####


IsLatestRelease(prog_ver, cur_ver, exclude_keys="beta|alpha")
{

	if RegExMatch(prog_ver, "(" exclude_keys ")")
		return 1

	StringSplit, prog_ver_array, prog_ver,`.
	StringSplit, cur_ver_array, cur_ver  ,`.

	Loop % cur_ver_array0
		if !( prog_ver_array%A_index% >= cur_ver_array%A_index% )
			return 0
	return 1
}
*/


;------------------------------------------------------------
ChangeButtonNames: 
;------------------------------------------------------------

IfWinNotExist, % l(lUpdateTitle, lAppName)
    return  ; Keep waiting.
SetTimer, ChangeButtonNames, Off 
WinActivate 
ControlSetText, Button3, %lButtonRemind%

return
;------------------------------------------------------------


;------------------------------------------------------------
Url2Var(strUrl)
;------------------------------------------------------------
{
	ComObjError(False)
	objWebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	objWebRequest.Open("GET", strUrl)
	objWebRequest.Send()

	Return objWebRequest.ResponseText()
}
;------------------------------------------------------------



;------------------------------------------------------------
; GUI2 About, Help and Options
;------------------------------------------------------------

;------------------------------------------------------------
GuiAbout:
;------------------------------------------------------------

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lAboutTitle, lAppName, lAppVersion)
Gui, 2:+Owner1
str32or64 := A_PtrSize  * 8
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Link, y10 vlblAboutText1, % L(lAboutText1, lAppName, lAppVersion, str32or64)
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Link, , % L(lAboutText2)
Gui, 2:Add, Link, , % L(lAboutText3)
Gui, 2:Font, s10 w400, Verdana
Gui, 2:Add, Link, , % L(lAboutText4)
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Button, x150 y+20 g2GuiClose, %lGui2Close%
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiHelp:
;------------------------------------------------------------

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lHelpTitle, lAppName, lAppVersion)
Gui, 2:+Owner1
intWidth := 450
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Text, x10 y10, %lAppName%
Gui, 2:Font, s10 w400, Verdana
Gui, 2:Add, Link, x10 w%intWidth%, %lHelpText1%
Gui, 2:Font, s8 w400, Verdana
; #### adapt hotkey
Gui, 2:Add, Link, w%intWidth%, %lHelpText2%
Gui, 2:Add, Link, w%intWidth%, %lHelpText3%
; #### adapt hotkey
Gui, 2:Add, Link, w%intWidth%, %lHelpText4%
; #### adapt hotkey
Gui, 2:Add, Link, w%intWidth%, %lHelpText5%
Gui, 2:Add, Link, w%intWidth%, %lHelpText6%
Gui, 2:Add, Link, w%intWidth%, %lHelpText7%
Gui, 2:Add, Link, w%intWidth%, %lHelpText8%
Gui, 2:Add, Link, w%intWidth%, %lHelpText9%
Gui, 2:Add, Button, x220 y+20 g2GuiClose, %lGui2Close%
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiOptions:
;------------------------------------------------------------

intGui1WinID := WinExist("A")

strMouseButtons := " |LButton|MButton|RButton|XButton1|XButton2|WheelUp|WheelDown|WheelLeft|WheelRight|"
; leave last | to enable default value on the last item

;---------------------------------------
; Gui hotkey titles
StringSplit, arrOptionsTitles, lOptionsTitles, |

;---------------------------------------
; Build Gui header
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lOptionsGuiTitle, lAppName, lAppVersion)
Gui, 2:+Owner1
Gui, 2:Font, s10 w700, Verdana
Gui, 2:Add, Text, x10 y10 w440 center, % L(lOptionsGuiTitle, lAppName)
Gui, 2:Font
Gui, 2:Add, Text, x10 y+10 w440 center, % L(lOptionsGuiIntro, lAppName)

; Build Hotkey Gui lines
loop, % arrIniKeyNames%0%
	GuiOptionsHotkey(A_Index)

; Build other options
Gui, 2:Add, Text, x10 y+5 w440 center, _______________________________________________________________

Gui, 2:Font, s8 w700
Gui, 2:Add, Text, x10 y+5 w440 center, %lOptionsOtherOptions%
Gui, 2:Font

Gui, 2:Add, CheckBox, y+20 x60 vblnOptionsRunAtStartup, %lOptionsRunAtStartup%
GuiControl, , blnOptionsRunAtStartup, % FileExist(A_Startup . "\" . lAppName . ".lnk") ? 1 : 0

Gui, 2:Add, CheckBox, yp x+60 vblnOptionsAlwaysTrayTip, %lOptionsAlwaysTrayTip%
GuiControl, , blnOptionsAlwaysTrayTip, % (intDisplayTrayTip < 0) ? 1 : 0

; Build Gui footer
Gui, 2:Add, Button, y+20 x170 vbtnOptionsSave gButtonOptionsSave, %lGuiSave%
Gui, 2:Add, Button, yp x+20 vbtnOptionsCancel gButtonOptionsCancel, %lGuiCancel%
Gui, 2:Add, Text
GuiControl, Focus, btnOptionsSave

;---------------------------------------
; Show until user click Save
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiOptionsHotkey(intIndex)
;------------------------------------------------------------
{
	global
	
	intType := 3
	if InStr(arrIniKeyNames%intIndex%, "Mouse")
		intType := 1
	if InStr(arrIniKeyNames%intIndex%, "Keyboard")
		intType := 2

	; Hotkey Header
	Gui, 2:Add, Text, % (intType = 3 ? "y+20 x5 w100 center" : "y+20 x5 w200 center"), %lOptionsTriggerFor%
	Gui, 2:Add, Text, % (intType = 3 ? "yp x108 w25 center" : "yp x+5 w25 center"), %lOptionsShift%
	Gui, 2:Add, Text, yp x+11 w25 center, %lOptionsCtrl%
	Gui, 2:Add, Text, yp x+10 w25 center, %lOptionsAlt%
	Gui, 2:Add, Text, yp x+10 w25 center, %lOptionsWin%
	if (intType <> 1)
		Gui, 2:Add, Text, yp x+10 w100 center, %lOptionsKeyboard%
	if (intType <> 2)
		Gui, 2:Add, Text, yp x+15 w80 center, %lOptionsMouse%

	Gui, 2:Font, s8 w700
	Gui, 2:Add, Text, % (intType = 3 ? "x5 y+5 w100 center" : "x5 y+5 w200 center"), % arrOptionsTitles%intIndex%
	Gui, 2:Font
	Gui, 2:Add, CheckBox, yp x+10 vblnOptionsShift%intIndex%, %A_Space%
		; if we have no text to put at the right of the ckeckbox
		; would it be possible to eliminate the dotted area ?
		; intIndex put a space becauyse this is what intIndex found the less visible.
	GuiControl, , blnOptionsShift%intIndex%, % InStr(strModifiers%intIndex%, "+") ? 1 : 0
	Gui, 2:Add, CheckBox, yp x+10 vblnOptionsCtrl%intIndex%, %A_Space%
	GuiControl, , blnOptionsCtrl%intIndex%, % InStr(strModifiers%intIndex%, "^") ? 1 : 0
	Gui, 2:Add, CheckBox, yp x+10 vblnOptionsAlt%intIndex%, %A_Space%
	GuiControl, , blnOptionsAlt%intIndex%, % InStr(strModifiers%intIndex%, "!") ? 1 : 0
	Gui, 2:Add, CheckBox, yp x+10 vblnOptionsWin%intIndex%, %A_Space%
	GuiControl, , blnOptionsWin%intIndex%, % InStr(strModifiers%intIndex%, "#") ? 1 : 0
	if (intType <> 1)
	{
		Gui, 2:Add, Hotkey, yp x+10 w90 vstrOptionsKey%intIndex% gOptionsHotkeyChanged
		GuiControl, , strOptionsKey%intIndex%, % strOptionsKey%intIndex%
	}
;	if (intType = 3)
;		Gui, 2:Add, Text, yp x+5, %A_Space%%A_Space%or
	if (intType <> 2)
		Gui, 2:Add, DropDownList, yp x+10 w90 vstrOptionsMouse%intIndex% gOptionsMouseChanged, % strMouseButtonsWithDefault%intIndex%
	Gui, 2:Add, Text, x10 y+5 w440, % lOptionsArrDescriptions%intIndex%
}
;------------------------------------------------------------


;------------------------------------------------------------
ButtonOptionsSave:
;------------------------------------------------------------
Gui, 2:Submit

StringSplit, arrIniVarNames, strIniKeyNames, |

Loop, % arrIniVarNames%0%
{
	strHotkey%A_Index% := Trim(strOptionsKey%A_Index% . strOptionsMouse%A_Index%)
	if StrLen(strHotkey%A_Index%)
	{
		Hotkey, % "$" . arrHotkeyVarNames%A_Index%, Off

		if (blnOptionsWin%A_Index%)
			strHotkey%A_Index% := "#" . strHotkey%A_Index%
		if (blnOptionsAlt%A_Index%)
			strHotkey%A_Index% := "!" . strHotkey%A_Index%
		if (blnOptionsShift%A_Index%)
			strHotkey%A_Index% := "+" . strHotkey%A_Index%
		if (blnOptionsCtrl%A_Index%)
			strHotkey%A_Index% := "^" . strHotkey%A_Index%
		IniWrite, % strHotkey%A_Index%, %strIniFile%, Global, % arrIniVarNames%A_Index%
	}
	else
		Oops(L(lOptionsNoKeyOrMouseSpecified, arrIniVarNames%A_Index%))
}
Gosub, LoadIniHotkeys ;  reload ini variables and reset hotkeys

IfExist, %A_Startup%\%lAppName%.lnk
	FileDelete, %A_Startup%\%lAppName%.lnk
if (blnOptionsRunAtStartup)
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%lAppName%.lnk
Menu, Tray, % blnOptionsRunAtStartup ? "Check" : "Uncheck", %lMenuRunAtStartup%

IniWrite, % blnOptionsAlwaysTrayTip ? -1 : 0, %strIniFile%, Global, DisplayTrayTip
	
Goto, 2GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonOptionsCancel:
;------------------------------------------------------------
Gosub, 2GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
OptionsHotkeyChanged:
;------------------------------------------------------------
strOptionsHotkeyControl := A_GuiControl ; hotkey var name
strOptionsHotkeyChanged := %strOptionsHotkeyControl% ; hotkey content

if !StrLen(strOptionsHotkeyChanged)
	return

SplitModifiersFromKey(strOptionsHotkeyChanged, strOptionsHotkeyChangedModifiers, strOptionsHotkeyChangedKey)

if StrLen(strOptionsHotkeyChangedModifiers) ; we have a modifier and we don't want it, reset keyboard to none and return
	GuiControl, , %A_GuiControl%, None
else ; we have a valid key, empty the mouse dropdown and return
{
	StringReplace, strOptionsMouseControl, strOptionsHotkeyControl, Key, Mouse ; get the matching mouse dropdown var
	GuiControl, ChooseString, %strOptionsMouseControl%, %A_Space%
}

return
;------------------------------------------------------------


;------------------------------------------------------------
OptionsMouseChanged:
;------------------------------------------------------------
strOptionsMouseControl := A_GuiControl ; mouse dropdown var name
StringReplace, strOptionsHotkeyControl, strOptionsMouseControl, Mouse, Key ; get the hotkey var

; we have a mouse button, empty the hotkey control
GuiControl, , %strOptionsHotkeyControl%, % ""

return
;------------------------------------------------------------


;------------------------------------------------------------
2GuiClose:
2GuiEscape:
;------------------------------------------------------------

Gui, 1:-Disabled
Gui, 2:Destroy
WinActivate, ahk_id %intGui1WinID%

return
;------------------------------------------------------------



;============================================================
; MIDDLESTUFF
;============================================================


;------------------------------------------------------------
LoadSettingsToGui:
;------------------------------------------------------------

GuiControlGet, blnSaveEnabled, Enabled, %lGuiSave%
if (blnSaveEnabled)
	return

Gui, 1:ListView, lvFoldersList
LV_Delete()
Gui, 1:ListView, lvDialogsList
LV_Delete()

Gui, 1:ListView, lvFoldersList
Loop, % arrGlobalFolders.MaxIndex()
{
	If !StrLen(arrGlobalFolders[A_Index].Name)
		LV_Add()
	else
		LV_Add(, arrGlobalFolders[A_Index].Name, arrGlobalFolders[A_Index].Path)
}
LV_Modify(1, "Select Focus")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

Gui, 1:ListView, lvDialogsList
Loop, % arrGlogalDialogs.MaxIndex()
{
	LV_Add(, arrGlogalDialogs[A_Index])
}
LV_Modify(1, "Select Focus")
LV_ModifyCol(1, "AutoHdr")

GuiControl, Focus, lvFoldersList

return
;------------------------------------------------------------


;------------------------------------------------------------
OpenFavorite:
;------------------------------------------------------------

strPath := GetPathFor(A_ThisMenuItem)

if (A_ThisHotkey = "$+MButton") or WindowIsDesktop(strGlobalClass) ;#### adapt to configurable triggers
	ComObjCreate("WScript.Shell").Exec("Explorer.exe /e /select," . strPath) ; ### TEST ON WIN XP - test when open the same folder in a new window
	; http://msdn.microsoft.com/en-us/library/bb774094http://msdn.microsoft.com/en-us/library/bb774094
	; ComObjCreate("Shell.Application").Explore(strPath)
	; ComObjCreate("WScript.Shell").Exec("Explorer.exe /e /select," . strPath)
	; ComObjCreate("Shell.Application").Open(strPath)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strGlobalClass)
	NavigateExplorer(strPath, strGlobalWinId)
else if WindowIsConsole(strGlobalClass)
	NavigateConsole(strPath, strGlobalWinId)
else
	NavigateDialog(strPath, strGlobalWinId, strGlobalClass)

return
;------------------------------------------------------------


;------------------------------------------------------------
OpenSpecialFolder:
;------------------------------------------------------------

; ShellSpecialFolderConstants: http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx
if (A_ThisMenuItem = lMenuDesktop)
	intSpecialFolder := 0
else if (A_ThisMenuItem = lMenuControlPanel)
	intSpecialFolder := 3
else if (A_ThisMenuItem = lMenuDocuments)
	intSpecialFolder := 5
else if (A_ThisMenuItem = lMenuRecycleBin)
	intSpecialFolder := 10
else if (A_ThisMenuItem = lMenuMyComputer)
	intSpecialFolder := 17
else if (A_ThisMenuItem = lMenuNetworkNeighborhood)
	intSpecialFolder := 18
else if (A_ThisMenuItem = lMenuPictures)
	intSpecialFolder := 39

if (A_ThisHotkey = "+MButton") or WindowIsDesktop(strGlobalClass)
	ComObjCreate("Shell.Application").Explore(intSpecialFolder)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strGlobalClass)
	NavigateExplorer(intSpecialFolder, strGlobalWinId)
else ; this is the console or a dialog box
{
	if (intSpecialFolder = 0)
		strPath := A_Desktop
	else if (intSpecialFolder = 5)
		strPath := A_MyDocuments
	else if (intSpecialFolder = 39)
	{
		; do not use: StringReplace, strPath, A_MyDocuments, Documents, Pictures
		; because A_MyDocument could contain a "Documents" string before the final folder
		StringLeft, strPath, A_MyDocuments, % StrLen(A_MyDocuments) - StrLen("Documents")
		strPath := strPath . "Pictures"
	}	
	else ; we do not support this special folder
		return

	if WindowIsConsole(strGlobalClass)
		NavigateConsole(strPath, strGlobalWinId)
	else
		NavigateDialog(strPath, strGlobalWinId, strGlobalClass)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
WinGetClassA()
;------------------------------------------------------------
{
	WinGetClass strClass, A
	return strClass
}


;------------------------------------------------------------
WinUnderMouseClass()
;------------------------------------------------------------
{
	WinGetClass strClass, % "ahk_id " . WinUnderMouseID()
	return strClass
}
;------------------------------------------------------------


;------------------------------------------------------------
WinUnderMouseID()
;------------------------------------------------------------
{
	MouseGetPos, , , strWinId
	return strWinId
}
;------------------------------------------------------------


;------------------------------------------------------------
GetPathFor(strName)
;------------------------------------------------------------
{
	Loop, % arrGlobalFolders.MaxIndex()
		if (strName = arrGlobalFolders[A_Index].Name)
			return arrGlobalFolders[A_Index].Path
}
;------------------------------------------------------------


;------------------------------------------------------------
CanOpenFavorite(strMouseOrKeyboard, ByRef strWinId, ByRef strClass)
;------------------------------------------------------------
; "CabinetWClass" and "ExploreWClass" -> Explorer
; "ProgMan" -> Desktop
; "WorkerW" -> Desktop
; "ConsoleWindowClass" -> Console (CMD)
; "#32770" -> Dialog
; "bosa_sdm_" (...) -> Dialog MS Office under WinXP
{
	if (strMouseOrKeyboard = "PopupMenuMouse")
		MouseGetPos, , , strWinId
	else
		strWinId := WinExist("A")
	WinGetClass strClass, % "ahk_id " . strWinId
	###_T(strMouseOrKeyboard, strClass)
	return WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or WindowIsConsole(strClass) or WindowIsDialog(strClass)
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsAnExplorer(strClass)
;------------------------------------------------------------
{
	return (strClass = "CabinetWClass") or (strClass = "ExploreWClass")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDesktop(strClass)
;------------------------------------------------------------
{
	return (strClass = "ProgMan") or (strClass = "WorkerW")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsConsole(strClass)
;------------------------------------------------------------
{
	return (strClass = "ConsoleWindowClass")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDialog(strClass)
;------------------------------------------------------------
{
	return (strClass = "#32770") or InStr(strClass, "bosa_sdm_")
}
;------------------------------------------------------------


;------------------------------------------------------------
DialogIsSupported(strWinId)
;------------------------------------------------------------
{
	WinGetTitle, strDialogTitle, ahk_id %strWinId%
	loop, % arrGlogalDialogs.MaxIndex()
		if InStr(strDialogTitle, arrGlogalDialogs[A_Index])
			return True

	return False
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateExplorer(varPath, strWinId)
;------------------------------------------------------------
/*
Excerpt and adapted from RMApp_Explorer_Navigate(FullPath, hwnd="") by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx
http://msdn.microsoft.com/en-us/library/aa752094
*/
{
	For pExp in ComObjCreate("Shell.Application").Windows
	{
		if (pExp.hwnd = strWinId)
			if varPath is integer ; ShellSpecialFolderConstant
			{
				try pExp.Navigate2(varPath)
				catch, objErr
					Oops(lNavigateSpecialError, varPath)
			}
			else
			{
				; try pExp.Navigate("file:///" . varPath) - removed to allow UNC (e.g. \\my.server.com@SSL\DavWWWRoot\Folder\Subfolder)
				try pExp.Navigate(varPath)
				catch, objErr
					Oops(lNavigateFileError, varPath)
			}
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateConsole(strPath, strWinId)
;------------------------------------------------------------
{
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
	SendInput, CD /D %strPath%{Enter}
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateDialog(strPath, strWinId, strClass)
;------------------------------------------------------------
/*
Excerpt from RMApp_Explorer_Navigate(FullPath, hwnd="") by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{
	if (strClass = "#32770")
		if ControlIsVisible("ahk_id " . strWinId, "Edit1")
			strControl := "Edit1"
			; in standard dialog windows, "Edit1" control is the right choice
		Else if ControlIsVisible("ahk_id " . strWinId, "Edit2")
			strControl := "Edit2"
			; but sometimes in MS office, if condition above fails, "Edit2" control is the right choice 
		Else ; if above fails - just return and do nothing.
			return
	Else if InStr(strClass, "bosa_sdm_") ; for some MS office dialog windows, which are not #32770 class
		if ControlIsVisible("ahk_id " . strWinId, "Edit1")
			strControl := "Edit1"
			; if "Edit1" control exists, it is the right choice
		Else if ControlIsVisible("ahk_id " . strWinId, "RichEdit20W2")
			strControl := "RichEdit20W2"
			; some MS office dialogs don't have "Edit1" control, but they have "RichEdit20W2" control, which is then the right choice.
		Else ; if above fails, just return and do nothing.
			return
	Else ; in all other cases, open a new Explorer and return from this function
	{
		ComObjCreate("Shell.Application").Explore(strPath)
		; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
		return
	}

	;===In this part (if we reached it), we'll send strPath to control and restore control's initial text after navigating to specified folder===
	ControlGetText, strPrevControlText, %strControl%, ahk_id %strWinId% ; we'll get and store control's initial text first
	
	ControlSetTextR(strControl, strPath, "ahk_id " . strWinId) ; set control's text to strPath
	ControlSetFocusR(strControl, "ahk_id " . strWinId) ; focus control
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
	
	;=== Avoid accidental hotkey & hotstring triggereing while doing SendInput - can be done simply by #UseHook, but do it if user doesn't have #UseHook in the script ===
	If (A_IsSuspended)
		blnWasSuspended := True
	if (!blnWasSuspended)
		Suspend, On
	SendInput, {End}{Space}{Backspace}{Enter} ; silly but necessary part - go to end of control, send dummy space, delete it, and then send enter
	if (!blnWasSuspended)
		Suspend, Off

	Sleep, 70 ; give some time to control after sending {Enter} to it
	ControlGetText, strControlTextAfterNavigation, %strControl%, ahk_id %strWinId% ; sometimes controls automatically restore their initial text
	if (strControlTextAfterNavigation <> strPrevControlText) ; if not
		ControlSetTextR(strControl, strPrevControlText, "ahk_id " . strWinId) ; we'll set control's text to its initial text
	
	if (WinExist("A") <> strWinId) ; sometimes initialy active window loses focus, so we'll activate it again
		WinActivate, ahk_id %strWinId%
}


;------------------------------------------------------------
ControlIsVisible(strWinTitle, strControlClass)
/*
Adapted from ControlIsVisible(WinTitle,ControlClass) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
;------------------------------------------------------------
{ ; used in Navigator
	ControlGet, blnIsControlVisible, Visible, , %strControlClass%, %strWinTitle%

	return blnIsControlVisible
}
;------------------------------------------------------------


;------------------------------------------------------------
ControlSetFocusR(strControl, strWinTitle = "", intTries = 3)
/*
Adapted from RMApp_ControlSetFocusR(Control, WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
;------------------------------------------------------------
{ ; used in Navigator. More reliable ControlSetFocus
	Loop, %intTries%
	{
		ControlFocus, %strControl%, %strWinTitle% ; focus control
		Sleep, 50
		ControlGetFocus, strFocusedControl, %strWinTitle% ; check
		if (strFocusedControl = strControl) ; if OK
			return True
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
ControlSetTextR(strControl, strNewText = "", strWinTitle = "", intTries = 3)
/*
Adapted from from RMApp_ControlSetTextR(Control, NewText="", WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
;------------------------------------------------------------
{ ; used in Navigator. More reliable ControlSetText
	Loop, %intTries%
	{
		ControlSetText, %strControl%, %strNewText%, %strWinTitle% ; set
		Sleep, 50
		ControlGetText, strCurControlText, %strControl%, %strWinTitle% ; check
		if (strCurControlText = strNewText) ; if OK
			return True
	}
}
;------------------------------------------------------------



;------------------------------------------------------------
SplitHotkey(strHotkey, strMouseButtons, ByRef strModifiers, ByRef strKey, ByRef strMouseButton, ByRef strMouseButtonsWithDefault)
;------------------------------------------------------------
{
	SplitModifiersFromKey(strHotkey, strModifiers, strKey)

	if InStr(strMouseButtons, "|" . strKey . "|") ;  we have a mouse button
	{
		strMouseButton := strKey
		strKey := ""
		StringReplace, strMouseButtonsWithDefault, strMouseButtons, %strMouseButton%|, %strMouseButton%|| ; with default value
	}
	else ; we have a key
		strMouseButtonsWithDefault := strMouseButtons ; no default value
}
;------------------------------------------------------------


;------------------------------------------------------------
SplitModifiersFromKey(strHotkey, ByRef strModifiers, ByRef strKey)
;------------------------------------------------------------
{
	intModifiersEnd := GetFirstNotModifier(strHotkey)
	StringLeft, strModifiers, strHotkey, %intModifiersEnd%
	StringMid, strKey, strHotkey, % (intModifiersEnd + 1)
}
;------------------------------------------------------------


;------------------------------------------------------------
GetFirstNotModifier(strHotkey)
;------------------------------------------------------------
{
	intPos := 0
	loop, Parse, strHotkey
		if (A_LoopField = "^") or (A_LoopField = "!") or (A_LoopField = "+") or (A_LoopField = "#")
			intPos := intPos + 1
		else
			return intPos
	return intPos
}
;------------------------------------------------------------



;============================================================
; TOOLS
;============================================================


; ------------------------------------------------
Oops(strMessage, objVariables*)
; ------------------------------------------------
{
	Gui, 1:+OwnDialogs
	MsgBox, 48, % L(lFuncOopsTitle, lAppName, lAppVersion), % L(strMessage, objVariables*)
}
; ------------------------------------------------



; ------------------------------------------------
L(strMessage, objVariables*)
; ------------------------------------------------
{
	Loop
	{
		if InStr(strMessage, "~" . A_Index . "~")
			StringReplace, strMessage, strMessage, ~%A_Index%~, % objVariables[A_Index]
 		else
			break
	}
	
	return strMessage
}
; ------------------------------------------------
