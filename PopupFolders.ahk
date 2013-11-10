;===============================================
/*
	PopupFolders
	Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
	By Jean Lalonde (JnLlnd on AHKScript.org forum), based on DirMenu v2 by Robert Ryan (rbrtryn on AutoHotkey.com forum)

	Version: PopupFolders v0.3 ALPHA
	- add NavigateConsole for console support (command prompt CMD)
	- change .ini filename to new app name
	- 

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

;@Ahk2Exe-SetName DirMenu3
;@Ahk2Exe-SetDescription Popup menu to jump from one folder to another
;@Ahk2Exe-SetVersion 0.1
;@Ahk2Exe-SetOrigFilename DirMenu.exe


;============================================================
; INITIALIZATION
;============================================================


#NoEnv
#SingleInstance force
#Include %A_ScriptDir%\PopupFolders_LANG.ahk
SetWorkingDir %A_ScriptDir%

global arrGlobalFolders := Object()
global arrGlogalDialogs := Object()
global strIniFile := A_ScriptDir . "\" . lAppName . ".ini"

;@Ahk2Exe-IgnoreBegin
	; Piece of code for developement phase only - won't be compiled
	if (A_ComputerName = "JEAN-PC") ; for my home PC
		strIniFile := A_ScriptDir . "\" . lAppName . "-HOME.ini"
	else if InStr(A_ComputerName, "STIC") ; for my work hotkeys
		strIniFile := A_ScriptDir . "\" . lAppName . "-WORK.ini"
	; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd

Gosub, LoadIniFile
Gosub, BuildGUI
Gosub, BuildTrayMenu
Gosub, BuildSpecialFoldersMenu
Gosub, BuildFoldersMenu
Gosub, BuildAddDialogMenu

if (blnDisplayTrayTip)
	TrayTip, % L(lTrayTipInstalledTitle, lAppName, lAppVersionLong)
	, % L(lTrayTipInstalledDetail, lAppName), , 1
return


;============================================================
; BACK END FUNCTIONS AND COMMANDS
;============================================================


; -----------------------------------------------------------
LoadIniFile:
; -----------------------------------------------------------
blnDebug := False

strSettingsHotkeyDefault := "^#f"

IfNotExist, %strIniFile%
	FileAppend,
		(LTrim Join`r`n
			[Global]
			SettingsHotkey=%strSettingsHotkeyDefault%
			DisplayTrayTip=1

			[Dialogs]
			Dialog1=Export
			Dialog2=Import
			Dialog3=Insert
			Dialog4=Open
			Dialog5=Save
			Dialog6=Select
			Dialog7=Upload

			[Folders]
			Folder1=C:\|C:\
			Folder2=Windows|%A_WinDir%
			Folder3=Program Files|%A_ProgramFiles%
		)
		, %strIniFile%
	
IniRead, blnDisplayTrayTip, %strIniFile%, Global, DisplayTrayTip
IniRead, strSettingsHotkey, %strIniFile%, Global, SettingsHotkey
if (strSettingsHotkey = "ERROR")
{
	IniWrite, %strSettingsHotkeyDefault%, %strIniFile%, Global, SettingsHotkey
	strSettingsHotkey := strSettingsHotkeyDefault
}
Hotkey, %strSettingsHotkey%, ShowGui

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
if (blnDebug)
	Loop, % arrGlobalFolders.MaxIndex()
		###_D("Folder" . A_Index . ": " . arrGlobalFolders[A_Index].Name . " -> " . arrGlobalFolders[A_Index].Path)
Loop
{
	IniRead, strIniLine, %strIniFile%, Dialogs, Dialog%A_Index%
	if (strIniLine = "ERROR")
		Break
	arrGlogalDialogs.Insert(strIniLine)
}
if (blnDebug)
	Loop, % arrGlogalDialogs.MaxIndex()
		###_D("Dialog" . A_Index . ": " . arrGlogalDialogs[A_Index])

blnDebug := False
return
;------------------------------------------------------------



;============================================================
; FRONT END FUNCTIONS AND COMMANDS
;============================================================


;------------------------------------------------------------
#If, CanOpenFavorite(strGlobalWinId, strGlobalClass)
MButton::
;------------------------------------------------------------
blnDebug := false

if (blnDebug)
	###_D("Yes: " . strGlobalWinId .  " " . strGlobalClass)

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

WinActivate, % "ahk_id " . strGlobalWinId
if (WindowIsAnExplorer(strGlobalClass) or WindowIsDesktop(strGlobalClass) or WindowIsConsole(strGlobalClass) or DialogIsSupported(strGlobalWinId))
{
	; Enable Add This Folder only if the mouse is over an Explorer (tested on WIN_XP and WIN_7) or a dialog box (works on WIN_7, not on WIN_XP)
	; Other tests shown that WIN_8 behaves like WIN_7. So, I assume WIN_8 to work. If someone could confirm (until I can test it myself)?
	Menu, menuFolders
		, % WindowIsAnExplorer(strGlobalClass) or (WindowIsDialog(strGlobalClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
		, %lMenuAddThisFolder%
	Menu, menuFolders, Show
}
else
	Menu, menuAddDialog, Show

blnDebug := false
return
#If
;------------------------------------------------------------


;------------------------------------------------------------
+MButton::
;------------------------------------------------------------
blnDebug := false

MouseGetPos, , , strGlobalWinId
WinGetClass strGlobalClass, % "ahk_id " . strGlobalWinId

; In case it was disabled while in a dialog box
Menu, menuSpecialFolders, Enable, %lMenuMyComputer%
Menu, menuSpecialFolders, Enable, %lMenuNetworkNeighborhood%
Menu, menuSpecialFolders, Enable, %lMenuControlPanel%
Menu, menuSpecialFolders, Enable, %lMenuRecycleBin%

; Enable Add This Folder only if the mouse is over an Explorer (tested on WIN_XP and WIN_7) or a dialog box (works on WIN_7, not on WIN_XP)
; Other tests shown that WIN_8 behaves like WIN_7. So, I assume WIN_8 to work. If someone could confirm (until I can test it myself)?
Menu, menuFolders
	, % WindowIsAnExplorer(strGlobalClass) or (WindowIsDialog(strGlobalClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
	, %lMenuAddThisFolder%
Menu, menuFolders, Show

blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
BuildTrayMenu:
;------------------------------------------------------------
Menu, Tray, Add
Menu, Tray, Add, %lMenuAbout%, GuiAbout
Menu, Tray, Add, %lMenuEditFoldersMenu%, ShowGui
Menu, Tray, Default, %lMenuEditFoldersMenu%
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
blnDebug := False

Menu, menuFolders, Add
Menu, menuFolders, DeleteAll
Loop, % arrGlobalFolders.MaxIndex()
{
	if (blnDebug)
		###_D(arrGlobalFolders[A_Index].Name)
	if (arrGlobalFolders[A_Index].Name = lMenuSeparator)
		Menu, menuFolders, Add
	else
		Menu, menuFolders, Add, % arrGlobalFolders[A_Index].Name, OpenFavorite
}
Menu, menuFolders, Add
Menu, menuFolders, Add, %lMenuSpecialFolders%, :menuSpecialFolders
Menu, menuFolders, Add
Menu, menuFolders, Add, %lMenuEditThisMenu%, ShowGui
Menu, menuFolders, Default, %lMenuEditThisMenu%
Menu, menuFolders, Add, %lMenuAddThisFolder%, AddThisFolder

if (blnDebug)
	Menu, menuFolders, Show

blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
BuildAddDialogMenu:
;------------------------------------------------------------
blnDebug := False

Menu, menuAddDialog, Add, %lMenuDialogNotSupported%, AddThisDialog
Menu, menuAddDialog, Disable, %lMenuDialogNotSupported%
Menu, menuAddDialog, Add, %lMenuAddThisDialog%, AddThisDialog
Menu, menuAddDialog, Add
Menu, menuAddDialog, Add, %lMenuEditFoldersMenu%, ShowGui
Menu, menuAddDialog, Default, %lMenuAddThisDialog%

if (blnDebug)
	Menu, menuAddDialog, Show

blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
BuildGui:
;------------------------------------------------------------
blnDebug := false

Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Picture, x10 y10, %A_ScriptDir%\ico\Folders-Likes-icon-48.png
Gui, 1:Add, Text, y20 w490 h25, %lAppName%
Gui, 1:Font, s8 w400, Verdana
Gui, Add, ListView, x10 w350 h220 Count32 -Multi NoSortHdr AltSubmit vlvFoldersList gGuiLvFoldersEvent, %lGuiLvFoldersHeader%
Gui, Add, Button, x+10 w75 r1 gGuiAddFolder, %lGuiAddFolder%
Gui, Add, Button, w75 r1 gGuiRemoveFolder, %lGuiRemoveFolder%
Gui, Add, Button, W75 r1 gGuiEditFolder, %lGuiEditFolder%
Gui, Add, Button, w75 r1 gGuiAddSeparator, %lGuiSeparator%
Gui, Add, Button, w75 r1 gGuiMoveFolderUp, %lGuiMoveFolderUp%
Gui, Add, Button, w75 r1 gGuiMoveFolderDown, %lGuiMoveFolderDown%

Gui, Add, ListView
	, x10 w350 h120 Count16 -Multi NoSortHdr +0x10 AltSubmit vlvDialogsList, %lGuiLvDialogsHeader%
Gui, Add, Button, x+10 w75 r1 gGuiAddDialog, %lGuiAddDialog%
Gui, Add, Button, w75 r1 gGuiRemoveDialog, %lGuiRemoveDialog%
Gui, Add, Button, w75 r1 gGuiEditDialog, %lGuiEditDialog%

Gui, Add, Button, x100 w75 r1 Disabled Default gGuiSave, %lGuiSave%
Gui, Add, Button, x+40 w75 r1 gGuiCancel, %lGuiCancel%
Gui, Add, Button, x+80 w75 r1 gGuiAbout, %lGuiAbout%

if (blnDebug)
	Gui, Show, w455 h455, % L(lGuiTitle, lAppName, lAppVersionLong)

blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiLvFoldersEvent:
;------------------------------------------------------------
blnDebug := false
if (BlnDebug)
	###_D(A_GuiEvent)

; REMOVE IF NOT USED
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFolder:
;------------------------------------------------------------
FileSelectFolder strNewPath, *C:\
if (strNewPath = "")
	return
AddFolder(strNewPath)
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveFolder:
;------------------------------------------------------------
Help(lNotImplementedYet)
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditFolder:
;------------------------------------------------------------
Help(lNotImplementedYet)
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddSeparator:
;------------------------------------------------------------
###_D(lNotImplementedYet . "`n" . lMenuSeparator)
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFolderUp:
;------------------------------------------------------------
Help(lNotImplementedYet)
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFolderDown:
;------------------------------------------------------------
Help(lNotImplementedYet)
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
Help(lNotImplementedYet)
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditDialog:
;------------------------------------------------------------
Help(lNotImplementedYet)
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiSave:
;------------------------------------------------------------
Help(lNotImplementedYet)
/*
	GuiControlGet, blnRevertEnabled, Enabled,  Re&vert
	if (blnRevertEnabled)
	{
		MsgBox, 36, %strAppName% v%strAppVersion%, Discard changes?
		IfMsgBox, No
			return
	}
	Gui Cancel
	LoadListViews()
*/
Gosub, BuildFoldersMenu
Gui, Cancel
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiCancel:
;------------------------------------------------------------
Help(lNotImplementedYet)
; EMPTY 2 LVs
Gui, ListView, lvFoldersList
LV_Delete()
Gui, ListView, lvDialogsList
LV_Delete()
Gui, Cancel
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAbout:
;------------------------------------------------------------
Help(lNotImplementedYet)
return
;------------------------------------------------------------


;------------------------------------------------------------
ShowGui:
;------------------------------------------------------------
Gosub, LoadSettingsToGui
Gui, Show, w455 h455, % L(lGuiTitle, lAppName, lAppVersionLong)
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
blnDebug := false

objPrevClipboard := ClipboardAll ; Save the entire clipboard
ClipBoard := ""

; Add This folder menu is active only if we are in Explorer (WIN_XP, WIN_7 or WIN_8) or in a Dialog box (WIN_7 or WIN_8).
; In all these OS, the key sequence {F4}{Esc} selects the current location of the window.
Loop, 3
{
	Sleep, 100 * A_Index
	Send {F4}{Esc} ; F4 move the caret the "Go To A Different Folder box" and {Esc} select it content ({Esc} could be replaced by ^a to Select All)
	Sleep, 100 * A_Index
	Send ^c ; Copy
	Sleep, 100 * A_Index
} Until (StrLen(ClipBoard))

strCurrentFolder := ClipBoard

If StrLen(strCurrentFolder)
{
	Gosub, ShowGui
	AddFolder(strCurrentFolder)
}
else
	Oops(lDialogCouldDetectCurrentFolder)

Clipboard := objPrevClipboard ; Restore the original clipboard
objPrevClipboard := "" ; Free the memory in case the clipboard was very large

blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
AddFolder(strPath)
;------------------------------------------------------------
{
	blnDebug := false
	Gui +OwnDialogs

	; suggest the deepest folder's name as default name for the added folder
	SplitPath, strPath, strDefaultName, , , , strDrive
	if (blnDebug)
		###_D("strPath: " . strPath . "`nstrDefaultName: " . strDefaultName . "`nstrDrive: " . strDrive)
	if !StrLen(strDefaultName) ; we are probably at the root of a drive
		strDefaultName := strDrive

	Loop
	{
		InputBox strName, % L(lDialogTitle, lAppName, lAppVersionLong) . lDialogFolderNameTitle, %lDialogFolderNamePrompt%, , 250, 120, , , , , %strDefaultName%
		if (ErrorLevel) or !StrLen(strName)
			return
	} until FolderNameIsNew(strName)
	
	Gui, ListView, lvFoldersList
	if (blnDebug)
		###_D("LV_GetCount: " . LV_GetCount() . "`nLV_GetNext: " . LV_GetNext() . "`nResult: " . LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 1) : 1)
	LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 1) : 1, "Select Focus", strName, strPath)
	LV_ModifyCol(1, "AutoHdr")
	LV_ModifyCol(2, "AutoHdr")
	GuiControl, Enable, %lGuiSave%
	GuiControl, Enable, %lGuiRemoveFolder%
	GuiControl, Enable, %lGuiEditFolder%

	blnDebug := false
}
;------------------------------------------------------------


;------------------------------------------------------------
FolderNameIsNew(strCandidateName)
;------------------------------------------------------------
{
	Gui, ListView, lvFoldersList
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
blnDebug := false

WinGetTitle, strDialogTitle, ahk_id %strGlobalWinId%
Gosub, ShowGui
AddDialog(strDialogTitle)


blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
AddDialog(strCurrentDialogTitle)
;------------------------------------------------------------
{
	Gui +OwnDialogs
	strDialog := strCurrentDialogTitle

	InputBox, strDialog, % L(lDialogAddDialogTitle, lAppName, lAppVersionLong), %lDialogAddDialogPrompt%, , 250, 150, , , , , %strDialog%
	if (ErrorLevel) or !StrLen(strDialog)
		return
	
	Gui, ListView, lvDialogsList
	Loop, % LV_GetCount()
	{
		LV_GetText(strThisName, A_Index, 1)
		if (strDialog = strThisName)
		{
			Oops(lDialogAddDialogAlready)
			return
		}
	}

	LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 1) : 1, "Select Focus", strDialog)
	LV_ModifyCol(1, "AutoHdr")
	GuiControl, Enable, %lGuiSave%
	GuiControl, Enable, %lGuiRemoveDialog%
	GuiControl, Enable, %lGuiEditDialog%
}

;------------------------------------------------------------



;============================================================
; MIDDLESTUFF
;============================================================


;------------------------------------------------------------
LoadSettingsToGui:
;------------------------------------------------------------
blnDebug := false

Gui, ListView, lvFoldersList
Loop, % arrGlobalFolders.MaxIndex()
{
	if (blnDebug)
		###_D(arrGlobalFolders[A_Index].Name . " " . arrGlobalFolders[A_Index].Path)
	LV_Add(, arrGlobalFolders[A_Index].Name, arrGlobalFolders[A_Index].Path)
}
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

Gui, ListView, lvDialogsList
Loop, % arrGlogalDialogs.MaxIndex()
{
	if (blnDebug)
		###_D(arrGlogalDialogs[A_Index])
	LV_Add(, arrGlogalDialogs[A_Index])
}
LV_ModifyCol(1, "AutoHdr")

blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
OpenFavorite:
;------------------------------------------------------------
blnDebug := false

strPath := GetPahtFor(A_ThisMenuItem)
if (blnDebug)
	###_D("strGlobalWinId: " . strGlobalWinId . "`nstrGlobalClass: " . strGlobalClass . "`nstrPath: " . strPath)

if (A_ThisHotkey = "+MButton") or WindowIsDesktop(strGlobalClass)
	ComObjCreate("Shell.Application").Explore(strPath)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strGlobalClass)
	NavigateExplorer(strPath, strGlobalWinId)
else if WindowIsConsole(strGlobalClass)
	NavigateConsole(strPath, strGlobalWinId)
else
	NavigateDialog(strPath, strGlobalWinId, strGlobalClass)
blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
OpenSpecialFolder:
;------------------------------------------------------------

blnDebug := false
if (blnDebug)
	###_D("strGlobalWinId: " . strGlobalWinId . "`nstrGlobalClass: " . strGlobalClass . "`nA_ThisMenuItem: " . A_ThisMenuItem)
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
		StringReplace, strPath, A_MyDocuments, Documents, Pictures
	else ; we do not support this special folder
		return

	if WindowIsConsole(strGlobalClass)
		NavigateConsole(strPath, strGlobalWinId)
	else
		NavigateDialog(strPath, strGlobalWinId, strGlobalClass)
}
blnDebug := false
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
GetPahtFor(strName)
;------------------------------------------------------------
{
	blnDebug := false

	if (blnDebug)
		Loop, % arrGlobalFolders.MaxIndex()
			###_D(strName . " = " . arrGlobalFolders[A_Index].Name)
	Loop, % arrGlobalFolders.MaxIndex()
		if (strName = arrGlobalFolders[A_Index].Name)
			return arrGlobalFolders[A_Index].Path

	blnDebug := false
	return
}
;------------------------------------------------------------


;------------------------------------------------------------
CanAddFolder(strClass)
;------------------------------------------------------------
{
}
;------------------------------------------------------------


;------------------------------------------------------------
CanOpenFavorite(ByRef strWinId, ByRef strClass)
;------------------------------------------------------------
; "CabinetWClass" -> Explorer
; "ProgMan" -> Desktop
; "WorkerW" -> Desktop
; "ConsoleWindowClass" -> Console (CMD)
; "#32770" -> Dialog
{
	blnDebug := false
	
	MouseGetPos, , , strWinId
	WinGetClass strClass, % "ahk_id " . strWinId

	if (blnDebug)
		if WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or WindowIsConsole(strClass) or WindowIsDialog(strClass)
			###_D(strClass . " is OK")
		else
			###_D(strClass . " is NOT OK")
		
	return WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or WindowIsConsole(strClass) or WindowIsDialog(strClass)
	blnDebug := false
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsAnExplorer(strClass)
;------------------------------------------------------------
{
	return (strClass = "CabinetWClass")
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
	blnDebug := false

	WinGetTitle, strDialogTitle, ahk_id %strWinId%
	if (blnDebug)
		loop, % arrGlogalDialogs.MaxIndex()
			###_D("DialogIsSupported? " . strDialogTitle . " " . arrGlogalDialogs[A_Index] . " " InStr(strDialogTitle, arrGlogalDialogs[A_Index]))
	loop, % arrGlogalDialogs.MaxIndex()
		if InStr(strDialogTitle, arrGlogalDialogs[A_Index])
			return True
	return False
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateExplorer(strPath, strWinId)
;------------------------------------------------------------
/*
Excerpt from RMApp_Explorer_Navigate(FullPath, hwnd="") by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx
http://msdn.microsoft.com/en-us/library/aa752094
*/
{
    For pExp in ComObjCreate("Shell.Application").Windows
        if (pExp.hwnd = strWinId) ; matching window found
            if strPath is integer  ; ShellSpecialFolderConstant
                pExp.Navigate2(strPath)
            else
                pExp.Navigate("file:///" . strPath)
            return
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateConsole(strPath, strWinId)
;------------------------------------------------------------
{
	blnDebug := false
	if (blnDebug)
		###_D("strPath: " . strPath . ";nstrWinId: " . strWinId)

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
	blnDebug := false
	if (blnDebug)
		###_D("Dialog: " . strPath)
	
	if (strClass = "#32770")
	{
		if ControlIsVisible("ahk_id " . strWinId, "Edit1")
			strControl := "Edit1"
			; in standard dialog windows, "Edit1" control is the right choice
		Else if ControlIsVisible("ahk_id " . strWinId, "Edit2")
			strControl := "Edit2"
			; but sometimes in MS office, if condition above fails, "Edit2" control is the right choice 
		Else ; if above fails - just return and do nothing.
		{
			if (blnDebug)
				###_D("strClass is #32770 but no valid control is visible.")
			Return
		}
	}
	Else if InStr(strClass, "bosa_sdm_") ; for some MS office dialog windows, which are not #32770 class
	{
		if ControlIsVisible("ahk_id " . strWinId, "Edit1")
			strControl := "Edit1"
			; if "Edit1" control exists, it is the right choice
		Else if ControlIsVisible("ahk_id " . strWinId, "RichEdit20W2")
			strControl := "RichEdit20W2"
			; some MS office dialogs don't have "Edit1" control, but they have "RichEdit20W2" control, which is then the right choice.
		Else ; if above fails, just return and do nothing.
		{
			if (blnDebug)
				###_D("strClass is bosa_sdm_ but no valid control is visible.")
			Return
		}
	}
	Else ; in all other cases, open a new Explorer and return from this function
	{
		if (blnDebug)
			###_D("Dialog box strClass is not in #32770 or bosa_sdm_. Open a new Explorer.")
		ComObjCreate("Shell.Application").Explore(strPath)
		; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
		Return
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


ControlIsVisible(strWinTitle, strControlClass)
/*
Adapted from ControlIsVisible(WinTitle,ControlClass) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{ ; used in Navigator
	ControlGet, blnIsControlVisible, Visible, , %strControlClass%, %strWinTitle%
	return blnIsControlVisible
}


ControlSetFocusR(strControl, strWinTitle = "", intTries = 3)
/*
Adapted from RMApp_ControlSetFocusR(Control, WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
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


ControlSetTextR(strControl, strNewText = "", strWinTitle = "", intTries = 3)
/*
Adapted from from RMApp_ControlSetTextR(Control, NewText="", WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{  ; used in Navigator. More reliable ControlSetText
	Loop, %intTries%
	{
		ControlSetText, %strControl%, %strNewText%, %strWinTitle% ; set
		Sleep, 50
		ControlGetText, strCurControlText, %strControl%, %strWinTitle% ; check
		if (strCurControlText = strNewText) ; if OK
			return True
	}
}



;============================================================
; TOOLS
;============================================================

; ### DELETE UNUSED FUNCTIONS BEDORE FINAL RELEASE

; ------------------------------------------------
Help(strMessage, objVariables*)
; ------------------------------------------------
{
	Gui, 1:+OwnDialogs 
	StringLeft, strTitle, strMessage, % InStr(strMessage, "$") - 1
	StringReplace, strMessage, strMessage, %strTitle%$
	MsgBox, 0, % L(lFuncHelpTitle, lAppName, lAppVersionLong, strTitle), % L(strMessage, objVariables*)
}
; ------------------------------------------------



; ------------------------------------------------
Oops(strMessage, objVariables*)
; ------------------------------------------------
{
	Gui, 1:+OwnDialogs
	MsgBox, 48, % L(lFuncOopsTitle, lAppName, lAppVersionLong), % L(strMessage, objVariables*)
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

###_D(str, blnAgain := 0)
{
   static blnSkip### := false
   if (blnSkip###)
      if (blnAgain)
         blnSkip### := false
      else
         return
   intOption := 6 + 512 ; 6 Cancel/Try Again/Continue + 512 Makes the 2nd button the default
   MsgBox, % intOption, ###_D(ébug), %str%
   IfMsgBox TryAgain
      ExitApp
   else IfMsgBox Cancel
      blnSkip### := true
   else
      return
}

