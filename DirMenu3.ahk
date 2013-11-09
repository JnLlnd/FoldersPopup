;===============================================
/*
DirMenu3
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By Jean Lalonde (JnLlnd on AHKScript.org forum), based on DirMenu v2 by Robert Ryan (rbrtryn on AutoHotkey.com forum)

	Version 0.1 ALPHA
	- init skeleton, read ini file and create arrays for folders menu and supported dialog boxes
	- create language file, build gui, tray menu and folder menu, skeleton for front end buttons and commands
	- create AddThisDialog menu, MButton condition, CanOpenFavorite improvements with WindowIsAnExplorer, WindowIsDesktop and DialogIsSupported
	- add SpecialFolders menu, OpenFavorite for Explorer and Desktop, NavigateExplorer

	Version:	DirMenu v2.2 (never released / not stable - base of a total rewrite to DirMenu3)
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
	
	Version:	DirMenu v2.1
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
; INITIALISATION
;============================================================


#NoEnv
#SingleInstance force
#Include %A_ScriptDir%\DirMenu3_LANG.ahk
SetWorkingDir %A_ScriptDir%

global bln###Debug := False
global strAppName := "DirMenu3"
global strAppVersion := "0.1 ALPHA"
global strIniFile := A_ScriptDir . "\DirMenu3.ini"

;@Ahk2Exe-IgnoreBegin
	; Piece of code for developement phase only - won't be compiled
	if (A_ComputerName = "JEAN-PC") ; for my home PC
		strIniFile := A_ScriptDir . "\DirMenu3-HOME.ini"
	else if InStr(A_ComputerName, "STIC") ; for my work hotkeys
		strIniFile := A_ScriptDir . "\DirMenu3-WORK.ini"
	; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd

Gosub, LoadIniFile
Gosub, BuildGUI
Gosub, BuildTrayMenu
Gosub, BuildSpecialFoldersMenu
Gosub, BuildFoldersMenu
Gosub, BuildAddDialogMenu

return


;============================================================
; BACK END FUNCTIONS AND COMMANDS
;============================================================


; -----------------------------------------------------------
LoadIniFile:
; -----------------------------------------------------------
blnDebug := False

IfNotExist, %strIniFile%
	FileAppend,
		(LTrim Join`r`n
			[Global]

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
			Folder3=Program Files|%A_Programs%
		)
		, %strIniFile%
		
arrFolders := Object()
Loop
{
	IniRead, strIniLine, %strIniFile%, Folders, Folder%A_Index%
	if (strIniLine = "ERROR")
		Break
	StringSplit, arrThisObject, strIniLine, |
	objFolder := Object()
	objFolder.Name := arrThisObject1
	objFolder.Path := arrThisObject2
	arrFolders.Insert(objFolder)
}
if (blnDebug)
	Loop, % arrFolders.MaxIndex()
		###_D("Folder" . A_Index . ": " . arrFolders[A_Index].Name . " -> " . arrFolders[A_Index].Path)
arrDialogs := Object()
Loop
{
	IniRead, strIniLine, %strIniFile%, Dialogs, Dialog%A_Index%
	if (strIniLine = "ERROR")
		Break
	arrDialogs.Insert(strIniLine)
}
if (blnDebug)
	Loop, % arrDialogs.MaxIndex()
		###_D("Dialog" . A_Index . ": " . arrDialogs[A_Index])

blnDebug := False
return
;------------------------------------------------------------



;============================================================
; FRONT END FUNCTIONS AND COMMANDS
;============================================================


;------------------------------------------------------------
+MButton::
;------------------------------------------------------------
blnDebug := false

Menu, menuFolders, Disable, &Add This Folder
Menu, menuFolders, Show

blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
#If, CanOpenFavorite(strWinId, strClass)
MButton::
;------------------------------------------------------------
blnDebug := false

if (blnDebug)
	###_D("Yes: " . strWinId .  " " . strClass)

if ((strClass = "#32770") or DialogIsSupported(strWinId))
{
	Menu, menuSpecialFolders, Disable, Control Panel
	Menu, menuSpecialFolders, Disable, Recycle Bin
}

WinActivate, % "ahk_id " . strWinId
if (WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or DialogIsSupported(strWinId))
	Menu, menuFolders, Show
else
	Menu, menuAddDialog, Show

blnDebug := false
return
#If
;------------------------------------------------------------


;------------------------------------------------------------
BuildTrayMenu:
;------------------------------------------------------------
Menu, Tray, Add
Menu, Tray, Add, A&bout, GuiAbout
Menu, Tray, Add, &Edit Folders Menu, ShowGui
Menu, Tray, Default, &Edit Folders Menu
return
;------------------------------------------------------------


;------------------------------------------------------------
BuildSpecialFoldersMenu:
;------------------------------------------------------------
Menu, menuSpecialFolders, Add, My Computer, OpenSpecialFolder
Menu, menuSpecialFolders, Add, Control Panel, OpenSpecialFolder
Menu, menuSpecialFolders, Add, Recycle Bin, OpenSpecialFolder
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

Gui, Add, Button, x100 w75 r1 gGuiSave Default, %lGuiSave%
Gui, Add, Button, x+40 w75 r1 gGuiCancel, %lGuiCancel%
Gui, Add, Button, x+80 w75 r1 gGuiAbout, %lGuiAbout%

if (blnDebug)
	Gui, Show, w455 h455, %lAppName% %lAppVersionLong% - Settings

blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiLvFoldersEvent:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFolder:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveFolder:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditFolder:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddSeparator:
;------------------------------------------------------------
; lMenuSeparator
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFolderUp:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFolderDown:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddDialog:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveDialog:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditDialog:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiSave:
;------------------------------------------------------------
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
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAbout:
;------------------------------------------------------------
return
;------------------------------------------------------------


;------------------------------------------------------------
ShowGui:
;------------------------------------------------------------
Gosub, LoadSettingsToGui
Gui, Show, w455 h455, %lAppName% %lAppVersionLong% - Settings
return
;------------------------------------------------------------


;------------------------------------------------------------
GuiClose:
;------------------------------------------------------------
GoSub, GuiSave
return
;------------------------------------------------------------



;============================================================
; MIDDLESTUFF
;============================================================


;------------------------------------------------------------
LoadSettingsToGui:
;------------------------------------------------------------
blnDebug := false

Gui, ListView, lvFoldersList
Loop, % arrFolders.MaxIndex()
{
	if (blnDebug)
		###_D(arrFolders[A_Index].Name . " " . arrFolders[A_Index].Path)
	LV_Add(, arrFolders[A_Index].Name, arrFolders[A_Index].Path)
}
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

Gui, ListView, lvDialogsList
Loop, % arrDialogs.MaxIndex()
{
	if (blnDebug)
		###_D(arrDialogs[A_Index])
	LV_Add(, arrDialogs[A_Index])
}
LV_ModifyCol(1, "AutoHdr")

blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
BuildFoldersMenu:
;------------------------------------------------------------
blnDebug := False

Menu, menuFolders, Add
Menu, menuFolders, DeleteAll
Loop, % arrFolders.MaxIndex()
{
	if (blnDebug)
		###_D(arrFolders[A_Index].Name)
	if (arrFolders[A_Index].Name = lMenuSeparator)
		Menu, menuFolders, Add
	else
		Menu, menuFolders, Add, % arrFolders[A_Index].Name, OpenFavorite
}
Menu, menuFolders, Add
Menu, menuFolders, Add, &Special Folders, :menuSpecialFolders
Menu, menuFolders, Add
Menu, menuFolders, Add, &Edit This Menu, ShowGui
Menu, menuFolders, Default, &Edit This Menu
Menu, menuFolders, Add, &Add This Folder, AddThisFolder

if (blnDebug)
	Menu, menuFolders, Show

blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
BuildAddDialogMenu:
;------------------------------------------------------------
blnDebug := False

Menu, menuAddDialog, Add, This dialog box type is not supported yet, AddThisDialog
Menu, menuAddDialog, Disable, This dialog box type is not supported yet
Menu, menuAddDialog, Add, &Add This Dialog box to the supported list, AddThisDialog
Menu, menuAddDialog, Add
Menu, menuAddDialog, Add, &Edit the Folders Menu, ShowGui
Menu, menuAddDialog, Default, &Add This Dialog box to the supported list

if (blnDebug)
	Menu, menuAddDialog, Show

blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
OpenFavorite:
;------------------------------------------------------------
blnDebug := true

strPath := GetPahtFor(A_ThisMenuItem)
if (blnDebug)
	###_D("strWinId: " . strWinId . "`nstrClass: " . strClass . "`nstrPath: " . strPath)

if (A_ThisHotkey = "+MButton")
	Run, Explorer.exe /n`,%strPath%
else if WindowIsDesktop(strClass)
	ComObjCreate("Shell.Application").Explore(strPath)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strClass)
	NavigateExplorer(strPath, strWinId)
else
{
	###_D("Dialog")
}
blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
OpenSpecialFolder:
;------------------------------------------------------------
blnDebug := true
if (blnDebug)
	###_D("strWinId: " . strWinId . "`nstrClass: " . strClass . "`nA_ThisMenuItem: " . A_ThisMenuItem)
if (A_ThisMenuItem = "Control Panel")
	NavigateExplorer(3, strWinId)
else if (A_ThisMenuItem = "Recycle Bin")
	NavigateExplorer(10, strWinId)
else if(A_ThisMenuItem = "My Computer")
	NavigateExplorer(17, strWinId)
blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
AddThisFolder:
;------------------------------------------------------------
blnDebug := true
blnDebug := False
return
;------------------------------------------------------------


;------------------------------------------------------------
AddThisDialog:
;------------------------------------------------------------
blnDebug := true
blnDebug := False
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
GetPahtFor(strMenu)
;------------------------------------------------------------
{
	blnDebug := false
	global arrFolders

	if (blnDebug)
		Loop, % arrFolders.MaxIndex()
			###_D(strMenu . " = " . arrFolders[A_Index].Name)
	Loop, % arrFolders.MaxIndex()
		if (strMenu = arrFolders[A_Index].Name)
			return arrFolders[A_Index].Path

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
; "ConsoleWindowClass" -> Console
; "#32770" -> Dialog
{
	blnDebug := false
	
	MouseGetPos, , , strWinId
	WinGetClass strClass, % "ahk_id " . strWinId

	if (blnDebug)
		if WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or (strClass = "#32770")
			###_D(strClass . " is OK")
		else
			###_D(strClass . " is NOT OK")
		
	if WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or (strClass = "#32770")
		return true
	else
	{
		if (blnDebug)
			###_D("Check if dialog title supported: " . strWindowTitle)
		return DialogIsSupported(strWinId)
	}
	blnDebug := false
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsAnExplorer(strClass)
;------------------------------------------------------------
{
	blnDebug := false
	if strClass in CabinetWClass,ConsoleWindowClass
		return True
	else
		return False
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDesktop(strClass)
;------------------------------------------------------------
{
	blnDebug := false
	if strClass in ProgMan,WorkerW
		return True
	else
		return False
}
;------------------------------------------------------------


;------------------------------------------------------------
DialogIsSupported(strWinId)
;------------------------------------------------------------
{
	blnDebug := false
	
	global arrDialogs
	
	WinGetTitle, strDialogTitle, ahk_id %strWinId%
	if (blnDebug)
		loop, % arrDialogs.MaxIndex()
			###_D("DialogIsSupported? " . strDialogTitle . " " . arrDialogs[A_Index] . " " InStr(strDialogTitle, arrDialogs[A_Index]))
	loop, % arrDialogs.MaxIndex()
		if InStr(strDialogTitle, arrDialogs[A_Index])
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


