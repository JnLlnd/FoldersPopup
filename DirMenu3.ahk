;===============================================
/*
DirMenu3
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By Jean Lalonde (JnLlnd on AHKScript.org forum), based on DirMenu v2 by Robert Ryan (rbrtryn on AutoHotkey.com forum)

	Version 0.1 ALPHA
	- init skeleton, read ini file and create arrays for folders menu and supported dialog boxes
	- create language file, build gui, tray menu and popup menu, skeleton for front end buttons and commands

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
Gosub, BuiltPopupMenu

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
			Dialog1=Insert
			Dialog2=Open
			Dialog3=Save
			Dialog4=Select

			[Folders]
			Folder1=C:\|C:\
			Folder2=Windows|C:\Windows
			Folder3=Program Files|C:\Program Files
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
blnDebug := true

Menu, menuFavorites, Disable, &Add This Folder
Menu, menuFavorites, Disable, Add This &Dialog
Menu, menuFavorites, Show

blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
#If, CanOpenFavorite(WinUnderMouseClass())
MButton::
;------------------------------------------------------------
blnDebug := true

if (blnDebug)
	###_D("MButton::`nClass under mouse -> " . WinUnderMouseClass() . "`nWinID -> " . WinUnderMouseID())
WinActivate, % "ahk_id " . WinUnderMouseID()
Menu, menuFavorites, Show

blnDebug := false
return
#If
;------------------------------------------------------------


;------------------------------------------------------------
OpenFavorite:
;------------------------------------------------------------
blnDebug := true

strPath := GetPahtFor(A_ThisMenu)

if (A_ThisHotkey = "+MButton")
	Run Explorer.exe /n`,%strPath%
else
{
	WinWaitActive, ahk_id %WinId%
	; Call[WinGetClass("A")].(Path)
}
blnDebug := false
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
Gosub, BuiltPopupMenu
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
BuildTrayMenu:
;------------------------------------------------------------
Menu, Tray, Add
Menu, Tray, Add, A&bout, GuiAbout
Menu, Tray, Add, &Edit Folders Menu, ShowGui
Menu, Tray, Default, &Edit Folders Menu
return
;------------------------------------------------------------


;------------------------------------------------------------
ShowGui:
;------------------------------------------------------------
Gosub, LoadSettings
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
LoadSettings:
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
BuiltPopupMenu:
;------------------------------------------------------------
blnDebug := False

Menu, menuFavorites, Add
Menu, menuFavorites, DeleteAll
Loop, % arrFolders.MaxIndex()
{
	if (blnDebug)
		###_D(arrFolders[A_Index].Name)
	if (arrFolders[A_Index].Name = lMenuSeparator)
		Menu menuFavorites, Add
	else
		Menu menuFavorites, Add, % arrFolders[A_Index].Name, OpenFavorite
}
Menu menuFavorites, Add
Menu menuFavorites, Add, &Edit This Menu, ShowGui
Menu menuFavorites, Default, &Edit This Menu
Menu menuFavorites, Add, &Add This Folder, AddThisFolder
Menu menuFavorites, Add, Add This &Dialog, AddThisDialog

if (blnDebug)
	Menu menuFavorites, Show

blnDebug := False
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
	blnDebug := true
	
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
CanOpenFavorite(strClass)
;------------------------------------------------------------
; "CabinetWClass" -> Explorer
; "ProgMan" -> Desktop
; "WorkerW" -> Desktop
; "ConsoleWindowClass" -> Console
; "#32770" -> Dialog
{
	blnDebug := false
	if (blnDebug)
		if strClass in CabinetWClass,ProgMan,WorkerW,ConsoleWindowClass,#32770
			###_D(strClass . " is OK")
		else
			###_D(strClass . " is NOT OK")
	if strClass in CabinetWClass,ProgMan,WorkerW,ConsoleWindowClass,#32770
		return true
	else
		return false
	blnDebug := false
}
;------------------------------------------------------------
