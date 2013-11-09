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
	- support MS Office dialog boxes on WinXP (bosa_sdm_), open special folders in explorers
	- NavigateDialog, add Desktop, Document and Pictures special folders, open these special menus in dialog boxes, enabling/disabling the appropriate menus in dialog boxes or explorers

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

global arrGlobalFolders := Object()
global arrGlogalDialogs := Object()
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
+MButton::
;------------------------------------------------------------
blnDebug := false

MouseGetPos, , , strGlobalWinId
WinGetClass strGlobalClass, % "ahk_id " . strGlobalWinId

; In case it was disabled while in a dialog box
Menu, menuSpecialFolders, Enable, My Computer
Menu, menuSpecialFolders, Enable, Network Neighborhood
Menu, menuSpecialFolders, Enable, Control Panel
Menu, menuSpecialFolders, Enable, Recycle Bin

; We won't detect the current folder, so we can't add it
Menu, menuFolders, Disable, &Add This Folder
Menu, menuFolders, Show

blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
#If, CanOpenFavorite(strGlobalWinId, strGlobalClass)
MButton::
;------------------------------------------------------------
blnDebug := true

if (blnDebug)
	###_D("Yes: " . strGlobalWinId .  " " . strGlobalClass)

; Can't find how to open a dialog box in My Computer or Network Neighborhood... need help ???
Menu, menuSpecialFolders
	, % ((strGlobalClass = "#32770") or InStr(strGlobalClass, "bosa_sdm_")) ? "Disable" : "Enable"
	, My Computer
Menu, menuSpecialFolders
	, % ((strGlobalClass = "#32770") or InStr(strGlobalClass, "bosa_sdm_")) ? "Disable" : "Enable"
	, Network Neighborhood

; There is no reason to open a dialog box in Control Panel or Recycle Bin
Menu, menuSpecialFolders
	, % ((strGlobalClass = "#32770") or InStr(strGlobalClass, "bosa_sdm_")) ? "Disable" : "Enable"
	, Control Panel
Menu, menuSpecialFolders
	, % ((strGlobalClass = "#32770") or InStr(strGlobalClass, "bosa_sdm_")) ? "Disable" : "Enable"
	, Recycle Bin
/*
if ((strGlobalClass = "#32770") or InStr(strGlobalClass, "bosa_sdm_"))
{
	; can't find how to open a dialog box in My Computer or Network Neighborhood... need help ???
	Menu, menuSpecialFolders, Disable, My Computer
	Menu, menuSpecialFolders, Disable, Network Neighborhood
	; there is no reason to open a dialog box in Control Panel or Recycle Bin
	Menu, menuSpecialFolders, Disable, Control Panel
	Menu, menuSpecialFolders, Disable, Recycle Bin
}
*/

WinActivate, % "ahk_id " . strGlobalWinId
if (WindowIsAnExplorer(strGlobalClass) or WindowIsDesktop(strGlobalClass) or DialogIsSupported(strGlobalWinId))
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
Menu, menuSpecialFolders, Add, Desktop, OpenSpecialFolder
Menu, menuSpecialFolders, Add, Documents, OpenSpecialFolder
Menu, menuSpecialFolders, Add, Pictures, OpenSpecialFolder
Menu, menuSpecialFolders, Add
Menu, menuSpecialFolders, Add, My Computer, OpenSpecialFolder
Menu, menuSpecialFolders, Add, Network Neighborhood, OpenSpecialFolder
Menu, menuSpecialFolders, Add
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
	###_D("strGlobalWinId: " . strGlobalWinId . "`nstrGlobalClass: " . strGlobalClass . "`nstrPath: " . strPath)

if (A_ThisHotkey = "+MButton") or WindowIsDesktop(strGlobalClass)
	ComObjCreate("Shell.Application").Explore(strPath)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strGlobalClass)
	NavigateExplorer(strPath, strGlobalWinId)
else
	NavigateDialog(strPath, strGlobalWinId, strGlobalClass)
blnDebug := false
return
;------------------------------------------------------------


;------------------------------------------------------------
OpenSpecialFolder:
;------------------------------------------------------------

blnDebug := true
if (blnDebug)
	###_D("strGlobalWinId: " . strGlobalWinId . "`nstrGlobalClass: " . strGlobalClass . "`nA_ThisMenuItem: " . A_ThisMenuItem)
; ShellSpecialFolderConstants: http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx
if (A_ThisMenuItem = "Desktop")
	intSpecialFolder := 0
else if (A_ThisMenuItem = "Control Panel")
	intSpecialFolder := 3
else if (A_ThisMenuItem = "Documents")
	intSpecialFolder := 5
else if (A_ThisMenuItem = "Recycle Bin")
	intSpecialFolder := 10
else if (A_ThisMenuItem = "My Computer")
	intSpecialFolder := 17
else if (A_ThisMenuItem = "Network Neighborhood")
	intSpecialFolder := 18
else if (A_ThisMenuItem = "Pictures")
	intSpecialFolder := 39

if (A_ThisHotkey = "+MButton") or WindowIsDesktop(strGlobalClass)
	ComObjCreate("Shell.Application").Explore(intSpecialFolder)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strGlobalClass)
	NavigateExplorer(intSpecialFolder, strGlobalWinId)
else ; this is a dialog box
{
	if (intSpecialFolder = 0)
		strPath := A_Desktop
	else if (intSpecialFolder = 5)
		strPath := A_MyDocuments
	else if (intSpecialFolder = 39)
		StringReplace, strPath, A_MyDocuments, Documents, Pictures
	else ; we do not support this special folder
		return
	NavigateDialog(strPath, strGlobalWinId, strGlobalClass)
}
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

	if (blnDebug)
		Loop, % arrGlobalFolders.MaxIndex()
			###_D(strMenu . " = " . arrGlobalFolders[A_Index].Name)
	Loop, % arrGlobalFolders.MaxIndex()
		if (strMenu = arrGlobalFolders[A_Index].Name)
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
; "ConsoleWindowClass" -> Console
; "#32770" -> Dialog
{
	blnDebug := false
	
	MouseGetPos, , , strWinId
	WinGetClass strClass, % "ahk_id " . strWinId

	if (blnDebug)
		if WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or (strClass = "#32770") or InStr(strClass, "bosa_sdm_")
			###_D(strClass . " is OK")
		else
			###_D(strClass . " is NOT OK")
		
	return WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or (strClass = "#32770") or InStr(strClass, "bosa_sdm_")
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
NavigateDialog(strPath, strWinId, strClass)
;------------------------------------------------------------
/*
Excerpt from RMApp_Explorer_Navigate(FullPath, strWinId="") by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{
	blnDebug := true
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
	Else if InStr(strClass, "bosa_sdm_") ; for some MS office dialog windows, which are not #32770 class.
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
	if (WinExist("A") != strWinId) ; in case that some window just popped out, and initialy active window lost focus
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
	
	;=== Avoid accidental hotkey & hotstring triggereing while doing SendInput - can be done simply by #UseHook, but do it if user doesn't have #UseHook in the script ===
	If (A_IsSuspended)
		blnWasSuspended := True
	if (!blnWasSuspended)
		Suspend, On
	SendInput, {End}{Space}{Backspace}{enter} ; silly but necessary part - go to end of control, send dummy space, delete it, and then send enter
	if (!blnWasSuspended)
		Suspend, Off

	Sleep, 70 ; give some time to control after sending {enter} to it
	ControlGetText, strControlTextAfterNavigation, %strControl%, ahk_id %strWinId% ; sometimes controls automatically restore their initial text
	if (strControlTextAfterNavigation <> strPrevControlText) ; if not
		ControlSetTextR(strControl, strPrevControlText, "ahk_id " . strWinId) ; we'll set control's text to its initial text
	
	if (WinExist("A") <> strWinId) ; sometimes initialy active window loses focus, so we'll activate it again
		WinActivate, ahk_id %strWinId%
}


ControlIsVisible(WinTitle,ControlClass)
/*
Adapted from ControlIsVisible(WinTitle,ControlClass) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{ ; used in Navigator
	ControlGet, IsControlVisible, Visible,, %ControlClass%, %WinTitle%
	return IsControlVisible
}

ControlSetFocusR(Control, WinTitle="", Tries=3)
/*
Adapted from RMApp_ControlSetFocusR(Control, WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{ ; used in Navigator. More reliable ControlSetFocus
	Loop, %Tries%
	{
		ControlFocus, %Control%, %WinTitle%			 ; focus control
		Sleep, 50
		ControlGetFocus, FocusedControl, %WinTitle%	 ; check
		if (FocusedControl = Control)				   ; if OK
			return 1
	}
}

ControlSetTextR(Control, NewText="", WinTitle="", Tries=3)
/*
Adapted from from RMApp_ControlSetTextR(Control, NewText="", WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{  ; used in Navigator. More reliable ControlSetText
	Loop, %Tries%
	{
		ControlSetText, %Control%, %NewText%, %WinTitle%			; set
		Sleep, 50
		ControlGetText, CurControlText, %Control%, %WinTitle%	   ; check
		if (CurControlText = NewText)							   ; if OK
			return 1
	}
}
