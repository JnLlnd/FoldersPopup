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
	- support MS Office dialog boxes on WinXP (bosa_sdm_), open special folders from desktop, NavigateDialog in progress

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
blnDebug := true

if (blnDebug)
	###_D("Yes: " . strWinId .  " " . strClass)

if ((strClass = "#32770") or InStr(strClass, "bosa_sdm_"))
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
	NavigateDialog(strPath, strWinId)

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
	intSpecialFolder := 3
else if (A_ThisMenuItem = "Recycle Bin")
	intSpecialFolder := 10
else if(A_ThisMenuItem = "My Computer")
	intSpecialFolder := 17

if WindowIsDesktop(strClass)
	ComObjCreate("Shell.Application").Explore(intSpecialFolder)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else 
	NavigateExplorer(intSpecialFolder, strWinId)
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


;------------------------------------------------------------
NavigateDialog(strPath, strWinId)
;------------------------------------------------------------
/*
Excerpt from RMApp_Explorer_Navigate(FullPath, strWinId="") by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{
    if (strClass = "#32770")
	{
        if ControlIsVisible("ahk_id " . strWinId, "Edit1")
            Control := "Edit1" ; in standard dialog windows, "Edit1" control is the right choice
        Else if ControlIsVisible("ahk_id " . strWinId, "Edit2")
            Control := "Edit2" ; but sometimes in MS office, if condition above fails, "Edit2" control is the right choice 
        Else ; if above fails - just return and do nothing.
            ###_D("strClass is #32770 but no valid control is visible")
    }
    Else if InStr(strClass, "bosa_sdm_") ; for some MS office dialog windows, which are not #32770 class.
    {
        if ControlIsVisible("ahk_id " . strWinId, "Edit1")
            Control := "Edit1" ; if "Edit1" control exists, it is the right choice
        Else if ControlIsVisible("ahk_id " . strWinId, "RichEdit20W2")
            Control := "RichEdit20W2" ; some MS office dialogs don't have "Edit1" control, but they have "RichEdit20W2" control, which is then the right choice.
        Else                            ; if above fails - just return and do nothing.
            Return
    }
    Else {  ; in all other cases, we'll explore FolderPath, and return from this function
        ComObjCreate("Shell.Application").Explore(FolderPath)   ; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
        Return
    }

    ;=== Refine ShellSpecialFolderConstant ===
    if FolderPath is integer
    {
        if (FolderPath = 17)            ; My Computer --> 17 or 0x11
            FolderPath := "::{20d04fe0-3aea-1069-a2d8-08002b30309d}"    ; because you can't navigate to "17" but you can navigate to "::{20d04fe0-3aea-1069-a2d8-08002b30309d}"
        else                            ; don't allow other ShellSpecialFolderConstants. For example - you can't navigate to Control panel while you're in standard "Open File" dialog box window.
            return
    }

    /*
    ShellSpecialFolderConstants:    http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx
    CSIDL:                          http://msdn.microsoft.com/en-us/library/windows/desktop/bb762494%28v=vs.85%29.aspx
    KNOWNFOLDERID:                  http://msdn.microsoft.com/en-us/library/windows/desktop/dd378457%28v=vs.85%29.aspx
    */

    
    ;===In this part (if we reached it), we'll send FolderPath to control and optionaly restore control's initial text after navigating to specified folder===  
    if (RestoreInitText = 1)    ; if we want to restore control's initial text after navigating to specified folder
        ControlGetText, InitControlText, %Control%, ahk_id %strWinId%   ; we'll get and store control's initial text first
    
    RMApp_ControlSetTextR(Control, FolderPath, "ahk_id " strWinId)  ; set control's text to FolderPath
    RMApp_ControlSetFocusR(Control, "ahk_id " strWinId)             ; focus control
    if (WinExist("A") != strWinId)          ; in case that some window just popped out, and initialy active window lost focus
        WinActivate, ahk_id %strWinId%      ; we'll activate initialy active window
    
    ;=== Avoid accidental hotkey & hotstring triggereing while doing SendInput - can be done simply by #UseHook, but do it if user doesn't have #UseHook in the script ===
    If (A_IsSuspended = 1)
        WasSuspended := 1
    if (WasSuspended != 1)
        Suspend, On
    SendInput, {End}{Space}{Backspace}{enter}   ; silly but necessary part - go to end of control, send dummy space, delete it, and then send enter
    if (WasSuspended != 1)
        Suspend, Off

    /*
    Question: Why not use ControlSetText, and then send enter to control via ControlSend, %Control%, {enter}, ahk_id %strWinId% ?
    Because in some "Save as"  dialogs in some programs, this causes auto saving file instead of navigating to specified folder! After a lot of testing, I concluded that most reliable method, which works and prevents this, is the one that looks weird & silly; after setting text via ControlSetText, control must be focused, then some dummy text must be sent to it via SendInput (in this case space, and then backspace which deletes it), and then enter, which causes navigation to specified folder.
    Question: Ok, but is "SendInput, {End}{Space}{Backspace}{enter}" really necessary? Isn't "SendInput, {enter}" sufficient?
    No. Sending "{End}{Space}{Backspace}{enter}" is definitely more reliable then just "{enter}". Sounds silly but tests showed that it's true.
    */
    
    if (RestoreInitText = 1) {  ; if we want to restore control's initial text after we navigated to specified folder
        Sleep, 70               ; give some time to control after sending {enter} to it
        ControlGetText, ControlTextAfterNavigation, %Control%, ahk_id %strWinId%    ; sometimes controls automatically restore their initial text
        if (ControlTextAfterNavigation != InitControlText)                      ; if not
            RMApp_ControlSetTextR(Control, InitControlText, "ahk_id " strWinId)     ; we'll set control's text to its initial text
    }
    if (WinExist("A") != strWinId)  ; sometimes initialy active window loses focus, so we'll activate it again
        WinActivate, ahk_id %strWinId%
    
    if (FocusedControl != "" and ControlIsVisible("ahk_id " strWinId, FocusedControl) = 1)
        RMApp_ControlSetFocusR(FocusedControl, "ahk_id " strWinId)              ; focus initialy focused control
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

RMApp_ControlSetFocusR(Control, WinTitle="", Tries=3)
/*
Adapted from RMApp_ControlSetFocusR(Control, WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{ ; used in Navigator. More reliable ControlSetFocus
    Loop, %Tries%
    {
        ControlFocus, %Control%, %WinTitle%             ; focus control
        Sleep, 50
        ControlGetFocus, FocusedControl, %WinTitle%     ; check
        if (FocusedControl = Control)                   ; if OK
            return 1
    }
}

RMApp_ControlSetTextR(Control, NewText="", WinTitle="", Tries=3)
/*
Adapted from from RMApp_ControlSetTextR(Control, NewText="", WinTitle="", Tries=3) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
{  ; used in Navigator. More reliable ControlSetText
    Loop, %Tries%
    {
        ControlSetText, %Control%, %NewText%, %WinTitle%            ; set
        Sleep, 50
        ControlGetText, CurControlText, %Control%, %WinTitle%       ; check
        if (CurControlText = NewText)                               ; if OK
            return 1
    }
}
 