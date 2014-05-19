/*
BUG

TODO
- remove the More submenu
- keep the Switch submenus in a "Switch" submenu
- add donate button in Settings and Options windows

*/

;===============================================
/*
	FoldersPopup
	Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
	By Jean Lalonde (JnLlnd on AHKScript.org forum), based on DirMenu v2 by Robert Ryan (rbrtryn on AutoHotkey.com forum)

	Version: v1.9 BETA (not to be released) (2014-05-nn)
	* fix strange bug with Special folders menu disappearing after Recent or Swith menus refresh because of DeleteAll command
	* fix bug missing error message and other language minor changes
	* reorder popup menu and place settings, add this folder and support freeware menus at the end of main menu
	* reorder checkboxes in GuiOptions

	Version: v1.8 ALPHA (not to be released) (2014-05-04)
	* add switch in dialog box to other explorer windows already opened
	* lMenuReservedShortcuts management with translations
	* sort folders button
	* folder up button
	* translated help to French
	* support freeware to popup menu
	* blnMenuReady before popup

	Version: v1.7 ALPHA (not to be released) (2014-04-27)
	* new settings dialog box layout with icons to add, edit or remove folders or dialog boxes
	* icons to open help, about and settings dialog boxes
	* dropdown to select the submenu to edit
	* left arrow to go back to edit the menu(s) previously displayed
	* double-click to edit folders or supported dialog boxes
	* adjustments to dialog boxes for German and French translation
	* updated about and help dialog boxes
	* solved a bug when Add this folder in some type of dialog boxes
	
	Version: v1.6 ALPHA (not to be released) (2014-04-19)
	* implement submenus ini file data structure and objects for folders
	* v1 ini file format automatic upgrade to v2 (all v1 folders placed in main menu)
	* load and save folders and submenus to ini file
	* display popup menus with submenus, disable empty submenus
	* add a dropdown menu to settings window to select the menu to edit
	* settings pugrade to add, edit, remove, move up or down submenus
	* add folder to the current submenu
	* add folder from popup menu to main menu
	* move folders or menus to other submenus
	* double-click an folder or submenu item in settings to edit it
	* update popup menus as settings are changed, backup available if user cancel settings changes
	* support numeric shortcuts for submenus
	* error checking: avoid duplicate names when moving an item to another submenu
	* error checking: avoid moving a submenu under itself

	Version: v1.5 ALPHA (not to be released) (2014-03-22)
	* add recent folders sub-menu
	* add ini variable RecentFolders
	* when blnDisplayMenuShortcuts reserve shortcut chars for app's items in menus
	* add GetDeepestFolderName as function
	* add ValueIsInObject function
	* add language dropdown
	* display full folder names in recent folders
	* add swith submenu to activate any other open Explorer
	* add DisplayRecentFolders and DisplaySwitchMenu options in Options dialog box and ini file

	Version: FoldersPopup v1.2.7 (2014-04-25)
	* Workaround to make the "Run Explorer" command work in rare configuration
	* Fix a bug in the check for update command

	Version: FoldersPopup v1.2.6 (2014-04-24)
	* Workaround for the hash (aka Sharp / "#") bug in Shell.Application that occurs only when navigatin in the current Explorer window to a subfolder including # in its parent path (eg.: C:\C#\Project)
	* Windows XP only: fix a bug when navigating to special folder "My Pictures" in dialog boxes

	Version: FoldersPopup v1.2.5 (2014-04-19)
	* Support for FreeCommander XE
	* Compatible with Clover (opens the folder in a new tab)
	* Fix wrong error message issue #28

	Version: FoldersPopup v1.2.4 (2014-04-17)
	* Fix shortcut (hotkey) assignements error (not a valid key namse error) on Windows system with keyboard regional settings supporting Cyrillic letters (Russian and others)
	
	Version: FoldersPopup v1.2.3 (2014-02-25)
	* Windows XP only: revert to pre-1.2.2 state due to a different behaviour of Winddows Explorer in XP

	Version: FoldersPopup v1.2.2 (2014-02-20)
	* opens new Explorer windows complying with the Explorer navigation pane setting

	Version: FoldersPopup v1.2.1 (2014-02-01)
	* fix a bug that added separator lines at the bottom of Tray Menu (one line added at each display of the popu menu)
	* improve diagnostic data collection (always at the user's discretion)

	Version: FoldersPopup v1.2 (2014-01-26)
	* add an option to add numeric keyboard shortcuts to launch folders in popup menu
	* add an option to display the popup menu at a fix position
	* add a diagnostic mode to collect support info (add DiagMode=1 under [Global] section in ini file)
	* redesign of the Options dialog box

	Version: FoldersPopup v1.01 (2013-12-24)
	* bug fix: mouse and keyboard triggers were disabled in non-explorer windows

	Version: FoldersPopup v1.0 (First official release, 2013-12-23)
	* configurable mouse button and keyboard triggers in a new "Options" dialog box
	* new keyboard triggers (by default, Windows-K and Shift-Windows-K) in addition to mouse button triggers (by default, Middle mouse and Shift-Middle mouse buttons)
	* add "Run at startup" checkbox to "Options" dialog box to launch Folders Popup automatically at Windows startup
	* add "Display the startup tray tip" checkbox to "Options" dialog box to display or hide the Folders popup's tray tip
	* add "Display Special Folders" checkbox to "Options" dialog box to enable/disable navigation to special folders (My Computer, Network, Recycle bion, etc.) in popup menu
	* better formated startup help tray tip
	* close "Settings" dialog box with Escape key

	Version: FoldersPopup v0.9 BETA (2013-11-11)
	* implemented startup option in tray and check4update
	* removed debugging code, prepare for compiler, removed external pictures
	* standardize dialog box titles, various text fixes
	* renamed the app FoldersPopup

	Version: PopupFolders v0.5 ALPHA (last alpha version, 2013-11-11)
	* implemented GuiAbout and GuiHelp, added About and Help to tray menu, tray tip displayed only 5 times
	* removed file:/// protocol prefix, added support for ExploreWClass, implemented try/catch to Explore shell method, offer to add manually when add folder failed

	Version: PopupFolders v0.4 ALPHA (2013-11-11)
	* add settings hotkey to ini file (default Crtl-Windows-F), enable AddThisFolder in all version Explorer and only in WIN_7/Win_8 dialog boxes (not working in WIN_XP)
	* add GuiSave, GuiCancel, RemoveFolder, EditFolder, AddSeparator, MoveFolderUp/Down, RemoveDialog, EditDialog, fix bug in GuiShow, add tray icon

	Version: PopupFolders v0.3 ALPHA (2013-11-10)
	* add NavigateConsole for console support (command prompt CMD)
	* change .ini filename to new app name

	Version: PopupFolders v0.2 ALPHA (2013-11-09)
	* renamed app PopupFolders, isolate text into language variables

	Version: DirMenu3 v0.1 ALPHA (2013-11-05)
	* init skeleton, read ini file and create arrays for folders menu and supported dialog boxes
	* create language file, build gui, tray menu and folder menu, skeleton for front end buttons and commands
	* create AddThisDialog menu, MButton condition, CanOpenFavorite improvements with WindowIsAnExplorer, WindowIsDesktop and DialogIsSupported
	* add SpecialFolders menu, OpenFavorite for Explorer and Desktop, NavigateExplorer
	* support MS Office dialog boxes on WinXP (bosa_sdm_), open special folders in explorers
	* NavigateDialog, add Desktop, Document and Pictures special folders, open these special menus in dialog boxes, enabling/disabling the appropriate menus in dialog boxes or explorers

	Version: DirMenu v2.2 (never released / not stable - base of a total rewrite to DirMenu3)
	* manage (add, modify or delete) supported dialog box titles in the Gui
	* suggest current dialog box when adding a name
	* save the supported dialog box names on the first line of the settings file (dirmenu.txt)
	* add "Add This Dialog" to the MButton menu to add the current dialog box name (need to desactivate when in an already supported dialog box)
	* added Win8 to the list of supported versions (assumed as equal to Win7 - could not test myself)
	* removed the "Menu File" button because not needed anymore
	* fixed an issue when 2 folders had the same name (now preventing the use of an existing name)
	* change default setting filename to "DirMenu2.txt" to avoid upgrade errors
	* upgrade previous versions settings files to v2.2
	* ask confirmation before discarding changes with Revert or Cancel buttons
	* replaces RegEx on strDialogNames with DialogIsSupported() function on the ListView
	
	Version: DirMenu v2.1
	* make it work with any locale (still working with English)
	* put supported dialog box titles in a variable (strDialogNames) at the top of the script for easy editing
	* put DirMenu data file name in a variable (strDirMenuFile) at the top of the script for easy editing
	* add "Add This Folder" to the MButton menu to add the current folder
	* add "Menu File" button to open de DirMenu.txt file for edition in Notepad
	* propose the deepest folder name as default name for a new folder

*/ 
;===============================================

; --- COMPILER DIRECTIVES ---

; Doc: http://fincs.ahk4.net/Ahk2ExeDirectives.htm
; Note: prefix comma with `

;@Ahk2Exe-SetName FoldersPopup
;@Ahk2Exe-SetDescription Popup menu to jump instantly from one folder to another. Freeware.
;@Ahk2Exe-SetVersion 1.9.2 BETA
;@Ahk2Exe-SetOrigFilename FoldersPopup.exe


;============================================================
; INITIALIZATION
;============================================================

#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off
SetWorkingDir, %A_ScriptDir%

strTempDir := "_temp"
FileCreateDir, %strTempDir%
FileInstall, FileInstall\FoldersPopup_LANG_DE.txt, %strTempDir%\FoldersPopup_LANG_DE.txt, 1
FileInstall, FileInstall\FoldersPopup_LANG_FR.txt, %strTempDir%\FoldersPopup_LANG_FR.txt, 1
FileInstall, FileInstall\about-32.png, %strTempDir%\about-32.png
FileInstall, FileInstall\add_image-26.png, %strTempDir%\add_image-26.png
FileInstall, FileInstall\add_property-48.png, %strTempDir%\add_property-48.png
FileInstall, FileInstall\delete_property-48.png, %strTempDir%\delete_property-48.png
FileInstall, FileInstall\separator-26.png, %strTempDir%\separator-26.png
FileInstall, FileInstall\down_circular-26.png, %strTempDir%\down_circular-26.png
FileInstall, FileInstall\edit_image-26.png, %strTempDir%\edit_image-26.png
FileInstall, FileInstall\edit_property-48.png, %strTempDir%\edit_property-48.png
FileInstall, FileInstall\generic_sorting2-26-grey.png, %strTempDir%\generic_sorting2-26-grey.png
FileInstall, FileInstall\help-32.png, %strTempDir%\help-32.png
FileInstall, FileInstall\left-12.png, %strTempDir%\left-12.png
FileInstall, FileInstall\remove_image-26.png, %strTempDir%\remove_image-26.png
FileInstall, FileInstall\settings-32.png, %strTempDir%\settings-32.png
FileInstall, FileInstall\up-12.png, %strTempDir%\up-12.png
FileInstall, FileInstall\up_circular-26.png, %strTempDir%\up_circular-26.png

Gosub, InitLanguageVariables

global strAppName := "FoldersPopup"
global strCurrentVersion := "1.9.2 BETA" ; "major.minor.bugs"
global strAppVersion := "v" . strCurrentVersion
global blnDiagMode := False
global strDiagFile := A_ScriptDir . "\" . strAppName . "-DIAG.txt"
global strIniFile := A_ScriptDir . "\" . strAppName . ".ini"
blnMenuReady := false

if InStr(A_ScriptDir, A_Temp) ; must be positioned after strAppName is created
; if the app runs from a zip file, the script directory is created under the system Temp folder
{
	Oops(lOopsZipFileError, strAppName)
	Gosub, CleanUpBeforeExit
}
	
;@Ahk2Exe-IgnoreBegin
; Piece of code for developement phase only - won't be compiled
; http://fincs.ahk4.net/Ahk2ExeDirectives.htm
if (A_ComputerName = "JEAN-PC") ; for my home PC
	strIniFile := A_ScriptDir . "\" . strAppName . "-HOME.ini"
else if InStr(A_ComputerName, "STIC") ; for my work hotkeys
	strIniFile := A_ScriptDir . "\" . strAppName . "-WORK.ini"
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd
global arrSubmenuStack := Object()

; Keep gosubs in this order
Gosub, InitSystemArrays
Gosub, InitLanguage
Gosub, LoadIniFile
Gosub, LoadTheme
; build even if blnDisplaySpecialFolders, blnDisplayRecentFolders or blnDisplaySwitchMenu are false because they could become true
Gosub, BuildSpecialFoldersMenu
Gosub, BuildRecentFoldersMenu
Gosub, BuildSwitchMenu
Gosub, BuildFoldersMenus ; need to be initialized here - will be updated at each call to popup menu
Gosub, BuildGui
Gosub, BuildAddDialogMenu
Gosub, Check4Update
Gosub, BuildTrayMenu
if (blnDiagMode)
	Gosub, InitDiagMode

IfExist, %A_Startup%\%strAppName%.lnk
{
	FileDelete, %A_Startup%\%strAppName%.lnk
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%strAppName%.lnk
	Menu, Tray, Check, %lMenuRunAtStartup%
}

if (blnDisplayTrayTip)
	TrayTip, % L(lTrayTipInstalledTitle, strAppName, strAppVersion)
		, % L(lTrayTipInstalledDetail, strAppName
			, Hotkey2Text(strModifiers1, strMouseButton1, strOptionsKey1)
			, Hotkey2Text(strModifiers3, strMouseButton3, strOptionsKey3)
			, Hotkey2Text(strModifiers2, strMouseButton2, strOptionsKey2)
			, Hotkey2Text(strModifiers4, strMouseButton4, strOptionsKey4))
		, , 1

OnExit, CleanUpBeforeExit

blnMenuReady := true

; gosub, GuiShow ; ### only when debugging Gui

return

#Include %A_ScriptDir%\FoldersPopup_LANG.ahk


;============================================================
; BACK END FUNCTIONS AND COMMANDS
;============================================================


;-----------------------------------------------------------
InitSystemArrays:
;-----------------------------------------------------------

; Hotkeys: ini names, hotkey variables name, default values, gosub label and Gui hotkey titles
strIniKeyNames := "PopupHotkeyMouse|PopupHotkeyNewMouse|PopupHotkeyKeyboard|PopupHotkeyNewKeyboard|SettingsHotkey"
StringSplit, arrIniKeyNames, strIniKeyNames, |
strHotkeyVarNames := "strPopupHotkeyMouse|strPopupHotkeyMouseNew|strPopupHotkeyKeyboard|strPopupHotkeyKeyboardNew|strSettingsHotkey"
StringSplit, arrHotkeyVarNames, strHotkeyVarNames, |
strHotkeyDefaults := "MButton|+MButton|#k|+#k|+#f"
StringSplit, arrHotkeyDefaults, strHotkeyDefaults, |
strHotkeyLabels := "PopupMenuMouse|PopupMenuNewWindowMouse|PopupMenuKeyboard|PopupMenuNewWindowKeyboard|GuiShow"
StringSplit, arrHotkeyLabels, strHotkeyLabels, |

strMouseButtons := " |LButton|MButton|RButton|XButton1|XButton2|WheelUp|WheelDown|WheelLeft|WheelRight|"
; leave last | to enable default value on the last item
StringSplit, arrMouseButtons, strMouseButtons, |
strMouseButtonsText := " |Left Mouse Button|Middle Mouse Button|Right Mouse Button|X Button1|X Button2|Wheel Up|Wheel Down|Wheel Left|Wheel Right|"
StringSplit, arrMouseButtonsText, strMouseButtonsText, |

return
;-----------------------------------------------------------


;-----------------------------------------------------------
LoadIniFile:
;-----------------------------------------------------------

; reinit after Settings save if already exist
Global arrMenus := Object() ; list of menus (Main and submenus), non-hierachical
arrMainMenu := Object() ; array of folders of the Main menu
arrMenus.Insert(lMainMenuName, arrMainMenu) ; lMainMenuName is used in the objects but not saved/restores in the ini file
arrDialogs := Object() ; list of supported dialog boxes

strPopupHotkeyMouseDefault := arrHotkeyDefaults1 ; "MButton"
strPopupHotkeyMouseNewDefault := arrHotkeyDefaults2 ; "+MButton"
strPopupHotkeyKeyboardDefault := arrHotkeyDefaults3 ; "#k"
strPopupHotkeyKeyboardNewDefault := arrHotkeyDefaults4 ; "+#k"
strSettingsHotkeyDefault := arrHotkeyDefaults5 ; "+#f"

IfNotExist, %strIniFile%
	FileAppend,
		(LTrim Join`r`n
			[Global]
			PopupHotkeyMouse=%strPopupHotkeyMouseDefault%
			PopupHotkeyNewMouse=%strPopupHotkeyMouseNewDefault%
			PopupHotkeyKeyboard=%strPopupHotkeyKeyboardDefault%
			PopupHotkeyNewKeyboard=%strPopupHotkeyKeyboardNewDefault%
			SettingsHotkey=%strSettingsHotkeyDefault%
			DisplayTrayTip=1
			DisplaySpecialFolders=1
			DisplayRecentFolders=1
			RecentFolders=5
			DisplaySwitchMenu=1
			DisplayMenuShortcuts=0
			PopupFix=0
			PopupFixPosition=20,20
			DiagMode=0
			Startups=1
			LanguageCode=EN
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
	
Gosub, LoadIniHotkeys

IniRead, blnDisplayTrayTip, %strIniFile%, Global, DisplayTrayTip, 1
IniRead, blnDisplaySpecialFolders, %strIniFile%, Global, DisplaySpecialFolders, 1
IniRead, blnDisplayRecentFolders, %strIniFile%, Global, DisplayRecentFolders, 1
IniRead, blnDisplaySwitchMenu, %strIniFile%, Global, DisplaySwitchMenu, 1
IniRead, blnPopupFix, %strIniFile%, Global, PopupFix, 0
IniRead, strPopupFixPosition, %strIniFile%, Global, PopupFixPosition, 20,20
StringSplit, arrPopupFixPosition, strPopupFixPosition, `,
IniRead, blnDisplayMenuShortcuts, %strIniFile%, Global, DisplayMenuShortcuts, 0
IniRead, strLatestSkipped, %strIniFile%, Global, LatestVersionSkipped, 0.0
IniRead, blnDiagMode, %strIniFile%, Global, DiagMode, 0
IniRead, intStartups, %strIniFile%, Global, Startups, 1
IniRead, blnDonator, %strIniFile%, Global, Donator, 0 ; Please, be fair. Don't cheat with this.
if !(blnDonator)
	lMenuReservedShortcuts := lMenuReservedShortcuts . lMenuReservedShortcutsDonate
IniRead, intRecentFolders, %strIniFile%, Global, RecentFolders, 5

Loop
{
	IniRead, strIniLine, %strIniFile%, Folders, Folder%A_Index%
	if (strIniLine = "ERROR")
		Break
	strIniLine := strIniLine . "|||" ; additional "|" to make sure we have all empty items
	; 1 FolderName, 2 FolderLocation, 3 MenuName, 4 SubmenuFullName
	StringSplit, arrThisFolder, strIniLine, |
	
	objFolder := Object() ; new menu item
	objFolder.FolderName := arrThisFolder1 ; display name of this menu item
	objFolder.FolderLocation := arrThisFolder2 ; path for this menu item
	objFolder.MenuName := lMainMenuName . arrThisFolder3 ; parent menu of this menu item, adding main menu name
	if StrLen(arrThisFolder4)
		objFolder.SubmenuFullName := lMainMenuName . arrThisFolder4 ; full name of the submenu, adding main menu name
	else
		objFolder.SubmenuFullName := ""

	if StrLen(objFolder.SubmenuFullName) ; then this is a new submenu
	{
		arrSubMenu := Object() ; create submenu
		arrMenus.Insert(objFolder.SubmenuFullName, arrSubMenu) ; add this submenu to the array of menus
	}
	arrMenus[objFolder.MenuName].Insert(objFolder) ; add this menu item to parent menu
}

Loop
{
	IniRead, strIniLine, %strIniFile%, Dialogs, Dialog%A_Index%
	if (strIniLine = "ERROR")
		Break
	arrDialogs.Insert(strIniLine)
}

IfNotExist, %strIniFile%
{
	Oops(lOopsWriteProtectedError, strAppName)
	Gosub, CleanUpBeforeExit
}

return
;------------------------------------------------------------


;-----------------------------------------------------------
LoadIniHotkeys:
;-----------------------------------------------------------
; Read the values and set hotkey shortcuts
loop, % arrIniKeyNames%0%
{
	; Prepare global arrays used by GuiHotkey function
	IniRead, arrHotkeyVarNames%A_Index%, %strIniFile%, Global, % arrIniKeyNames%A_Index%, % arrHotkeyDefaults%A_Index%
	SplitHotkey(arrHotkeyVarNames%A_Index%, strMouseButtons
		, strModifiers%A_Index%, strOptionsKey%A_Index%, strMouseButton%A_Index%, strMouseButtonsWithDefault%A_Index%)
	; example: Hotkey, $MButton, PopupMenuMouse
	Hotkey, % "$" . arrHotkeyVarNames%A_Index%, % arrHotkeyLabels%A_Index%, On UseErrorLevel
	if (ErrorLevel)
		Oops(lDialogInvalidHotkey, Hotkey2Text(strModifiers%A_Index%, strMouseButton%A_Index%, strOptionsKey%A_Index%), strAppName, arrOptionsTitles%A_Index%)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
InitLanguage:
;------------------------------------------------------------

IniRead, strLanguageCode, %strIniFile%, Global, LanguageCode, EN
strLanguageFile := strTempDir . "\" . strAppName . "_LANG_" . strLanguageCode . ".txt"

if FileExist(strLanguageFile)
{
	FileRead, strLanguageStrings, %strLanguageFile%
	Loop, Parse, strLanguageStrings, `n, `r
	{
		StringSplit, arrLanguageBit, A_LoopField, `t
		if SubStr(arrLanguageBit1, 1, 1) = "l"
			%arrLanguageBit1% := arrLanguageBit2
		StringReplace, %arrLanguageBit1%, %arrLanguageBit1%, ``n, `n, All
	}
}
else
	strLanguageCode := "EN"

; Init Language Arrays
StringSplit, arrOptionsTitles, lOptionsTitles, |
StringSplit, arrOptionsTitlesLong, lOptionsTitlesLong, |
strOptionsLanguageCodes := "EN|FR|DE"
StringSplit, arrOptionsLanguageCodes, strOptionsLanguageCodes, |
StringSplit, arrOptionsLanguageLabels, lOptionsLanguageLabels, |

loop, %arrOptionsLanguageCodes0%
	if (arrOptionsLanguageCodes%A_Index% = strLanguageCode)
		{
			strLanguageLabel := arrOptionsLanguageLabels%A_Index%
			break
		}

return
;------------------------------------------------------------


;------------------------------------------------------------
InitDiagMode:
;------------------------------------------------------------

MsgBox, 52, %strAppName%, % L(lDiagModeCaution, strAppName, strDiagFile)
IfMsgBox, No
{
	blnDiagMode := False
	IniWrite, 0, %strIniFile%, Global, DiagMode
	return
}
	
if !FileExist(strDiagFile)
{
	FileAppend, DateTime`tType`tData`n, %strDiagFile%
	Diag("DIAGNOSTIC FILE", lDiagModeIntro)
	Diag("AppName", strAppName)
	Diag("AppVersion", strAppVersion)
	Diag("A_ScriptFullPath", A_ScriptFullPath)
	Diag("A_AhkVersion", A_AhkVersion)
	Diag("A_OSVersion", A_OSVersion)
	Diag("A_Is64bitOS", A_Is64bitOS)
	Diag("A_Language", A_Language)
	Diag("A_IsAdmin", A_IsAdmin)
}
else
	FileAppend, `n, %strDiagFile% ; required when the last line of the existing file ends with "

FileRead, strDiag, %strIniFile%
StringReplace, strDiag, strDiag, `", `"`"
Diag("IniFile", """" . strDiag . """`n")

return
;------------------------------------------------------------


;-----------------------------------------------------------
CleanUpBeforeExit:
;-----------------------------------------------------------

FileRemoveDir, %strTempDir%, 1 ; Remove all files and subdirectories

if (blnDiagMode)
{
	MsgBox, 52, %strAppName%, % L(lDiagModeExit, strAppName, strDiagFile) . "`n`n" . lDiagModeIntro . "`n`n" . lDiagModeSee
	IfMsgBox, Yes
		Run, %strDiagFile%
}
ExitApp
;-----------------------------------------------------------



;============================================================
; FRONT END FUNCTIONS AND COMMANDS
;============================================================


;------------------------------------------------------------
PopupMenuMouse: ; default MButton
PopupMenuKeyboard: ; default #k
;------------------------------------------------------------

if !(blnMenuReady)
	return

If !CanOpenFavorite(A_ThisLabel, strTargetWinId, strTargetClass)
{
	StringReplace, strThisHotkey, A_ThisHotkey, $ ; remove $ from hotkey
	if (A_ThisLabel = "PopupMenuMouse")
		Send, {%strThisHotkey%} ; for example {MButton}
	else
	{
		StringLower, strThisHotkey, strThisHotkey
		Send, %strThisHotkey% ; for example #k
	}
	return
}

blnMouse := InStr(A_ThisLabel, "Mouse")
blnNewWindow := false
Gosub, SetMenuPosition

; Can't find how to navigate a dialog box to My Computer or Network Neighborhood... is this is feasible?
if (blnDisplaySpecialFolders)
{
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass) or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuMyComputer%
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass) or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuNetworkNeighborhood%

	; There is no point to navigate a dialog box or console to Control Panel or Recycle Bin, unknown for FreeCommander
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass) or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuControlPanel%
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass) or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuRecycleBin%
}

if (blnDisplayRecentFolders)
{
	Gosub, BuildRecentFoldersMenu
	Menu, menuMore
		, % (intRecentFoldersIndex ? "Enable" : "Disable")
		, %lMenuRecentFolders%
}

if (blnDisplaySwitchMenu)
{
	Gosub, BuildSwitchMenu
	Menu, menuMore
		, % (intExplorersIndex ? "Enable" : "Disable")
		, %lMenuSwitchExplorer%
	Menu, menuMore
		, % (intExplorersIndex and DialogIsSupported(strTargetWinId)? "Enable" : "Disable")
		, %lMenuSwitchDialog%
}

if (WindowIsAnExplorer(strTargetClass) or WindowIsDesktop(strTargetClass) or WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass) or DialogIsSupported(strTargetWinId))
{
	; Enable Add This Folder only if the mouse is over an Explorer (tested on WIN_XP and WIN_7) or a dialog box (works on WIN_7, not on WIN_XP)
	; Other tests shown that WIN_8 behaves like WIN_7. So, I assume WIN_8 to work. If someone could confirm (until I can test it myself)?
	if (blnDiagMode)
		Diag("ShowMenu", "Folders Menu " . (WindowIsAnExplorer(strTargetClass) or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "WITH" : "WITHOUT") . " Add this folder")
	Menu, %lMainMenuName%
		, % WindowIsAnExplorer(strTargetClass) or WindowIsFreeCommander(strTargetClass) or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
		, %lMenuAddThisFolder%
	Menu, %lMainMenuName%, Show, %intMenuPosX%, %intMenuPosY% ; mouse pointer if mouse button, 20x20 offset of active window if keyboard shortcut
}
else
{
	if (blnDiagMode)
		Diag("ShowMenu", "Add Dialog")
	Menu, menuAddDialog, Show
}

return
;------------------------------------------------------------


;------------------------------------------------------------
PopupMenuNewWindowMouse: ; default +MButton::
PopupMenuNewWindowKeyboard: ; default +#k
;------------------------------------------------------------

if !(blnMenuReady)
	return

blnMouse := InStr(A_ThisLabel, "Mouse")
blnNewWindow := true
Gosub, SetMenuPosition

WinGetClass strTargetClass, % "ahk_id " . strTargetWinId

if (blnDiagMode)
{
	Diag("MouseOrKeyboard", A_ThisLabel)
	WinGetTitle strDiag, % "ahk_id " . strTargetWinId
	Diag("WinTitle", strDiag)
	Diag("WinId", strTargetWinId)
	Diag("Class", strTargetClass)
	Diag("ShowMenu", "Folders Menu " . (WindowIsAnExplorer(strTargetClass) or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "WITH" : "WITHOUT") . " Add this folder")
}

if (blnDisplaySpecialFolders)
{
	; In case it was disabled while in a dialog box
	Menu, menuSpecialFolders, Enable, %lMenuMyComputer%
	Menu, menuSpecialFolders, Enable, %lMenuNetworkNeighborhood%
	Menu, menuSpecialFolders, Enable, %lMenuControlPanel%
	Menu, menuSpecialFolders, Enable, %lMenuRecycleBin%
}

if (blnDisplayRecentFolders)
{
	Gosub, BuildRecentFoldersMenu
	Menu, menuMore
		, % (intRecentFoldersIndex ? "Enable" : "Disable")
		, %lMenuRecentFolders%
}

if (blnDisplaySwitchMenu)
{
	Gosub, BuildSwitchMenu
	Menu, menuMore
		, % (intExplorersIndex ? "Enable" : "Disable")
		, %lMenuSwitchExplorer%
	Menu, menuMore
		, % (intExplorersIndex and DialogIsSupported(strTargetWinId)? "Enable" : "Disable")
		, %lMenuSwitchDialog%
}

; Enable "Add This Folder" only if the target window is an Explorer (tested on WIN_XP and WIN_7)
; or a dialog box under WIN_7 (does not work under WIN_XP).
; Other tests shown that WIN_8 behaves like WIN_7.
Menu, %lMainMenuName%
	, % WindowIsAnExplorer(strTargetClass) or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
	, %lMenuAddThisFolder%
Menu, %lMainMenuName%, Show, %intMenuPosX%, %intMenuPosY% ; mouse pointer if mouse button, 20x20 offset of active window if keyboard shortcut

return
;------------------------------------------------------------


;------------------------------------------------------------
SetMenuPosition:
;------------------------------------------------------------

CoordMode, Menu, % (blnPopupFix ? "Screen" : "Window")

; will be overridden if not fix location
intMenuPosX := arrPopupFixPosition1
intMenuPosY := arrPopupFixPosition2

if (blnMouse)
{
	if (blnNewWindow)
		MouseGetPos, , , strTargetWinId
	else
		WinActivate, % "ahk_id " . strTargetWinId
	if (!blnPopupFix)
	{
		; display menu at mouse pointer location
		intMenuPosX :=
		intMenuPosY :=
	}
}
else
{
	if (blnNewWindow)
		strTargetWinId := WinExist("A")
	if (!blnPopupFix)
	{
		; display menu at an offset of 20x20 pixel from top-left client area
		intMenuPosX := 20
		intMenuPosY := 20
	}
}

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildTrayMenu:
;------------------------------------------------------------

;@Ahk2Exe-IgnoreBegin
; Piece of code for developement phase only - won't be compiled
Menu, Tray, Icon, %A_ScriptDir%\Folders-Likes-icon-192-RED-center.ico, 1
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd
Menu, Tray, Add
Menu, Tray, Add, % L(lMenuSettings, strAppName), GuiShow
Menu, Tray, Add
Menu, Tray, Add, %lMenuRunAtStartup%, RunAtStartup
Menu, Tray, Add
Menu, Tray, Add, %lMenuUpdate%, Check4Update
Menu, Tray, Add, %lMenuHelp%, GuiHelp
Menu, Tray, Add, %lMenuAbout%, GuiAbout
Menu, Tray, Default, % L(lMenuSettings, strAppName)

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
BuildRecentFoldersMenu:
;------------------------------------------------------------

Menu, menuRecentFolders, Add
; Menu, menuRecentFolders, DeleteAll ; do not use, because it makes the Special menu to disappear 1/2 times
Menu, menuRecentFolders, Delete

objRecentFolders := Object()
intRecentFoldersIndex := 0 ; used in PopupMenu... to check if we disable the menu when empty

if (A_OSVersion = "WIN_XP")
{
	strRecentsFolder := SubStr(A_AppData, 1, InStr(A_AppData, "\", , 0) - 1) . "\Recent"
	if !FileExist(strRecentsFolder)
		strRecentsFolder := SubStr(A_AppData, 1, InStr(A_AppData, "\", , 0) - 1) . "\My Recent Documents"
}
else
	strRecentsFolder := A_AppData . "\Microsoft\Windows\Recent"

blnDisplayRecentFolders := (StrLen(FileExist(strRecentsFolder)) > 0) ; turn recent menu off if recent folder not found

strDirList := ""
	
Loop, %strRecentsFolder%\*.* ; tried to limit to number of recent but no good because not sorted chronologically
	strDirList := strDirList . A_LoopFileTimeModified . "`t`" . A_LoopFileFullPath . "`n"
Sort, strDirList, R

intShortcut := 0

Loop, parse, strDirList, `n
{
	if !StrLen(A_LoopField) ; last line is empty
		continue

	arrTargetFullName := StrSplit(A_LoopField, A_Tab)
	strTargetFullName := arrTargetFullName[2]
	FileGetShortcut, %strTargetFullName%, strOutTarget
	if (errorlevel) ; hidden or system files (like desktop.ini) returns an error
		continue
	
	FileGetAttrib, strAttributes, %strOutTarget%
	if !InStr(strAttributes, "D")
		SplitPath, strOutTarget, , strOutTarget ; for files, retreive the folder containing the file

	if !ValueIsInObject(strOutTarget, objRecentFolders)
	{
		intRecentFoldersIndex := intRecentFoldersIndex + 1
		objRecentFolders.Insert(intRecentFoldersIndex, strOutTarget)
		
		Menu, menuRecentFolders, Add, % (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, false) . " " : "") . strOutTarget, OpenRecentFolder
		
		if (intRecentFoldersIndex >= intRecentFolders)
			break
	}
}

if !(intRecentFoldersIndex) ; recreate the menu because it was deleted
	Menu, menuRecentFolders, Add
Menu, menuMore, Add, %lMenuRecentFolders%, :menuRecentFolders ; re-add to More menu because it was deleted

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildSwitchMenu:
;------------------------------------------------------------

Menu, menuSwitchExplorer, Add
; Menu, menuSwitchExplorer, DeleteAll ; do not use, because it makes the Special menu to disappear
Menu, menuSwitchExplorer, Delete
Menu, menuSwitchDialog, Add
; Menu, menuSwitchDialog, DeleteAll ; do not use, because it makes the Special menu to disappear
Menu, menuSwitchDialog, Delete

objExplorersWinId := Object()
objExplorersLocation := Object()
intExplorersIndex := 0 ; used in PopupMenu... to check if we disable the menu when empty
intShortcut := 0

For pExp in ComObjCreate("Shell.Application").Windows
{
	strType := ""
	try
		strType := pExp.Type
	if !StrLen(strType)
	{
		intExplorersIndex := intExplorersIndex + 1
		objExplorersWinId.Insert(intExplorersIndex, pExp.HWND)
		objExplorersLocation.Insert(intExplorersIndex, pExp.LocationURL) ; also .LocationName see http://msdn.microsoft.com/en-us/library/aa752084#properties
		
		strLocation :=  UriDecode(pExp.LocationURL)
		StringReplace, strLocation, strLocation, file:///
		StringReplace, strLocation, strLocation, /, \, A
		if !StrLen(strLocation)
			strLocation := pExp.LocationName

		Menu, menuSwitchExplorer, Add
			, % (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, false) . " " : "") . strLocation
			, SwitchExplorer
		Menu, menuSwitchDialog, Add
			, % (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, false) . " " : "") . strLocation
			, SwitchDialog
	}
}

if !(intExplorersIndex) ; recreate the menu because it was deleted
{
	Menu, menuSwitchExplorer, Add
	Menu, menuSwitchDialog, Add
}
Menu, menuMore, Add, %lMenuSwitchExplorer%, :menuSwitchExplorer ; re-add to More menu because it was deleted
Menu, menuMore, Add, %lMenuSwitchDialog%, :menuSwitchDialog ; re-add to More menu because it was deleted

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildFoldersMenus:
;------------------------------------------------------------

Menu, %lMainMenuName%, Add
Menu, %lMainMenuName%, DeleteAll
; Menu, %lMainMenuName%, Color, %strMenuBackgroundColor% ; only later when icons will be added

BuildOneMenu(lMainMenuName) ; and recurse for submenus

Menu, menuMore, Add
Menu, menuMore, DeleteAll

if (blnDisplaySpecialFolders)
	Menu, menuMore, Add, %lMenuSpecialFolders%, :menuSpecialFolders
if (blnDisplayRecentFolders)
	Menu, menuMore, Add, %lMenuRecentFolders%, :menuRecentFolders
if (blnDisplaySwitchMenu)
{
	Menu, menuMore, Add, %lMenuSwitchExplorer%, :menuSwitchExplorer
	Menu, menuMore, Add, %lMenuSwitchDialog%, :menuSwitchDialog
}

Menu, %lMainMenuName%, Add
Menu, %lMainMenuName%, Add, %lMenuMore%, :menuMore
Menu, %lMainMenuName%, Add
Menu, %lMainMenuName%, Add, % L(lMenuSettings, strAppName), GuiShow
Menu, %lMainMenuName%, Default, % L(lMenuSettings, strAppName)
Menu, %lMainMenuName%, Add, %lMenuAddThisFolder%, AddThisFolder

if !(blnDonator)
{
	Menu, %lMainMenuName%, Add
	Menu, %lMainMenuName%, Add, %lDonateMenu%, GuiDonate
}

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildOneMenu(strMenu)
; recursive function
;------------------------------------------------------------
{
	global blnDisplayMenuShortcuts
	intShortcut := 0
	
	; try because at first execution strMenu does not exist and produces an error,
	; but DeleteAll is required later for menu updates
	try Menu, %strMenu%, DeleteAll
	
	arrThisMenu := arrMenus[strMenu]
	Loop, % arrThisMenu.MaxIndex()
		if StrLen(arrThisMenu[A_Index].SubmenuFullName) ; this is a submenu
		{
			strSubMenuFullName := arrThisMenu[A_Index].SubmenuFullName
			strSubMenuDisplayName := arrThisMenu[A_Index].FolderName
			strSubMenuParent := arrThisMenu[A_Index].MenuName
			
			BuildOneMenu(strSubMenuFullName) ; recursive call
			
			try Menu, %strSubMenuParent%, Add
				, % (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, (strMenu = lMainMenuName)) . " " : "")
				. strSubMenuDisplayName, % ":" . strSubMenuFullName
			catch e
			{
				Menu, % arrThisMenu[A_Index].MenuName, Add, % arrThisMenu[A_Index].FolderName, OpenFavorite ; will never be called because disabled
				Menu, % arrThisMenu[A_Index].MenuName, Disable, % arrThisMenu[A_Index].FolderName
			}
		}
		else if (arrThisMenu[A_Index].FolderName = lMenuSeparator)
			Menu, % arrThisMenu[A_Index].MenuName, Add
		else
			Menu, % arrThisMenu[A_Index].MenuName , Add
				, % (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, (strMenu = lMainMenuName)) . " " : "")
				. arrThisMenu[A_Index].FolderName, OpenFavorite
}
;------------------------------------------------------------


;------------------------------------------------------------
BuildAddDialogMenu:
;------------------------------------------------------------

Menu, menuAddDialog, Add, %lMenuDialogNotSupported%, AddThisDialog
Menu, menuAddDialog, Disable, %lMenuDialogNotSupported%
Menu, menuAddDialog, Add, %lMenuAddThisDialog%, AddThisDialog
Menu, menuAddDialog, Add
Menu, menuAddDialog, Add, % L(lMenuSettings, strAppName), GuiShow
Menu, menuAddDialog, Default, %lMenuAddThisDialog%

return
;------------------------------------------------------------


;------------------------------------------------------------
LoadTheme:
; colors for themes
; skin: WindowColor=FFC495
; Note: implement menu background color only when icons are implemented
;------------------------------------------------------------

IniRead, strGuiWindowColor, %strIniFile%, Gui, WindowColor, E0E0E0
IniRead, strGuiListviewBackgroundColor, %strIniFile%, Gui, ListviewBackground, FFFFFF
IniRead, strGuiListviewTextColor, %strIniFile%, Gui, ListviewText, 000000
IniRead, strMenuBackgroundColor, %strIniFile%, Gui, MenuBackgroundColor, FFFFFF

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildGui:
;------------------------------------------------------------

Gui, 1:New, , % L(lGuiTitle, strAppName, strAppVersion)
Gui, Margin, 10, 10
Gui, Color, %strGuiWindowColor%, %strGuiWindowColor%

intCol := 480
intWidth := 460

Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Text, xm ym w%intWidth% h25, %strAppName% %strAppVersion%
Gui, 1:Font, s10 w400, Verdana
Gui, 1:Add, Text, xm y+1, %lAppTagline%
Gui, 1:Font, s8 w400, Verdana
Gui, 1:Add, Text, xm+30, %lGuiSubmenuDropdownLabel%
Gui, 1:Add, Picture, xm y+5 gGuiGotoPreviousMenu vpicPreviousMenu hidden, %strTempDir%\left-12.png
Gui, 1:Add, Picture, xm+15 yp gGuiGotoUpMenu vpicUpMenu hidden, %strTempDir%\up-12.png
Gui, 1:Add, DropDownList, xm+30 yp w320 vdrpMenusList gGuiMenusListChanged ; Sort

Gui, 1:Font, s8 w400, Verdana
Gui, 1:Add, ListView, xm+30 w320 h220 Count32 -Multi NoSortHdr LV0x10 c%strGuiListviewTextColor% Background%strGuiListviewBackgroundColor% vlvFoldersList gGuiFoldersListEvent, %lGuiLvFoldersHeader%|Hidden Menu|Hidden Submenu
LV_ModifyCol(3, 0) ; hide 3rd column
LV_ModifyCol(4, 0) ; hide 4th column

Gui, 1:Add, Text, Section x+0 yp

Gui, 1:Add, Picture, xm ys+25 gGuiMoveFolderUp, %strTempDir%\up_circular-26.png
Gui, 1:Add, Picture, xm ys+55 gGuiMoveFolderDown, %strTempDir%\down_circular-26.png
Gui, 1:Add, Picture, xm ys+85 gGuiAddSeparator, %strTempDir%\separator-26.png
Gui, 1:Add, Picture, xm+1 ys+175 gGuiSortFolders, %strTempDir%\generic_sorting2-26-grey.png

Gui, 1:Add, Text, Section xs ys

Gui, 1:Add, Picture, xs+10 ys gGuiAddFolder, %strTempDir%\add_property-48.png
Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs y+0 w68 center gGuiAddFolder, %lGuiAddFolder%

Gui, 1:Add, Picture, xs+10 gGuiEditFolder, %strTempDir%\edit_property-48.png
Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs y+0 w68 center gGuiEditFolder, %lGuiEditFolder%

Gui, 1:Add, Picture, xs+10 gGuiRemoveFolder, %strTempDir%\delete_property-48.png
Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs y+0 w68 center gGuiRemoveFolder, %lGuiRemoveFolder%

Gui, 1:Add, Text, xs+90 ys h220 0x11

GuiControlGet, arrLvPos, Pos, lvFoldersList

Gui, 1:Font, s8 w400, Verdana
Gui, 1:Add, Text, Section x%intCol% y%arrLvPosY% w220, %lGuiLvDialogsHeader%
Gui, 1:Add, ListView, xs w220 h150 Count16 -Hdr -Multi NoSortHdr +0x10 LV0x10 c%strGuiListviewTextColor% Background%strGuiListviewBackgroundColor% vlvDialogsList gGuiDialogsListEvent, Header

Gui, 1:Add, Picture, xs+70 y+5 gGuiAddDialog, %strTempDir%\add_image-26.png
Gui, 1:Add, Picture, x+10 yp gGuiEditDialog, %strTempDir%\edit_image-26.png
Gui, 1:Add, Picture, x+10 yp gGuiRemoveDialog, %strTempDir%\remove_image-26.png

Gui, 1:Add, Text, Section xs+60 ym

Gui, 1:Add, Picture, xs+10 ym gGuiHelp, %strTempDir%\help-32.png
Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs y+0 w52 center gGuiHelp, %lGuiHelp%

Gui, 1:Add, Picture, Section x+10 ym gGuiAbout, %strTempDir%\about-32.png
Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs-10 y+0 w52 center, %lGuiAbout%

Gui, 1:Add, Picture, Section x+10 ym gGuiOptions, %strTempDir%\settings-32.png
Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs-10 y+0 w52 center gGuiOptions, %lGuiOptions%

Gui, 1:Add, Text, xm y350 h1 w690
Gui, 1:Font, s9 w600, Verdana
Gui, 1:Add, Button, x260 w100 h48 Disabled Default vbtnGuiSave gGuiSave, %lGuiSave%
Gui, 1:Add, Button, x+20 w100 h48 vbtnGuiCancel gGuiCancel, %lGuiClose% ; Close until changes occur
Gui, 1:Font, s6 w400, Verdana

return
;------------------------------------------------------------


;------------------------------------------------------------
BackupMenuObjects:
; in case of Gui Cancel to restore objects to original state
;------------------------------------------------------------

arrMenusBK := Object()
for strMenuName, arrMenu in arrMenus
{
	arrMenuBK := Object()
	for intIndex, objFolder in arrMenu
	{
		objFolderBK := Object()
		objFolderBK.MenuName := objFolder.MenuName
		objFolderBK.FolderName := objFolder.FolderName
		objFolderBK.FolderLocation := objFolder.FolderLocation
		objFolderBK.SubmenuFullName := objFolder.SubmenuFullName
		arrMenuBK.Insert(objFolderBK)
	}
	arrMenusBK.Insert(strMenuName, arrMenuBK) ; add this submenu to the array of menus
}

return
;------------------------------------------------------------


;------------------------------------------------------------
RestoreBackupMenuObjects:
; when Gui Cancel to restore objects to original state
;------------------------------------------------------------

arrMenus := Object() ; reinit
for strMenuName, arrMenuBK in arrMenusBK
{
	arrMenu := Object()
	for intIndex, objFolderBK in arrMenuBK
	{
		objFolder := Object()
		objFolder.MenuName := objFolderBK.MenuName
		objFolder.FolderName := objFolderBK.FolderName
		objFolder.FolderLocation := objFolderBK.FolderLocation
		objFolder.SubmenuFullName := objFolderBK.SubmenuFullName
		arrMenu.Insert(objFolder)
	}
	arrMenus.Insert(strMenuName, arrMenu) ; add this submenu to the array of menus
}

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiFoldersListEvent:
;------------------------------------------------------------
Gui, 1:ListView, lvFoldersList

if (A_GuiEvent = "DoubleClick")
	gosub, GuiEditFolder

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiDialogsListEvent:
;------------------------------------------------------------
Gui, 1:ListView, lvDialogsList

if (A_GuiEvent = "DoubleClick")
	gosub, GuiEditDialog

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMenusListChanged:
GuiGotoUpMenu:
GuiGotoPreviousMenu:
;------------------------------------------------------------

Gosub, SaveCurrentListviewToMenuObject ; save current LV before changing strCurrentMenu

strSavedMenu := strCurrentMenu
if (A_ThisLabel = "GuiMenusListChanged")
{
	arrSubmenuStack.Insert(1, strSavedMenu) ; push the current menu to the left arrow stack
	GuiControlGet, strCurrentMenu, , drpMenusList
}
else if (A_ThisLabel = "GuiGotoUpMenu")
{
	arrSubmenuStack.Insert(1, strSavedMenu) ; push the current menu to the left arrow stack
	strCurrentMenu := SubStr(strCurrentMenu, 1, InStr(strCurrentMenu, lGuiSubmenuSeparator, , 0) - 1) 
}
else ; GuiGotoPreviousMenu
{
	strCurrentMenu := arrSubmenuStack[1] ; pull the top menu from the left arrow stack
	arrSubmenuStack.Remove(1) ; remove the top menu from the left arrow stack
}
strPreviousMenu := strSavedMenu

GuiControl, % arrSubmenuStack.MaxIndex() ? "Show" : "Hide", picPreviousMenu
GuiControl, % arrSubmenuStack.MaxIndex() ? "Show" : "Hide", picUpMenu

Gosub, LoadOneMenuToGui

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

LV_GetText(strSubmenuFullName, intItemToRemove, 4)
if StrLen(strSubmenuFullName)
{
	MsgBox, 52, % L(lDialogFolderRemoveTitle, strAppName), % L(lDialogFolderRemovePrompt, strSubmenuFullName)
	IfMsgBox, No
		return
	RemoveAllSubMenus(strSubmenuFullName)
}

LV_Delete(intItemToRemove)
LV_Modify(intItemToRemove, "Select Focus")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

if StrLen(strSubmenuFullName)
{
	Gosub, SaveCurrentListviewToMenuObject ; save current LV before update the dropdown menu
	GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
}

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
RemoveAllSubMenus(strSubmenuFullName)
; recursive function
;------------------------------------------------------------
{
	Loop, % arrMenus[strSubmenuFullName].MaxIndex()
		if StrLen(arrMenus[strSubmenuFullName][A_Index].SubmenuFullName)
			RemoveAllSubMenus(arrMenus[strSubmenuFullName][A_Index].SubmenuFullName) ; recursive call
	arrMenus.Remove(strSubmenuFullName)
}
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddSeparator:
;------------------------------------------------------------

GuiControl, Focus, lvFoldersList
Gui, 1:ListView, lvFoldersList
LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1, "Select Focus", lMenuSeparator, lMenuSeparator . lMenuSeparator, strCurrentMenu)
LV_Modify(LV_GetNext(), "Vis")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiSortFolders:
;------------------------------------------------------------

Gui, 1:ListView, lvFoldersList
LV_ModifyCol(1, "Sort")

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

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
LV_GetText(strThisLocation, intSelectedRow, 2)
LV_GetText(strThisMenu, intSelectedRow, 3)
LV_GetText(strThisSubmenu, intSelectedRow, 4)

LV_GetText(strPriorName, intSelectedRow - 1, 1)
LV_GetText(strPriorLocation, intSelectedRow - 1, 2)
LV_GetText(strPriorMenu, intSelectedRow - 1, 3)
LV_GetText(strPriorSubmenu, intSelectedRow - 1, 4)

LV_Modify(intSelectedRow, "", strPriorName, strPriorLocation, strPriorMenu, strPriorSubmenu)
LV_Modify(intSelectedRow - 1, "Select Focus Vis", strThisName, strThisLocation, strThisMenu, strThisSubmenu)

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

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
LV_GetText(strThisLocation, intSelectedRow, 2)
LV_GetText(strThisMenu, intSelectedRow, 3)
LV_GetText(strThisSubmenu, intSelectedRow, 4)

LV_GetText(strNextName, intSelectedRow + 1, 1)
LV_GetText(strNextLocation, intSelectedRow + 1, 2)
LV_GetText(strNextMenu, intSelectedRow + 1, 3)
LV_GetText(strNextSubmenu, intSelectedRow + 1, 4)
	
LV_Modify(intSelectedRow, "", strNextName, strNextLocation, strNextMenu, strNextSubmenu)
LV_Modify(intSelectedRow + 1, "Select Focus Vis", strThisName, strThisLocation, strThisMenu, strThisSubmenu)

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

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

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

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

InputBox strNewDialog, % L(lDialogEditDialogTitle, strAppName, strAppVersion)
	, %lDialogEditDialogPrompt%, , 350, 150, , , , , %strCurrentDialog%
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

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiSave:
;------------------------------------------------------------

blnMenuReady := false

Gosub, SaveCurrentListviewToMenuObject ; save current LV before saving

IniDelete, %strIniFile%, Folders
Gui, 1:ListView, lvFoldersList

intIndex := 0
SaveOneMenu(lMainMenuName) ; recursive function

IniDelete, %strIniFile%, Dialogs
Gui, 1:ListView, lvDialogsList
Loop % LV_GetCount()
{
	LV_GetText(strDialog, A_Index, 1)
	IniWrite, %strDialog%, %strIniFile%, Dialogs, Dialog%A_Index%
}

Gosub, LoadIniFile
Gosub, BuildFoldersMenus
GuiControl, Disable, %lGuiSave%
GuiControl, , %lGuiCancel%, %lGuiClose%

Gosub, GuiCancel
blnMenuReady := true

return
;------------------------------------------------------------


;------------------------------------------------------------
SaveOneMenu(strMenu)
; recursive function
;------------------------------------------------------------
{
	global strIniFile
	global intIndex
	
	Loop, % arrMenus[strMenu].MaxIndex()
	{
		strValue := arrMenus[strMenu][A_Index].FolderName 
			. "|" . arrMenus[strMenu][A_Index].FolderLocation 
			. "|" . SubStr(arrMenus[strMenu][A_Index].MenuName, StrLen(lMainMenuName) + 1) ; remove main menu name for ini file
			. "|" . SubStr(arrMenus[strMenu][A_Index].SubmenuFullName, StrLen(lMainMenuName) + 1) ; remove main menu name for ini file
		intIndex := intIndex + 1
		IniWrite, %strValue%, %strIniFile%, Folders, Folder%intIndex%
		if StrLen(arrMenus[strMenu][A_Index].SubmenuFullName)
			SaveOneMenu(arrMenus[strMenu][A_Index].SubmenuFullName) ; recursive call
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
GuiShow:
;------------------------------------------------------------

strCurrentMenu := lMainMenuName
Gosub, BackupMenuObjects
Gosub, LoadSettingsToGui
Gui, 1:Show ; , w455 ; , h455

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiCancel:
;------------------------------------------------------------

GuiControlGet, blnSaveEnabled, Enabled, btnGuiSave
if (blnSaveEnabled)
{
	Gui, 1:+OwnDialogs
	MsgBox, 36, % L(lDialogCancelTitle, strAppName, strAppVersion), %lDialogCancelPrompt%
	IfMsgBox, Yes
	{
		blnMenuReady := false
		Gosub, RestoreBackupMenuObjects
		
		; restore popup menu
		Gosub, BuildSpecialFoldersMenu
		Gosub, BuildRecentFoldersMenu
		Gosub, BuildSwitchMenu
		Gosub, BuildFoldersMenus ; need to be initialized here - will be updated at each call to popup menu

		GuiControl, Disable, btnGuiSave
		GuiControl, , btnGuiCancel, %lGuiClose%
		blnMenuReady := true
	}
	IfMsgBox, No
		return
}
Gui, 1:Cancel

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiClose:
GuiEscape:
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
; In all these OS, with Explorer, the key sequence {F4}{Esc} selects the current location of the window.
; With dialog boxes, the key sequence {F4}{Esc} generally selects the current location of the window. But, in some
; dialog boxes, the {Esc} key closes the dialog box. We will check window title to detect this behavior.

WinGetTitle, strWindowTitle, A ; to check later if this window is closed unexpectedly

if (strTargetClass = "#32770")
	intWaitTimeIncrement := 300 ; time allowed for dialog boxes
else
	intWaitTimeIncrement := 150 ; time allowed for Explorer

if (blnDiagMode)
	intTries := 8
else
	intTries := 3

if WindowIsFreeCommander(strTargetClass)
{
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
	{
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
		Sleep, intWaitTimeIncrement
	}
	MouseGetPos, , , , strCommanderControl, 1
	Sleep, intWaitTimeIncrement
	ControlFocus, %strCommanderControl%
	Loop, %intTries%
	{
		Sleep, intWaitTimeIncrement * A_Index
		Send, !g
		Sleep, intWaitTimeIncrement * A_Index
		Send, ^c
		intTries := A_Index
		Sleep, intWaitTimeIncrement
	} Until (StrLen(ClipBoard))
}
else
	Loop, %intTries%
	{
		Sleep, intWaitTimeIncrement * A_Index
		Send {F4}{Esc} ; F4 move the caret the "Go To A Different Folder box" and {Esc} select it content ({Esc} could be replaced by ^a to Select All)
		Sleep, intWaitTimeIncrement * A_Index
		Send ^c ; Copy
		Sleep, intWaitTimeIncrement * A_Index
		intTries := A_Index
		WinGetTitle, strWindowThisTitle, A ; to check if the window was closed unexpectedly
	} Until (StrLen(ClipBoard) or (strWindowTitle <> strWindowThisTitle))

strCurrentLocation := ClipBoard
Clipboard := objPrevClipboard ; Restore the original clipboard
objPrevClipboard := "" ; Free the memory in case the clipboard was very large

if (blnDiagMode)
{
	Diag("Menu", A_ThisLabel)
	Diag("Class", strTargetClass)
	Diag("Tries", intTries)
	Diag("AddedFolder", strCurrentLocation)
}

If !StrLen(strCurrentLocation) or !InStr(strCurrentLocation, ":")  or (strWindowTitle <> strWindowThisTitle)
{
	Gui, 1:+OwnDialogs 
	MsgBox, 52, % L(lDialogAddFolderManuallyTitle, strAppName, strAppVersion), %lDialogAddFolderManuallyPrompt%
	IfMsgBox, Yes
	{
		Gosub, GuiShow
		Gosub, GuiAddFolder
	}
}
else
{
	Gosub, GuiShow
	Gosub, GuiAddFromPopup
}

return
;------------------------------------------------------------


;------------------------------------------------------------
AddThisDialog:
;------------------------------------------------------------

WinGetTitle, strDialogTitle, ahk_id %strTargetWinId%
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

	InputBox, strNewDialog, % L(lDialogAddDialogTitle, strAppName, strAppVersion), %lDialogAddDialogPrompt%, , 350, 150, , , , , %strCurrentDialogTitle%
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

	GuiControl, Enable, btnGuiSave
	GuiControl, , btnGuiCancel, %lGuiCancel%
}
;------------------------------------------------------------


;------------------------------------------------------------
RunAtStartup:
;------------------------------------------------------------
; Startup code adapted from Avi Aryan Ryan in Clipjump

Menu, Tray, Togglecheck, %lMenuRunAtStartup%
IfExist, %A_Startup%\%strAppName%.lnk
	FileDelete, %A_Startup%\%strAppName%.lnk
else
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%strAppName%.lnk

return
;------------------------------------------------------------


;------------------------------------------------------------
Check4Update:
;------------------------------------------------------------
Gui, 1:+OwnDialogs

if Time2Donate(intStartups, blnDonator)
{
	MsgBox, 36, % l(lDonateCheckTitle, intStartups, strAppName), % l(lDonateCheckPrompt, strAppName, intStartups)
	IfMsgBox, Yes
		Gosub, GuiDonate
}
IniWrite, % (intStartups + 1), %strIniFile%, Global, Startups

strLatestVersion := Url2Var("https://raw.github.com/JnLlnd/FoldersPopup/master/latest-version.txt")
strLatestVersion := Trim(strLatestVersion, "`n`l") ; remove en-of-line if present

if RegExMatch(strCurrentVersion, "(alpha|beta)")
	or (FirstVsSecondIs(strLatestSkipped, strLatestVersion) >= 0 and (A_ThisMenuItem <> lMenuUpdate))
	return

if FirstVsSecondIs(strLatestVersion, strCurrentVersion) = 1
{
	Gui, 1:+OwnDialogs
	SetTimer, ChangeButtonNames, 50

	MsgBox, 3, % l(lUpdateTitle, strAppName), % l(lUpdatePrompt, strAppName, strCurrentVersion, strLatestVersion), 30
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
	MsgBox, 4, % l(lUpdateTitle, strAppName), % l(lUpdateYouHaveLatest, strAppVersion, strAppName), 30
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


;------------------------------------------------------------
ChangeButtonNames: 
;------------------------------------------------------------

IfWinNotExist, % l(lUpdateTitle, strAppName)
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
Time2Donate(intStartups, blnDonator)
;------------------------------------------------------------
{
	if (intStartups > 100)
		intDivisor := 25
	else if (intStartups > 60)
		intDivisor := 20
	else if (intStartups > 30)
		intDivisor := 15
	else intDivisor := 10
		
	return !Mod(intStartups, intDivisor) and !(blnDonator)
}
;------------------------------------------------------------


;------------------------------------------------------------
; GUI2 About, Help and Options
;------------------------------------------------------------

;------------------------------------------------------------
GuiAddFolder:
GuiAddFromPopup:
;------------------------------------------------------------

intRowToEdit := 0 ;  used when saving to flag to insert a new row
if (A_ThisLabel = "GuiAddFromPopup")
{
	strCurrentName := GetDeepestFolderName(strCurrentLocation) ; value received from AddThisFolder
	blnRadioFolder := true
	blnRadioSubmenu := false
}
else
{
	strCurrentName := ""
	strCurrentLocation := ""
}
strCurrentSubmenuFullName := ""

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lDialogAddEditFolderTitle, lDialogAdd, strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:+OwnDialogs

if (A_ThisLabel = "GuiAddFolder")
{
	Gui, 2:Add, Text, x10, Add:
	Gui, 2:Add, Radio, x+10 yp vblnRadioFolder checked gRadioButtonsChanged, %lDialogFolderLabel%
	Gui, 2:Add, Radio, x+10 yp vblnRadioSubmenu gRadioButtonsChanged, %lDialogSubmenuLabel%
}
else
{
	blnRadioFolder := true
	blnRadioSubmenu := false
}

Gui, 2:Add, Text, x10 y+10 vlblShortName, %lDialogFolderShortName%
Gui, 2:Add, Edit, x10 w200 vstrFolderShortName, %strCurrentName%

Gui, 2:Add, Text, x10 vlblFolder, %lDialogFolderLabel%
Gui, 2:Add, Edit, x10 w300 vstrFolderLocation, %strCurrentLocation%
Gui, 2:Add, Button, x+10 yp vbtnSelectFolderLocation gButtonSelectFolderLocation default, %lDialogBrowseButton%

Gui, 2:Add, Button, x150 gGuiAddFolderSave, %lDialogAdd%
Gui, 2:Add, Button, x+20 yp gGuiAddFolderCancel, %lGuiCancel%
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFolderSave:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

blnMenuReady := false

if !StrLen(strFolderShortName)
{
	Oops(lDialogFolderNameEmpty)
	return
}

if  (blnRadioFolder and !StrLen(strFolderLocation))
{
	Oops(lDialogFolderLocationEmpty)
	return
}

if !FolderNameIsNew(strFolderShortName)
	and (strFolderShortName <> strCurrentName)
{
	Oops(lDialogFolderNameNotNew, strFolderShortName)
	return
}

if (blnRadioSubmenu)
{
	if InStr(strFolderShortName, lGuiSubmenuSeparator)
	{
		Oops(L(lDialogFolderNameNoSeparator, lGuiSubmenuSeparator))
		return
	}
	strFolderLocation := lGuiSubmenuLocation
	strSubmenuFullName := strCurrentMenu . lGuiSubmenuSeparator . strFolderShortName
	
	arrNewMenu := Object() ; array of folders of the new menu
	arrMenus.Insert(strSubmenuFullName, arrNewMenu)
}
else
	strSubmenuFullName := ""

Gosub, 2GuiClose

Gui, 1:Default
GuiControl, 1:Focus, lvFoldersList
Gui, 1:ListView, lvFoldersList

; insert new item at the top of the Gui
LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1, "Select Focus", strFolderShortName, strFolderLocation, strCurrentMenu, strSubmenuFullName)
LV_Modify(LV_GetNext(), "Vis")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

Gosub, SaveCurrentListviewToMenuObject ; save current LV tbefore update the dropdown menu
GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
Gosub, BuildFoldersMenus ; update menus

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lDialogCancelButton%

blnMenuReady := true

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditFolder:
;------------------------------------------------------------

Gui, 1:ListView, lvFoldersList
intRowToEdit := LV_GetNext()
LV_GetText(strCurrentName, intRowToEdit, 1)

if !StrLen(strCurrentName)
{
	Oops(lDialogSelectItemToEdit)
	return
}
LV_GetText(strCurrentLocation, intRowToEdit, 2)
LV_GetText(strCurrentSubmenuFullName, intRowToEdit, 4)
blnRadioSubmenu := StrLen(strCurrentSubmenuFullName)
blnRadioFolder := !blnRadioSubmenu

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lDialogAddEditFolderTitle, lDialogEdit, strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:+OwnDialogs

Gui, 2:Add, Text, x10 y10 vlblFolderParentMenu, %lDialogFolderParentMenu%
Gui, 2:Add, DropDownList, x10 w250 vdrpParentMenu, % BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu, strCurrentSubmenuFullName) . "|"

Gui, 2:Add, Text, x10 y+10 vlblShortName, % (StrLen(strCurrentSubmenuFullName) ? lDialogSubmenuShortName : lDialogFolderShortName)
Gui, 2:Add, Edit, x10 w200 vstrFolderShortName, %strCurrentName%

Gui, 2:Add, Text, % "x10 vlblFolder " . (strCurrentLocation = lGuiSubmenuLocation ? "hidden" : ""), %lDialogFolderLabel%
Gui, 2:Add, Edit, % "x10 w300 vstrFolderLocation " . (strCurrentLocation = lGuiSubmenuLocation ? "hidden" : ""), %strCurrentLocation%
Gui, 2:Add, Button, % "x+10 yp vbtnSelectFolderLocation gButtonSelectFolderLocation " . (strCurrentLocation = lGuiSubmenuLocation ? "hidden" : "default"), %lDialogBrowseButton%

Gui, 2:Add, Button, x100 gGuiEditFolderSave, %lDialogSave%
Gui, 2:Add, Button, x+20 yp gGuiEditFolderCancel, %lGuiCancel%
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiEditFolderSave:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

blnMenuReady := false

GuiControlGet, strParentMenu, , drpParentMenu
	
if !StrLen(strFolderShortName)
{
	Oops(lDialogFolderNameEmpty)
	return
}

if  (blnRadioFolder and !StrLen(strFolderLocation))
{
	Oops(lDialogFolderLocationEmpty)
	return
}

if !FolderNameIsNew(strFolderShortName, (strParentMenu = strCurrentMenu ? "" : strParentMenu))
	if ((strParentMenu <> strCurrentMenu) or (strFolderShortName <> strCurrentName))
	{
		Oops(lDialogFolderNameNotNew, strFolderShortName)
		return
	}

if (blnRadioSubmenu)
{
	if InStr(strFolderShortName, lGuiSubmenuSeparator)
	{
		Oops(L(lDialogFolderNameNoSeparator, lGuiSubmenuSeparator))
		return
	}
	strNewSubmenuFullName := strParentMenu . lGuiSubmenuSeparator . strFolderShortName
	
	UpdateMenuNameInSubmenus(strCurrentSubmenuFullName, strNewSubmenuFullName) ; change names in arrMenus and arrMenu objects	
}
else
	strNewSubmenuFullName := ""

Gosub, 2GuiClose

Gui, 1:Default
GuiControl, 1:Focus, lvFoldersList
Gui, 1:ListView, lvFoldersList

if (strParentMenu = strCurrentMenu) ; modify selected row
	LV_Modify(intRowToEdit, "Select Focus", strFolderShortName, strFolderLocation, strCurrentMenu, strNewSubmenuFullName)
else ; move menu item to another menu: add to the new menu's object and remove item from this menu in the Gui
{
	objFolder := Object() ; new menu item
	objFolder.MenuName := strParentMenu ; parent menu of this menu item
	objFolder.FolderName := strFolderShortName ; display name of this menu item
	objFolder.FolderLocation := strFolderLocation ; path for this menu item
	objFolder.SubmenuFullName := strNewSubmenuFullName ; full name of the submenu
	arrMenus[objFolder.MenuName].Insert(objFolder) ; add this menu item to the new menu
	
	LV_Delete(intRowToEdit)
	LV_Modify(intRowToEdit, "Select Focus")
}

LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

Gosub, SaveCurrentListviewToMenuObject ; save current LV tbefore update the dropdown menu
GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
Gosub, BuildFoldersMenus ; update menus

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lDialogCancelButton%

blnMenuReady := true

return
;------------------------------------------------------------


;------------------------------------------------------------
UpdateMenuNameInSubmenus(strOldMenu, strNewMenu)
; recursive function
;------------------------------------------------------------
{
	arrMenus.Insert(strNewMenu, arrMenus[strOldMenu])
	Loop, % arrMenus[strNewMenu].MaxIndex()
	{
		arrMenus[strNewMenu][A_Index].MenuName := strNewMenu
		if StrLen(arrMenus[strNewMenu][A_Index].SubmenuFullName)
		{
			arrMenus[strNewMenu][A_Index].SubmenuFullName := strNewMenu . lGuiSubmenuSeparator . arrMenus[strNewMenu][A_Index].FolderName
			UpdateMenuNameInSubmenus(strOldMenu . lGuiSubmenuSeparator . arrMenus[strOldMenu][A_Index].FolderName
				, arrMenus[strNewMenu][A_Index].SubmenuFullName) ; recursive call
		}
	}
	arrMenus.Remove(strOldMenu)
}
;------------------------------------------------------------


;------------------------------------------------------------
RadioButtonsChanged:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

GuiControl, , lblShortName, % (blnRadioSubmenu ? lDialogSubmenuShortName : lDialogFolderShortName)
GuiControl, % (blnRadioSubmenu ? "Hide" : "Show"), lblFolder
GuiControl, % (blnRadioSubmenu ? "Hide" : "Show"), strFolderLocation
GuiControl, % (blnRadioSubmenu ? "Hide" : "Show"), btnSelectFolderLocation
GuiControl, % (blnRadioSubmenu ? "-" : "-") . "Default", btnSelectFolderLocation

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonSelectFolderLocation:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

FileSelectFolder, strNewLocation, *%strCurrentLocation%, 3, %lDialogAddFolderSelect%
if !(StrLen(strNewLocation))
	return
GuiControl, 2:, strFolderLocation, %strNewLocation%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFolderCancel:
GuiEditFolderCancel:
;------------------------------------------------------------
gosub, 2GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
FolderNameIsNew(strCandidateName, strMenu := "")
;------------------------------------------------------------
{
	if !StrLen(strMenu)
	{
		Gui, 1:Default
		Gui, 1:ListView, lvFoldersList
		Loop, % LV_GetCount()
		{
			LV_GetText(strThisName, A_Index, 1)
			if (strCandidateName = strThisName)
				return False
		}
	}
	else
		Loop, % arrMenus[strMenu].MaxIndex()
			if (strCandidateName = arrMenus[strMenu][A_Index].FolderName)
				return False

	return True
}
;------------------------------------------------------------


;------------------------------------------------------------
GuiAbout:
;------------------------------------------------------------

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lAboutTitle, strAppName, strAppVersion)
Gui, 2:+Owner1
str32or64 := A_PtrSize  * 8
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Link, y10 w350 vlblAboutText1, % L(lAboutText1, strAppName, strAppVersion, str32or64)
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Link, , % L(lAboutText2, strAppName)
Gui, 2:Add, Link, , % L(lAboutText3, chr(169))
Gui, 2:Font, s10 w400, Verdana
Gui, 2:Add, Link, , % L(lAboutText4)
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Button, x115 y+20 gGuiDonate, %lDonateButton%
Gui, 2:Add, Button, x150 y+20 g2GuiClose vbtnAboutClose, %lGui2Close%
GuiControl, Focus, btnAboutClose
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiDonate:
;------------------------------------------------------------

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lDonateTitle, strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Link, y10 w420, % L(lDonateText1, strAppName)
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Link, x175 w185 y+10, % L(lDonateText2, "http://code.jeanlalonde.ca/support-freeware/")
loop, 3
{
	Gui, 2:Add, Button, % (A_Index = 1 ? "y+10 Default vbtnDonateDefault " : "") . " xm w150 gButtonDonate" . A_Index, % lDonatePlatformName%A_Index%
	Gui, 2:Add, Link, x+10 w235 yp, % L(lDonatePlatformComment%A_Index%, (A_Index = 3 ? "https://flattr.com/" : ""))
}

Gui, 2:Font, s10 w700, Verdana
Gui, 2:Add, Link, xm y+20 w420, %lDonateText3%
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Link, xm y+10 w420, % L(lDonateText4, strAppName)

strDonateReviewUrl1 := "http://download.cnet.com/FoldersPopup/3000-2344_4-76062382.html"
strDonateReviewUrl2 := "http://www.portablefreeware.com/index.php?id=2557"
strDonateReviewUrl3 := "http://www.softpedia.com/get/System/OS-Enhancements/FoldersPopup.shtml"
strDonateReviewUrl4 := "http://fileforum.betanews.com/detail/Folders-Popup/1385175626/1"
strDonateReviewUrl5 := "http://www.filecluster.com/System-Utilities/Other-Utilities/Download-FoldersPopup.html"

loop, 5
	Gui, 2:Add, Link, % (A_Index = 1 ? "y+10" : "") . " x175 w185", % "<a href=""" . strDonateReviewUrl%A_Index% . """>" . lDonateReviewName%A_Index% . "</a>"

Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Button, x175 y+20 g2GuiClose vbtnDonateClose, %lGui2Close%
GuiControl, Focus, btnDonateDefault
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonDonate1:
ButtonDonate2:
ButtonDonate3:
;------------------------------------------------------------

strDonatePlatformUrl1 := "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=AJNCXKWKYAXLCV"
strDonatePlatformUrl2 := "http://www.shareit.com/product.html?productid=300628012"
strDonatePlatformUrl3 := "http://code.jeanlalonde.ca/?flattrss_redirect&id=19&md5=e1767c143c9bde02b4e7f8d9eb362b71"

StringReplace, strButton, A_ThisLabel, ButtonDonate
Run, % strDonatePlatformUrl%strButton%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiHelp:
;------------------------------------------------------------

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lHelpTitle, strAppName, strAppVersion)
Gui, 2:+Owner1
intWidth := 600
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Text, x10 y10, %strAppName%
Gui, 2:Font, s10 w400, Verdana
Gui, 2:Add, Link, x10 w%intWidth%, %lHelpTextLead%
Gui, 2:Font, s8 w400, Verdana
loop, 6
	Gui, 2:Add, Link, w%intWidth%, % lHelpText%A_Index%
Gui, 2:Add, Link, w%intWidth%, % L(lHelpText7, chr(169))
Gui, 2:Add, Button, x180 y+20 gGuiDonate, %lDonateButton%
Gui, 2:Add, Button, x+80 yp g2GuiClose vbtnHelpClose, %lGui2Close%
GuiControl, Focus, btnHelpClose
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiOptions:
;------------------------------------------------------------

intGui1WinID := WinExist("A")

;---------------------------------------
; Build Gui header
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lOptionsGuiTitle, strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:Font, s10 w700, Verdana
Gui, 2:Add, Text, x10 y10 w430 center, % L(lOptionsGuiTitle, strAppName)
Gui, 2:Font
Gui, 2:Add, Text, x10 y+10 w430 center, % L(lOptionsGuiIntro, strAppName)

; Build Hotkey Gui lines
loop, % arrIniKeyNames%0%
{
	Gui, 2:Font, s8 w700
	Gui, 2:Add, Text, x15 y+10, % arrOptionsTitles%A_Index%
	Gui, 2:Font, s9 w500, Courier New
	Gui, 2:Add, Text, x200 yp w220 h20 center 0x1000 vlblHotkeyText%A_Index%, % Hotkey2Text(strModifiers%A_Index%, strMouseButton%A_Index%, strOptionsKey%A_Index%)
	Gui, 2:Font
	Gui, 2:Add, Button, h20 yp x400 vbtnChangeHotkey%A_Index% gButtonOptionsChangeHotkey%A_Index%, %lOptionsChangeHotkey%
}

Gui, 2:Add, Text, x10 y+15 h2 w410 0x10 ; Horizontal Line > Etched Gray

Gui, 2:Font, s8 w700
Gui, 2:Add, Text, x10 y+5 w410 center, %lOptionsOtherOptions%
Gui, 2:Font

Gui, 2:Add, Text, y+10 x40, %lOptionsLanguage%
Gui, 2:Add, DropDownList, yp x+10 vdrpLanguage Sort, %lOptionsLanguageLabels%
GuiControl, ChooseString, drpLanguage, %strLanguageLabel%

Gui, 2:Add, CheckBox, y+10 x40 vblnOptionsRunAtStartup, %lOptionsRunAtStartup%
GuiControl, , blnOptionsRunAtStartup, % FileExist(A_Startup . "\" . strAppName . ".lnk") ? 1 : 0

Gui, 2:Add, CheckBox, yp x250 vblnDisplaySpecialFolders, %lOptionsDisplaySpecialFolders%
GuiControl, , blnDisplaySpecialFolders, %blnDisplaySpecialFolders%

Gui, 2:Add, CheckBox, y+10 x40 vblnDisplayMenuShortcuts, %lOptionsDisplayMenuShortcuts%
GuiControl, , blnDisplayMenuShortcuts, %blnDisplayMenuShortcuts%

Gui, 2:Add, CheckBox, yp x250 vblnDisplayRecentFolders, %lOptionsDisplayRecentFolders%
GuiControl, , blnDisplayRecentFolders, %blnDisplayRecentFolders%

Gui, 2:Add, CheckBox, y+10 x40 vblnDisplayTrayTip, %lOptionsTrayTip%
GuiControl, , blnDisplayTrayTip, %blnDisplayTrayTip%

Gui, 2:Add, CheckBox, yp x250 vblnDisplaySwitchMenu, %lOptionsDisplaySwitchMenu%
GuiControl, , blnDisplaySwitchMenu, %blnDisplaySwitchMenu%

Gui, 2:Add, CheckBox, y+10 x40 vblnPopupFix gPopupFixClicked, %lOptionsPopupFix%
GuiControl, , blnPopupFix, %blnPopupFix%

Gui, 2:Add, Text, % "y+10 x58 vlblPopupFixPositionX " . (blnPopupFix ? "" : "hidden"), %lOptionsPopupFixPositionX%
Gui, 2:Add, Edit, % "yp x+5 w36 h15 vstrPopupFixPositionX center " . (blnPopupFix ? "" : "hidden"), %arrPopupFixPosition1%
Gui, 2:Add, Text, % "yp x+5 vlblPopupFixPositionY " . (blnPopupFix ? "" : "hidden"), %lOptionsPopupFixPositionY%
Gui, 2:Add, Edit, % "yp x+5 w36 h15 vstrPopupFixPositionY center " . (blnPopupFix ? "" : "hidden"), %arrPopupFixPosition2%

; Build Gui footer
Gui, 2:Add, Button, y+20 x180 vbtnOptionsSave gButtonOptionsSave, %lGuiSave%
Gui, 2:Add, Button, yp x+15 vbtnOptionsCancel gButtonOptionsCancel, %lGuiCancel%
Gui, 2:Add, Text
GuiControl, Focus, btnOptionsSave

Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonOptionsChangeHotkey1:
ButtonOptionsChangeHotkey2:
ButtonOptionsChangeHotkey3:
ButtonOptionsChangeHotkey4:
ButtonOptionsChangeHotkey5:
;------------------------------------------------------------

StringReplace, intIndex, A_ThisLabel, ButtonOptionsChangeHotkey
intGui2WinID := WinExist("A")

Gui, 2:Submit, NoHide
Gui, 3:New, , % L(lOptionsChangeHotkeyTitle, strAppName, strAppVersion)
Gui, 3:+Owner2
Gui, 3:Font, s10 w700, Verdana
Gui, 3:Add, Text, x10 y10 w350 center, % L(lOptionsChangeHotkeyTitle, strAppName)
Gui, 3:Font

intType := 3
if InStr(arrIniKeyNames%intIndex%, "Mouse")
	intType := 1
if InStr(arrIniKeyNames%intIndex%, "Keyboard")
	intType := 2

Gui, 3:Add, Text, y+15 x10 , %lOptionsTriggerFor%
Gui, 3:Font, s8 w700
Gui, 3:Add, Text, x+5 yp , % arrOptionsTitles%intIndex%
Gui, 3:Font

Gui, 3:Add, Text, x10 y+5 w350, % lOptionsArrDescriptions%intIndex%

Gui, 3:Add, CheckBox, y+20 x100 vblnOptionsShift, %lOptionsShift%
GuiControl, , blnOptionsShift, % InStr(strModifiers%intIndex%, "+") ? 1 : 0
GuiControlGet, posTop, Pos, blnOptionsShift
Gui, 3:Add, CheckBox, y+10 x100 vblnOptionsCtrl, %lOptionsCtrl%
GuiControl, , blnOptionsCtrl, % InStr(strModifiers%intIndex%, "^") ? 1 : 0
Gui, 3:Add, CheckBox, y+10 x100 vblnOptionsAlt, %lOptionsAlt%
GuiControl, , blnOptionsAlt, % InStr(strModifiers%intIndex%, "!") ? 1 : 0
Gui, 3:Add, CheckBox, y+10 x100 vblnOptionsWin, %lOptionsWin%
GuiControl, , blnOptionsWin, % InStr(strModifiers%intIndex%, "#") ? 1 : 0

if (intType = 1)
	Gui, 3:Add, DropDownList, % "y" . posTopY . " x200 w100 vstrOptionsMouse gOptionsMouseChanged", % strMouseButtonsWithDefault%intIndex%
if (intType <> 1)
{
	Gui, 3:Add, Hotkey, % "y" . posTopY . " x200 w100 vstrOptionsKey gOptionsHotkeyChanged"
	GuiControl, , strOptionsKey, % strOptionsKey%intIndex%
}
if (intType = 3)
	Gui, 3:Add, DropDownList, % "y" . posTopY + 50 . " x200 w100 vstrOptionsMouse gOptionsMouseChanged", % strMouseButtonsWithDefault%intIndex%

Gui, 3:Add, Button, y220 x140 vbtnChangeHotkeySave gButtonChangeHotkeySave%intIndex%, %lGuiSave%
Gui, 3:Add, Button, yp x+20 vbtnChangeHotkeyCancel gButtonChangeHotkeyCancel, %lGuiCancel%
Gui, 3:Add, Text
GuiControl, Focus, btnChangeHotkeySave

Gui, 3:Show, AutoSize Center
Gui, 2:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
PopupFixClicked:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

GuiControl, % (blnPopupFix ? "Show" : "Hide"), lblPopupFixPositionX
GuiControl, % (blnPopupFix ? "Show" : "Hide"), strPopupFixPositionX
GuiControl, % (blnPopupFix ? "Show" : "Hide"), lblPopupFixPositionY
GuiControl, % (blnPopupFix ? "Show" : "Hide"), strPopupFixPositionY

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonOptionsSave:
;------------------------------------------------------------
Gui, 2:Submit

blnMenuReady := false

IfExist, %A_Startup%\%strAppName%.lnk
	FileDelete, %A_Startup%\%strAppName%.lnk
if (blnOptionsRunAtStartup)
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%strAppName%.lnk
Menu, Tray, % blnOptionsRunAtStartup ? "Check" : "Uncheck", %lMenuRunAtStartup%

IniWrite, %blnDisplayTrayTip%, %strIniFile%, Global, DisplayTrayTip
IniWrite, %blnDisplaySpecialFolders%, %strIniFile%, Global, DisplaySpecialFolders
IniWrite, %blnDisplayRecentFolders%, %strIniFile%, Global, DisplayRecentFolders
IniWrite, %blnDisplaySwitchMenu%, %strIniFile%, Global, DisplaySwitchMenu
IniWrite, %blnPopupFix%, %strIniFile%, Global, PopupFix
IniWrite, %blnDisplayMenuShortcuts%, %strIniFile%, Global, DisplayMenuShortcuts

IniWrite, %strPopupFixPositionX%`,%strPopupFixPositionY%, %strIniFile%, Global, PopupFixPosition
arrPopupFixPosition1 := strPopupFixPositionX
arrPopupFixPosition2 := strPopupFixPositionY

strLanguageCodePrev := strLanguageCode
strLanguageLabel := drpLanguage
loop, %arrOptionsLanguageLabels0%
	if (arrOptionsLanguageLabels%A_Index% = strLanguageLabel)
		{
			strLanguageCode := arrOptionsLanguageCodes%A_Index%
			break
		}
IniWrite, %strLanguageCode%, %strIniFile%, Global, LanguageCode

; if language changed, offer to restart the app
if (strLanguageCodePrev <> strLanguageCode)
{
	MsgBox, 52, %strAppName%, % L(lReloadPrompt, strLanguageLabel, strAppName)
	IfMsgBox, Yes
	{
		Gosub, RestoreBackupMenuObjects
		Reload
	}
}	

; else rebuild Folders menus w/ or w/o optional folders and shortcuts

for strMenuName, arrMenu in arrMenus
{
	Menu, %strMenuName%, Add
	Menu, %strMenuName%, DeleteAll
	arrMenu := ; free object's memory
}
Gosub, BuildFoldersMenus

Gosub, 2GuiClose

blnMenuReady := true

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonOptionsCancel:
;------------------------------------------------------------
Gosub, 2GuiClose

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
ButtonChangeHotkeySave1:
ButtonChangeHotkeySave2:
ButtonChangeHotkeySave3:
ButtonChangeHotkeySave4:
ButtonChangeHotkeySave5:
;------------------------------------------------------------
Gui, 3:Submit

StringReplace, intIndex, A_ThisLabel, ButtonChangeHotkeySave

StringSplit, arrIniVarNames, strIniKeyNames, |

strHotkey := Trim(strOptionsKey . strOptionsMouse)

if StrLen(strHotkey)
{
	Hotkey, % "$" . arrHotkeyVarNames%intIndex%, Off, UseErrorLevel ; UseErrorLevel in case we had an invalid key name in the ini file

	if (blnOptionsWin)
		strHotkey := "#" . strHotkey
	if (blnOptionsAlt)
		strHotkey := "!" . strHotkey
	if (blnOptionsShift)
		strHotkey := "+" . strHotkey
	if (blnOptionsCtrl)
		strHotkey := "^" . strHotkey
	IniWrite, % strHotkey, %strIniFile%, Global, % arrIniVarNames%intIndex%
	
}
else
	Oops(L(lOptionsNoKeyOrMouseSpecified, arrIniVarNames%intIndex%))

Gosub, LoadIniHotkeys ; reload ini variables and reset hotkeys

GuiControl, 2:, lblHotkeyText%intIndex%, % Hotkey2Text(strModifiers%intIndex%, strMouseButton%intIndex%, strOptionsKey%intIndex%)

Gosub, 3GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonChangeHotkeyCancel:
;------------------------------------------------------------
Gosub, 3GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
3GuiClose:
3GuiEscape:
;------------------------------------------------------------

Gui, 2:-Disabled
Gui, 3:Destroy
WinActivate, ahk_id %intGui2WinID%

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

Gosub, LoadOneMenuToGui

Gui, 1:ListView, lvDialogsList
LV_Delete()
Loop, % arrDialogs.MaxIndex()
	LV_Add(, arrDialogs[A_Index])
LV_Modify(1, "Select Focus")
LV_ModifyCol(1, "AutoHdr")

GuiControl, Focus, lvFoldersList

return
;------------------------------------------------------------


;------------------------------------------------------------
LoadOneMenuToGui:
;------------------------------------------------------------

Gui, 1:ListView, lvFoldersList
LV_Delete()

Loop, % arrMenus[strCurrentMenu].MaxIndex()
	if StrLen(arrMenus[strCurrentMenu][A_Index].SubmenuFullName) ; this is a submenu
		LV_Add(, arrMenus[strCurrentMenu][A_Index].FolderName, lGuiSubmenuLocation, arrMenus[strCurrentMenu][A_Index].MenuName, arrMenus[strCurrentMenu][A_Index].SubmenuFullName)
	else ; this is a folder
		LV_Add(, arrMenus[strCurrentMenu][A_Index].FolderName, arrMenus[strCurrentMenu][A_Index].FolderLocation, arrMenus[strCurrentMenu][A_Index].MenuName, "")
LV_Modify(1, "Select Focus")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

GuiControl, , drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"

return
;------------------------------------------------------------


;------------------------------------------------------------
SaveCurrentListviewToMenuObject:
;------------------------------------------------------------

arrMenus[strCurrentMenu] := Object() ; re-init current menu array
Gui, 1:ListView, lvFoldersList

Loop % LV_GetCount()
{
	LV_GetText(strName, A_Index, 1)
	LV_GetText(strLocation, A_Index, 2)
	LV_GetText(strMenu, A_Index, 3)
	LV_GetText(strSubmenu, A_Index, 4)

	objFolder := Object() ; new menu item
	objFolder.FolderName := strName ; display name of this menu item
	objFolder.FolderLocation := strLocation ; path for this menu item
	objFolder.MenuName := strCurrentMenu ; parent menu of this menu item - do not use strMenu because lack of lMainMenuName
	objFolder.SubmenuFullName := strSubmenu ; if menu, full name of the submenu
	arrMenus[strCurrentMenu].Insert(objFolder) ; add this menu item to parent menu - do not use strMenu because lack of lMainMenuName
}

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildMenuTreeDropDown(strMenu, strDefaultMenu, strSkipSubmenu := "")
; recursive function
;------------------------------------------------------------
{
	strList := strMenu
	if (strMenu = strDefaultMenu)
		strList := strList . "|" ; default value

	Loop, % arrMenus[strMenu].MaxIndex()
		if StrLen(arrMenus[strMenu][A_Index].SubmenuFullName) ; this is a submenu
			if (arrMenus[strMenu][A_Index].SubmenuFullName <> strSkipSubmenu) ; skip if under edited submenu
				strList := strList . "|" . BuildMenuTreeDropDown(arrMenus[strMenu][A_Index].SubmenuFullName, strDefaultMenu, strSkipSubmenu) ; recursive call
	return strList
}
;------------------------------------------------------------


;------------------------------------------------------------
OpenFavorite:
;------------------------------------------------------------

if (blnDisplayMenuShortcuts)
	StringTrimLeft, strThisMenu, A_ThisMenuItem, 3 ; remove "&1 " from menu item
else
	strThisMenu := A_ThisMenuItem
strLocation := GetPathFor(A_ThisMenu, strThisMenu)

if (blnDiagMode)
{
	Diag("A_ThisHotkey", A_ThisHotkey)
	Diag("Navigate", "RegularFolder")
	Diag("Path", strLocation)
}

if InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "Shell.Application")
	
	if (A_OSVersion = "WIN_XP")
		ComObjCreate("Shell.Application").Explore(strLocation)
	else
		Run, Explorer %strLocation%
	; http://msdn.microsoft.com/en-us/library/bb774094
	; ComObjCreate("Shell.Application").Explore(strLocation)
	; ComObjCreate("WScript.Shell").Exec("Explorer.exe /e /select," . strLocation) ; not tested on XP
	; ComObjCreate("Shell.Application").Open(strLocation)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
	
	if (blnDiagMode)
		Diag("NavigateResult", ErrorLevel)
}
else if WindowIsAnExplorer(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateExplorer")
	
	NavigateExplorer(strLocation, strTargetWinId)
	
	if (blnDiagMode)
		Diag("NavigateResult", ErrorLevel)
}
else if WindowIsConsole(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateConsole")
	
	NavigateConsole(strLocation, strTargetWinId)

	if (blnDiagMode)
		Diag("NavigateResult", ErrorLevel)
}
else if WindowIsFreeCommander(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateFreeCommander")
	
	NavigateFreeCommander(strLocation, strTargetWinId)

	if (blnDiagMode)
		Diag("NavigateResult", ErrorLevel)
}
else
{
	if (blnDiagMode)
	{
		Diag("Navigate", "NavigateDialog")
		Diag("TargetClass", strTargetClass)
	}
	
	NavigateDialog(strLocation, strTargetWinId, strTargetClass)

	if (blnDiagMode)
		Diag("NavigateResult", ErrorLevel)
}

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

if (blnDiagMode)
{
	Diag("A_ThisHotkey", A_ThisHotkey)
	Diag("Navigate", "SpecialFolder")
	Diag("SpecialFolder", intSpecialFolder)
}

if InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass)
	ComObjCreate("Shell.Application").Explore(intSpecialFolder)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strTargetClass)
	NavigateExplorer(intSpecialFolder, strTargetWinId)
else ; this is the console, FreeCommander or a dialog box
{
	if (intSpecialFolder = 0)
		strLocation := A_Desktop
	else if (intSpecialFolder = 5)
		strLocation := A_MyDocuments
	else if (intSpecialFolder = 39)
	{
		if (A_OSVersion = "WIN_XP")
		{
			strLanguage := SubStr(A_Language, 3, 2) ; "040C" -> "0C"
			; For A_Language values locale list see http://msdn.microsoft.com/en-ca/goglobal/bb895996.aspx
			; "09" English, "0C" French, "0A" spanish, "07" German
			if (strLanguage = "0C") ; French
				strLocation := A_MyDocuments . "\Mes images"
			else if (strLanguage = "0A") ; Spanish
				strLocation := A_MyDocuments . "\My imágenes"
			else if (strLanguage = "09") ; English
				strLocation := A_MyDocuments . "\My Pictures"
			else ; will work only for locales using "My Pictures" - check if folder exists
				strLocation := A_MyDocuments . "\My Pictures"
		}
		else ; WIN_7+
		{
			; do not use: StringReplace, strPath, A_MyDocuments, Documents, Pictures
			; because A_MyDocument could contain a "Documents" string before the final folder
			StringLeft, strLocation, A_MyDocuments, % StrLen(A_MyDocuments) - StrLen("Documents")
			strLocation := strLocation . "Pictures" ; should work with all languages - tested in En and Fr
		}
	}	
	else ; we do not support this special folder
	{
		if (blnDiagMode)
			Diag("NavigateResult", "SpecialFolderNotSupported")
		return
	}

	if !StrLen(FileExist(strLocation))
		strLocation := A_MyDocuments
		; Replace wrong path bby My Documents path. 
		; Draft of message for user:
		; This special folder is not supported in your current language settings.
		; To improve this, you could post a message to ~1~ support mentioning your language code (~2~) and the path of the selected special folder on your system.

	if WindowIsConsole(strTargetClass)
		NavigateConsole(strLocation, strTargetWinId)
	else if WindowIsFreeCommander(strTargetClass)
		NavigateFreeCommander(strLocation, strTargetWinId)
	else
		NavigateDialog(strLocation, strTargetWinId, strTargetClass)
}

if (blnDiagMode)
	Diag("NavigateResult", ErrorLevel)

return
;------------------------------------------------------------


;------------------------------------------------------------
OpenRecentFolder:
;------------------------------------------------------------

if InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass)
	ComObjCreate("Shell.Application").Explore(objRecentFolders[A_ThisMenuItemPos])
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else if WindowIsAnExplorer(strTargetClass)
	NavigateExplorer(objRecentFolders[A_ThisMenuItemPos], strTargetWinId)
else ; this is the console, FreeCommander or a dialog box
	if WindowIsConsole(strTargetClass)
		NavigateConsole(objRecentFolders[A_ThisMenuItemPos], strTargetWinId)
	else if WindowIsFreeCommander(strTargetClass)
		NavigateFreeCommander(objRecentFolders[A_ThisMenuItemPos], strTargetWinId)
	else
		NavigateDialog(objRecentFolders[A_ThisMenuItemPos], strTargetWinId, strTargetClass)

return
;------------------------------------------------------------


;------------------------------------------------------------
SwitchExplorer:
;------------------------------------------------------------

if InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass)
	ComObjCreate("Shell.Application").Explore(objExplorersLocation[A_ThisMenuItemPos])
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
else
	WinActivate, % "ahk_id " . objExplorersWinId[A_ThisMenuItemPos]

return
;------------------------------------------------------------


;------------------------------------------------------------
SwitchDialog:
;------------------------------------------------------------

NavigateDialog(objExplorersLocation[A_ThisMenuItemPos], strTargetWinId, strTargetClass)

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
GetPathFor(strMenu, strName)
;------------------------------------------------------------
{
	global
	
	Loop, % arrMenus[strMenu].MaxIndex()
		if (strName = arrMenus[strMenu][A_Index].FolderName)
			return arrMenus[strMenu][A_Index].FolderLocation
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

	blnCanOpenFavorite := WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or WindowIsConsole(strClass) or WindowIsDialog(strClass) or WindowIsFreeCommander(strClass)

	if (blnDiagMode)
	{
		Diag("MouseOrKeyboard", strMouseOrKeyboard)
		WinGetTitle strDiag, % "ahk_id " . strWinId
		Diag("WinTitle", strDiag)
		Diag("WinId", strWinId)
		Diag("Class", strClass)
		Diag("CanOpenFavorite", blnCanOpenFavorite)
	}
	return blnCanOpenFavorite
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
WindowIsFreeCommander(strClass)
;------------------------------------------------------------
{
	return InStr(strClass, "FreeCommanderXE")
}
;------------------------------------------------------------


;------------------------------------------------------------
DialogIsSupported(strWinId)
;------------------------------------------------------------
{
	global
	
	WinGetTitle, strDialogTitle, ahk_id %strWinId%
	loop, % arrDialogs.MaxIndex()
		if InStr(strDialogTitle, arrDialogs[A_Index])
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
	if !InStr(varPath, "#") ; prevent the hash bug in Shell.Application
	{
		intCountMatch := 0
		For pExp in ComObjCreate("Shell.Application").Windows
		{
			if (pExp.hwnd = strWinId)
				if varPath is integer ; ShellSpecialFolderConstant
				{
					intCountMatch := intCountMatch + 1
					try pExp.Navigate2(varPath)
					catch, objErr
						Oops(lNavigateSpecialError, varPath)
				}
				else
				{
					intCountMatch := intCountMatch + 1
					; Was "try pExp.Navigate("file:///" . varPath)" - removed "file:///" to allow UNC (e.g. \\my.server.com@SSL\DavWWWRoot\Folder\Subfolder)
					try pExp.Navigate(varPath)
					catch, objErr
						; Note: If an error occurs in Navigate, the error message is given by Navigate itself and this script does not
						; receive an error notification. From my experience, the following line would never be executed.
						Oops(lNavigateFileError, varPath)
				}
		}
		if !(intCountMatch) ; for Explorer add-ons like Clover (verified - it now opens the folder in a new tab), others?
			if varPath is integer ; ShellSpecialFolderConstant
				ComObjCreate("Shell.Application").Explore(varPath)
			else
				Run, Explorer %varPath%
	}
	else
	{
		; Workaround for the hash (aka Sharp / "#") bug in Shell.Application - occurs only when navigatin in the current Explorer window
		; see http://stackoverflow.com/questions/22868546/navigate-shell-command-not-working-when-the-path-includes-an-hash
		Send, {F4}{Esc}
		Sleep, 500
		ControlSetText, Edit1, %varPath%, A
		ControlSend, Edit1, {Enter}, A
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateConsole(strLocation, strWinId)
;------------------------------------------------------------
{
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
	SendInput, CD /D %strLocation%{Enter}
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateFreeCommander(strLocation, strWinId)
;------------------------------------------------------------
{
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
	{
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
		Sleep, 200
	}
	MouseGetPos, , , , strCommanderControl, 1
	Sleep, 200
	ControlFocus, %strCommanderControl%
	SetKeyDelay, 200
	Send, !g
	SetKeyDelay, 0
	Send, %strLocation%
	Sleep, 200
	Send, {Enter}
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateDialog(strLocation, strWinId, strClass)
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
		{
			if (blnDiagMode)
				Diag("NavigateDialog", "Error: #32770 Edit1 and Edit2 controls not visible")
			return
		}
	Else if InStr(strClass, "bosa_sdm_") ; for some MS office dialog windows, which are not #32770 class
		if ControlIsVisible("ahk_id " . strWinId, "Edit1")
			strControl := "Edit1"
			; if "Edit1" control exists, it is the right choice
		Else if ControlIsVisible("ahk_id " . strWinId, "RichEdit20W2")
			strControl := "RichEdit20W2"
			; some MS office dialogs don't have "Edit1" control, but they have "RichEdit20W2" control, which is then the right choice.
		Else ; if above fails, just return and do nothing.
		{
			if (blnDiagMode)
				Diag("NavigateDialog", "Error: bosa_sdm Edit1 and RichEdit20W2 controls not visible")
			return
		}
	Else ; in all other cases, open a new Explorer and return from this function
	{
		if (A_OSVersion = "WIN_XP")
			ComObjCreate("Shell.Application").Explore(strLocation)
		else
			Run, Explorer %strLocation%
		; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
		if (blnDiagMode)
			Diag("NavigateDialog", "Not #32770 or bosa_sdm: open New Explorer")
		return
	}

	if (blnDiagMode)
	{
		Diag("NavigateDialogControl", strControl)
		Diag("NavigateDialogPath", strLocation)
	}
			
	;===In this part (if we reached it), we'll send strLocation to control and restore control's initial text after navigating to specified folder===
	ControlGetText, strPrevControlText, %strControl%, ahk_id %strWinId% ; we'll get and store control's initial text first
	
	ControlSetTextR(strControl, strLocation, "ahk_id " . strWinId) ; set control's text to strLocation
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

	if (blnDiagMode)
		Diag("NavigateDialog", "Finished")
}


;------------------------------------------------------------
ControlIsVisible(strWinTitle, strControlClass)
/*
Adapted from ControlIsVisible(WinTitle,ControlClass) by Learning One
http://ahkscript.org/boards/viewtopic.php?f=5&t=526&start=20#p4673
*/
;------------------------------------------------------------
{
	; used in Navigator
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
{
	; used in Navigator. More reliable ControlSetFocus
	Loop, %intTries%
	{
		ControlFocus, %strControl%, %strWinTitle% ; focus control
		Sleep, % (50 * intTries) ; JL added "* intTries"
		ControlGetFocus, strFocusedControl, %strWinTitle% ; check
		if (strFocusedControl = strControl) ; if OK
		{
			if (blnDiagMode)
				Diag("ControlSetFocusR Tries", A_Index)
			return True
		}
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
{
	; used in Navigator. More reliable ControlSetText
	Loop, %intTries%
	{
		ControlSetText, %strControl%, %strNewText%, %strWinTitle% ; set
		Sleep, % (50 * intTries) ; JL added "* intTries"
		ControlGetText, strCurControlText, %strControl%, %strWinTitle% ; check
		if (strCurControlText = strNewText) ; if OK
		{
			if (blnDiagMode)
				Diag("ControlSetTextR Tries", A_Index)
			return True
		}
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


;------------------------------------------------------------
Hotkey2Text(strModifiers, strMouseButton, strOptionKey)
;------------------------------------------------------------
{
	global
	
	str := ""
	loop, parse, strModifiers
	{
		if (A_LoopField = "!")
			str := str . lOptionsAlt . "+"
		if (A_LoopField = "^")
			str := str . lOptionsCtrl . "+"
		if (A_LoopField = "+")
			str := str . lOptionsShift . "+"
		if (A_LoopField = "#")
			str := str . lOptionsWin . "+"
	}
	if StrLen(strMouseButton)
		str := str . GetText4MouseButton(strMouseButton)
	if StrLen(strOptionKey)
	{
		StringUpper, strOptionKey, strOptionKey
		str := str . strOptionKey
	}

	return str
}
;------------------------------------------------------------


;------------------------------------------------------------
GetText4MouseButton(strSource)
; Returns the string in arrMouseButtonsText at the same position of strSource in arrMouseButtons
;------------------------------------------------------------
{
	global
	
	loop, %arrMouseButtons0%
		if (strSource = arrMouseButtons%A_Index%)
			return arrMouseButtonsText%A_Index%
}
;------------------------------------------------------------


;------------------------------------------------------------
GetIniName4Hotkey(strSource)
; Returns the string in arrIniKeyNames at the same position of strSource in arrHotkeyVarNames content
;------------------------------------------------------------
{
	global

	loop, %arrHotkeyVarNames0%
		if (strSource = "$" . arrHotkeyVarNames%A_Index%)
			return arrIniKeyNames%A_Index%
}
;------------------------------------------------------------



;============================================================
; TOOLS
;============================================================

;------------------------------------------------
Oops(strMessage, objVariables*)
;------------------------------------------------
{
	Gui, 1:+OwnDialogs
	MsgBox, 48, % L(lOopsTitle, strAppName, strAppVersion), % L(strMessage, objVariables*)
}
; ------------------------------------------------


;------------------------------------------------
L(strMessage, objVariables*)
;------------------------------------------------
{
	Loop
	{
		if InStr(strMessage, "~" . A_Index . "~")
			StringReplace, strMessage, strMessage, ~%A_Index%~, % objVariables[A_Index], A
 		else
			break
	}
	
	return strMessage
}
;------------------------------------------------


;------------------------------------------------------------
GetDeepestFolderName(strLocation)
;------------------------------------------------------------
{
	SplitPath, strLocation, strDeepestName, , , , strDrive
	if !StrLen(strDeepestName) ; we are probably at the root of a drive
		return strDrive
	else
		return strDeepestName
}
;------------------------------------------------------------


;------------------------------------------------------------
ValueIsInObject(strValue, obj)
;------------------------------------------------------------
{
	loop, % obj.MaxIndex()
		if (strValue = obj[A_Index])
			return true
		
	return false
}
;------------------------------------------------------------


;------------------------------------------------
Diag(strName, strData)
;------------------------------------------------
{
	FormatTime, strNow, %A_Now%, yyyyMMdd@HH:mm:ss
	loop
	{
		FileAppend, %strNow%.%A_MSec%`t%strName%`t%strData%`n, %strDiagFile%
		if ErrorLevel
			Sleep, 20
	}
	until !ErrorLevel or (A_Index > 50) ; after 1 second (20ms x 50), we have a problem
}
;------------------------------------------------


;------------------------------------------------------------
NextMenuShortcut(ByRef intShortcut, blnMainMenu)
;------------------------------------------------------------
{
	if (intShortcut < 10)
		strShortcut := intShortcut ; 0 .. 9
	else
	{
		while (blnMainMenu and InStr(lMenuReservedShortcuts, Chr(intShortcut + 55)))
			intShortcut := intShortcut + 1
		strShortcut := Chr(intShortcut + 55) ; Chr(10 + 55) = "A" .. Chr(35 + 55) = "Z"
	}
	intShortcut := intShortcut + 1
	return strShortcut
}
;------------------------------------------------------------


;------------------------------------------------
UriDecode(str)
; by polyethene
; http://www.autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters/?p=112822
;------------------------------------------------
{
	Loop
		If RegExMatch(str, "i)(?<=%)[\da-f]{1,2}", hex)
			StringReplace, str, str, `%%hex%, % Chr("0x" . hex), All
		Else
			Break
	return str
}
