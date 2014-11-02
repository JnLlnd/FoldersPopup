/*
Bugs:

To-do for v4:
- Write DOpus add-in to list folders including special folders
- Save groups with special folders in DOpus to ini file
- Load groups with special folders in DOpus from ini file
- adapt website to new default hotkey Win-A and Shift-Win-A
- adapt website to new standard setup procedure
*/
;===============================================
/*
	FoldersPopup
	Written using AutoHotkey_L v1.1.09.03+ (http://ahkscript.org/)
	By Jean Lalonde (JnLlnd on AHKScript.org forum)
	
	Based on DirMenu v2 by Robert Ryan (rbrtryn on AutoHotkey.com forum)
	http://www.autohotkey.com/board/topic/91109-favorite-folders-popup-menu-with-gui/
	who was maybe inspired by Savage's script FavoriteFolders
	http://www.autohotkey.com/docs/scripts/FavoriteFolders.htm
	or Rexx version Folder Menu
	http://www.autohotkey.com/board/topic/13392-folder-menu-a-popup-menu-to-quickly-change-your-folders/


	Version: 3.9.1 BETA (2014-10-XX)
	* New setup procedure with standard Install / Uninstall procedures (using Inno Setup) - keeping a separate zip file for portable version
	* Adapt Run at startup shortcut for Inno Setup by using the working directory instead of the script directory
	* Create a unique environment code (mutex) to allow Inno Setup to detect if FP is running before uninstall or update
	* Changed default FP hotkeys Windows-K and Shift-Windows-K to Windows-A and Shift-Windows-A (Windows-K is a reserved shortcut in Win 8.1) - configs of actual users are not changed
	* Add the option "Use tabs" for DirectoryOpus users to choose to open new folders in new tab (new default) or in a new lister
	* After DOpusRt opens a folder in a new tab, activate that window
	* Change Group menu label to "Group of folders"
	* Support Group menu of Explorer and DOpus windows containing the same folder
	* Support saving multiple windows (Explorers or DOpus) containing the same folder
	* Create objects to get special folders class id by name and name by class id
	* Save groups with special folders to ini file
	* Load groups with special folders from ini file
	* Fix a bug with labels when changing the hotkey for Recent folders menu and Settings windows


	Version: 3.3 (2014-10-24)
	
	TotalCommander integration
	* automatic detection for Total Commander support
	* Total Commander configuration in Options
	* ini configuration for TotalCommander window
	* add a checkbox in options to let Total Commander users choose to open new folders (Shift-Middle-Mouse) in a new tab or in a new window
	* new TotalCommanderUseTabs and TotalCommanderNewTabOrWindow switches in ini file
	* show popup menu in TotalCommander windows
	* add this folder from Total Commander window
	* navigate regular and special folder in TotalCommander existing window
	* open regular and special folder in new TotalCommander window or tab according to TotalCommanderNewTabOrWindow
	* inserts small delays when opening TC special folders to improve reliability
	* disable Switch menu first time TC is enabled until the tabs issue is resolved in TC
	
	Other changes
	* addition of Swedish language, thanks to Åke Engelbrektson
	* fix a bug when user select a hotkey replacement for Middle-mouse button that involves a modifier key (e.g. Shift+Right-click)
	* fix bug with icons on Windows Server (disable icons)
	* fix bug making new folders opening in Explorer instead of Total Commander (or Directory Opus) when called from the Tray left-click menu
	* change DOpus command to open a new lister to Go with NEW parameter

	Retained for v4
	* save group GUI with selector, group name of load options
	* save group GUI improvements
	* load group and read replace setting
	* close other Explorers, TC and DO before loading a group when group setting is replace
	* open Explorers instances in a group
	* open Directory Opus instances in a group

	Version: 3.2.7.2 BETA (2014-10-13)
	Published in v3.3 prod
	* add a checkbox in options to let Total Commander users choose to open new folders (Shift-Middle-Mouse) in a new tab or in a new window
	* fix bug making new folders opening in Explorer instead of Total Commander (or Directory Opus) when called from the Tray left-click menu
	* inserts small delays when opening TC special folders to improve reliability
	
	Version: 3.2.7.1 BETA (2014-10-12) starting version number used for beta testing phase

	Published in v3.3 prod
	* automatic detection for Total Commander support
	* Total Commander configuration in Options
	* ini configuration for TotalCommander window
	* special TotalCommanderNewTabOrWindow switch in ini file
	* show popup menu in TotalCommander windows
	* add this folder from Total Commander window
	* navigate regular and special folder in TotalCommander existing window
	* open regular and special folder in new TotalCommander window or tab according to TotalCommanderNewTabOrWindow
	* change DOpus command to open a new lister to Go with NEW parameter

	Retained for v4
	* reorg Switch menu taking SwitchExplorer to Switch menu level, with Switch in dialog box at the bottom of Switch menu
	* rename Switch to Explorers
	* integrate with DOpus listers in Explorers menu
	* save and restore groups of Explorers
	* save and restore groups with positions
	* add Save this group and Load a group menus to Explorers menu
	* add Groups button to main Gui
	* TotalCommander support in Explorers menu (if tabs issue resolved)

	Version: 3.2.2 (2014-10-02)
	* fix layout in options gui
	* remove support for MS Office 2003/2007 file dialog boxes
	* German language update

	Version: 3.2.1 (2014-09-20)
	* When Explorer replacement activated in DOpus, ghost Explorer in the Switch Explorer menu skipped
	* Removed Flattr from donation platforms
	* Remove Switch Explorer support for DOpus listers containing an FTP folder (until issue resolved - https://github.com/JnLlnd/FoldersPopup/issues/84)
	* Addition of the korean language - thanks to Om Il-Sung (Dollnamul)

	Version: 3.2 (2014-09-16)
	
	Directory Opus integration
	* collect info about opened DOpus listers using DOpusRt
	* collect info about opened Explorers and DOpus listers in two objects, merge the two sets of folders, remove duplicates and build Switch menus
	* adapt SwitchExplorer and SwitchDialog to new object model
	* switch explorer in DOpus using DOpusRt, switch to DOpus if 2 panes or multiple tabs
	* handling coll:// DOpus windows like search results in Switch Menu
	* use DOpus icons for listers in Switch Explorer
	* enable special folders menus when target window is Directory Opus and navigate to special folders using DOpusRt and built-in aliases
	* navigate folders and recent folders in current lister using DOpusRt
	* open new lister using DOpusRt
	* prompt at startup to activate DOpusRt if DOpus found under Program Files
	* when Add This Folder, read current folder using DOpusRt
	
	Other changes
	* new option to show the popup menu near the mouse pointer, in the active window or at a fix position
	* prevent intermittent Windows bug showing an error when building recent folders menu if an external drive has been removed
	* setting the image and recent items special folders reading the Registry for a solution working in all Windows locales
	* fix bug when showing special folders names in Switch menus
	* fix bug when duplicate folders were found in Switch menus
	* prevent paths longer than 260 chars in Switch menu from causing an error
	* limit menu name to 250 chars maximum in add/edit folder dialog box
	* different ini variable LatestVersionSkippedBeta to remember latest skipped version in beta mode

	Version: 3.1.3 (2014-09-07)
	* bug fix: make all special folders menu items work when popup menu is activated from the tray icon
	* improve handling of the hash (aka Sharp / "#") bug in Shell.Application (see v1.2.6)
	* fix bug when navigating in a CMD window with path including AHK reserved chars
	
	Version: 3.1.2 (2014-09-03)
	* Menu icons now supporting Windows Vista
	* Stop building recent folders menu at startup (unnecessary since this menu is refreshed on demand)
	
	Version: 3.1.1 (2014-08-30)
	* Fix a bug that created the diag file even when diagnostic mode was off
	* Stop forcing the working directory to be the app's directory.
	* Note 1: User can set the working directory of his choice by creating a Windows shortcut and setting the "Start in:" option.
	* Note 2: By default and when user enable the "Run at startup" option, the working directory is the app's directory.
	
	Version: 3.1 (2014-08-29)
	* First public release of Folders Popup v3
	* Fix a bug in Switch in dialog box menu
	
	Version: 3.0.12 (2014-08-27)
	* German and Dutch translation update (Thanks to Edgar "Fast Edi" Hoffmann and Pieter Dejonghe)
	* Left click on Tray icon to show favorites menu
	
	Version: 3.0.11 (2014-08-24)
	* fix an icon error under WinXP
	
	Version: 3.0.10 (2014-08-23)
	* fix bug when selecting a mouse hotkey after None was selected for that hotkey
	* in Change Hotkey, unselect modifiers when None is selected as mouse trigger
	* additional text to clarify triggers in Settings, Options
	* new menu icon for submenus

	Version: 3.0.9 (2014-08-22)
	* replaces Send command with SendInput
	* fix bug when navigating to network folder in DOpus
	* add popup menu and color to tray menu
	
	Version: 3.0.8 (2014-08-20)
	* add type of favorites for links, display default browser icon for link favorites and open links in default browser
	* fix bug with DOpus when path includes AHK reserved chars
	* better support of DOpus when in dual listers
	
	Version: 3.0.7 (2014-08-18)
	* make display icons optional, refactor Add Menu commands in a centralized function
	* allow to select no mouse trigger for popup menu, add None to the dropdown list in Change hotkey window
	* add mouse or keyboard hotkey to open the recent folders list
	* fix error when icon location contains %1
	* fix error when assigning color to an empty submenu
	* fix a v2 bug with shortcuts numbers increment in Switch menus
	
	Version: 3.0.6 (2014-07-26)
	* Redesign of buttons in Settings
	* Addition to ini file of themes with colors for dialog boxes and menu
	* Implementation of colors to menus and dialog boxes
	* Add option in Settings/Options to select theme
	
	Version: 3.0.5 (2014-07-23)
	* fix a v2 bug allowing editing in Settings with no item selected
	* fix a v3.0.2 bug when adding an item to a menu other than the current menu in Settings
	* change cursor to hand for all buttons in Settings
	* refactor (merge) Add and Edit favorites GUI and Save commands (no change visible to users)
	
	Version: 3.0.4 (2014-07-21)
	* fix a bug when adding a menu and numeric shortcuts are active
	* lighter tray tip message after menu is updated in settings
	* fix a bug when retrieving icons for documents
	* change cursor for an hand for all buttons in Gui
	* support icons for document being executable files
	
	Version: 3.0.3 (2014-07-20)
	* remove "supported dialog boxes" management
	* in gui remove listview, add/edit/remove buttons, reposition other buttons
	* remove add dialog box menu, save dialog box, dialog is supported function
	
	Version: 3.0.2 (2014-07-19)
	* add favorite type "F" folder, "D" document or "S" submenu and refactor all
	* remove or add ... to main menu items
	* manage icons resource at init, supporting XP and Win7+
	* include parent menu dropdown list when add favorite
	* fix old 2.0 bug not detecting name already used when adding from add this folder
	* menu icon size default size to 16 for XP and 24 for other OS
	
	Version: 3.0.1 (2014-07-15)
	* do not check if network favorites exist
	* error icon when local favorite does not exist (removed feature)
	* error message when unavailable local favorite is selected in popup menu
	* traytip status when refreshing menus
	
	Version: 3.0.0 (2014-07-14)
	* support favorite documents as popup menu items, add Document radio button to add dialog box
	* when adding document, suggest short name for menu
	* when menu item is a document, launch it with Run
	* add icons to folders menu, submenus, documents and special folders
	* add Settings Option for menu icon size, default size to 24
	* keep the regular tray icon when suspended
	* implement Exit tray menu
	* disable separator editing
	* adapt labels to "favorites" instead of "folders"
	* build function to auto-center action buttons in Gui
	
	Version: 2.2.1 (2014-07-11)
	* fix bug when adding a folder to a submenu using drag and drop
	* add an incentive message about drag and drop at the bottom of Settings window
	* ignore submenu change in Settings when user select the current menu

	Version: 2.2 (2014-07-06)
	* support drag and drop to add favorite
	* make the cursor change to a hand when the mouse pointer is over buttons or clickable text in Settings dialog box (tried to also implement tooltips but even with a timer, it flickers too much)
	* Recent folders menu now shown in a detached menu, at the calling popup menu location, refreshed each time it is opened, with tooltip while refreshing
	* fix a bug with number of Recent folders hide/display in Settings, Options
	* fix layout bug in edit folder dialog box
	* fix bug with Switch to Explorer opening a new window
	* replace PCAstuces review URL with Freewares & Tutos
	
	Version: 2.1.1 (2014-06-25)
	* complete translation of mouse button names
	* fix bug when changing Settings shortcut
	* fix PCAstuces URL missing
	
	Version: 2.1 (2014-06-17)
	* when adding this folder, select in which menu to add the new folder
	* new button when edit menu entry to open this menu
	* in edit folder dialog box, set focus to and select folder name
	* on-demand recent folders update to keep the popup menu snappy regardless of the number of recent items to parse or the performance of the PC
	* option in settings to choose the number of recent folders in popup menu, now default to 10
	* refactor (code merge) of GuiAddFavoriteSave and GuiEditFavoriteSave
	* allow to add this folder from a network folder starting with "\\"
	* fix bug with up arrow to go to parent menu
	* addition of Dutch translation (thanks to Pieter Dejonghe!)
	* fix missing translations
	
	Version: 2.0.3 (2014-06-06)
	* fix bugs with switch folders and recent folders options
	* update german translation
	
	Version: 2.0.2 (2014-06-03)
	* improve performance of Recent Folders menu building, process only recent folders in recent items
	* fix bug when a recent folder is not available (only XP?)
	* fix header bug in diagnostic mode
	
	Version: 2.0.1 (2014-06-01)
	* complete german translation
	* fix language typos
	
	Version: 2.0.0 (2014-05-28)
	* see all additions from v1.5 ALPHA to v1.9 BETA
	
	Version: v1.9 BETA (not to be released) (2014-05-27)
	* fix bug missing error message and other language minor changes
	* reorder popup menu and place settings, add this folder and support freeware menus at the end of main menu
	* reorder checkboxes in GuiOptions
	* support recent folders on Win XP
	* loading language files and images to the exe files
	* create a "Switch..." submenu for "Switch to Explorer" and "Switch in dialog box"options
	* allow "Switch to Explorer" menu to open a new window (with combining Middle mouse button with the Shift key)
	* prevent app from running directly from the zip file or running in a write-protected folder
	* new "Support freeware" dialog box and options
	* internal changes in the check for update function
	* better error handling if error occurs during ComObjCreate, situation occurring when Directory Opus is running (tested with v11.4)
	* basic support for Directory Opus v11.4 (navigate and add this folder) similar to FreeCommander XE, fix bugs in FreeCommanderXE support

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
	* add switch submenu to activate any other open Explorer
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
;========================================================================================================================
; --- COMPILER DIRECTIVES ---
;========================================================================================================================

; Doc: http://fincs.ahk4.net/Ahk2ExeDirectives.htm
; Note: prefix comma with `

;@Ahk2Exe-SetName FoldersPopup
;@Ahk2Exe-SetDescription Folders Popup (freeware) - Move like a breeze between your frequently used folders and documents!
;@Ahk2Exe-SetVersion 3.9.1 BETA
;@Ahk2Exe-SetOrigFilename FoldersPopup.exe


;========================================================================================================================
; INITIALIZATION
;========================================================================================================================

#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off
ComObjError(False) ; we will do our own error handling

; avoid error message when shortcut destination is missing
; see http://ahkscript.org/boards/viewtopic.php?f=5&t=4477&p=25239#p25236
DllCall("SetErrorMode", "uint", SEM_FAILCRITICALERRORS := 1)

; By default, the A_WorkingDir is A_ScriptDir.
; When the shortcut is created by Inno Setup, the working is set to the folder under {userappdata}.
; In portable mode, the user can set the working directory in his own Windows shortcut.
; If user enable "Run at startup", the "Start in:" shortcut option is set to the current A_WorkingDir.

; Force A_WorkingDir to A_ScriptDir if uncomplied (development phase)
;@Ahk2Exe-IgnoreBegin
; Piece of code for development phase only - won't be compiled
; see http://fincs.ahk4.net/Ahk2ExeDirectives.htm
SetWorkingDir, %A_ScriptDir%
ListLines, On
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd

Gosub, InitFileInstall
Gosub, InitLanguageVariables

global strAppName := "FoldersPopup"
global strCurrentVersion := "3.9.1" ; "major.minor.bugs" or "major.minor.beta.release"
global strCurrentBranch := "beta" ; "prod" or "beta", always lowercase for filename
global strAppVersion := "v" . strCurrentVersion . (strCurrentBranch = "beta" ? " " . strCurrentBranch : "")
global str32or64 := A_PtrSize * 8
global blnDiagMode := False
global strDiagFile := A_WorkingDir . "\" . strAppName . "-DIAG.txt"
global strIniFile := A_WorkingDir . "\" . strAppName . ".ini"
global blnMenuReady := false
global arrSubmenuStack := Object()
global objIconsFile := Object()
global objIconsIndex := Object()
global strHotkeyNoneModifiers := ">^!+#" ; right-control/atl/shift/windows impossible keys combination
global strHotkeyNoneKey := "9"

if InStr(A_ScriptDir, A_Temp) ; must be positioned after strAppName is created
; if the app runs from a zip file, the script directory is created under the system Temp folder
{
	Oops(lOopsZipFileError, strAppName)
	Gosub, CleanUpBeforeExit
}

;@Ahk2Exe-IgnoreBegin
; Piece of code for developement phase only - won't be compiled
if (A_ComputerName = "JEAN-PC") ; for my home PC
	strIniFile := A_WorkingDir . "\" . strAppName . "-HOME.ini"
else if InStr(A_ComputerName, "STIC") ; for my work hotkeys
	strIniFile := A_WorkingDir . "\" . strAppName . "-WORK.ini"
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd

; Keep gosubs in this order
Gosub, InitSystemArrays
Gosub, InitLanguage
Gosub, InitLanguageArrays
Gosub, InitSpecialFolders
Gosub, LoadIniFile
if (blnDiagMode)
	Gosub, InitDiagMode
Gosub, LoadTheme

; build even if blnDisplaySpecialFolders or blnDisplaySwitchMenu are false because they could become true
; no need to build Recent folders menu at startup since this menu is refreshed on demand
Gosub, BuildSpecialFoldersMenu
Gosub, BuildGroupMenu ; need to be initialized here - will be updated at each call to popup menu
Gosub, BuildFavoritesMenus
Gosub, BuildGui
Gosub, Check4Update
Gosub, BuildTrayMenu

IfExist, %A_Startup%\%strAppName%.lnk ; update the shortcut in case the exe filename changed
{
	FileDelete, %A_Startup%\%strAppName%.lnk
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%strAppName%.lnk, %A_WorkingDir%
	Menu, Tray, Check, %lMenuRunAtStartup%
}

if (blnDisplayTrayTip)
	TrayTip, % L(lTrayTipInstalledTitle, strAppName, strAppVersion)
		, % L(lTrayTipInstalledDetail, strAppName
			, Hotkey2Text(strModifiers1, strMouseButton1, strOptionsKey1)
			, Hotkey2Text(strModifiers3, strMouseButton3, strOptionsKey3)
			, Hotkey2Text(strModifiers2, strMouseButton2, strOptionsKey2)
			, Hotkey2Text(strModifiers4, strMouseButton4, strOptionsKey4))
		, , 17 ; 1 info icon + 16 no sound)

OnExit, CleanUpBeforeExit

blnMenuReady := true

; Load the cursor and start the "hook"
objCursor := DllCall("LoadCursor", "UInt", NULL, "Int", 32649, "UInt") ; IDC_HAND
OnMessage(0x200, "WM_MOUSEMOVE")

; To popup menu when left click on the tray icon - See AHK_NOTIFYICON function below
OnMessage(0x404, "AHK_NOTIFYICON")

; Create a mutex to allow Inno Setup to detect if FP is running before uninstall or update
DllCall("CreateMutex", "uint", 0, "int", false, "str", strAppName . "Mutex")

; ### only when debugging Gui
; Gosub, GuiShow
; Gosub, GuiOptions
; Gosub, GuiAddFavorite
; Gosub, GuiAddFromPopup
; Gosub, GuiAddFromDropFiles
; Gosub, GuiEditFavorite
; Gosub, PopupMenuNewWindowKeyboard
; Gosub, BuildGroupMenu
; Gosub, GuiGroupSave
; Gosub, GuiManageGroups

return

#Include %A_ScriptDir%\FoldersPopup_LANG.ahk


;-----------------------------------------------------------
InitFileInstall:
;-----------------------------------------------------------

strTempDir := A_WorkingDir . "\_temp"
FileCreateDir, %strTempDir%

FileInstall, FileInstall\FoldersPopup_LANG_DE.txt, %strTempDir%\FoldersPopup_LANG_DE.txt, 1
FileInstall, FileInstall\FoldersPopup_LANG_FR.txt, %strTempDir%\FoldersPopup_LANG_FR.txt, 1
FileInstall, FileInstall\FoldersPopup_LANG_NL.txt, %strTempDir%\FoldersPopup_LANG_NL.txt, 1
FileInstall, FileInstall\FoldersPopup_LANG_KO.txt, %strTempDir%\FoldersPopup_LANG_KO.txt, 1
FileInstall, FileInstall\FoldersPopup_LANG_SV.txt, %strTempDir%\FoldersPopup_LANG_SV.txt, 1
FileInstall, FileInstall\default_browser_icon.html, %strTempDir%\default_browser_icon.html, 1

FileInstall, FileInstall\about-32.png, %strTempDir%\about-32.png
FileInstall, FileInstall\add_property-48.png, %strTempDir%\add_property-48.png
FileInstall, FileInstall\delete_property-48.png, %strTempDir%\delete_property-48.png
FileInstall, FileInstall\channel_mosaic-48.png, %strTempDir%\channel_mosaic-48.png
FileInstall, FileInstall\separator-26.png, %strTempDir%\separator-26.png
FileInstall, FileInstall\down_circular-26.png, %strTempDir%\down_circular-26.png
FileInstall, FileInstall\edit_property-48.png, %strTempDir%\edit_property-48.png
FileInstall, FileInstall\generic_sorting2-26-grey.png, %strTempDir%\generic_sorting2-26-grey.png
FileInstall, FileInstall\help-32.png, %strTempDir%\help-32.png
FileInstall, FileInstall\left-12.png, %strTempDir%\left-12.png
FileInstall, FileInstall\settings-32.png, %strTempDir%\settings-32.png
FileInstall, FileInstall\up-12.png, %strTempDir%\up-12.png
FileInstall, FileInstall\up_circular-26.png, %strTempDir%\up_circular-26.png

FileInstall, FileInstall\thumbs_up-32.png, %strTempDir%\thumbs_up-32.png
FileInstall, FileInstall\solutions-32.png, %strTempDir%\solutions-32.png
FileInstall, FileInstall\handshake-32.png, %strTempDir%\handshake-32.png
FileInstall, FileInstall\conference-32.png, %strTempDir%\conference-32.png
FileInstall, FileInstall\gift-32.png, %strTempDir%\gift-32.png

return
;-----------------------------------------------------------


;-----------------------------------------------------------
InitSystemArrays:
;-----------------------------------------------------------

; Hotkeys: ini names, hotkey variables name, default values, gosub label and Gui hotkey titles
strIniKeyNames := "PopupHotkeyMouse|PopupHotkeyNewMouse|PopupHotkeyKeyboard|PopupHotkeyNewKeyboard|RecentsHotkey|SettingsHotkey"
StringSplit, arrIniKeyNames, strIniKeyNames, |
strHotkeyVarNames := "strPopupHotkeyMouse|strPopupHotkeyMouseNew|strPopupHotkeyKeyboard|strPopupHotkeyKeyboardNew|strRecentsHotkey|strSettingsHotkey"
StringSplit, arrHotkeyVarNames, strHotkeyVarNames, |
strHotkeyDefaults := "MButton|+MButton|#a|+#a|+#r|+#f"
StringSplit, arrHotkeyDefaults, strHotkeyDefaults, |
strHotkeyLabels := "PopupMenuMouse|PopupMenuNewWindowMouse|PopupMenuKeyboard|PopupMenuNewWindowKeyboard|RefreshRecentFolders|GuiShow"
StringSplit, arrHotkeyLabels, strHotkeyLabels, |
strMouseButtons := "None|LButton|MButton|RButton|XButton1|XButton2|WheelUp|WheelDown|WheelLeft|WheelRight|"
; leave last | to enable default value on the last item
StringSplit, arrMouseButtons, strMouseButtons, |

strIconsMenus := "lMenuDesktop|lMenuDocuments|lMenuPictures|lMenuMyComputer|lMenuNetworkNeighborhood|lMenuControlPanel|lMenuRecycleBin"
	. "|menuRecentFolders|menuGroupDialog|menuGroupExplorer|lMenuSpecialFolders|lMenuGroup"
	. "|lMenuRecentFolders|lMenuSettings|lMenuAddThisFolder|lDonateMenu|Submenu|Network|UnknownDocument|Folder"
	. "|menuGroupSave|menuGroupLoad"

if (A_OSVersion = "WIN_XP")
{
	strIconsFile := "shell32|shell32|shell32|shell32|shell32|shell32|shell32"
				. "|shell32|shell32|shell32|shell32|shell32"
				. "|shell32|shell32|shell32|shell32|shell32|shell32|shell32|shell32"
				. "|shell32|shell32"
	strIconsIndex := "35|127|118|16|19|22|33"
				. "|4|147|4|4|147"
				. "|214|166|111|161|85|10|3|4"
				. "|7|7"
}
else if (A_OSVersion = "WIN_VISTA")
{
	strIconsFile := "imageres|imageres|imageres|imageres|imageres|imageres|imageres"
				. "|imageres|imageres|imageres|imageres|shell32"
				. "|imageres|imageres|shell32|shell32|shell32|imageres|imageres|imageres"
				. "|shell32|shell32"
	strIconsIndex := "105|85|67|104|114|22|49"
				. "|112|174|3|3|251|"
				. "112|109|88|161|85|28|2|3"
				. "|259|259"
}
else
{
	strIconsFile := "imageres|imageres|imageres|imageres|imageres|imageres|imageres"
				. "|imageres|imageres|imageres|imageres|shell32"
				. "|imageres|imageres|imageres|imageres|shell32|imageres|imageres|imageres"
				. "|shell32|shell32"
	strIconsIndex := "106|189|68|105|115|23|50"
				. "|113|176|203|203|99|"
				. "113|110|217|208|298|29|3|4"
				. "|297|46"
}

StringSplit, arrIconsFile, strIconsFile, |
StringSplit, arrIconsIndex, strIconsIndex, |

Loop, Parse, strIconsMenus, |
{
	objIconsFile[A_LoopField] := A_WinDir . "\System32\" . arrIconsFile%A_Index% . ".dll"
	objIconsIndex[A_LoopField] := arrIconsIndex%A_Index%
}
; example: objIconsFile["lMenuPictures"] and objIconsIndex["lMenuPictures"]

return
;-----------------------------------------------------------


;-----------------------------------------------------------
LoadIniFile:
;-----------------------------------------------------------

; reinit after Settings save if already exist
Global arrMenus := Object() ; list of menus (Main and submenus), non-hierachical
arrMainMenu := Object() ; array of favorites of the Main menu
arrMenus.Insert(lMainMenuName, arrMainMenu) ; lMainMenuName is used in the objects but not saved/restores in the ini file

strPopupHotkeyMouseDefault := arrHotkeyDefaults1 ; "MButton"
strPopupHotkeyMouseNewDefault := arrHotkeyDefaults2 ; "+MButton"
strPopupHotkeyKeyboardDefault := arrHotkeyDefaults3 ; "#a"
strPopupHotkeyKeyboardNewDefault := arrHotkeyDefaults4 ; "+#a"
strRecentsHotkeyDefault := arrHotkeyDefaults5 ; "+#r"
strSettingsHotkeyDefault := arrHotkeyDefaults6 ; "+#f"

intIconSize := (A_OSVersion = "WIN_XP" ? 16 : 24)

IfNotExist, %strIniFile%
	FileAppend,
		(LTrim Join`r`n
			[Global]
			PopupHotkeyMouse=%strPopupHotkeyMouseDefault%
			PopupHotkeyNewMouse=%strPopupHotkeyMouseNewDefault%
			PopupHotkeyKeyboard=%strPopupHotkeyKeyboardDefault%
			PopupHotkeyNewKeyboard=%strPopupHotkeyKeyboardNewDefault%
			RecentsHotkey=%strRecentsHotkeyDefault%
			SettingsHotkey=%strSettingsHotkeyDefault%
			DisplayTrayTip=1
			DisplayIcons=1
			DisplaySpecialFolders=1
			DisplayRecentFolders=1
			RecentFolders=10
			DisplaySwitchMenu=1
			DisplayMenuShortcuts=0
			PopupMenuPosition=1
			PopupFixPosition=20,20
			DiagMode=0
			Startups=1
			LanguageCode=EN
			DirectoryOpusPath=
			IconSize=%intIconSize%
			[Folders]
			Folder1=C:\|C:\
			Folder2=Windows|%A_WinDir%
			Folder3=Program Files|%A_ProgramFiles%

)
		, %strIniFile%

Gosub, LoadIniHotkeys

; fix language error
IniRead, blnDonor, %strIniFile%, Global, Donator, -1
if (blnDonor >= 0)
{
	IniWrite, %blnDonor%, %strIniFile%, Global, Donor
	IniDelete, %strIniFile%, Global, Donator
}
; backward compatibility with option PopupFix (until v3.1.3)
IniRead, blnPopupFix, %strIniFile%, Global, PopupFix, -1
if (blnPopupFix >= 0)
{
	IniWrite, % (blnPopupFix ? 3 : 1) , %strIniFile%, Global, PopupMenuPosition
	IniDelete, %strIniFile%, Global, PopupFix
}

IniRead, blnDisplayTrayTip, %strIniFile%, Global, DisplayTrayTip, 1
IniRead, blnDisplayIcons, %strIniFile%, Global, DisplayIcons, 1
blnDisplayIcons := (blnDisplayIcons and OSVersionIsWorkstation())
IniRead, blnDisplaySpecialFolders, %strIniFile%, Global, DisplaySpecialFolders, 1
IniRead, blnDisplayRecentFolders, %strIniFile%, Global, DisplayRecentFolders, 1
IniRead, blnDisplayGroupMenu, %strIniFile%, Global, DisplaySwitchMenu, 1 ; keep "Switch" in label instead of "Group" for backward compatibility
IniRead, intPopupMenuPosition, %strIniFile%, Global, PopupMenuPosition, 1
IniRead, strPopupFixPosition, %strIniFile%, Global, PopupFixPosition, 20,20
StringSplit, arrPopupFixPosition, strPopupFixPosition, `,
IniRead, blnDisplayMenuShortcuts, %strIniFile%, Global, DisplayMenuShortcuts, 0
IniRead, strLatestSkipped, %strIniFile%, Global, % "LatestVersionSkipped" . (strCurrentBranch = "beta" ? "Beta" : ""), 0.0
IniRead, blnDiagMode, %strIniFile%, Global, DiagMode, 0
IniRead, intStartups, %strIniFile%, Global, Startups, 1
IniRead, intRecentFolders, %strIniFile%, Global, RecentFolders, 10
IniRead, intIconSize, %strIniFile%, Global, IconSize, 24
IniRead, strGroups, %strIniFile%, Global, Groups, %A_Space% ; empty string if not found

IniRead, blnDonor, %strIniFile%, Global, Donor, 0 ; Please, be fair. Don't cheat with this.
if !(blnDonor)
	lMenuReservedShortcuts := lMenuReservedShortcuts . lMenuReservedShortcutsDonate

IniRead, strDirectoryOpusPath, %strIniFile%, Global, DirectoryOpusPath, %A_Space% ; empty string if not found
IniRead, blnDirectoryOpusUseTabs, %strIniFile%, Global, DirectoryOpusUseTabs, 1 ; use tabs by default
if StrLen(strDirectoryOpusPath)
	blnUseDirectoryOpus := FileExist(strDirectoryOpusPath)
if (blnUseDirectoryOpus)
	Gosub, SetDOpusRt
else
	if (strDirectoryOpusPath <> "NO")
		Gosub, CheckDirectoryOpus

IniRead, strTotalCommanderPath, %strIniFile%, Global, TotalCommanderPath, %A_Space% ; empty string if not found
IniRead, blnTotalCommanderUseTabs, %strIniFile%, Global, TotalCommanderUseTabs, 1 ; use tabs by default
if StrLen(strTotalCommanderPath)
	blnUseTotalCommander := FileExist(strTotalCommanderPath)
if (blnUseTotalCommander)
	Gosub, SetTCCommand
else
	if (strTotalCommanderPath <> "NO")
		Gosub, CheckTotalCommander

IniRead, strTheme, %strIniFile%, Global, Theme
if (strTheme = "ERROR") ; if Theme not found, we have a v1 or v2 ini file - add the themes to the ini file
{
	strTheme := "Grey"
	strAvailableThemes := "Grey|Yellow|Light Blue|Light Red|Light Green"
	IniWrite, %strTheme%, %strIniFile%, Global, Theme
	IniWrite, %strAvailableThemes%, %strIniFile%, Global, AvailableThemes
	FileAppend,
		(LTrim Join`r`n
			[Gui-Grey]
			WindowColor=E0E0E0
			TextColor=000000
			ListviewBackground=FFFFFF
			ListviewText=000000
			MenuBackgroundColor=FFFFFF
			[Gui-Yellow]
			WindowColor=f9ffc6
			TextColor=000000
			ListviewBackground=fcffe0
			ListviewText=000000
			MenuBackgroundColor=fcffe0
			[Gui-Light Blue]
			WindowColor=e8e7fa
			TextColor=000000
			ListviewBackground=e7f0fa
			ListviewText=000000
			MenuBackgroundColor=e7f0fa
			[Gui-Light Red]
			WindowColor=fddcd7
			TextColor=000000
			ListviewBackground=fef1ef
			ListviewText=000000
			MenuBackgroundColor=fef1ef
			[Gui-Light Green]
			WindowColor=d6fbde
			TextColor=000000
			ListviewBackground=edfdf1
			ListviewText=000000
			MenuBackgroundColor=edfdf1

)
		, %strIniFile%
}
else
	IniRead, strAvailableThemes, %strIniFile%, Global, AvailableThemes
	
Loop
{
	IniRead, strIniLine, %strIniFile%, Folders, Folder%A_Index% ; keep "Folders" label instead of "Favorite" for backward compatibility
	if (strIniLine = "ERROR")
		Break
	strIniLine := strIniLine . "|||" ; additional "|" to make sure we have all empty items
	; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType
	StringSplit, arrThisFavorite, strIniLine, |

	objFavorite := Object() ; new menu item
	objFavorite.FavoriteName := arrThisFavorite1 ; display name of this menu item
	objFavorite.FavoriteLocation := arrThisFavorite2 ; path for this menu item
	objFavorite.MenuName := lMainMenuName . arrThisFavorite3 ; parent menu of this menu item, adding main menu name

	if StrLen(arrThisFavorite4)
		objFavorite.SubmenuFullName := lMainMenuName . arrThisFavorite4 ; full name of the submenu, adding main menu name
	else
		objFavorite.SubmenuFullName := ""
	if StrLen(arrThisFavorite5)
		
		objFavorite.FavoriteType := arrThisFavorite5 ; "F" folder, "D" document, "U" URL or "S" submenu
		
	else ; for upward compatibility from v1 and v2 ini files
		if StrLen(objFavorite.SubmenuFullName)
			objFavorite.FavoriteType := "S" ; "S" submenu
		else ; for upward compatibility from v1 ini files
			objFavorite.FavoriteType := "F" ; "F" folder

	if (objFavorite.FavoriteType = "S") ; then this is a new submenu
	{
		arrSubMenu := Object() ; create submenu
		arrMenus.Insert(objFavorite.SubmenuFullName, arrSubMenu) ; add this submenu to the array of menus
	}
	arrMenus[objFavorite.MenuName].Insert(objFavorite) ; add this menu item to parent menu
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
	; Prepare global arrays used by SplitHotkey function
	IniRead, arrHotkeyVarNames%A_Index%, %strIniFile%, Global, % arrIniKeyNames%A_Index%, % arrHotkeyDefaults%A_Index%
	SplitHotkey(arrHotkeyVarNames%A_Index%, strMouseButtons
		, strModifiers%A_Index%, strOptionsKey%A_Index%, strMouseButton%A_Index%, strMouseButtonsWithDefault%A_Index%)
	; example: Hotkey, $MButton, PopupMenuMouse
	if (arrHotkeyVarNames%A_Index% = "None") ; do not compare with lOptionsMouseNone because it is translated
		Hotkey, % "$" . strHotkeyNoneModifiers . strHotkeyNoneKey, % arrHotkeyLabels%A_Index%, On UseErrorLevel
	else
		Hotkey, % "$" . arrHotkeyVarNames%A_Index%, % arrHotkeyLabels%A_Index%, On UseErrorLevel
	if (ErrorLevel)
		Oops(lDialogInvalidHotkey, Hotkey2Text(strModifiers%A_Index%, strMouseButton%A_Index%, strOptionsKey%A_Index%), strAppName, arrOptionsTitles%A_Index%)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
LoadTheme:
;------------------------------------------------------------

IniRead, strGuiWindowColor, %strIniFile%, Gui-%strTheme%, WindowColor, E0E0E0
IniRead, strTextColor, %strIniFile%, Gui-%strTheme%, TextColor, 000000
IniRead, strGuiListviewBackgroundColor, %strIniFile%, Gui-%strTheme%, ListviewBackground, FFFFFF
IniRead, strGuiListviewTextColor, %strIniFile%, Gui-%strTheme%, ListviewText, 000000
IniRead, strMenuBackgroundColor, %strIniFile%, Gui-%strTheme%, MenuBackgroundColor, FFFFFF

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

return
;------------------------------------------------------------


;------------------------------------------------------------
InitLanguageArrays:
;------------------------------------------------------------
StringSplit, arrOptionsTitles, lOptionsTitles, |
StringSplit, arrOptionsTitlesSub, lOptionsTitlesSub, |
StringSplit, arrOptionsTitlesLong, lOptionsTitlesLong, |
strOptionsLanguageCodes := "EN|FR|DE|NL|KO|SV"
StringSplit, arrOptionsLanguageCodes, strOptionsLanguageCodes, |
StringSplit, arrOptionsLanguageLabels, lOptionsLanguageLabels, |

loop, %arrOptionsLanguageCodes0%
	if (arrOptionsLanguageCodes%A_Index% = strLanguageCode)
		{
			strLanguageLabel := arrOptionsLanguageLabels%A_Index%
			break
		}

lOptionsMouseButtonsText := lOptionsMouseNone . "|" . lOptionsMouseButtonsText ; use lOptionsMouseNone because this is displayed
StringSplit, arrMouseButtonsText, lOptionsMouseButtonsText, |

return
;------------------------------------------------------------


;------------------------------------------------------------
InitSpecialFolders:
;------------------------------------------------------------

global objClassIdByLocalizedName := Object()
global objLocalizedNameByClassId := Object()

; do not process special folders where pExplorer has a folder in .LocationURL
; InitClassId("{450d8fba-ad25-11d0-98a8-0800361b1103}") ;	My Documents / Mes documents
; InitClassId("{4336a54d-038b-4685-ab02-99bb52d3fb8b}") ; Public Folder / Public
; InitClassId("{59031a47-3f72-44a7-89c5-5595fe6b30ee}") ; (User files) <User name> / <User name>

; special folders with only localized name in .LocationName actually used in popup menu
; these can be open with "explorer ::{CLSID}", but also with "explorer.exe shell:::{CLSID}"
InitClassId("{20d04fe0-3aea-1069-a2d8-08002b30309d}") ; My Computer / Ordinateur
InitClassId("{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}") ;	Computer and devices / Réseau
InitClassId("{645ff040-5081-101b-9f08-00aa002f954e}") ;	Recycle Bin / Corbeille
InitClassId("{26EE0668-A00A-44D7-9371-BEB064C98683}") ;	Control Panel / Panneau de configuration

; other special folders with only localized name in .LocationName
; these can be open with "explorer ::{CLSID}", but also with "explorer.exe shell:::{CLSID}"
InitClassId("{031E4825-7B94-4dc3-B131-E946B44C8DD5}") ; Libraries / Bibliothèques
InitClassId("{208d2c60-3aea-1069-a2d7-08002b30309d}") ;	My Network Places / Réseau (WORKGROUP)
InitClassId("{21EC2020-3AEA-1069-A2DD-08002B30309D}") ;	All Control Panel items / Tous les Panneaux de configuration
InitClassId("{7007acc7-3202-11d1-aad2-00805fc1270e}") ;	Network Connections / Connexions réseau
InitClassId("{E7DE9B1A-7533-4556-9484-B26FB486475E}") ;	Network Map / Mappage réseau
InitClassId("{2227a280-3aea-1069-a2de-08002b30309d}") ;	Printers / Imprimantes

; other special folders with only localized name in .LocationName
; these can only be open with "explorer.exe shell:::{CLSID}"
InitClassId("{22877a6d-37a1-461a-91b0-dbda5aaebc99}") ;	Recent Places / Emplacements récents
InitClassId("{992CFFA0-F557-101A-88EC-00DD010CCC48}") ;	Network Connections / Connexions réseau

return
;------------------------------------------------------------


;------------------------------------------------------------
InitClassId(strClassId)
;------------------------------------------------------------
{
    strLocalizedName := GetLocalizedNameForClassId(strClassId)
    objClassIdByLocalizedName.Insert(strLocalizedName, strClassId)
    objLocalizedNameByClassId.Insert(strClassId, strLocalizedName)
}
;------------------------------------------------------------


;------------------------------------------------------------
GetLocalizedNameForClassId(strClassId)
;------------------------------------------------------------
{
    /*
        Question 1: What's the best Registry key should I use for this to work on Win XP, 7, 8+ ?
    
        HKEY_CLASSES_ROOT\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}
        HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}
        HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}
        HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Wow6432Node\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}
        HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}
		
		Seems to be HKEY_CLASSES_ROOT for XP compatibility accorging to
		http://msdn.microsoft.com/en-us/library/windows/desktop/ms724475(v=vs.85).aspx
    */
    RegRead, strLocalizedString, HKEY_CLASSES_ROOT, CLSID\%strClassId%, LocalizedString
    ; strLocalizedString example: "@%SystemRoot%\system32\shell32.dll,-9216"

    StringSplit, arrLocalizedString, strLocalizedString, `,
    intDllNameStart := InStr(arrLocalizedString1, "\", , 0)
    StringRight, strDllFile, arrLocalizedString1, % StrLen(arrLocalizedString1) - intDllNameStart
    strDllIndex := arrLocalizedString2
    strTranslatedName := TranslateMUI(strDllFile, Abs(strDllIndex))
    
    /*
    MsgBox, % ""
        . "strClassId: " . strClassId . "`n"
        . "strLocalizedString: " . strLocalizedString . "`n"
        . "strDllFile: " . strDllFile . "`n"
        . "strDllIndex: " . strDllIndex . "`n"
        . "strTranslatedName: " . strTranslatedName . "`n"
    */
    
    return strTranslatedName
}
;------------------------------------------------------------


;------------------------------------------------------------
TranslateMUI(resDll, resID)
; source: 7plus (https://github.com/7plus/7plus/blob/master/MiscFunctions.ahk)
;------------------------------------------------------------
{
	VarSetCapacity(buf, 256) 
	hDll := DllCall("LoadLibrary", "str", resDll, "Ptr") 
	Result := DllCall("LoadString", "Ptr", hDll, "uint", resID, "str", buf, "int", 128)
	return buf
}
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
	Diag("A_WorkingDir", A_WorkingDir)
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


;------------------------------------------------------------
SplitHotkey(strHotkey, strMouseButtons, ByRef strModifiers, ByRef strKey, ByRef strMouseButton, ByRef strMouseButtonsWithDefault)
;------------------------------------------------------------
{
	if (strHotkey = "None") ; do not compare with lOptionsMouseNone because it is translated
	{
		strMouseButton := "None" ; do not use lOptionsMouseNone because it is translated
		strKey := ""
		StringReplace, strMouseButtonsWithDefault, lOptionsMouseButtonsText, % lOptionsMouseNone . "|", % lOptionsMouseNone . "||" ; use lOptionsMouseNone because this is displayed
	}
	else 
	{
		SplitModifiersFromKey(strHotkey, strModifiers, strKey)
		if InStr(strMouseButtons, "|" . strKey . "|") ;  we have a mouse button
		{
			strMouseButton := strKey
			strKey := ""
			StringReplace, strMouseButtonsWithDefault, lOptionsMouseButtonsText, % GetText4MouseButton(strMouseButton) . "|", % GetText4MouseButton(strMouseButton) . "||" ; with default value
		}
		else ; we have a key
			strMouseButtonsWithDefault := lOptionsMouseButtonsText ; no default value
	}
}
;------------------------------------------------------------


;------------------------------------------------
WM_MOUSEMOVE(wParam, lParam)
; "hook" for image buttons cursor
; see http://www.autohotkey.com/board/topic/70261-gui-buttons-hover-cant-change-cursor-to-hand/
;------------------------------------------------
{
	Global objCursor
	
	MouseGetPos, , , , strControl ; Static1, StaticN, Button1, ButtonN
	StringReplace, strControl, strControl, Static

	If InStr(".3.4.6.7.9.10.11.12.14.15.16.17.18.19.20.21.25.26.27.28.29.30.Button1.Button2.", "." . strControl . ".")
		DllCall("SetCursor", "UInt", objCursor)

	return
}
;------------------------------------------------


;------------------------------------------------------------
AHK_NOTIFYICON(wParam, lParam) 
; Adapted from Lexikos http://www.autohotkey.com/board/topic/11250-mouseover-trayicon-triggering-an-event/#entry153388
; To popup menu when left click on the tray icon - See the OnMessage command in the init section
;------------------------------------------------------------
{ 
	if (lParam = 0x202) ; WM_LBUTTONUP
	{ 
		SetTimer, PopupMenuNewWindowMouse, -1 
		return 0 
	} 
} 
;------------------------------------------------------------



;========================================================================================================================
; EXIT
;========================================================================================================================


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



;========================================================================================================================
; BUILD
;========================================================================================================================


;------------------------------------------------------------
BuildTrayMenu:
;------------------------------------------------------------

Menu, Tray, Icon, , , 1 ; last 1 to freeze icon during pause or suspend
Menu, Tray, NoStandard
;@Ahk2Exe-IgnoreBegin
; Piece of code for developement phase only - won't be compiled
Menu, Tray, Icon, %A_ScriptDir%\Folders-Likes-icon-192-RED-center.ico, 1, 1 ; last 1 to freeze icon during pause or suspend
Menu, Tray, Standard
Menu, Tray, Add
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd
; Menu, Tray, Add, % L(lMenuFPMenu, strAppName, lMenuMenu), :%lMainMenuName% ; REMOVED seems to cause a BUG in submenu display (first display only) - unexplained...
Menu, Tray, Add, % L(lMenuSettings, strAppName), GuiShow
Menu, Tray, Add
Menu, Tray, Add, %lMenuRunAtStartup%, RunAtStartup
Menu, Tray, Add, %lMenuSuspendHotkeys%, SuspendHotkeys
Menu, Tray, Add
Menu, Tray, Add, %lMenuUpdate%, Check4Update
Menu, Tray, Add, %lMenuHelp%, GuiHelp
Menu, Tray, Add, %lMenuAbout%, GuiAbout
Menu, Tray, Add, %lDonateMenu%, GuiDonate
Menu, Tray, Add
Menu, Tray, Add, % L(lMenuExitFoldersPopup, strAppName), ExitFoldersPopup
Menu, Tray, Default, % L(lMenuSettings, strAppName)
Menu, Tray, Color, %strMenuBackgroundColor%
Menu, Tray, Tip, % strAppName . " " . strAppVersion . " (" . str32or64 . "-bit)`n" . (blnDonor ? lDonateThankyou : lDonateButton)

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildSpecialFoldersMenu:
;------------------------------------------------------------

/*
Additions

/common
A_DesktopCommon en enlevant le dernier dossier

/downloads
RegRead, str, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, {374DE290-123F-4565-9164-39C4925E467B}

/mymusic
A_Desktop et remplacer Desktop par Music
RegRead, str, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Music

/mymusic
A_Desktop et remplacer Desktop par Music
RegRead, str, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Video

/recent
RegRead, str, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, Recent

/temp
A_Temp

/favorites
RegRead, str, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, Favorites

/templates
RegRead, str, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, Templates
*/

Menu, menuSpecialFolders, Add
Menu, menuSpecialFolders, DeleteAll ; had problem with DeleteAll making the Special menu to disappear 1/2 times - now OK
Menu, menuSpecialFolders, Color, %strMenuBackgroundColor%

AddMenuIcon("menuSpecialFolders", lMenuDesktop, "OpenSpecialFolder", "lMenuDesktop")
AddMenuIcon("menuSpecialFolders", lMenuDocuments, "OpenSpecialFolder", "lMenuDocuments")
AddMenuIcon("menuSpecialFolders", lMenuPictures, "OpenSpecialFolder", "lMenuPictures")
Menu, menuSpecialFolders, Add
AddMenuIcon("menuSpecialFolders", lMenuMyComputer, "OpenSpecialFolder", "lMenuMyComputer")
AddMenuIcon("menuSpecialFolders", lMenuNetworkNeighborhood, "OpenSpecialFolder", "lMenuNetworkNeighborhood")
Menu, menuSpecialFolders, Add
AddMenuIcon("menuSpecialFolders", lMenuControlPanel, "OpenSpecialFolder", "lMenuControlPanel")
AddMenuIcon("menuSpecialFolders", lMenuRecycleBin, "OpenSpecialFolder", "lMenuRecycleBin")

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildRecentFoldersMenu:
;------------------------------------------------------------

Menu, menuRecentFolders, Add
Menu, menuRecentFolders, DeleteAll ; had problem with DeleteAll making the Special menu to disappear 1/2 times - now OK
Menu, menuRecentFolders, Color, %strMenuBackgroundColor%

objRecentFolders := Object()
intRecentFoldersIndex := 0 ; used in PopupMenu... to check if we disable the menu when empty

RegRead, strRecentsFolder, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, Recent

/*
Alternative to collect recent files *** DO NOT WORK with XP
See: post from Skan http://ahkscript.org/boards/viewtopic.php?f=5&t=4477#p25261
Implement for Win7+ if FileGetShortcut still produce Windows errors when external drive is not available (despite DllCall in initialization)

strWinPathRecent := RegExReplace(SubStr(strRecentsFolder, 3) . "\", "\\", "\\")
strRecentFoldersList := ""
for ObjItem in ComObjGet("winmgmts:")
	.ExecQuery("Select * from Win32_ShortcutFile where path = '" . strWinPathRecent . "'")
	strRecentFoldersList := strRecentFoldersList . ObjItem.LastModified . A_Tab . ObjItem.Target . "`n"
*/

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
	if (blnDiagMode)
		Diag("BuildRecentFoldersMenu", strTargetFullName)
	
	FileGetShortcut, %strTargetFullName%, strOutTarget
	
	if (errorlevel) ; hidden or system files (like desktop.ini) returns an error
		continue
	if !FileExist(strOutTarget) ; if folder/document was delete or on a removable drive
	{
		if (blnDiagMode)
			Diag("RecentShortcut NO", strOutTarget)
		continue
	}
	else
		if (blnDiagMode)
			Diag("RecentShortcut YES", strOutTarget)
	
	if LocationIsDocument(strOutTarget) ; not a folder
		continue

	intRecentFoldersIndex := intRecentFoldersIndex + 1
	objRecentFolders.Insert(intRecentFoldersIndex, strOutTarget)
	
	strMenuName := (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, false) . " " : "") . strOutTarget
	AddMenuIcon("menuRecentFolders", strMenuName, "OpenRecentFolder", "Folder")

	if (intRecentFoldersIndex >= intRecentFolders)
		break
}

return
;------------------------------------------------------------


;------------------------------------------------------------
RefreshRecentFolders:
;------------------------------------------------------------

ToolTip, %lMenuRefreshRecent%...
Gosub, BuildRecentFoldersMenu
ToolTip
CoordMode, Menu, % (intPopupMenuPosition = 2 ? "Window" : "Screen")
Menu, menuRecentFolders, Show, %intMenuPosX%, %intMenuPosY% ; same position as the calling popup menu

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildGroupMenu:
;------------------------------------------------------------

if (blnUseDirectoryOpus)
{
	FileDelete, %strDOpusTempFilePath%
	RunDOpusRt("/info", strDOpusTempFilePath, ",paths") ; list opened listers in a text file
	; Run, "%strDirectoryOpusRtPath%" /info "%strDOpusTempFilePath%"`,paths
	loop, 10
		if FileExist(strDOpusTempFilePath)
			Break
		else
			Sleep, 50 ; was 10 and had some gliches - 50 enough?
	FileRead, strListText, %strDOpusTempFilePath%

	objDOpusListers := Object()
	CollectDOpusListersList(objDOpusListers, strListText) ; list all listers, excluding special folders like Recycle Bin
}

/* when TC tabs resolved
if (blnUseTotalCommander)
{
	objTCLists := Object()
	; List available folders. Need solution from TC developer to list folders in all tabs:
	; see http://www.ghisler.ch/board/viewtopic.php?t=41198
	CollectTCLists(objTCLists)
}
*/

objExplorersWindows := Object()
CollectExplorers(objExplorersWindows, ComObjCreate("Shell.Application").Windows)

objGroupMenuExplorers := Object()

intExplorersIndex := 0
intExplorersExist := 0 ; used in PopupMenu and SaveGroup to check if we disable menu or button when empty

if (blnUseDirectoryOpus)
	for intIndex, objLister in objDOpusListers
	{
		; if popup menu is for dialog box and we have no path or have a collection, skip it
		if (WindowIsDialog(strTargetClass) and !(blnNewWindow))
			and (!StrLen(objLister.Path) or InStr(objLister.Path, "coll://"))
			continue
		
		if !NameAndTypeIsInObject(objLister.Path, "DO", objGroupMenuExplorers)
			intExplorersIndex := intExplorersIndex + 1
			
		objGroupMenuExplorer := Object()
		objGroupMenuExplorer.Path := objLister.Path
		if InStr(objLister.Path, "Lister-Quick-Find-Results")
			objGroupMenuExplorer.Name := "Directory Opus Quick Find Results (" . intExplorersIndex . ")"
		else if InStr(objLister.Path, "coll://")
			objGroupMenuExplorer.Name := "Directory Opus Collection (" . intExplorersIndex . ")"
		else
			objGroupMenuExplorer.Name := objLister.Path
		objGroupMenuExplorer.WindowId := objLister.Lister
		objGroupMenuExplorer.TabId := objLister.Tab
		objGroupMenuExplorer.Position := objLister.Position
		objGroupMenuExplorer.MinMax := objLister.MinMax
		objGroupMenuExplorer.Pane := (objLister.Pane = 0 ? 1 : objLister.Pane) ; consider pane 0 as pane 1
		objGroupMenuExplorer.WindowType := "DO"
		if FileExist(objGroupMenuExplorer.Name)
			intExplorersExist := intExplorersExist + 1
		
		objGroupMenuExplorers.Insert(intExplorersIndex, objGroupMenuExplorer)
	}

/* when TC tabs resolved
if (blnUseTotalCommander)
	for intIndex, objTCList in objTCLists
	{
		if !NameIsInObject(objTCList.Path, objGroupMenuExplorers)
		{
			intExplorersIndex := intExplorersIndex + 1
			
			objGroupMenuExplorer := Object()
			objGroupMenuExplorer.Path := objTCList.Path
			objGroupMenuExplorer.Name := objTCList.Path
			objGroupMenuExplorer.WindowId := objTCList.WindowId
			objGroupMenuExplorer.Position := objTCList.Position
			objGroupMenuExplorer.MinMax := objTCList.MinMax
			objGroupMenuExplorer.Pane := objTCList.Pane
			objGroupMenuExplorer.WindowType := "TC"
			if FileExist(objGroupMenuExplorer.Name)
				intExplorersExist := intExplorersExist + 1
			
			objGroupMenuExplorers.Insert(intExplorersIndex, objGroupMenuExplorer)
		}
	}
*/

for intIndex, objExplorer in objExplorersWindows
{
	; if popup menu is for dialog box and we have no path, skip it
	if (WindowIsDialog(strTargetClass) and !(blnNewWindow))
		and !StrLen(objExplorer.LocationURL)
		continue
		
	if !NameAndTypeIsInObject(objExplorer.LocationName, "EX", objGroupMenuExplorers)
		intExplorersIndex := intExplorersIndex + 1
		
	objGroupMenuExplorer := Object()
	objGroupMenuExplorer.Path := objExplorer.LocationURL
	objGroupMenuExplorer.Name := objExplorer.LocationName
	objGroupMenuExplorer.IsSpecialFolder := objExplorer.IsSpecialFolder
	objGroupMenuExplorer.WindowId := objExplorer.WindowId
	objGroupMenuExplorer.Position := objExplorer.Position
	objGroupMenuExplorer.MinMax := objExplorer.MinMax
	objGroupMenuExplorer.WindowType := "EX"
	if FileExist(objGroupMenuExplorer.Name) or (objGroupMenuExplorer.IsSpecialFolder)
		intExplorersExist := intExplorersExist + 1

	objGroupMenuExplorers.Insert(intExplorersIndex, objGroupMenuExplorer)
}

Menu, menuGroup, Add
Menu, menuGroup, DeleteAll
Menu, menuGroup, Color, %strMenuBackgroundColor%

intShortcutGroup := 0

for intIndex, objGroupMenuExplorer in objGroupMenuExplorers
{
	strMenuName := (blnDisplayMenuShortcuts and (intShortcutGroup <= 35) ? "&" . NextMenuShortcut(intShortcutGroup, false) . " " : "") . objGroupMenuExplorer.Name
	if !WindowIsDialog(strTargetClass) or (blnNewWindow)
		AddMenuIcon("menuGroup"
			, strMenuName . (objGroupMenuExplorer.WindowType = "DO" ? " (DOpus)" : (objGroupMenuExplorer.WindowType = "TC" ? " (TC)" : ""))
			, "SwitchExplorer"
			, (objGroupMenuExplorer.WindowType = "DO" ? "DirectoryOpus" : (objGroupMenuExplorer.WindowType = "TC" ? "TotalCommander" : "menuGroupExplorer")))
	else
		AddMenuIcon("menuGroup", strMenuName, "SwitchDialog", "menuGroupDialog")
}

Menu, menuGroupsList, Add
Menu, menuGroupsList, DeleteAll
Menu, menuGroupsList, Color, %strMenuBackgroundColor%

Loop, Parse, strGroups, |
{
	IniRead, blnReplaceWhenRestoringThisGroup, %striniFile%, Group-%A_LoopField%, ReplaceWhenRestoringThisGroup, %A_Space% ; empty if not found
	AddMenuIcon("menuGroupsList", A_LoopField . " (" . (blnReplaceWhenRestoringThisGroup ? lMenuGroupReplace : lMenuGroupAdd) . ")" , "GroupLoad", "lMenuGroup")
}
if (intExplorersIndex)
	Menu, menuGroup, Add
AddMenuIcon("menuGroup", lMenuGroupSave, "GuiGroupSave", "menuGroupSave")
AddMenuIcon("menuGroup", lMenuGroupLoad, ":menuGroupsList", "menuGroupLoad")

return
;------------------------------------------------------------


;------------------------------------------------------------
CollectExplorers(objExplorers, pExplorers)
;------------------------------------------------------------
{
	For pExplorer in pExplorers
	; see http://msdn.microsoft.com/en-us/library/windows/desktop/aa752084(v=vs.85).aspx
	{
		if (A_LastError)
			; an error occurred during ComObjCreate (A_LastError probably is E_UNEXPECTED = -2147418113 #0x8000FFFFL)
			break

		strType := ""
		try
			strType := pExplorer.Type ; Gets the user type name of the contained document object. "Document HTML" for IE windows. Should be empty for file Explorer windows.
		
		if !StrLen(strType)
		{
			/*
			###_D(""
				. "pExplorer.LocationURL: !" . pExplorer.LocationURL . "!`n"
				. "pExplorer.Type: !" . pExplorer.Type . "!`n"
				. "pExplorer.LocationName: !" . pExplorer.LocationName . "!`n"
				. "pExplorer.HWND: !" . pExplorer.HWND . "!`n"
				. "")
			*/
			objExplorer := Object()
			objExplorer.Position := pExplorer.Left . "|" . pExplorer.Top . "|" . pExplorer.Width . "|" . pExplorer.Height
			
			objExplorer.IsSpecialFolder := !StrLen(pExplorer.LocationURL) ; empty for special folders like Recycle bin
			if  (objExplorer.IsSpecialFolder)
				objExplorer.LocationName := pExplorer.LocationName ; see http://msdn.microsoft.com/en-us/library/aa752084#properties
			else
			{
				objExplorer.LocationURL := pExplorer.LocationURL
				strLocationName :=  UriDecode(pExplorer.LocationURL)
				StringReplace, strLocationName, strLocationName, file:///
				StringReplace, strLocationName, strLocationName, /, \, A
				objExplorer.LocationName := strLocationName
			}
			
			objExplorer.WindowId := pExplorer.HWND
			WinGet, intMinMax, MinMax, % "ahk_id " . pExplorer.HWND
			objExplorer.MinMax := intMinMax
			
			if StrLen(pExplorer.HWND)
				objExplorers.Insert(A_Index, objExplorer)
		}
	}
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
;------------------------------------------------


/* when TC tabs resolved
;------------------------------------------------------------
CollectTCLists(objLists)
;------------------------------------------------------------
{
	objClipboardBK := ClipboardAll

	cm_CopySrcPathToClip := 2029
	cm_CopyTrgPathToClip := 2030
	intTCWindows := 0

	WinGet, arrIDs, List, ahk_class TTOTAL_CMD
	Loop, %arrIDs%
	{
		strWindowId := arrIDs%A_Index%
		WinGetPos, intX, intY, intW, intH, ahk_id %strWindowId%
		WinGet, intMinMax, MinMax, ahk_id %strWindowId%
		ClipBoard := ""
		sleep, 50
		SendMessage, 0x433, %cm_CopySrcPathToClip%, , , ahk_id %strWindowId%
		ClipWait, 2, 1 ; wait up to 2 seconds for text (1) in the Clipboard
		sleep, 10 ; wait 10 additional seconds to improve reliability
		if StrLen(ClipBoard)
		{
			objTCList := Object()
			objTCList.Path := ClipBoard
			objTCList.WindowId := strWindowId
			objTCList.Position := intX . "|" . intY . "|" . intW . "|" . intH
			objTCList.MinMax := intMinMax
			objTCList.Pane := 1
			objLists.Insert(intTCWindows, objTCList)
			ClipBoard := ""
			sleep, 50
		}
		SendMessage, 0x433, %cm_CopyTrgPathToClip%, , ,  ahk_id %strWindowId%
		ClipWait, 2, 1 ; wait up to 2 seconds for text (1) in the Clipboard
		sleep, 10 ; wait 10 additional seconds to improve reliability
		if StrLen(ClipBoard)
		{
			objTCList := Object()
			objTCList.Path := ClipBoard
			objTCList.WindowId := strWindowId
			objTCList.Position := intX . "|" . intY . "|" . intW . "|" . intH
			objTCList.MinMax := intMinMax
			objTCList.Pane := 2
			objLists.Insert(intTCWindows, objTCList)
		}
	}
	
	Clipboard := objClipboardBK
	objClipboardBK :=
}
;------------------------------------------------------------
*/


;------------------------------------------------------------
CollectDOpusListersList(objListers, strList)
; list all DirectoryOpus listers, excluding special folders like Recycle Bin, Network because they are not included in dopus-list.txt
;------------------------------------------------------------
{
	strList := SubStr(strList, InStr(strList, "<path"))
	Loop
	{
		objLister := Object()
		
		strList := SubStr(strList, InStr(strList, "<path"))
		strSubStr := SubStr(strList, InStr(strList, "<path"))
		strSubStr := SubStr(strSubStr, 1, InStr(strSubStr, "</path>") - 1)
		
		if (StrLen(strSubStr))
		{
			objLister.Active_lister := ParseDOpusListerProperty(strSubStr, "active_lister")
			objLister.Active_tab := ParseDOpusListerProperty(strSubStr, "active_tab")
			objLister.Lister := ParseDOpusListerProperty(strSubStr, "lister")
			objLister.Side := ParseDOpusListerProperty(strSubStr, "side")
			objLister.Tab := ParseDOpusListerProperty(strSubStr, "tab")
			objLister.Tab_state := ParseDOpusListerProperty(strSubStr, "tab_state")
			objLister.Path := SubStr(strSubStr, InStr(strSubStr, ">") + 1)

			WinGetPos, intX, intY, intW, intH, % "ahk_id " . objLister.lister
			objLister.Position := intX . "|" . intY . "|" . intW . "|" . intH
			WinGet, intMinMax, MinMax, % "ahk_id " . objLister.lister
			objLister.MinMax := intMinMax
			objLister.Pane := objLister.Side
			
			if !InStr(objLister.Path, "ftp://")
				; Swith Explorer to DOpus FTP folder not supported (see https://github.com/JnLlnd/FoldersPopup/issues/84)
				objListers.Insert(A_Index, objLister)
				
			strList := SubStr(strList, StrLen(strSubStr))
		}
	} until	(!StrLen(strSubStr))
}
;------------------------------------------------------------


;------------------------------------------------------------
ParseDOpusListerProperty(strSource, strProperty)
;------------------------------------------------------------
{
	intStartPos := InStr(strSource, " " . strProperty . "=")
	if !(intStartPos)
		return ""
	strSource := SubStr(strSource, intStartPos + StrLen(strProperty) + 3)
	intEndPos := InStr(strSource, """")
	return SubStr(strSource, 1, intEndPos - 1)
}
;------------------------------------------------------------


;------------------------------------------------------------
NameAndTypeIsInObject(strName, strType, obj)
;------------------------------------------------------------
{
	loop, % obj.MaxIndex()
		if (strName = obj[A_Index].Name) and (strType = obj[A_Index].WindowType)
			return true
		
	return false
}
;------------------------------------------------------------


;------------------------------------------------------------
BuildFavoritesMenus:
BuildFavoritesMenusWithStatus:
;------------------------------------------------------------

if (A_ThisLabel = "BuildFavoritesMenusWithStatus")
	TrayTip, % L(lTrayTipWorkingTitle, strAppName, strAppVersion)
		, %lTrayTipWorkingDetail%, , 1

Menu, %lMainMenuName%, Add
Menu, %lMainMenuName%, DeleteAll
Menu, %lMainMenuName%, Color, %strMenuBackgroundColor%

BuildOneMenu(lMainMenuName) ; and recurse for submenus

Menu, %lMainMenuName%, Add

if (blnDisplaySpecialFolders)
	AddMenuIcon(lMainMenuName, lMenuSpecialFolders, ":menuSpecialFolders", "lMenuSpecialFolders")

if (blnDisplayGroupMenu)
{
	AddMenuIcon(lMainMenuName, lMenuGroup, ":menuGroup", "lMenuGroup")
	Menu, menuGroup, Color, %strMenuBackgroundColor%
}

if (blnDisplayRecentFolders)
	AddMenuIcon(lMainMenuName, lMenuRecentFolders . "...", "RefreshRecentFolders", "lMenuRecentFolders")

if (blnDisplaySpecialFolders or blnDisplayRecentFolders or blnDisplayGroupMenu)
	Menu, %lMainMenuName%, Add

AddMenuIcon(lMainMenuName, L(lMenuSettings, strAppName) . "...", "GuiShow", "lMenuSettings")
Menu, %lMainMenuName%, Default, % L(lMenuSettings, strAppName) . "..."
AddMenuIcon(lMainMenuName, lMenuAddThisFolder . "...", "AddThisFolder", "lMenuAddThisFolder")

if !(blnDonor)
{
	Menu, %lMainMenuName%, Add
	AddMenuIcon(lMainMenuName, lDonateMenu . "...", "GuiDonate", "lDonateMenu")
}

if (A_ThisLabel = "BuildFavoritesMenusWithStatus")
	TrayTip, % L(lTrayTipInstalledTitle, strAppName, strAppVersion)
		, %lTrayTipWorkingDetailFinished%, , 1

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildOneMenu(strMenu)
; recursive function
;------------------------------------------------------------
{
	global blnDisplayMenuShortcuts
	global blnDisplayIcons
	global intIconSize
	global strMenuBackgroundColor
	
	intShortcut := 0
	
	; try because at first execution strMenu does not exist and produces an error,
	; but DeleteAll is required later for menu updates
	try Menu, %strMenu%, DeleteAll
	
	arrThisMenu := arrMenus[strMenu]
	Loop, % arrThisMenu.MaxIndex()
		
		if (arrThisMenu[A_Index].FavoriteType = "S") ; this is a submenu
		{
			strSubMenuFullName := arrThisMenu[A_Index].SubmenuFullName
			strSubMenuDisplayName := arrThisMenu[A_Index].FavoriteName
			strSubMenuParent := arrThisMenu[A_Index].MenuName
			
			BuildOneMenu(strSubMenuFullName) ; recursive call
			Try Menu, %strSubMenuParent%, Color, %strMenuBackgroundColor% ; Try because this can fail if submenu is empty
			
			strMenuName := (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, (strMenu = lMainMenuName)) . " " : "") . strSubMenuDisplayName
			Try Menu, %strSubMenuParent%, Add, %strMenuName%, % ":" . strSubMenuFullName
			catch e ; when menu is empty
			{
				Menu, % arrThisMenu[A_Index].MenuName, Add, %strMenuName%, OpenFavorite ; will never be called because disabled
				Menu, % arrThisMenu[A_Index].MenuName, Disable, %strMenuName%
			}
			if (blnDisplayIcons)
				Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%
					, % objIconsFile["Submenu"], % objIconsIndex["Submenu"], %intIconSize%
		}
		
		else if (arrThisMenu[A_Index].FavoriteName = lMenuSeparator) ; this is a separator

			Menu, % arrThisMenu[A_Index].MenuName, Add
			
		else ; this is a favorite (folder, document or URL)
		{
			strMenuName := (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut, (strMenu = lMainMenuName)) . " " : "")
				. arrThisMenu[A_Index].FavoriteName
			Menu, % arrThisMenu[A_Index].MenuName, Add, %strMenuName%, OpenFavorite


			if (blnDisplayIcons)
			{
				if (arrThisMenu[A_Index].FavoriteType = "F") ; this is a folder
					
					Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%
						, % objIconsFile["Folder"], % objIconsIndex["Folder"], %intIconSize%
				
				else ;  this is a document or an URL
				{
					if (arrThisMenu[A_Index].FavoriteType = "U") ; this is an URL
						GetIcon4Location(strTempDir . "\default_browser_icon.html", strDefaultIcon, intDefaultIcon)
						; not sure it is required to have a physical file with .html extension - but keep it as is by safety
					else ; this is a document
						GetIcon4Location(arrThisMenu[A_Index].FavoriteLocation, strDefaultIcon, intDefaultIcon)
					
					if StrLen(strDefaultIcon) and !InStr(strDefaultIcon, "%")
						Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%, %strDefaultIcon%, %intDefaultIcon%, %intIconSize%
					else
						Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%
							, % objIconsFile["UnknownDocument"], % objIconsIndex["UnknownDocument"], %intIconSize%
				}
			}
		}
}
;------------------------------------------------------------


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


;------------------------------------------------------------
AddMenuIcon(strMenuName, ByRef strMenuItemName, strLabel, strIconValue)
;------------------------------------------------------------
{
	global intIconSize
	global blnDisplayIcons
	global objIconsFile
	global objIconsIndex

	/*
	###_D(""
		. "strMenuName: " . strMenuName . "`n"
		. "strMenuItemName: " . strMenuItemName . "`n"
		. "strLabel: " . strLabel . "`n"
		. "strIconValue: " . strIconValue . "`n"
		. "objIconsFile[strIconValue]: " . objIconsFile[strIconValue] . "`n"
		. "objIconsIndex[strIconValue]: " . objIconsIndex[strIconValue] . "`n"
		. "")
	*/
	
	if !StrLen(strMenuItemName)
		return
	
	; The names of menus and menu items can be up to 260 characters long.
	if StrLen(strMenuItemName) > 260
		strMenuItemName := SubStr(strMenuItemName, 1, 256) . "..." ; minus one for the luck ;-)
	
	Menu, %strMenuName%, Add, %strMenuItemName%, %strLabel%
	if (blnDisplayIcons)
		Menu, %strMenuName%, Icon, %strMenuItemName%, % objIconsFile[strIconValue], % objIconsIndex[strIconValue], %intIconSize%
}
;------------------------------------------------------------


;------------------------------------------------------------
GetIcon4Location(strLocation, ByRef strDefaultIcon, ByRef intDefaultIcon)
; get icon, extract from kiu http://www.autohotkey.com/board/topic/8616-kiu-icons-manager-quickly-change-icon-files/
;------------------------------------------------------------
{
	global blnDiagMode
	
	SplitPath, strLocation, , , strExtension
	RegRead, strHKeyClassRoot, HKEY_CLASSES_ROOT, .%strExtension%
	RegRead, strDefaultIcon, HKEY_CLASSES_ROOT, %strHKeyClassRoot%\DefaultIcon
	if (blnDiagMode)
	{
		Diag("BuildOneMenuIcon", strLocation)
		Diag("strHKeyClassRoot", strHKeyClassRoot)
		Diag("strDefaultIcon-1", strDefaultIcon)
	}
	
	if (strDefaultIcon = "%1") ; use the file itself (for executable)
		strDefaultIcon := strLocation
	
	IfInString, strDefaultIcon, `, ; this is for icongroup files
	{
		StringSplit, arrDefaultIcon, strDefaultIcon, `,
		strDefaultIcon := arrDefaultIcon1
		intDefaultIcon := arrDefaultIcon2
	}
	else
		intDefaultIcon := 1

	if (blnDiagMode)
	{
		Diag("strDefaultIcon-2", strDefaultIcon)
		Diag("intDefaultIcon", intDefaultIcon)
	}
}
;------------------------------------------------------------



;========================================================================================================================
; POPUP MNEU
;========================================================================================================================


;------------------------------------------------------------
PopupMenuMouse: ; default MButton
PopupMenuKeyboard: ; default #a
;------------------------------------------------------------

if !(blnMenuReady)
	return

If !CanOpenFavorite(A_ThisLabel, strTargetWinId, strTargetClass, strTargetControl)
{
	StringReplace, strThisHotkey, A_ThisHotkey, $ ; remove $ from hotkey
	if (A_ThisLabel = "PopupMenuMouse")
	{
		SplitModifiersFromKey(strThisHotkey, strThisHotkeyModifiers, strThisHotkeyKey)
		SendInput, %strThisHotkeyModifiers%{%strThisHotkeyKey%} ; for example {MButton} or +{LButton}
	}
	else
	{
		StringLower, strThisHotkey, strThisHotkey
		SendInput, %strThisHotkey% ; for example #a
	}
	return
}

blnMouse := InStr(A_ThisLabel, "Mouse")
blnNewWindow := false ; used in SetMenuPosition and BuildGroupMenu
Gosub, SetMenuPosition ; sets strTargetWinId or activate the window strTargetWinId set by CanOpenFavorite

if (A_ThisLabel = "PopupMenuMouse") and (WindowIsDirectoryOpus(strTargetClass) or WindowIsTotalCommander(strTargetClass))
{
	Click ; to make sure the DOpus lister or TC pane under the mouse become active
	Sleep, 20
}

; Can't find how to navigate a dialog box to My Computer or Network Neighborhood... is this is feasible?
if (blnDisplaySpecialFolders)
{
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass)
		or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuMyComputer%
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass)
		or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuNetworkNeighborhood%

	; There is no point to navigate a dialog box or console to Control Panel or Recycle Bin, unknown for FreeCommander
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass)
		or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuControlPanel%
	Menu, menuSpecialFolders
		, % WindowIsConsole(strTargetClass) or WindowIsFreeCommander(strTargetClass)
		or WindowIsDialog(strTargetClass) ? "Disable" : "Enable"
		, %lMenuRecycleBin%
}

if (blnDisplayGroupMenu)
{
	Gosub, BuildGroupMenu
	Menu, menuGroup
		, % (!intExplorersExist ? "Disable" : "Enable") ; disable Save group menu if no Explorer
		, %lMenuGroupSave%
}

if (WindowIsAnExplorer(strTargetClass) or WindowIsDesktop(strTargetClass) or WindowIsConsole(strTargetClass)
	or WindowIsFreeCommander(strTargetClass) or WindowIsTotalCommander(strTargetClass)
	or WindowIsDirectoryOpus(strTargetClass) or WindowIsDialog(strTargetClass))
{
	; Enable Add This Folder only if the mouse is over an Explorer (tested on WIN_XP and WIN_7) or a dialog box (works on WIN_7, not on WIN_XP)
	; Other tests shown that WIN_8 behaves like WIN_7. So, I assume WIN_8 to work. If someone could confirm (until I can test it myself)?
	if (blnDiagMode)
		Diag("ShowMenu", "Favorites Menu " 
			. (WindowIsAnExplorer(strTargetClass)  or WindowIsFreeCommander(strTargetClass)
			or WindowIsTotalCommander(strTargetClass) or WindowIsDirectoryOpus(strTargetClass)
			or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "WITH" : "WITHOUT")
			. " Add this folder")
	Menu, %lMainMenuName%
		, % WindowIsAnExplorer(strTargetClass) or WindowIsFreeCommander(strTargetClass)
		or WindowIsTotalCommander(strTargetClass) or WindowIsDirectoryOpus(strTargetClass)
		or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
		, %lMenuAddThisFolder%...
	Menu, %lMainMenuName%, Show, %intMenuPosX%, %intMenuPosY% ; at mouse pointer if option 1, 20x20 offset of active window if option 2 and fix location if option 3
}

return
;------------------------------------------------------------


;------------------------------------------------------------
PopupMenuNewWindowMouse: ; default +MButton::
PopupMenuNewWindowKeyboard: ; default +#a
;------------------------------------------------------------

if !(blnMenuReady)
	return

blnMouse := InStr(A_ThisLabel, "Mouse")
blnNewWindow := true ; used in SetMenuPosition and BuildGroupMenu
Gosub, SetMenuPosition ; sets strTargetWinId

WinGetClass strTargetClass, % "ahk_id " . strTargetWinId

if (A_ThisLabel = "PopupMenuMouse") and WindowIsTotalCommander(strTargetClass)
{
	Click ; to make sure the TC pane under the mouse become active before creating a new tab
	Sleep, 20
}

if (blnDiagMode)
{
	Diag("MouseOrKeyboard", A_ThisLabel)
	WinGetTitle strDiag, % "ahk_id " . strTargetWinId
	Diag("WinTitle", strDiag)
	Diag("WinId", strTargetWinId)
	Diag("Class", strTargetClass)
	Diag("ShowMenu", "Favorites Menu " . (WindowIsAnExplorer(strTargetClass) or WindowIsFreeCommander(strTargetClass) 
	or WindowIsTotalCommander(strTargetClass) or WindowIsDirectoryOpus(strTargetClass)
		or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "WITH" : "WITHOUT") . " Add this folder")
}

if (blnDisplaySpecialFolders)
{
	; In case it was disabled while in a dialog box
	Menu, menuSpecialFolders, Enable, %lMenuMyComputer%
	Menu, menuSpecialFolders, Enable, %lMenuNetworkNeighborhood%
	Menu, menuSpecialFolders, Enable, %lMenuControlPanel%
	Menu, menuSpecialFolders, Enable, %lMenuRecycleBin%
}

if (blnDisplayGroupMenu)
{
	Gosub, BuildGroupMenu
	Menu, menuGroup
		, % (!intExplorersExist ? "Disable" : "Enable") ; disable Save group menu if no Explorer
		, %lMenuGroupSave%
}

; Enable "Add This Folder" only if the target window is an Explorer (tested on WIN_XP and WIN_7)
; or a dialog box under WIN_7 (does not work under WIN_XP).
; Other tests shown that WIN_8 behaves like WIN_7.
Menu, %lMainMenuName%
	, % WindowIsAnExplorer(strTargetClass) or WindowIsFreeCommander(strTargetClass)
	or WindowIsTotalCommander(strTargetClass) or WindowIsDirectoryOpus(strTargetClass)
	or (WindowIsDialog(strTargetClass) and InStr("WIN_7|WIN_8", A_OSVersion)) ? "Enable" : "Disable"
	, %lMenuAddThisFolder%...
Menu, %lMainMenuName%, Show, %intMenuPosX%, %intMenuPosY% ; at mouse pointer if option 1, 20x20 offset of active window if option 2 and fix location if option 3

return
;------------------------------------------------------------


;------------------------------------------------------------
SetMenuPosition:
;------------------------------------------------------------

; relative to active window if option intPopupMenuPosition = 2
CoordMode, Mouse, % (intPopupMenuPosition = 2 ? "Window" : "Screen")
CoordMode, Menu, % (intPopupMenuPosition = 2 ? "Window" : "Screen")

if (intPopupMenuPosition = 1) ; display menu near mouse pointer location
	MouseGetPos, intMenuPosX, intMenuPosY
else if (intPopupMenuPosition = 2) ; display menu at an offset of 20x20 pixel from top-left of active window area
{
	intMenuPosX := 20
	intMenuPosY := 20
}
else ; (intPopupMenuPosition =  3) - fix position - use the intMenuPosX and intMenuPosY values from the ini file
{
	intMenuPosX := arrPopupFixPosition1
	intMenuPosY := arrPopupFixPosition2
}

; not related to set position but this is a good place to execute it ;-)
if (blnMouse)
	if (blnNewWindow)
		MouseGetPos, , , strTargetWinId ; sets strTargetWinId for PopupMenuNewWindowMouse
	else
		WinActivate, % "ahk_id " . strTargetWinId ; activate for PopupMenuMouse
else (keyboard)
	if (blnNewWindow)
		strTargetWinId := WinExist("A") ; sets strTargetWinId for PopupMenuNewWindowKeyboard

return
;------------------------------------------------------------



;========================================================================================================================
; CLASS
;========================================================================================================================


;------------------------------------------------------------
CanOpenFavorite(strMouseOrKeyboard, ByRef strWinId, ByRef strClass, ByRef strControlID)
;------------------------------------------------------------
; "CabinetWClass" and "ExploreWClass" -> Explorer
; "ProgMan" -> Desktop
; "WorkerW" -> Desktop
; "ConsoleWindowClass" -> Console (CMD)
; "#32770" -> Dialog
; "bosa_sdm_" (...) -> Dialog MS Office under WinXP
{
	
	global blnUseDirectoryOpus
	global blnUseTotalCommander
	
	if (strMouseOrKeyboard = "PopupMenuMouse")
		MouseGetPos, , , strWinId, strControlID
	else
	{
		strWinId := WinExist("A")
		strControlID := ""
	}
	WinGetClass strClass, % "ahk_id " . strWinId

	blnCanOpenFavorite := WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or WindowIsConsole(strClass)
		or WindowIsDialog(strClass) or WindowIsFreeCommander(strClass)
		or (WindowIsTotalCommander(strClass) and blnUseTotalCommander)
		or (WindowIsDirectoryOpus(strClass) and blnUseDirectoryOpus)

	if (blnDiagMode)
	{
		Diag("MouseOrKeyboard", strMouseOrKeyboard)
		WinGetTitle strDiag, % "ahk_id " . strWinId
		Diag("WinTitle", strDiag)
		Diag("WinId", strWinId)
		Diag("ControlID", strControlID)
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
	return (strClass = "ProgMan") or (strClass = "WorkerW") or (strClass = "Shell_TrayWnd")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsTray(strClass)
;------------------------------------------------------------
{
	return (strClass = "Shell_TrayWnd") or (strClass = "NotifyIconOverflowWindow")
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
	return (strClass = "#32770")
	; or InStr(strClass, "bosa_sdm_")
	; Removed 2014-09-27  (see http://code.jeanlalonde.ca/folderspopup/#comment-7912)
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
WindowIsTotalCommander(strClass)
;------------------------------------------------------------
{
	return InStr(strClass, "TTOTAL_CMD")
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDirectoryOpus(strClass)
;------------------------------------------------------------
{
	return InStr(strClass, "dopus")
}
;------------------------------------------------------------



;========================================================================================================================
; MENU ACTIONS
;========================================================================================================================


;------------------------------------------------------------
OpenFavorite:
;------------------------------------------------------------

if (blnDisplayMenuShortcuts)
	StringTrimLeft, strThisMenu, A_ThisMenuItem, 3 ; remove "&1 " from menu item
else
	strThisMenu := A_ThisMenuItem
strLocation := GetLocationFor(A_ThisMenu, strThisMenu)
strFavoriteType := GetFavoriteTypeFor(A_ThisMenu, strThisMenu)

if (blnDiagMode)
{
	Diag("A_ThisHotkey", A_ThisHotkey)
	Diag("Navigate", "OpenFavorite")
	Diag("Path", strLocation)
	Diag("FavoriteType", strFavoriteType)
}

if !FileExist(strLocation) ; this favorite does not exist
	if (strFavoriteType <> "U") ;  and this is not an URL
	{
		Gui, 1:+OwnDialogs
		MsgBox, 0, % L(lDialogFavoriteDoesNotExistTitle, strAppName)
			, % L(lDialogFavoriteDoesNotExistPrompt, strLocation)
			, 30
		return
	}

if (strFavoriteType = "D" or strFavoriteType = "U") ; this is a document or an URL
{
	Run, %strLocation%
	return
}

; else this is a folder

if InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "RunNewWindow")
	
	if (blnUseDirectoryOpus)
	{
		RunDOpusRt("/acmd Go ", strLocation, " " . strDirectoryOpusNewTabOrWindow) ; open in a new lister or tab
		WinActivate, ahk_class dopus.lister
	}
	else if (blnUseTotalCommander)
		NewTotalCommander(strLocation, strTargetWinId, strTargetControl)
	else
		if (A_OSVersion = "WIN_XP")
			ComObjCreate("Shell.Application").Explore(strLocation)
		else
			Run, Explorer %strLocation%
	; http://msdn.microsoft.com/en-us/library/bb774094
	; ComObjCreate("Shell.Application").Explore(strLocation)
	; ComObjCreate("WScript.Shell").Exec("Explorer.exe /e /select," . strLocation) ; not tested on XP
	; ComObjCreate("Shell.Application").Open(strLocation)
	; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
}
else if WindowIsAnExplorer(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateExplorer")
	
	NavigateExplorer(strLocation, strTargetWinId)
}
else if WindowIsConsole(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateConsole")
	
	NavigateConsole(strLocation, strTargetWinId)
}
else if WindowIsDirectoryOpus(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "DirectoryOpus")

	NavigateDirectoryOpus(strLocation, strTargetWinId)
}
else if WindowIsFreeCommander(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateFreeCommander")
	
	NavigateFreeCommander(strLocation, strTargetWinId, strTargetControl)
}
else if WindowIsTotalCommander(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateTotalCommander")
	
	NavigateTotalCommander(strLocation, strTargetWinId, strTargetControl)
}
else ; dialog box
{
	if (blnDiagMode)
	{
		Diag("Navigate", "NavigateDialog")
		Diag("TargetClass", strTargetClass)
	}
	
	NavigateDialog(strLocation, strTargetWinId, strTargetClass)
}

if (blnDiagMode)
	Diag("NavigateResult", ErrorLevel)

return
;------------------------------------------------------------


;------------------------------------------------------------
GetLocationFor(strMenu, strName)
;------------------------------------------------------------
{
	global
	
	Loop, % arrMenus[strMenu].MaxIndex()
		if (strName = arrMenus[strMenu][A_Index].FavoriteName)
			return arrMenus[strMenu][A_Index].FavoriteLocation
}
;------------------------------------------------------------


;------------------------------------------------------------
GetFavoriteTypeFor(strMenu, strName)
;------------------------------------------------------------
{
	global
	
	Loop, % arrMenus[strMenu].MaxIndex()
		if (strName = arrMenus[strMenu][A_Index].FavoriteName)
			return arrMenus[strMenu][A_Index].FavoriteType
}
;------------------------------------------------------------


;------------------------------------------------------------
OpenSpecialFolder:
;------------------------------------------------------------

if (blnDiagMode)
{
	Diag("A_ThisHotkey", A_ThisHotkey)
	Diag("Navigate", "SpecialFolder")
}

strLocation := ""
blnNewSpecialWindow := InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass) or WindowIsTray(strTargetClass)

if (blnNewSpecialWindow and blnUseDirectoryOpus) or WindowIsDirectoryOpus(strTargetClass)
{
	if (A_ThisMenuItem = lMenuDesktop)
		strDOpusAlias := "desktop"
	else if (A_ThisMenuItem = lMenuControlPanel)
		strDOpusAlias := "controls"
	else if (A_ThisMenuItem = lMenuDocuments)
		strDOpusAlias := "mydocuments"
	else if (A_ThisMenuItem = lMenuRecycleBin)
		strDOpusAlias := "trash"
	else if (A_ThisMenuItem = lMenuMyComputer)
		strDOpusAlias := "mycomputer"
	else if (A_ThisMenuItem = lMenuNetworkNeighborhood)
		strDOpusAlias := "network"
	else if (A_ThisMenuItem = lMenuPictures)
		strDOpusAlias := "mypictures"
	
	if (blnNewSpecialWindow)
	{
		RunDOpusRt("/acmd Go ", "/" . strDOpusAlias, " " . strDirectoryOpusNewTabOrWindow) ; open special folder in a new lister or tab
		WinActivate, ahk_class dopus.lister
	}
	else
		NavigateDirectoryOpus("/" . strDOpusAlias, strTargetWinId)
}
else if (blnNewSpecialWindow and blnUseTotalCommander) or WindowIsTotalCommander(strTargetClass)
{
	intTCCommand := 0
	if (A_ThisMenuItem = lMenuDesktop)
		intTCCommand := 2121 ; cm_OpenDesktop
	else if (A_ThisMenuItem = lMenuControlPanel)
		intTCCommand := 2123 ; cm_OpenControls
	else if (A_ThisMenuItem = lMenuDocuments)
		strLocation := A_MyDocuments
	else if (A_ThisMenuItem = lMenuRecycleBin)
		intTCCommand := 2127 ; cm_OpenRecycled
	else if (A_ThisMenuItem = lMenuMyComputer)
		intTCCommand := 2122 ; cm_OpenDrives
	else if (A_ThisMenuItem = lMenuNetworkNeighborhood)
		intTCCommand := 2125 ; cm_OpenNetwork
	else if (A_ThisMenuItem = lMenuPictures)
		RegRead, strLocation, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Pictures
	
	if (intTCCommand = 2123)
		Run, Control ; workaround for command 2123 cm_OpenControls not working in my tests with TC v8.51a
	else
	{
		if (intTCCommand)
			if (blnNewSpecialWindow)
			{
				if !WinExist("ahk_class TTOTAL_CMD") ; open a first instance and get PID
					or InStr(strTotalCommanderNewTabOrWindow, "/N") ; open a new instance and get PID
					{
						Run, %strTotalCommanderPath%
						WinWait, A, , 10
						Sleep, 200 ; wait additional time to improve SendMessage reliability
					}
				if !InStr(strTotalCommanderNewTabOrWindow, "/N") ; open the folder in a new tab
				{
					intTCCommandOpenNewTab := 3001 ; cm_OpenNewTab
					Sleep, 100 ; wait to improve SendMessage reliability
					SendMessage, 0x433, %intTCCommandOpenNewTab%, , , ahk_class TTOTAL_CMD
				}
				Sleep, 100 ; wait to reduce risk SendMessage not working
				SendMessage, 0x433, %intTCCommand%, , , ahk_class TTOTAL_CMD
				Sleep, 100 ; wait to improve SendMessage reliability
				WinActivate, ahk_class TTOTAL_CMD
			}
			else
			{
				Sleep, 100 ; wait to improve SendMessage reliability
				SendMessage, 0x433, %intTCCommand%, , , ahk_class TTOTAL_CMD
				Sleep, 100 ; wait to improve SendMessage reliability
				WinActivate, ahk_class TTOTAL_CMD
			}
		else ; with strLocation
			if (blnNewSpecialWindow)
				NewTotalCommander(strLocation, strTargetWinId, strTargetControl)
			else
			{
				NavigateTotalCommander(strLocation, strTargetWinId, strTargetControl)
				WinActivate, ahk_class TTOTAL_CMD
			}
	}
}
else ; this is Explorer, Console, FreeCommander or a dialog box
{
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
	
	if (blnNewSpecialWindow)
		ComObjCreate("Shell.Application").Explore(intSpecialFolder)
		; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
	else
	{
		if WindowIsAnExplorer(strTargetClass)
			NavigateExplorer(intSpecialFolder, strTargetWinId)
		else ; this is Console, FreeCommander or a dialog box
		{
			if (intSpecialFolder = 0)
				strLocation := A_Desktop
			else if (intSpecialFolder = 5)
				strLocation := A_MyDocuments
			else if (intSpecialFolder = 39)
				RegRead, strLocation, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Pictures
			else
				; We cannot open the special folders lMenuMyComputer, lMenuNetworkNeighborhood, lMenuControlPanel and lMenuRecycleBin
				; in the current window. Need to open an Explorer with ComObjCreate
				blnNewSpecialWindow := true
			
			if (blnNewSpecialWindow)
				ComObjCreate("Shell.Application").Explore(intSpecialFolder)
			else if WindowIsConsole(strTargetClass)
				NavigateConsole(strLocation, strTargetWinId)
			else if WindowIsFreeCommander(strTargetClass)
				NavigateFreeCommander(strLocation, strTargetWinId, strTargetControl)
			else
				NavigateDialog(strLocation, strTargetWinId, strTargetClass)
		}
	}
}

if (blnDiagMode)
{
	Diag("A_ThisHotkey", A_ThisHotkey)
	Diag("Navigate", "SpecialFolder")
	Diag("SpecialFolderWindows", intSpecialFolder)
	Diag("SpecialFolderDOpus", strDOpusAlias)
	Diag("SpecialFolderTCCommand", intTCCommand)
	Diag("SpecialFolderLocation", strLocation)
	Diag("SpecialFolderNewWindow", blnNewSpecialWindow)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
OpenRecentFolder:
;------------------------------------------------------------

if InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass)
	if (blnUseDirectoryOpus)
	{
		RunDOpusRt("/acmd Go ", objRecentFolders[A_ThisMenuItemPos], " " . strDirectoryOpusNewTabOrWindow) ; open in a new lister or tab
		WinActivate, ahk_class dopus.lister
	}
	else if (blnUseTotalCommander)
		NewTotalCommander(objRecentFolders[A_ThisMenuItemPos], strTargetWinId, strTargetControl)
	else
		; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
		ComObjCreate("Shell.Application").Explore(objRecentFolders[A_ThisMenuItemPos])
else if WindowIsAnExplorer(strTargetClass)
	NavigateExplorer(objRecentFolders[A_ThisMenuItemPos], strTargetWinId)
else if WindowIsDirectoryOpus(strTargetClass)
	if (blnUseDirectoryOpus)
		RunDOpusRt("/aCmd Go", objRecentFolders[A_ThisMenuItemPos]) ; navigate the current lister
	else
		NavigateExplorer(objRecentFolders[A_ThisMenuItemPos], strTargetWinId)
else ; this is the console, FreeCommander or a dialog box
	if WindowIsConsole(strTargetClass)
		NavigateConsole(objRecentFolders[A_ThisMenuItemPos], strTargetWinId)
	else if WindowIsFreeCommander(strTargetClass)
		NavigateFreeCommander(objRecentFolders[A_ThisMenuItemPos], strTargetWinId, strTargetControl)
	else if WindowIsTotalCommander(strTargetClass)
	if (blnUseTotalCommander)
		NavigateTotalCommander(objRecentFolders[A_ThisMenuItemPos], strTargetWinId, strTargetControl)
	else
		NavigateExplorer(objRecentFolders[A_ThisMenuItemPos], strTargetWinId)
	else
		NavigateDialog(objRecentFolders[A_ThisMenuItemPos], strTargetWinId, strTargetClass)

return
;------------------------------------------------------------


;------------------------------------------------------------
SwitchExplorer:
;------------------------------------------------------------

if (objGroupMenuExplorers[A_ThisMenuItemPos].WindowType = "DO") ; this is a DOpus lister
{
	if InStr(objGroupMenuExplorers[A_ThisMenuItemPos].Path, "%")
	{
		; double % for DOpusRT (http://resource.dopus.com/viewtopic.php?f=3&t=23013#p124395)
		strDOpusPathOK := objGroupMenuExplorers[A_ThisMenuItemPos].Path
		StringReplace, strDOpusPathOK, strDOpusPathOK, % "%", % "%%", A
		objGroupMenuExplorers[A_ThisMenuItemPos].Path := strDOpusPathOK
	}
	RunDOpusRt("/acmd Go ", objGroupMenuExplorers[A_ThisMenuItemPos].Path, " EXISTINGLISTER") ; activate an existing lister listing this path
}
else ; Explorer
	WinActivate, % "ahk_id " . objGroupMenuExplorers[A_ThisMenuItemPos].WindowId

return
;------------------------------------------------------------


;------------------------------------------------------------
SwitchDialog:
;------------------------------------------------------------

NavigateDialog(objGroupMenuExplorers[A_ThisMenuItemPos].Path, strTargetWinId, strTargetClass)

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupSave:
;------------------------------------------------------------

Gui, 2:New, , % L(lGuiGroupSaveTitle, strAppName, strAppVersion)
Gui, 2:Margin, 10, 10
Gui, 2:Color, %strGuiWindowColor%
Gui, 2:+OwnDialogs

Gui, 2:Font, s10 w700, Verdana
Gui, 2:Add, Text, x10 y10 w670 center, %lMenuGroupSave%
Gui, 2:Font

Gui, 2:Add, Text, x10 y+15 w670 center, %lGuiGroupSaveSelect%

Gui, 2:Add, ListView
	, xm w680 h200 Checked Count32 -Multi NoSortHdr LV0x10 c%strGuiListviewTextColor% Background%strGuiListviewBackgroundColor% vlvGroupList
	, %lGuiGroupSaveLvHeader%|Hidden: Path|IsSpecialFolder|WindowId|Position|TabId
Loop, 4
	LV_ModifyCol(A_Index + 3, "Right")

for intIndex, objGroupMenuExplorer in objGroupMenuExplorers
{
	arrExplorerPosition := objGroupMenuExplorer.Position
	StringSplit, arrExplorerPosition, arrExplorerPosition, "|"
	LV_Add("Check", objGroupMenuExplorer.Name
		, (objGroupMenuExplorer.WindowType = "DO" ? "Directory Opus" : (objGroupMenuExplorer.WindowType = "TC" ? "Total Commander" : "Windows Explorer"))
		, (objGroupMenuExplorer.MinMax = -1 ? "Minimized" : (objGroupMenuExplorer.MinMax = "1" ? "Maximized" : "Normal"))
		, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition1 : "-")
		, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition2 : "-")
		, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition3 : "-")
		, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition4 : "-")
		, (objGroupMenuExplorer.Pane ? objGroupMenuExplorer.Pane : "-")
		, objGroupMenuExplorer.Path, objGroupMenuExplorer.IsSpecialFolder, objGroupMenuExplorer.WindowId, objGroupMenuExplorer.Position, objGroupMenuExplorer.TabId)
}

LV_Modify(1, "Select Focus")
Loop, 12
	LV_ModifyCol(A_Index, "AutoHdr")
Loop, % (blnUseDirectoryOpus ? 5 : 6)
	LV_ModifyCol(A_Index + (blnUseDirectoryOpus ? 8 : 7), 0) ; hide 8th-12th columns

Gui, 2:Add, Text, x10 y+10 w250 section, %lGuiGroupSaveShortName%
Gui, 2:Add, Edit, x10 y+5 w250 Limit248 vstrGroupSaveName, %strGroupSaveName% ; maximum length of ini section name is 255 ("Group-" is prepended)

Gui, 2:Add, Text, ys x300, %lGuiGroupSaveRestoreOption%:
Gui, 2:Add, Radio, x+10 yp section vblnRadioReplaceWindows checked, %lGuiGroupSaveReplaceWindowsLabel%
Gui, 2:Add, Radio, xs y+5 vblnRadioAddWindows, %lGuiGroupSaveAddWindowsLabel%

Gui, 2:Add, Button, y+20 vbtnGroupSave gButtonGroupSave, %lGuiSave%
Gui, 2:Add, Button, yp vbtnGroupSaveCancel gButtonGroupSaveCancel, %lGuiCancel%
GuiCenterButtons(L(lGuiGroupSaveTitle, lMenuGroupSave, strAppName, strAppVersion), 10, 5, 20, , "btnGroupSave", "btnGroupSaveCancel")

GuiControl, Focus, strGroupSaveName
GuiControl, +Default, btnGroupSave
Gui, 2:Show, AutoSize Center

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonGroupSave:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

if !StrLen(strGroupSaveName)
{
	Oops(lGuiGroupSaveNameEmpty)
	return
}

if GroupExists(strGroupSaveName)
{
	MsgBox, 52, % L(lGuiGroupSaveTitle, lMenuGroupSave, strAppName), % L(lGuiGroupSaveReplaceGroup, strGroupSaveName)
	IfMsgBox, No
		return
}

if GroupExists(strGroupSaveName)
	IniDelete, %strIniFile%, Group-%strGroupSaveName%
else
{
	strGroups := strGroups . (StrLen(strGroups) ? "|" : "") . strGroupSaveName
	IniWrite, %strGroups%, %strIniFile%, Global, Groups
}

IniWrite, %blnRadioReplaceWindows%, %strIniFile%, Group-%strGroupSaveName%, ReplaceWhenRestoringThisGroup
intRow := 0
Loop
{
	intRow := LV_GetNext(intRow, "Checked")
    if !(intRow)
        break
	IniWrite, % objGroupMenuExplorers[intRow].Name
		. "|" . objGroupMenuExplorers[intRow].WindowType
		. "|" . objGroupMenuExplorers[intRow].MinMax
		. "|" . objGroupMenuExplorers[intRow].Position
		. "|" . objGroupMenuExplorers[intRow].Pane
		. "|" . objGroupMenuExplorers[intRow].WindowId
		. "|" . objGroupMenuExplorers[intRow].IsSpecialFolder
		, %strIniFile%, Group-%strGroupSaveName%, Explorer%A_Index%
}

strUseMenuLoad := lMenuGroup . " > " . lMenuGroupLoad
StringReplace, strUseMenuLoad, strUseMenuLoad, &, , All
MsgBox, 0, % L(lGuiGroupSaveTitle, lMenuGroupSave, strAppName)
	, % L(lGuiGroupSaved, strGroupSaveName, strUseMenuLoad)
	, 30

Gosub, ButtonGroupSaveCancel
strGroupSaveName := ""

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonGroupSaveCancel:
;------------------------------------------------------------

Gui, 2:Destroy

return
;------------------------------------------------------------


;------------------------------------------------------------
GroupExists(strGroup)
;------------------------------------------------------------
{
	global strGroups
	
	Loop, Parse, strGroups, |
		if (A_LoopField = strGroup)
			return True
		
	return False
}
;------------------------------------------------------------


;------------------------------------------------------------
GroupLoad:
GroupLoadFromManage:
;------------------------------------------------------------

intSleepAfterWindowCommand := 400

if (A_ThisLabel = "GroupLoad")
{
	strGroup := "Group-" . A_ThisMenuItem
	StringLeft, strGroup, strGroup, % InStr(strGroup, "(") - 2 ; remove " (add)" or " (replace)"
}
else
	strGroup := "Group-" . strSelectedGroup

IniRead, blnReplaceWhenRestoringThisGroup, %strIniFile%, %strGroup%, ReplaceWhenRestoringThisGroup, 0 ; false if not found

if (blnReplaceWhenRestoringThisGroup)
	Gosub, CloseBeforeRestoringGroup

objIniExplorersInGroup := Object()
Gosub, Group2Object
intTotalWindowsInIni := objIniExplorersInGroup.MaxIndex()
intActualWindowInIni := 1

while, intExplorer := WindowOfType("EX") ; returns the index of the first Explorer saved window in the group
{
	Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni%
	if (objIniExplorersInGroup[intExplorer].IsSpecialFolder)
		strExplorerLocationOrClassId := "shell:::" . objClassIdByLocalizedName[objIniExplorersInGroup[intExplorer].Name]
	else
		strExplorerLocationOrClassId := objIniExplorersInGroup[intExplorer].Name
	Run, % "explorer.exe " . strExplorerLocationOrClassId,
		, % (objIniExplorersInGroup[intExplorer].MinMax = -1 ? "Min" : (objIniExplorersInGroup[intExplorer].MinMax = 1 ? "Max" : ""))
		
	Sleep, %intSleepAfterWindowCommand%
	WinWaitActive, A, , 5
	Sleep, 200
	
	if !(objIniExplorersInGroup[intExplorer].MinMax) ; window normal, not min nor max
		WinMove, A, 
			, % objIniExplorersInGroup[intExplorer].Left
			, % objIniExplorersInGroup[intExplorer].Top
			, % objIniExplorersInGroup[intExplorer].Width
			, % objIniExplorersInGroup[intExplorer].Height
	
	objIniExplorersInGroup.Remove(intExplorer) ; remove the first Explorer saved window from the group
	intActualWindowInIni := intActualWindowInIni + 1
}

while, intDOWindow := WindowOfType("DO") ; returns the index of the first DOpus saved window in the group
{
	strFirstWindowId := objIniExplorersInGroup[intDOWindow].WindowId
	strNewWindowId := ""
	
	while, intDOIndex := DOpusWindowOfId(strFirstWindowId) ; returns the paths of the first group of DOpus windows (grouped by WindowsId)
	{
		if !StrLen(strNewWindowId) ; the window does not exist, create it with the first path in the pane 1
		{
			Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni%
			RunDOpusRt("/acmd Go ", objIniExplorersInGroup[intDOWindow].Name, " NEW=nodual") ; open in a new lister
			Sleep, %intSleepAfterWindowCommand%
			WinWaitActive, A, , 5
			Sleep, 200
			
			strNewWindowId := WinExist("A")
			if !(objIniExplorersInGroup[intDOWindow].MinMax) ; window normal, not min nor max
				WinMove, A,
					, % objIniExplorersInGroup[intDOWindow].Left
					, % objIniExplorersInGroup[intDOWindow].Top
					, % objIniExplorersInGroup[intDOWindow].Width
					, % objIniExplorersInGroup[intDOWindow].Height
			else
				if (objIniExplorersInGroup[intDOWindow].MinMax = -1)
					WinMinimize, A
				else
					WinMaximize, A

			objIniExplorersInGroup.Remove(intDOWindow) ; remove the first DOpus saved window from the group
			intActualWindowInIni := intActualWindowInIni + 1
		}
		else ; the window exist, add other path of pane tabs or pane 2
		{
			while, intDOIndexPane := DOpusWindowOfPane(1, strFirstWindowId) ; returns the paths of the windows of the left pane for this group of windows
			{
				Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni%
				RunDOpusRt("/acmd Go ", objIniExplorersInGroup[intDOIndexPane].Name, " NEWTAB") ; open in a new tab of pane 1
				Sleep, %intSleepAfterWindowCommand%
				objIniExplorersInGroup.Remove(intDOIndexPane) ; remove the first DOpus saved window from the group
				intActualWindowInIni := intActualWindowInIni + 1
			}
			blnNewPane2 := true
			while, intDOIndexPane := DOpusWindowOfPane(2, strFirstWindowId) ; returns the paths of the windows of the right pane for this group of windows
			{
				Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni%
				if (blnNewPane2)
				{
					RunDOpusRt("/acmd Go ", objIniExplorersInGroup[intDOIndexPane].Name, " OPENINRIGHT") ; open in a first tab of pane 2
					Sleep, %intSleepAfterWindowCommand%
					blnNewPane2 := false
				}
				else
				{
					RunDOpusRt("/acmd Go ", objIniExplorersInGroup[intDOIndexPane].Name, " OPENINRIGHT NEWTAB") ; open in a new tab of pane 2
					Sleep, %intSleepAfterWindowCommand%
				}
				objIniExplorersInGroup.Remove(intDOIndexPane) ; remove the first DOpus saved window from the group
				intActualWindowInIni := intActualWindowInIni + 1
			}
		}
	}
}
Tooltip ; clear tooltip

return
;------------------------------------------------------------


;------------------------------------------------------------
CloseBeforeRestoringGroup:
;------------------------------------------------------------

intSleepTime := 67 ; for visual effect only...
Tooltip, %lGuiGroupClosing%

strWindowsId := ""
for objExplorer in ComObjCreate("Shell.Application").Windows
	; do not close in this loop as it mess up the handlers
	strWindowsId := strWindowsId . objExplorer.HWND . "|"
StringTrimRight, strWindowsId, strWindowsId, 1
Loop, Parse, strWindowsId, |
{
	WinClose, ahk_id %A_LoopField%
	Sleep, %intSleepTime%
}

if (blnUseTotalCommander)
{
	WinGet, arrIDs, List, ahk_class TTOTAL_CMD
	Loop, %arrIDs%
	{
		WinClose, % "ahk_id " . arrIDs%A_Index%
		Sleep, %intSleepTime%
	}
}

if (blnUseDirectoryOpus)
{
	WinGet, arrIDs, List, ahk_class dopus.lister
	Loop, %arrIDs%
	{
		WinClose, % "ahk_id " . arrIDs%A_Index%
		Sleep, %intSleepTime%
	}
}
Tooltip ; clear tooltip
	
return
;------------------------------------------------------------


;------------------------------------------------------------
Group2Object:
;------------------------------------------------------------

Loop
{
	IniRead, strExplorer, %strIniFile%, %strGroup%, Explorer%A_Index%
	if (strExplorer = "ERROR")
		Break

	StringSplit, arrThisExplorer, strExplorer, |

	;  1	objGroupMenuExplorers[intRow].Name
	;  2	objGroupMenuExplorers[intRow].WindowType
	;  3	objGroupMenuExplorers[intRow].MinMax
	;  4	objGroupMenuExplorers[intRow].Left
	;  5	objGroupMenuExplorers[intRow].Top
	;  6	objGroupMenuExplorers[intRow].Width
	;  7	objGroupMenuExplorers[intRow].Height
	;  8	objGroupMenuExplorers[intRow].Pane
	;  9	objGroupMenuExplorers[intRow].WindowId
	; 10	objGroupMenuExplorers[intRow].IsSpecialFolder
	
	objIniEntry := Object()
	objIniEntry.Name := arrThisExplorer1
	objIniEntry.WindowType := arrThisExplorer2
	objIniEntry.MinMax := arrThisExplorer3
	objIniEntry.Left := arrThisExplorer4
	objIniEntry.Top := arrThisExplorer5
	objIniEntry.Width := arrThisExplorer6
	objIniEntry.Height := arrThisExplorer7
	objIniEntry.Pane := arrThisExplorer8
	objIniEntry.WindowId := arrThisExplorer9
	objIniEntry.IsSpecialFolder := arrThisExplorer10
	
	objIniExplorersInGroup.Insert(A_Index, objIniEntry)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
WindowOfType(strType)
;------------------------------------------------------------
{
	global objIniExplorersInGroup

	intFound := 0
	for intIndex, intEntry in objIniExplorersInGroup
		if (objIniExplorersInGroup[intIndex].WindowType = strType)
		{
			intFound := intIndex
			break
		}
			
	return intFound
}

;------------------------------------------------------------


;------------------------------------------------------------
DOpusWindowOfId(strId)
;------------------------------------------------------------
{
	global objIniExplorersInGroup

	intFound := 0
	for intIndex, intEntry in objIniExplorersInGroup
		if (objIniExplorersInGroup[intIndex].WindowId = strId)
		{
			intFound := intIndex
			break
		}
			
	return intFound
}

;------------------------------------------------------------


;------------------------------------------------------------
DOpusWindowOfPane(intPane, strId)
; intPane can be 0 or 1 if in left side, or 2 if in right side
;------------------------------------------------------------
{
	global objIniExplorersInGroup

	intFound := 0
	for intIndex, intEntry in objIniExplorersInGroup
		if (objIniExplorersInGroup[intIndex].Pane = intPane and objIniExplorersInGroup[intIndex].WindowId = strId)
			{
				intFound := intIndex
				break
			}
			
	return intFound
}

;------------------------------------------------------------



;========================================================================================================================
; NAVIGATE
;========================================================================================================================


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
	if !Regexmatch(varPath, "#.*\\") ; prevent the hash bug in Shell.Application - when a hash in path is followed by a backslash like in "c:\abc#xyz\abc")
	{
		intCountMatch := 0
		For pExplorer in ComObjCreate("Shell.Application").Windows
		{
			if (pExplorer.hwnd = strWinId)
				if varPath is integer ; ShellSpecialFolderConstant
				{
					intCountMatch := intCountMatch + 1
					try pExplorer.Navigate2(varPath)
					catch, objErr
						Oops(lNavigateSpecialError, varPath)
				}
				else
				{
					intCountMatch := intCountMatch + 1
					; Was "try pExplorer.Navigate("file:///" . varPath)" - removed "file:///" to allow UNC (e.g. \\my.server.com@SSL\DavWWWRoot\Folder\Subfolder)
					try pExplorer.Navigate(varPath)
					catch, objErr
						; Note: If an error occurs in Navigate, the error message is given by Navigate itself and this script does not
						; receive an error notification. From my experience, the following line would never be executed.
						Oops(lNavigateFileError, varPath)
				}
		}
		if !(intCountMatch)
		; for Explorer add-ons like Clover (verified - it now opens the folder in a new tab), others?
		; also when strWinId is DOpus window and DOpus is not used
			if varPath is integer ; ShellSpecialFolderConstant
				ComObjCreate("Shell.Application").Explore(varPath)
			else
				Run, Explorer %varPath%
	}
	else
		; Workaround for the hash (aka Sharp / "#") bug in Shell.Application - occurs only when navigating in the current Explorer window
		; see http://stackoverflow.com/questions/22868546/navigate-shell-command-not-working-when-the-path-includes-an-hash
		; and http://ahkscript.org/boards/viewtopic.php?f=5&t=526&p=25287#p25274
		SendInput, {F4}{Esc}{Raw}%varPath%`n
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateConsole(strLocation, strWinId)
;------------------------------------------------------------
{
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
	SendInput, {Raw}CD /D %strLocation%
	Sleep, 200
	SendInput, {Enter}
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateDirectoryOpus(strLocation, strWinId)
;------------------------------------------------------------
{
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window

	RunDOpusRt("/aCmd Go", strLocation) ; navigate the current lister
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateFreeCommander(strLocation, strWinId, strControl)
;------------------------------------------------------------
{
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
	{
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
		Sleep, 200
	}
	; to make sure the file list under the mouse become active
	SetKeyDelay, 200
	ControlFocus, %strControl%
	Sleep, 200
	SendInput, !g ; Alt-G goto shortcut for FreeCommander
	Sleep, 200
	SendInput, {Raw}%strLocation%
	Sleep, 200
	SendInput, {Enter}
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateTotalCommander(strLocation, strWinId, strControl)
;------------------------------------------------------------
{
	global strTotalCommanderPath
	
	/*
	###_D(""
		. "strLocation: " . strLocation . "`n"
		. "strWinId: " . strWinId . "`n"
		. "strControl: " . strControl . "`n"
		. "")
	*/
	
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
	{
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
		Sleep, 200
	}
	Run, %strTotalCommanderPath% /O /S "/L=%strLocation%" ; /O existing file list, /S source-dest /L=source (active pane) - change folder in the active pane/tab
}
;------------------------------------------------------------


;------------------------------------------------------------
NewTotalCommander(strLocation, strWinId, strControl)
;------------------------------------------------------------
{
	global strTotalCommanderPath
	global strTotalCommanderNewTabOrWindow
	
	/*
	###_D(""
		. "strLocation: " . strLocation . "`n"
		. "strWinId: " . strWinId . "`n"
		. "strControl: " . strControl . "`n"
		. "")
	*/
	
	; strTotalCommanderNewTabOrWindow in ini file should contain "/O /T" to open in an new tab of the existing file list (default), or "/N" to open in a new file list
	Run, %strTotalCommanderPath% %strTotalCommanderNewTabOrWindow% /S "/L=%strLocation%" ; /L= left pane of the new window
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
	/*
	Remove support for Office 2003/2007 file dialog boxes
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
	*/
	Else ; in all other cases, open a new Explorer and return from this function
	{
		if (A_OSVersion = "WIN_XP")
			ComObjCreate("Shell.Application").Explore(strLocation)
		else
			Run, Explorer %strLocation%
		; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
		if (blnDiagMode)
			Diag("NavigateDialog", "Not #32770: open New Explorer")
		return
	}

	if (blnDiagMode)
	{
		Diag("NavigateDialogControl", strControl)
		Diag("NavigateDialogPath", strLocation)
	}
			
	;===In this part (if we reached it), we'll send strLocation to control and restore control's initial text after navigating to specified folder===
	ControlGetText, strPrevControlText, %strControl%, ahk_id %strWinId% ; we'll get and store control's initial text first
	
	if !ControlSetTextR(strControl, strLocation, "ahk_id " . strWinId) ; set control's text to strLocation
		return ; abort if control is not set
	if InStr(strClass, "bosa_sdm_") ; additional delay for XPs running MS Word or XL 
		Sleep, 100
	if !ControlSetFocusR(strControl, "ahk_id " . strWinId) ; focus control
		return
	if InStr(strClass, "bosa_sdm_") ; additional delay for XPs running MS Word or XL 
		Sleep, 100
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
	if InStr(strClass, "bosa_sdm_") ; additional delay for XPs running MS Word or XL 
		Sleep, 100
	
	;=== Avoid accidental hotkey & hotstring triggereing while doing SendInput - can be done simply by #UseHook, but do it if user doesn't have #UseHook in the script ===
	If (A_IsSuspended)
		blnWasSuspended := True
	if (!blnWasSuspended)
		Suspend, On
	SendInput, {End}{Space}{Backspace}{Enter} ; silly but necessary part - go to end of control, send dummy space, delete it, and then send enter
	if (!blnWasSuspended)
		Suspend, Off

	Sleep, 100 ; give some time to control after sending {Enter} to it
	ControlGetText, strControlTextAfterNavigation, %strControl%, ahk_id %strWinId% ; sometimes controls automatically restore their initial text
	if (strControlTextAfterNavigation <> strPrevControlText) ; if not
		ControlSetTextR(strControl, strPrevControlText, "ahk_id " . strWinId) ; we'll set control's text to its initial text
	
	if (WinExist("A") <> strWinId) ; sometimes initialy active window loses focus, so we'll activate it again
		WinActivate, ahk_id %strWinId%

	if (blnDiagMode)
		Diag("NavigateDialog", "Finished")
}
;------------------------------------------------------------


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
		Sleep, % (100 * A_Loop) ; JL added "* A_Loop"
		ControlGetFocus, strFocusedControl, %strWinTitle% ; check
		if (strFocusedControl = strControl) ; if OK
		{
			if (blnDiagMode)
				Diag("ControlSetFocusR Tries", A_Index)
			return True
		}
	}
	if (blnDiagMode)
		Diag("ControlSetFocusR", "Unable to set focus")
	return false
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
		Sleep, % (100 * A_Loop) ; JL added "* A_Loop"
		ControlGetText, strCurControlText, %strControl%, %strWinTitle% ; check
		if (strCurControlText = strNewText) ; if OK
		{
			if (blnDiagMode)
				Diag("ControlSetTextR Tries", A_Index)
			return True
		}
	}
	if (blnDiagMode)
		Diag("ControlSetTextR", "Unable to set text")
	return false
}
;------------------------------------------------------------



;========================================================================================================================
; TRAY MENU ACTIONS
;========================================================================================================================


;------------------------------------------------------------
RunAtStartup:
;------------------------------------------------------------
; Startup code adapted from Avi Aryan Ryan in Clipjump

Menu, Tray, Togglecheck, %lMenuRunAtStartup%
IfExist, %A_Startup%\%strAppName%.lnk
	FileDelete, %A_Startup%\%strAppName%.lnk
else
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%strAppName%.lnk, %A_WorkingDir%

return
;------------------------------------------------------------


;------------------------------------------------------------
SuspendHotkeys:
;------------------------------------------------------------

if (A_IsSuspended)
	Suspend, Off
else
	Suspend, On

Menu, Tray, % (A_IsSuspended ? "check" : "uncheck"), %lMenuSuspendHotkeys%

return
;------------------------------------------------------------


;------------------------------------------------------------
ExitFoldersPopup:
;------------------------------------------------------------

ExitApp

;------------------------------------------------------------


;------------------------------------------------------------
Check4Update:
;------------------------------------------------------------
Gui, 1:+OwnDialogs

if Time2Donate(intStartups, blnDonor)
{
	MsgBox, 36, % l(lDonateCheckTitle, intStartups, strAppName), % l(lDonateCheckPrompt, strAppName, intStartups)
	IfMsgBox, Yes
		Gosub, GuiDonate
}
IniWrite, % (intStartups + 1), %strIniFile%, Global, Startups

strLatestVersion := Url2Var("http://code.jeanlalonde.ca/ahk/folderspopup/latest-version/latest-version-" . strCurrentBranch . ".php")

if !StrLen(strLatestVersion)
	if (A_ThisMenuItem = lMenuUpdate)
	{
		Oops(lUpdateError)
		return ; an error occured during ComObjCreate
	}
strLatestVersion := SubStr(strLatestVersion, InStr(strLatestVersion, "[[") + 2) 
strLatestVersion := SubStr(strLatestVersion, 1, InStr(strLatestVersion, "]]") - 1) 
strLatestVersion := Trim(strLatestVersion, "`n`l") ; remove en-of-line if present

Loop, Parse, strLatestVersion, , 0123456789. ; strLatestVersion should only contain digits and dots
	; if we get here, the content returned by the URL above is wrong
	if (A_ThisMenuItem <> lMenuUpdate)
		return ; return silently
	else
	{
		Oops(lUpdateError) ; return with an error message
		return
	}

if (FirstVsSecondIs(strLatestSkipped, strLatestVersion) >= 0 and (A_ThisMenuItem <> lMenuUpdate))
	return

if FirstVsSecondIs(strLatestVersion, strCurrentVersion) = 1
{
	Gui, 1:+OwnDialogs
	SetTimer, ChangeButtonNamesUpdate, 50

	MsgBox, 3, % l(lUpdateTitle, strAppName)
		, % l(lUpdatePrompt, strAppName
			, strCurrentVersion . (strCurrentBranch = "beta" ? " BETA" : "")
			, strLatestVersion . (strCurrentBranch = "beta" ? " BETA" : ""))
		, 30
	IfMsgBox, Yes
		if (strCurrentBranch = "prod")
			Run, http://code.jeanlalonde.ca/folderspopup/
		else
			Run, http://code.jeanlalonde.ca/beta-testers-wanted-for-folders-popup-v3-2/
	IfMsgBox, No
		IniWrite, %strLatestVersion%, %strIniFile%, Global, % "LatestVersionSkipped" . (strCurrentBranch = "beta" ? "Beta" : "")
	IfMsgBox, Cancel ; Remind me
		IniWrite, 0.0, %strIniFile%, Global, % "LatestVersionSkipped" . (strCurrentBranch = "beta" ? "Beta" : "")
	IfMsgBox, TIMEOUT ; Remind me
		IniWrite, 0.0, %strIniFile%, Global, % "LatestVersionSkipped" . (strCurrentBranch = "beta" ? "Beta" : "")
}
else if (A_ThisMenuItem = lMenuUpdate)
{
	MsgBox, 4, % l(lUpdateTitle, strAppName), % l(lUpdateYouHaveLatest, strAppVersion, strAppName), 30
	IfMsgBox, Yes
		if (strCurrentBranch = "prod")
			Run, http://code.jeanlalonde.ca/folderspopup/
		else
			Run, http://code.jeanlalonde.ca/folders-popup-compatible-with-total-commander-beta-testers-wanted/
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
ChangeButtonNamesUpdate:
;------------------------------------------------------------

IfWinNotExist, % l(lUpdateTitle, strAppName)
    return  ; Keep waiting.
SetTimer, ChangeButtonNamesUpdate, Off
WinActivate 
ControlSetText, Button3, %lUpdateButtonRemind%

return
;------------------------------------------------------------


;------------------------------------------------------------
Time2Donate(intStartups, blnDonor)
;------------------------------------------------------------
{
	if (intStartups > 100)
		intDivisor := 25
	else if (intStartups > 60)
		intDivisor := 20
	else if (intStartups > 30)
		intDivisor := 15
	else intDivisor := 10
		
	return !Mod(intStartups, intDivisor) and !(blnDonor)
}
;------------------------------------------------------------



;========================================================================================================================
; MAIN GUI
;========================================================================================================================


;------------------------------------------------------------
BuildGui:
;------------------------------------------------------------

Gui, 1:New, , % L(lGuiTitle, strAppName, strAppVersion)
Gui, Margin, 10, 10
Gui, 1:Color, %strGuiWindowColor%

intWidth := 460

Gui, 1:Font, s12 w700 c%strTextColor%, Verdana
Gui, 1:Add, Text, xm ym h25, %strAppName% %strAppVersion%
Gui, 1:Font, s9 w400, Verdana
Gui, 1:Add, Text, xm y+1 w517, %lAppTagline%


Gui, 1:Add, Picture, x+10 ym Section gGuiOptions, %strTempDir%\settings-32.png ; Static3
Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs-10 y+0 w52 center gGuiOptions, %lGuiOptions% ; Static4

Gui, 1:Font, s8 w400, Verdana
Gui, 1:Add, Text, xm+30, %lGuiSubmenuDropdownLabel%
Gui, 1:Add, Picture, xm y+5 gGuiGotoPreviousMenu vpicPreviousMenu hidden, %strTempDir%\left-12.png ; Static6
Gui, 1:Add, Picture, xm+15 yp gGuiGotoUpMenu vpicUpMenu hidden, %strTempDir%\up-12.png ; Static7
Gui, 1:Add, DropDownList, xm+30 yp w480 vdrpMenusList gGuiMenusListChanged

Gui, 1:Font, s8 w400, Verdana
Gui, 1:Add, ListView
	, xm+30 w480 h240 Count32 -Multi NoSortHdr LV0x10 c%strGuiListviewTextColor% Background%strGuiListviewBackgroundColor% vlvFavoritesList gGuiFavoritesListEvent
	, %lGuiLvFavoritesHeader%|Hidden Menu|Hidden Submenu|Hidden FavoriteType
Loop, 3
	LV_ModifyCol(A_Index + 2, 0) ; hide 3rd-5th columns

Gui, 1:Add, Text, Section x+0 yp

Gui, 1:Add, Picture, xm ys+25 gGuiMoveFavoriteUp, %strTempDir%\up_circular-26.png ; 9
Gui, 1:Add, Picture, xm ys+55 gGuiMoveFavoriteDown, %strTempDir%\down_circular-26.png ; Static10
Gui, 1:Add, Picture, xm ys+85 gGuiAddSeparator, %strTempDir%\separator-26.png ; Static11
Gui, 1:Add, Picture, xm+1 ys+205 gGuiSortFavorites, %strTempDir%\generic_sorting2-26-grey.png ; Static12

Gui, 1:Add, Text, Section xs ys

Gui, 1:Font, s8 w400, Arial ; button legend

Gui, 1:Add, Picture, xs+10 ys+7 gGuiAddFavorite, %strTempDir%\add_property-48.png ; Static14
Gui, 1:Add, Text, xs y+0 w68 center gGuiAddFavorite, %lGuiAddFavorite% ; Static15

Gui, 1:Add, Picture, xs+10 gGuiEditFavorite, %strTempDir%\edit_property-48.png ; Static16
Gui, 1:Add, Text, xs y+0 w68 center gGuiEditFavorite, %lGuiEditFavorite% ; Static17

Gui, 1:Add, Picture, xs+10 gGuiRemoveFavorite, %strTempDir%\delete_property-48.png ; Static18
Gui, 1:Add, Text, xs y+0 w68 center gGuiRemoveFavorite, %lGuiRemoveFavorite% ; Static19

Gui, 1:Add, Picture, xs+10 y+35 gGuiManageGroups, %strTempDir%\channel_mosaic-48.png ; Static20
Gui, 1:Add, Text, xs y+5 w68 center gGuiManageGroups, %lDialogGroups% ; Static21

Gui, 1:Add, Text, Section x185 ys+250

Gui, 1:Font, s8 w600 italic, Verdana
Gui, 1:Add, Text, xm yp w520 center, %lGuiDropFilesIncentive%

Gui, 1:Add, Text, xm y+60
Gui, 1:Font, s9 w600, Verdana
Gui, 1:Add, Button, ys+70 Disabled Default vbtnGuiSave gGuiSave, %lGuiSave% ; Button1
Gui, 1:Add, Button, yp vbtnGuiCancel gGuiCancel, %lGuiClose% ; Close until changes occur - Button2
Gui, 1:Font, s6 w400, Verdana
GuiCenterButtons(L(lGuiTitle, strAppName, strAppVersion), 50, 30, 40, -80, "btnGuiSave", "btnGuiCancel")

Gui, 1:Add, Picture, x490 yp+13 gGuiAbout Section, %strTempDir%\about-32.png ; Static25
Gui, 1:Add, Picture, x540 yp gGuiHelp, %strTempDir%\help-32.png ; Static26
if !(blnDonor)
{
	strDonateButtons := "thumbs_up|solutions|handshake|conference|gift"
	StringSplit, arrDonateButtons, strDonateButtons, |
	Random, intDonateButton, 1, 5

	Gui, 1:Add, Picture, xm+20 yp gGuiDonate, % strTempDir . "\" . arrDonateButtons%intDonateButton% . "-32.png" ; Static27
}

Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, xs-20 ys+35 w68 center gGuiAbout, %lGuiAbout% ; Static28
Gui, 1:Add, Text, xs+40 ys+35 w52 center gGuiHelp, %lGuiHelp% ; Static29
if !(blnDonor)
	Gui, 1:Add, Text, xm+10 ys+35 w52 center gGuiDonate, %lGuiDonate% ; Static30

return
;------------------------------------------------------------


;------------------------------------------------------------
LoadSettingsToGui:
;------------------------------------------------------------

GuiControlGet, blnSaveEnabled, Enabled, %lGuiSave%
if (blnSaveEnabled)
	return

Gosub, LoadOneMenuToGui

GuiControl, Focus, lvFavoritesList

return
;------------------------------------------------------------


;------------------------------------------------------------
LoadOneMenuToGui:
;------------------------------------------------------------

Gui, 1:ListView, lvFavoritesList
LV_Delete()

Loop, % arrMenus[strCurrentMenu].MaxIndex()
	if (arrMenus[strCurrentMenu][A_Index].FavoriteType = "S") ; this is a submenu
		LV_Add(, arrMenus[strCurrentMenu][A_Index].FavoriteName, lGuiSubmenuLocation, arrMenus[strCurrentMenu][A_Index].MenuName
			, arrMenus[strCurrentMenu][A_Index].SubmenuFullName, arrMenus[strCurrentMenu][A_Index].FavoriteType)
	else ; this is a folder or document
		LV_Add(, arrMenus[strCurrentMenu][A_Index].FavoriteName, arrMenus[strCurrentMenu][A_Index].FavoriteLocation, arrMenus[strCurrentMenu][A_Index].MenuName
			, "", arrMenus[strCurrentMenu][A_Index].FavoriteType)
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
Gui, 1:ListView, lvFavoritesList

Loop % LV_GetCount()
{
	LV_GetText(strName, A_Index, 1)
	LV_GetText(strLocation, A_Index, 2)
	LV_GetText(strMenu, A_Index, 3)
	LV_GetText(strSubmenu, A_Index, 4)
	LV_GetText(strFavoriteType, A_Index, 5)

	objFavorite := Object() ; new menu item
	objFavorite.FavoriteName := strName ; display name of this menu item
	objFavorite.FavoriteLocation := strLocation ; path for this menu item
	objFavorite.MenuName := strCurrentMenu ; parent menu of this menu item - do not use strMenu because lack of lMainMenuName
	objFavorite.SubmenuFullName := strSubmenu ; if menu, full name of the submenu
	objFavorite.FavoriteType := strFavoriteType ; "F" folder, "D" document or "S" submenu
	arrMenus[strCurrentMenu].Insert(objFavorite) ; add this menu item to parent menu - do not use strMenu because lack of lMainMenuName
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
		if (arrMenus[strMenu][A_Index].FavoriteType = "S") ; this is a submenu
			if (arrMenus[strMenu][A_Index].SubmenuFullName <> strSkipSubmenu) ; skip if under edited submenu
				strList := strList . "|" . BuildMenuTreeDropDown(arrMenus[strMenu][A_Index].SubmenuFullName, strDefaultMenu, strSkipSubmenu) ; recursive call
	return strList
}
;------------------------------------------------------------


;------------------------------------------------------------
GuiFavoritesListEvent:
;------------------------------------------------------------
Gui, 1:ListView, lvFavoritesList

if (A_GuiEvent = "DoubleClick")
	gosub, GuiEditFavorite

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMenusListChanged:
GuiGotoUpMenu:
GuiGotoPreviousMenu:
OpenMenuFromEditForm:
;------------------------------------------------------------

if (A_ThisLabel = "GuiMenusListChanged")
{
	GuiControlGet, strNewDropdownMenu, , drpMenusList
	if (strNewDropdownMenu = strCurrentMenu) ; user selected the current menu in the dropdown
		return
}

Gosub, SaveCurrentListviewToMenuObject ; save current LV before changing strCurrentMenu

strSavedMenu := strCurrentMenu
if (A_ThisLabel = "GuiMenusListChanged")
{
	arrSubmenuStack.Insert(1, strSavedMenu) ; push the current menu to the left arrow stack
	strCurrentMenu := strNewDropdownMenu
}
else if (A_ThisLabel = "GuiGotoUpMenu")
{
	arrSubmenuStack.Insert(1, strSavedMenu) ; push the current menu to the left arrow stack
	strCurrentMenu := SubStr(strCurrentMenu, 1, InStr(strCurrentMenu, lGuiSubmenuSeparator, , 0) - 1) 
}
else if (A_ThisLabel = "GuiGotoPreviousMenu")
{
	strCurrentMenu := arrSubmenuStack[1] ; pull the top menu from the left arrow stack
	arrSubmenuStack.Remove(1) ; remove the top menu from the left arrow stack
}
else if (A_ThisLabel = "OpenMenuFromEditForm")
{
	arrSubmenuStack.Insert(1, strSavedMenu) ; push the current menu to the left arrow stack
	strCurrentMenu := strCurrentSubmenuFullName
}
else
	return ; should not occur

strPreviousMenu := strSavedMenu

GuiControl, % (arrSubmenuStack.MaxIndex() ? "Show" : "Hide"), picPreviousMenu
GuiControl, % (strCurrentMenu <> lMainMenuName ? "Show" : "Hide"), picUpMenu

Gosub, LoadOneMenuToGui

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiManageGroups:
;------------------------------------------------------------

intWidth := 350

Gosub, BuildGroupMenu ; refresh explorers object and intExplorersExist counter

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:+OwnDialogs
Gui, 2:Color, %strGuiWindowColor%

Gui, 2:Font, w600 
Gui, 2:Add, Text, x10 y10, About Groups
Gui, 2:Font

Gui, 2:Add, Text, x10 y+10 w%intWidth%, %lDialogGroupManageIntro%

Gui, 2:Font, w600 
Gui, 2:Add, Text, x10 y+20, %lDialogGroupManageCreatingTitle%
Gui, 2:Font 

strUseMenuSave := lMenuGroup . " > " . lMenuGroupSave
Gui, 2:Add, Text, x10 y+10 w%intWidth%, % L(lDialogGroupManageCreatingPrompt, lDialogGroupNew, strUseMenuSave)
Gui, 2:Add, Button, x10 y+10 vbtnGroupManageNew gGuiGroupManageNew, %lDialogGroupNew%
GuiControl, % (!intExplorersExist ? "Disable" : "Enable") ; disable Save group menu if no Explorer
	, btnGroupManageNew
GuiCenterButtons(L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion), , , , , "btnGroupManageNew")
if !(intExplorersExist)
	Gui, 2:Add, Text, x10 y+10 w%intWidth%, %lDialogGroupManageCannotSave%

Gui, 2:Font, w600 
Gui, 2:Add, Text, x10 y+20, %lDialogGroupManageManagingTitle%
Gui, 2:Font

Gui, 2:Add, DropDownList, x10 y+10 w%intWidth% vdrpGroupsList, %lDialogGroupSelect%||%strGroups%

Gui, 2:Add, Button, x10 y+10 vbtnGroupManageLoad  gGuiGroupManageLoad, %lDialogGroupLoad%
Gui, 2:Add, Button, x10 yp vbtnGroupManageDelete gGuiGroupManageDelete, %lDialogGroupDelete%
GuiCenterButtons(L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion), , , , , "btnGroupManageLoad", "btnGroupManageDelete")

Gui, 2:Add, Button, x+10 y+30 vbtnGroupManageClose g2GuiClose, %lGui2Close%
GuiCenterButtons(L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion), , , , , "btnGroupManageClose")
Gui, 2:Add, Text, x10, %A_Space%

Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupManageDelete:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

if !StrLen(drpGroupsList) or (drpGroupsList = lDialogGroupSelect)
{
	Oops(lDialogGroupSelectError, lDialogGroupDeleteError)
	return
}

MsgBox, 52, % L(lDialogGroupDeleteTitle, strAppName), % L(lDialogGroupDeletePrompt, drpGroupsList)
IfMsgBox, No
	return

strGroups := strGroups . "|"
StringReplace, strGroups, strGroups, %drpGroupsList%|
StringTrimRight, strGroups, strGroups, 1
GuiControl, 2:, drpGroupsList, |%lDialogGroupSelect%||%strGroups%

IniDelete, %strIniFile%, Group-%drpGroupsList%
IniWrite, %strGroups%, %strIniFile%, Global, Groups

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupManageLoad:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

if !StrLen(drpGroupsList) or (drpGroupsList = lDialogGroupSelect)
{
	Oops(lDialogGroupSelectError, lDialogGroupLoadError)
	return
}

strSelectedGroup := drpGroupsList

Gosub, 2GuiClose
Gosub, GuiClose
Gosub, GroupLoadFromManage

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupManageNew:
;------------------------------------------------------------

Gosub, 2GuiClose
Gosub, GuiClose
Gosub, GuiGroupSave

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddSeparator:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList
LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1, "Select Focus", lMenuSeparator, lMenuSeparator . lMenuSeparator, strCurrentMenu)
LV_Modify(LV_GetNext(), "Vis")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiSortFavorites:
;------------------------------------------------------------

Gui, 1:ListView, lvFavoritesList
LV_ModifyCol(1, "Sort")

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFavoriteUp:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList
intSelectedRow := LV_GetNext()
if (intSelectedRow = 1)
	return

LV_GetText(strThisName, intSelectedRow, 1)
LV_GetText(strThisLocation, intSelectedRow, 2)
LV_GetText(strThisMenu, intSelectedRow, 3)
LV_GetText(strThisSubmenu, intSelectedRow, 4)
LV_GetText(strThisFavoriteType, intSelectedRow, 5)

LV_GetText(strPriorName, intSelectedRow - 1, 1)
LV_GetText(strPriorLocation, intSelectedRow - 1, 2)
LV_GetText(strPriorMenu, intSelectedRow - 1, 3)
LV_GetText(strPriorSubmenu, intSelectedRow - 1, 4)
LV_GetText(strPriorFavoriteType, intSelectedRow - 1, 5)

LV_Modify(intSelectedRow, "", strPriorName, strPriorLocation, strPriorMenu, strPriorSubmenu, strPriorFavoriteType)
LV_Modify(intSelectedRow - 1, "Select Focus Vis", strThisName, strThisLocation, strThisMenu, strThisSubmenu, strThisFavoriteType)

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFavoriteDown:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList
intSelectedRow := LV_GetNext()
if (intSelectedRow = LV_GetCount())
	return

LV_GetText(strThisName, intSelectedRow, 1)
LV_GetText(strThisLocation, intSelectedRow, 2)
LV_GetText(strThisMenu, intSelectedRow, 3)
LV_GetText(strThisSubmenu, intSelectedRow, 4)
LV_GetText(strThisFavoriteType, intSelectedRow, 5)

LV_GetText(strNextName, intSelectedRow + 1, 1)
LV_GetText(strNextLocation, intSelectedRow + 1, 2)
LV_GetText(strNextMenu, intSelectedRow + 1, 3)
LV_GetText(strNextSubmenu, intSelectedRow + 1, 4)
LV_GetText(strNextFavoriteType, intSelectedRow + 1, 5)
	
LV_Modify(intSelectedRow, "", strNextName, strNextLocation, strNextMenu, strNextSubmenu, strNextFavoriteType)
LV_Modify(intSelectedRow + 1, "Select Focus Vis", strThisName, strThisLocation, strThisMenu, strThisSubmenu, strThisFavoriteType)

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiSave:
;------------------------------------------------------------

blnMenuReady := false

Gosub, SaveCurrentListviewToMenuObject ; save current LV before saving

IniDelete, %strIniFile%, Folders ; keep "Folders" label instead of "Favorite" for backward compatibility
Gui, 1:ListView, lvFavoritesList

intIndex := 0
SaveOneMenu(lMainMenuName) ; recursive function

Gosub, LoadIniFile
Gosub, BuildFavoritesMenusWithStatus
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
		strValue := arrMenus[strMenu][A_Index].FavoriteName
			. "|" . arrMenus[strMenu][A_Index].FavoriteLocation
			. "|" . SubStr(arrMenus[strMenu][A_Index].MenuName, StrLen(lMainMenuName) + 1) ; remove main menu name for ini file
			. "|" . SubStr(arrMenus[strMenu][A_Index].SubmenuFullName, StrLen(lMainMenuName) + 1) ; remove main menu name for ini file
			. "|" . arrMenus[strMenu][A_Index].FavoriteType
		intIndex := intIndex + 1
		IniWrite, %strValue%, %strIniFile%, Folders, Folder%intIndex% ; keep "Folders" label instead of "Favorite" for backward compatibility
		if (arrMenus[strMenu][A_Index].FavoriteType = "S")
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
Gui, 1:Show, AutoSize Center

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
		Gosub, BuildGroupMenu
		Gosub, BuildFavoritesMenus ; need to be initialized here - will be updated at each call to popup menu

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
GuiDropFiles:
;------------------------------------------------------------

Loop, parse, A_GuiEvent, `n
{
    strCurrentLocation = %A_LoopField%
    Break
}

Gosub, GuiAddFromDropFiles

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
	for intIndex, objFavorite in arrMenu
	{
		objFavoriteBK := Object()
		objFavoriteBK.MenuName := objFavorite.MenuName
		objFavoriteBK.FavoriteName := objFavorite.FavoriteName
		objFavoriteBK.FavoriteLocation := objFavorite.FavoriteLocation
		objFavoriteBK.SubmenuFullName := objFavorite.SubmenuFullName
		objFavoriteBK.FavoriteType := objFavorite.FavoriteType
		arrMenuBK.Insert(objFavoriteBK)
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
	for intIndex, objFavoriteBK in arrMenuBK
	{
		objFavorite := Object()
		objFavorite.MenuName := objFavoriteBK.MenuName
		objFavorite.FavoriteName := objFavoriteBK.FavoriteName
		objFavorite.FavoriteLocation := objFavoriteBK.FavoriteLocation
		objFavorite.SubmenuFullName := objFavoriteBK.SubmenuFullName
		objFavorite.FavoriteType := objFavoriteBK.FavoriteType
		arrMenu.Insert(objFavorite)
	}
	arrMenus.Insert(strMenuName, arrMenu) ; add this submenu to the array of menus
}

return
;------------------------------------------------------------



;========================================================================================================================
; ADD, EDIT, REMOVE FAVORITES
;========================================================================================================================


;------------------------------------------------------------
AddThisFolder:
;------------------------------------------------------------

strCurrentLocation := ""

if WindowIsDirectoryOpus(strTargetClass)
{
	objDOpusListers := Object()
	CollectDOpusListersList(objDOpusListers, strListText) ; list all listers, excluding special folders like Recycle Bin
	/*
		From leo @ GPSoftware (http://resource.dopus.com/viewtopic.php?f=3&t=23013):
		Lines will have active_lister="1" if they represent tabs from the active lister.
		To get the active tab you want the line with active_lister="1" and tab_state="1".
		tab_state="1" means it's the selected tab, on the active side of the lister.
		tab_state="2" means it's the selected tab, on the inactive side of a dual-display lister.
		Tabs which are not visible (because another tab is selected on top of them) don't get a tab_state attribute at all.
	*/

	for intIndex, objLister in objDOpusListers
		if (objLister.active_lister = "1" and objLister.tab_state = "1")
		{
			strCurrentLocation := objLister.path
			break
		}
}
else
{
	objPrevClipboard := ClipboardAll ; Save the entire clipboard
	ClipBoard := ""

	; Add This folder menu is active only if we are in Explorer (WIN_XP, WIN_7 or WIN_8) or in a Dialog box (WIN_7 or WIN_8).
	; In all these OS, with Explorer, the key sequence {F4}{Esc} selects the current location of the window.
	; With dialog boxes, the key sequence {F4}{Esc} generally selects the current location of the window. But, in some
	; dialog boxes, the {Esc} key closes the dialog box. We will check window title to detect this behavior.

	if (strTargetClass = "#32770")
		intWaitTimeIncrement := 300 ; time allowed for dialog boxes
	else
		intWaitTimeIncrement := 150 ; time allowed for Explorer

	if (blnDiagMode)
		intTries := 8
	else
		intTries := 3

	strWindowTitle := ""
	Loop, %intTries%
	{
		Sleep, intWaitTimeIncrement * A_Index
		WinGetTitle, strWindowTitle, A ; to check later if this window is closed unexpectedly
	} Until (StrLen(strWindowTitle))

	if WindowIsTotalCommander(strTargetClass)
	{
		cm_CopySrcPathToClip := 2029
		SendMessage, 0x433, %cm_CopySrcPathToClip%, , , ahk_class TTOTAL_CMD ; 
		WinGetTitle, strWindowThisTitle, A ; to check if the window was closed unexpectedly
	}
	else if WindowIsFreeCommander(strTargetClass)
	{
		if (WinExist("A") <> strTargetWinId) ; in case that some window just popped out, and initialy active window lost focus
		{
			WinActivate, ahk_id %strTargetWinId% ; we'll activate initialy active window
			Sleep, intWaitTimeIncrement
		}
		if StrLen(strTargetControl)
			ControlFocus, %strTargetControl%, ahk_id %strTargetWinId%
		Loop, %intTries%
		{
			Sleep, intWaitTimeIncrement * A_Index
			SendInput, !g ; goto shortcut for FreeCommander
			Sleep, intWaitTimeIncrement * A_Index
			SendInput, ^c
			Sleep, intWaitTimeIncrement * A_Index
			intTries := A_Index
			WinGetTitle, strWindowThisTitle, A ; to check if the window was closed unexpectedly
		} Until (StrLen(ClipBoard))
	}
	else ; Explorer
		Loop, %intTries%
		{
			Sleep, intWaitTimeIncrement * A_Index
			SendInput, {F4}{Esc} ; F4 move the caret the "Go To A Different Folder box" and {Esc} select it content ({Esc} could be replaced by ^a to Select All)
			Sleep, intWaitTimeIncrement * A_Index
			SendInput, ^c ; Copy
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
}

If !StrLen(strCurrentLocation) or !(InStr(strCurrentLocation, ":") or InStr(strCurrentLocation, "\\")) or (strWindowTitle <> strWindowThisTitle)
{
	Gui, 1:+OwnDialogs 
	MsgBox, 52, % L(lDialogAddFolderManuallyTitle, strAppName, strAppVersion), %lDialogAddFolderManuallyPrompt%
	IfMsgBox, Yes
	{
		Gosub, GuiShow
		Gosub, GuiAddFavorite
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
GuiAddFavorite:
GuiAddFromPopup:
GuiAddFromDropFiles:
GuiEditFavorite:
;------------------------------------------------------------

if (A_ThisLabel = "GuiEditFavorite")
{
	Gui, 1:ListView, lvFavoritesList
	intRowToEdit := LV_GetNext()
	LV_GetText(strCurrentName, intRowToEdit, 1)
	
	if (strCurrentName = lMenuSeparator)
		return

	if !StrLen(strCurrentName) or (intRowToEdit = 0)
	{
		Oops(lDialogSelectItemToEdit)
		return
	}
	LV_GetText(strCurrentLocation, intRowToEdit, 2)
	LV_GetText(strCurrentSubmenuFullName, intRowToEdit, 4)
	LV_GetText(strFavoriteType, intRowToEdit, 5)
	
	blnRadioSubmenu := (strFavoriteType = "S")
	blnRadioFolder := (strFavoriteType = "F")
	blnRadioFile := (strFavoriteType = "D")
	blnRadioURL := (strFavoriteType = "U")
}
else
{
	intRowToEdit := 0 ;  used when saving to flag to insert a new row
	strCurrentName := "" ; make sure it is empty
	strCurrentSubmenuFullName := "" ;  make sure it is empty
	strFavoriteType := "" ;  make sure it is empty
	
	if (A_ThisLabel = "GuiAddFromPopup" or A_ThisLabel = "GuiAddFromDropFiles")
		; strCurrentLocation is received from AddThisFolder or GuiDropFiles
		strFavoriteShortName := GetDeepestFolderName(strCurrentLocation)
	else
	{
		;  make sure these variables are empty
		strCurrentLocation := ""
		strFavoriteShortName := ""
	}

	blnRadioSubmenu := false
	blnRadioFolder := false
	blnRadioFile := false
	blnRadioURL := false
	
	if (A_ThisLabel = "GuiAddFromPopup")
		blnRadioFolder := true
	else if (A_ThisLabel = "GuiAddFromDropFiles")
	{
		blnRadioFile := LocationIsDocument(strCurrentLocation)
		blnRadioFolder := not blnRadioFile
	}
}

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lDialogAddEditFavoriteTitle, (A_ThisLabel = "GuiEditFavorite" ? lDialogEdit : lDialogAdd), strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:+OwnDialogs
Gui, 2:Color, %strGuiWindowColor%

Gui, 2:Add, Text, % x10 y10 vlblFavoriteParentMenu, % (blnRadioSubmenu ? lDialogSubmenuParentMenu : lDialogFavoriteParentMenu)
Gui, 2:Add, DropDownList, x10 w300 vdrpParentMenu, % BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu, strCurrentSubmenuFullName) . "|"

if (A_ThisLabel = "GuiAddFavorite")
{
	Gui, 2:Add, Text, x10, %lDialogAdd%:
	Gui, 2:Add, Radio, x+10 yp vblnRadioFolder checked gRadioButtonsChanged, %lDialogFolderLabel%
	Gui, 2:Add, Radio, x+10 yp vblnRadioFile gRadioButtonsChanged, %lDialogFileLabel%
	Gui, 2:Add, Radio, x+10 yp vblnRadioURL gRadioButtonsChanged, %lDialogURLLabel%
	Gui, 2:Add, Radio, x+10 yp vblnRadioSubmenu gRadioButtonsChanged, %lDialogSubmenuLabel%
}

Gui, 2:Add, Text, x10 y+10 w300 vlblShortName, % (blnRadioSubmenu ? lDialogSubmenuShortName : (blnRadioFile ? lDialogFileShortName : (blnRadioURL ? lDialogURLShortName : lDialogFolderShortName)))
Gui, 2:Add, Edit, x10 w300 Limit250 vstrFolderShortName, % (A_ThisLabel = "GuiEditFavorite" ? strCurrentName : strFavoriteShortName)

if (blnRadioSubmenu)
	Gui, 2:Add, Button, x+10 yp vbnlEditFolderOpenMenu gGuiOpenThisMenu, %lDialogOpenThisMenu%
else
{
	Gui, 2:Add, Text, x10 w300 vlblFolder, % (blnRadioFile ? lDialogFileLabel : (blnRadioURL ? lDialogURLLabel : lDialogFolderLabel))
	Gui, 2:Add, Edit, x10 w300 h20 vstrFolderLocation, %strCurrentLocation%
	if !(blnRadioURL)
		Gui, 2:Add, Button, x+10 yp vbtnSelectFolderLocation gButtonSelectFolderLocation default, %lDialogBrowseButton%
}

if (A_ThisLabel = "GuiEditFavorite")
{
	Gui, 2:Add, Button, y+20 vbtnEditFolderSave gGuiEditFavoriteSave, %lDialogSave%
	Gui, 2:Add, Button, yp vbtnEditFolderCancel gGuiEditFavoriteCancel, %lGuiCancel%
	GuiCenterButtons(L(lDialogAddEditFavoriteTitle, lDialogEdit, strAppName, strAppVersion), 10, 5, 20, , "btnEditFolderSave", "btnEditFolderCancel")
}
else
{
	Gui, 2:Add, Button, y+20 vbtnAddFolderAdd gGuiAddFavoriteSave, %lDialogAdd%
	Gui, 2:Add, Button, yp vbtnAddFolderCancel gGuiAddFavoriteCancel, %lGuiCancel%
	GuiCenterButtons(L(lDialogAddEditFavoriteTitle, lDialogAdd, strAppName, strAppVersion), 10, 5, 20, , "btnAddFolderAdd", "btnAddFolderCancel")
}

GuiControl, 2:Focus, strFavoriteShortName
if (A_ThisLabel = "GuiEditFavorite")
	SendInput, ^a
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiOpenThisMenu:
;------------------------------------------------------------
Gosub, 2GuiClose

Gui, 1:Default
GuiControl, 1:Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList

Gosub, OpenMenuFromEditForm

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFavoriteSave:
GuiEditFavoriteSave:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

GuiControlGet, strParentMenu, , drpParentMenu

if !StrLen(strFavoriteShortName)
{
	Oops(blnRadioSubmenu ? lDialogSubmenuNameEmpty : lDialogFavoriteNameEmpty)
	return
}

if  ((blnRadioFolder or blnRadioFile or blnRadioURL) and !StrLen(strFolderLocation))
{
	Oops(lDialogFavoriteLocationEmpty)
	return
}

if !FolderNameIsNew(strFavoriteShortName, (strParentMenu = strCurrentMenu ? "" : strParentMenu))
	if ((strParentMenu <> strCurrentMenu) or (strFavoriteShortName <> strCurrentName)) ; same name is OK in current menu
	{
		Oops(lDialogFavoriteNameNotNew, strFavoriteShortName)
		return
	}

if (blnRadioSubmenu)
{
	if InStr(strFavoriteShortName, lGuiSubmenuSeparator)
	{
		Oops(L(lDialogFavoriteNameNoSeparator, lGuiSubmenuSeparator))
		return
	}
	
	strNewSubmenuFullName := strParentMenu . lGuiSubmenuSeparator . strFavoriteShortName
	
	if (A_ThisLabel = "GuiAddFavoriteSave")
	{
		strFolderLocation := lGuiSubmenuLocation
		
		arrNewMenu := Object() ; array of folders of the new menu
		arrMenus.Insert(strNewSubmenuFullName, arrNewMenu)
	}
	else ; GuiEditFavoriteSave
		UpdateMenuNameInSubmenus(strCurrentSubmenuFullName, strNewSubmenuFullName) ; change names in arrMenus and arrMenu objects
}
else
	strNewSubmenuFullName := ""

strFavoriteType := (blnRadioSubmenu ? "S" : (blnRadioURL ? "U" : (blnRadioFile ? "D" : (blnRadioURL ? "U" : "F")))) ; submenu, document, URL or folder

Gosub, 2GuiClose

Gui, 1:Default
GuiControl, 1:Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList

if (strParentMenu = strCurrentMenu) ; add as top row to current menu
{
	if (A_ThisLabel = "GuiAddFavoriteSave")
	{
		LV_Insert(LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1, "Select Focus"
			, strFavoriteShortName, strFolderLocation, strCurrentMenu, strNewSubmenuFullName, strFavoriteType)
		LV_Modify(LV_GetNext(), "Vis")
	
		Gosub, SaveCurrentListviewToMenuObject ; save current LV tbefore update the dropdown menu
		GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
	}
	else ; GuiEditFavoriteSave
	{
		LV_Modify(intRowToEdit, "Select Focus"
			, strFavoriteShortName, strFolderLocation, strCurrentMenu, strNewSubmenuFullName, strFavoriteType)
	}
}
else ; add menu item to selected menu object
{
	objFavorite := Object() ; new menu item
	objFavorite.MenuName := strParentMenu ; parent menu of this menu item
	objFavorite.FavoriteName := strFavoriteShortName ; display name of this menu item
	objFavorite.FavoriteLocation := strFolderLocation ; path for this menu item
	objFavorite.SubmenuFullName := strNewSubmenuFullName ; full name of the submenu
	objFavorite.FavoriteType := strFavoriteType ; "F" folder, "D" document, "U" for URL or "S" submenu
	arrMenus[objFavorite.MenuName].Insert(1, objFavorite) ; add this menu item to the new menu

	if (A_ThisLabel = "GuiEditFavoriteSave")
	{
		LV_Delete(intRowToEdit)
		LV_Modify(intRowToEdit, "Select Focus")
	}
}

GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"

LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

if (A_ThisLabel = "GuiEditFavoriteSave")
{
	Gosub, SaveCurrentListviewToMenuObject ; save current LV tbefore update the dropdown menu
	GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
}

Gosub, BuildFavoritesMenusWithStatus ; update menus

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
		if (arrMenus[strNewMenu][A_Index].FavoriteType = "S")
		{
			arrMenus[strNewMenu][A_Index].SubmenuFullName := strNewMenu . lGuiSubmenuSeparator . arrMenus[strNewMenu][A_Index].FavoriteName
			UpdateMenuNameInSubmenus(strOldMenu . lGuiSubmenuSeparator . arrMenus[strOldMenu][A_Index].FavoriteName
				, arrMenus[strNewMenu][A_Index].SubmenuFullName) ; recursive call
		}
	}
	arrMenus.Remove(strOldMenu)
}
;------------------------------------------------------------


;------------------------------------------------------------
FolderNameIsNew(strCandidateName, strMenu := "")
;------------------------------------------------------------
{
	if !StrLen(strMenu)
	{
		Gui, 1:Default
		Gui, 1:ListView, lvFavoritesList
		Loop, % LV_GetCount()
		{
			LV_GetText(strThisName, A_Index, 1)
			if (strCandidateName = strThisName)
				return False
		}
	}
	else
		Loop, % arrMenus[strMenu].MaxIndex()
			if (strCandidateName = arrMenus[strMenu][A_Index].FavoriteName)
				return False

	return True
}
;------------------------------------------------------------


;------------------------------------------------------------
RadioButtonsChanged:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

if (blnRadioFolder)
{
	GuiControl, , lblShortName, %lDialogFolderShortName%
	GuiControl, , lblFolder, %lDialogFolderLabel%
}
else if (blnRadioFile)
{
	GuiControl, , lblShortName, %lDialogFileShortName%
	GuiControl, , lblFolder, %lDialogFileLabel%
}
else if (blnRadioURL)
{
	GuiControl, , lblShortName, %lDialogURLShortName%
	GuiControl, , lblFolder, %lDialogURLLabel%
}
else ; blnRadioSubmenu
{
	GuiControl, , lblShortName, %lDialogSubmenuShortName%
	GuiControl, , strFolderLocation ; empty control
}

GuiControl, % (blnRadioSubmenu ? "Hide" : "Show"), lblFolder
GuiControl, % (blnRadioSubmenu ? "Hide" : "Show"), strFolderLocation
GuiControl, % (blnRadioSubmenu or blnRadioURL ? "Hide" : "Show"), btnSelectFolderLocation

if (blnRadioSubmenu or blnRadioURL)
	GuiControl, +Default, btnAddFolderAdd
else
	GuiControl, +Default, btnSelectFolderLocation

GuiControl, Focus, strFavoriteShortName

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonSelectFolderLocation:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

if (blnRadioFolder)
	FileSelectFolder, strNewLocation, *%strCurrentLocation%, 3, %lDialogAddFolderSelect%
else
	FileSelectFile, strNewLocation, S3, %strCurrentLocation%, %lDialogAddFileSelect%

if !(StrLen(strNewLocation))
	return

GuiControl, 2:, strFolderLocation, %strNewLocation%

if !StrLen(strFavoriteShortName)
	GuiControl, 2:, strFavoriteShortName, % GetDeepestFolderName(strNewLocation)

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFavoriteCancel:
GuiEditFavoriteCancel:
;------------------------------------------------------------
Gosub, 2GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveFavorite:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList
intItemToRemove := LV_GetNext()
if !(intItemToRemove)
{
	Oops(lDialogSelectItemToRemove)
	return
}

LV_GetText(strFavoriteType, intItemToRemove, 5)
if (strFavoriteType = "S") ; this is a submneu
{
	LV_GetText(strSubmenuFullName, intItemToRemove, 4)
	MsgBox, 52, % L(lDialogFavoriteRemoveTitle, strAppName), % L(lDialogFavoriteRemovePrompt, strSubmenuFullName)
	IfMsgBox, No
		return
	RemoveAllSubMenus(strSubmenuFullName)
}

LV_Delete(intItemToRemove)
LV_Modify(intItemToRemove, "Select Focus")
LV_ModifyCol(1, "AutoHdr")
LV_ModifyCol(2, "AutoHdr")

if (strFavoriteType = "S")
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
		if (arrMenus[strSubmenuFullName][A_Index].FavoriteType = "S") ; this is a submenu
			RemoveAllSubMenus(arrMenus[strSubmenuFullName][A_Index].SubmenuFullName) ; recursive call
	arrMenus.Remove(strSubmenuFullName)
}
;------------------------------------------------------------



;========================================================================================================================
; OPTIONS
;========================================================================================================================


;------------------------------------------------------------
GuiOptions:
;------------------------------------------------------------

intGui1WinID := WinExist("A")

;---------------------------------------
; Build Gui header
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lOptionsGuiTitle, strAppName, strAppVersion)
Gui, 2:Color, %strGuiWindowColor%
Gui, 2:+Owner1
Gui, 2:Font, s10 w700, Verdana
Gui, 2:Add, Text, x10 y10 w595 center, % L(lOptionsGuiTitle, strAppName)
Gui, 2:Font
Gui, 2:Add, Text, x10 y+10 w595 center, % L(lOptionsGuiIntro, strAppName)

; Build Hotkey Gui lines
loop, % arrIniKeyNames%0%
{
	Gui, 2:Font, s8 w700
	Gui, 2:Add, Text, Section x15 y+10, % arrOptionsTitles%A_Index%
	Gui, 2:Font, s9 w500, Courier New
	Gui, 2:Add, Text, x260 ys+5 w280 h23 center 0x1000 vlblHotkeyText%A_Index% gButtonOptionsChangeHotkey%A_Index%, % Hotkey2Text(strModifiers%A_Index%, strMouseButton%A_Index%, strOptionsKey%A_Index%)
	Gui, 2:Font
	Gui, 2:Add, Button, yp x555 vbtnChangeHotkey%A_Index% gButtonOptionsChangeHotkey%A_Index%, %lOptionsChangeHotkey%
	Gui, 2:Font, s8 w500
	Gui, 2:Add, Text, x15 ys+15, % arrOptionsTitlesSub%A_Index%
}

Gui, 2:Add, Text, x15 y+15 h2 w595 0x10 ; Horizontal Line > Etched Gray

Gui, 2:Font, s8 w700
Gui, 2:Add, Text, x10 y+5 w595 center, %lOptionsOtherOptions%
Gui, 2:Font

; column 1
Gui, 2:Add, Text, y+10 x15 Section, %lOptionsLanguage%
Gui, 2:Add, DropDownList, ys x+10 w120 vdrpLanguage Sort, %lOptionsLanguageLabels%
GuiControl, ChooseString, drpLanguage, %strLanguageLabel%

Gui, 2:Add, CheckBox, y+10 xs vblnOptionsRunAtStartup, %lOptionsRunAtStartup%
GuiControl, , blnOptionsRunAtStartup, % FileExist(A_Startup . "\" . strAppName . ".lnk") ? 1 : 0

Gui, 2:Add, CheckBox, y+10 xs vblnDisplayMenuShortcuts, %lOptionsDisplayMenuShortcuts%
GuiControl, , blnDisplayMenuShortcuts, %blnDisplayMenuShortcuts%

Gui, 2:Add, CheckBox, y+10 xs vblnDisplayTrayTip, %lOptionsTrayTip%
GuiControl, , blnDisplayTrayTip, %blnDisplayTrayTip%

Gui, 2:Add, CheckBox, y+10 xs vblnDisplayIcons, %lOptionsDisplayIcons%
GuiControl, , blnDisplayIcons, %blnDisplayIcons%
if !OSVersionIsWorkstation()
	GuiControl, Disable, blnDisplayIcons

; Build Gui footer

Gui, 2:Add, Text, x15 y+15 h2 w595 0x10 ; Horizontal Line > Etched Gray

Gui, 2:Font, s8 w700
Gui, 2:Add, Text, y+5 xs, % L(lOptionsThirdPartyTitle, "Directory Opus")
Gui, 2:Font
Gui, 2:Add, Text, y+5 xs, % L(lOptionsThirdPartyDetail, "Directory Opus")
Gui, 2:Add, Text, y+10 xs, %lOptionsThirdPartyPrompt%
Gui, 2:Add, Edit, x+10 yp w300 h20 vstrDirectoryOpusPath, %strDirectoryOpusPath%
Gui, 2:Add, Button, x+10 yp vbtnSelectDOpusPath gButtonSelectDOpusPath, %lDialogBrowseButton%
Gui, 2:Add, Checkbox, x+10 yp vblnDirectoryOpusUseTabs, %lOptionsDirectoryOpusUseTabs%
GuiControl, , blnDirectoryOpusUseTabs, %blnDirectoryOpusUseTabs%

Gui, 2:Font, s8 w700
Gui, 2:Add, Text, y+5 xs, % L(lOptionsThirdPartyTitle, "Total Commander")
Gui, 2:Font
Gui, 2:Add, Text, y+5 xs,  % L(lOptionsThirdPartyDetail, "Total Commander")
Gui, 2:Add, Text, y+10 xs, %lOptionsThirdPartyPrompt%
Gui, 2:Add, Edit, x+10 yp w300 h20 vstrTotalCommanderPath, %strTotalCommanderPath%
Gui, 2:Add, Button, x+10 yp vbtnSelectTCPath gButtonSelectTCPath, %lDialogBrowseButton%
Gui, 2:Add, Checkbox, x+10 yp vblnTotalCommanderUseTabs, %lOptionsTotalCommanderUseTabs%
GuiControl, , blnTotalCommanderUseTabs, %blnTotalCommanderUseTabs%

Gui, 2:Add, Button, y+20 vbtnOptionsSave gButtonOptionsSave, %lGuiSave%
Gui, 2:Add, Button, yp vbtnOptionsCancel gButtonOptionsCancel, %lGuiCancel%
Gui, 2:Add, Button, yp vbtnOptionsDonate gGuiDonate, %lDonateButton%
GuiCenterButtons(L(lOptionsGuiTitle, strAppName, strAppVersion), 10, 5, 20, , "btnOptionsSave", "btnOptionsCancel", "btnOptionsDonate")

Gui, 2:Add, Text
GuiControl, Focus, btnOptionsSave

; column 2
Gui, 2:Add, Text, ys x240 Section, %lOptionsIconSize%
Gui, 2:Add, DropDownList, ys x+10 w40 vdrpIconSize Sort, 16|24|32|64
GuiControl, ChooseString, drpIconSize, %intIconSize%

Gui, 2:Add, CheckBox, y+10 xs vblnDisplaySpecialFolders, %lOptionsDisplaySpecialFolders%
GuiControl, , blnDisplaySpecialFolders, %blnDisplaySpecialFolders%

Gui, 2:Add, CheckBox, y+10 xs vblnDisplayGroupMenu, %lOptionsDisplayGroupMenu%
GuiControl, , blnDisplayGroupMenu, %blnDisplayGroupMenu%

Gui, 2:Add, CheckBox, y+10 xs vblnDisplayRecentFolders gDisplayRecentFoldersClicked, %lOptionsDisplayRecentFolders%
GuiControl, , blnDisplayRecentFolders, %blnDisplayRecentFolders%

Gui, 2:Add, Edit, % "y+10 xs w36 h17 vintRecentFolders center " . (blnDisplayRecentFolders ? "" : "hidden"), %intRecentFolders%
Gui, 2:Add, Text, % "yp x+10 vlblRecentFolders " . (blnDisplayRecentFolders ? "" : "hidden"), %lOptionsRecentFolders%

; column 3

Gui, 2:Add, Text, ys x430 Section, %lOptionsTheme%
Gui, 2:Add, DropDownList, ys x+10 w120 vdrpTheme Sort, %strAvailableThemes%
GuiControl, ChooseString, drpTheme, %strTheme%

Gui, 2:Add, Text, y+7 xs Section, %lOptionsMenuPositionPrompt%

Gui, 2:Add, Radio, % "y+5 xs vradPopupMenuPosition1 gPopupMenuPositionClicked Group " . (intPopupMenuPosition = 1 ? "Checked" : ""), %lOptionsMenuNearMouse%
Gui, 2:Add, Radio, % "y+5 xs vradPopupMenuPosition2 gPopupMenuPositionClicked " . (intPopupMenuPosition = 2 ? "Checked" : ""), %lOptionsMenuActiveWindow%
Gui, 2:Add, Radio, % "y+5 xs vradPopupMenuPosition3 gPopupMenuPositionClicked " . (intPopupMenuPosition = 3 ? "Checked" : ""), %lOptionsMenuFixPosition%

Gui, 2:Add, Text, % "y+5 xs+18 vlblPopupFixPositionX " . (intPopupMenuPosition = 3 ? "" : "hidden"), %lOptionsPopupFixPositionX%
Gui, 2:Add, Edit, % "yp x+5 w36 h17 vstrPopupFixPositionX center " . (intPopupMenuPosition = 3 ? "" : "hidden"), %arrPopupFixPosition1%
Gui, 2:Add, Text, % "yp x+5 vlblPopupFixPositionY " . (intPopupMenuPosition = 3 ? "" : "hidden"), %lOptionsPopupFixPositionY%
Gui, 2:Add, Edit, % "yp x+5 w36 h17 vstrPopupFixPositionY center " . (intPopupMenuPosition = 3 ? "" : "hidden"), %arrPopupFixPosition2%


Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
DisplayRecentFoldersClicked:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

GuiControl, % (blnDisplayRecentFolders ? "Show" : "Hide"), intRecentFolders
GuiControl, % (blnDisplayRecentFolders ? "Show" : "Hide"), lblRecentFolders

return
;------------------------------------------------------------


;------------------------------------------------------------
PopupMenuPositionClicked:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

GuiControl, % (radPopupMenuPosition3 ? "Show" : "Hide"), lblPopupFixPositionX
GuiControl, % (radPopupMenuPosition3 ? "Show" : "Hide"), strPopupFixPositionX
GuiControl, % (radPopupMenuPosition3 ? "Show" : "Hide"), lblPopupFixPositionY
GuiControl, % (radPopupMenuPosition3 ? "Show" : "Hide"), strPopupFixPositionY

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonSelectDOpusPath:
;------------------------------------------------------------

if StrLen(strDirectoryOpusPath) and (strDirectoryOpusPath <> "NO")
	strCurrentDOpusLocation := strDirectoryOpusPath
else
	strCurrentDOpusLocation := A_ProgramFiles . "\GPSoftware\Directory Opus\dopus.exe"

FileSelectFile, strNewDOpusLocation, 3, %strCurrentDOpusLocation%, %lDialogAddFolderSelect%

if !(StrLen(strNewDOpusLocation))
	return

GuiControl, 2:, strDirectoryOpusPath, %strNewDOpusLocation%

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonSelectTCPath:
;------------------------------------------------------------

if StrLen(strTotalCommanderPath) and (strTotalCommanderPath <> "NO")
	strCurrentTCLocation := strTotalCommanderPath
else
	strCurrentTCLocation := GetTotalCommanderPath()

FileSelectFile, strNewTCLocation, 3, %strCurrentTCLocation%, %lDialogAddFolderSelect%

if !(StrLen(strNewTCLocation))
	return

GuiControl, 2:, strTotalCommanderPath, %strNewTCLocation%

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
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%strAppName%.lnk, %A_WorkingDir%
Menu, Tray, % blnOptionsRunAtStartup ? "Check" : "Uncheck", %lMenuRunAtStartup%

IniWrite, %blnDisplayTrayTip%, %strIniFile%, Global, DisplayTrayTip
IniWrite, %blnDisplayIcons%, %strIniFile%, Global, DisplayIcons
IniWrite, %blnDisplaySpecialFolders%, %strIniFile%, Global, DisplaySpecialFolders
IniWrite, %blnDisplayRecentFolders%, %strIniFile%, Global, DisplayRecentFolders
IniWrite, %intRecentFolders%, %strIniFile%, Global, RecentFolders
IniWrite, %blnDisplayGroupMenu%, %strIniFile%, Global, DisplaySwitchMenu ; keep "Switch" for backward compatibility
IniWrite, %blnDisplayMenuShortcuts%, %strIniFile%, Global, DisplayMenuShortcuts

if (radPopupMenuPosition1)
	intPopupMenuPosition := 1
else if (radPopupMenuPosition2)
	intPopupMenuPosition := 2
else
	intPopupMenuPosition := 3
IniWrite, %intPopupMenuPosition%, %strIniFile%, Global, PopupMenuPosition
	
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

strThemePrev := strTheme
strTheme := drpTheme
IniWrite, %strTheme%, %strIniFile%, Global, Theme

intIconSize := drpIconSize
IniWrite, %intIconSize%, %strIniFile%, Global, IconSize

IniWrite, %strDirectoryOpusPath%, %strIniFile%, Global, DirectoryOpusPath
IniWrite, %blnDirectoryOpusUseTabs%, %strIniFile%, Global, DirectoryOpusUseTabs
blnUseDirectoryOpus := StrLen(strDirectoryOpusPath)
if (blnUseDirectoryOpus)
{
	blnUseDirectoryOpus := FileExist(strDirectoryOpusPath)
	if (blnUseDirectoryOpus)
		Gosub, SetDOpusRt
}
if (blnDirectoryOpusUseTabs)
	strDirectoryOpusNewTabOrWindow := "NEWTAB" ; open new folder in a new lister tab
else
	strDirectoryOpusNewTabOrWindow := "NEW" ; open new folder in a new DOpus lister (instance)

IniWrite, %strTotalCommanderPath%, %strIniFile%, Global, TotalCommanderPath
IniWrite, %blnTotalCommanderUseTabs%, %strIniFile%, Global, TotalCommanderUseTabs
blnUseTotalCommander := StrLen(strTotalCommanderPath)
if (blnUseTotalCommander)
{
	blnUseTotalCommander := FileExist(strTotalCommanderPath)
	if (blnUseTotalCommander)
		Gosub, SetTCCommand
}
if (blnTotalCommanderUseTabs)
	strTotalCommanderNewTabOrWindow := "/O /T" ; open new folder in a new tab
else
	strTotalCommanderNewTabOrWindow := "/N" ; open new folder in a new window (TC instance)

; if language or theme changed, offer to restart the app
if (strLanguageCodePrev <> strLanguageCode) or (strThemePrev <> strTheme)
{
	MsgBox, 52, %strAppName%, % L(lReloadPrompt, (strLanguageCodePrev <> strLanguageCode ? lOptionsLanguage : lOptionsTheme), (strLanguageCodePrev <> strLanguageCode ? strLanguageLabel : strTheme), strAppName)
	IfMsgBox, Yes
	{
		Gosub, RestoreBackupMenuObjects
		Reload
	}
}	

; else rebuild special and Group menus
Gosub, BuildSpecialFoldersMenu
Gosub, BuildGroupMenu

; and rebuild Folders menus w/ or w/o optional folders and shortcuts
for strMenuName, arrMenu in arrMenus
{
	Menu, %strMenuName%, Add
	Menu, %strMenuName%, DeleteAll
	arrMenu := ; free object's memory
}
Gosub, BuildFavoritesMenusWithStatus

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
ButtonOptionsChangeHotkey1:
ButtonOptionsChangeHotkey2:
ButtonOptionsChangeHotkey3:
ButtonOptionsChangeHotkey4:
ButtonOptionsChangeHotkey5:
ButtonOptionsChangeHotkey6:
;------------------------------------------------------------

StringReplace, intIndex, A_ThisLabel, ButtonOptionsChangeHotkey
intGui2WinID := WinExist("A")

Gui, 2:Submit, NoHide
Gui, 3:New, , % L(lOptionsChangeHotkeyTitle, strAppName, strAppVersion)
Gui, 3:Color, %strGuiWindowColor%
Gui, 3:+Owner2
Gui, 3:Font, s10 w700, Verdana
Gui, 3:Add, Text, x10 y10 w350 center, % L(lOptionsChangeHotkeyTitle, strAppName)
Gui, 3:Font

intHotkeyType := 3 ; Settings and Recent folders
if InStr(arrIniKeyNames%intIndex%, "Mouse")
	intHotkeyType := 1
if InStr(arrIniKeyNames%intIndex%, "Keyboard")
	intHotkeyType := 2

Gui, 3:Add, Text, y+15 x10 , %lOptionsTriggerFor%
Gui, 3:Font, s8 w700
Gui, 3:Add, Text, x+5 yp , % arrOptionsTitles%intIndex%
Gui, 3:Font

Gui, 3:Add, Text, x10 y+5 w350, % lOptionsArrDescriptions%intIndex%

Gui, 3:Add, CheckBox, y+20 x50 vblnOptionsShift, %lOptionsShift%
GuiControl, , blnOptionsShift, % InStr(strModifiers%intIndex%, "+") ? 1 : 0
GuiControlGet, posTop, Pos, blnOptionsShift
Gui, 3:Add, CheckBox, y+10 x50 vblnOptionsCtrl, %lOptionsCtrl%
GuiControl, , blnOptionsCtrl, % InStr(strModifiers%intIndex%, "^") ? 1 : 0
Gui, 3:Add, CheckBox, y+10 x50 vblnOptionsAlt, %lOptionsAlt%
GuiControl, , blnOptionsAlt, % InStr(strModifiers%intIndex%, "!") ? 1 : 0
Gui, 3:Add, CheckBox, y+10 x50 vblnOptionsWin, %lOptionsWin%
GuiControl, , blnOptionsWin, % InStr(strModifiers%intIndex%, "#") ? 1 : 0

; initialize or we may have an error if another hotkey was changed before
strOptionsMouse := ""
strOptionsKey := ""

if (intHotkeyType = 1)
	Gui, 3:Add, DropDownList, % "y" . posTopY . " x150 w200 vstrOptionsMouse gOptionsMouseChanged", % strMouseButtonsWithDefault%intIndex%
if (intHotkeyType <> 1)
{
	Gui, 3:Add, Text, % "y" . posTopY . " x150 w60", %lOptionsKeyboard%
	Gui, 3:Add, Hotkey, yp x+10 w130 vstrOptionsKey gOptionsHotkeyChanged
	GuiControl, , strOptionsKey, % strOptionsKey%intIndex%
}
if (intHotkeyType = 3)
	Gui, 3:Add, DropDownList, % "y" . posTopY + 50 . " x150 w200 vstrOptionsMouse gOptionsMouseChanged", % strMouseButtonsWithDefault%intIndex%

Gui, 3:Add, Button, x10 y220 vbtnResetHotkey gButtonResetHotkey%intIndex%, %lGuiResetDefault%
GuiCenterButtons(L(lOptionsChangeHotkeyTitle, strAppName, strAppVersion), 10, 5, 20, , "btnResetHotkey")
Gui, 3:Add, Button, y+20 x10 vbtnChangeHotkeySave gButtonChangeHotkeySave%intIndex%, %lGuiSave%
Gui, 3:Add, Button, yp x+20 vbtnChangeHotkeyCancel gButtonChangeHotkeyCancel, %lGuiCancel%
GuiCenterButtons(L(lOptionsChangeHotkeyTitle, strAppName, strAppVersion), 10, 5, 20, , "btnChangeHotkeySave", "btnChangeHotkeyCancel")

Gui, 3:Add, Text
GuiControl, Focus, btnChangeHotkeySave
Gui, 3:Show, AutoSize Center
Gui, 2:+Disabled

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
	GuiControl, Choose, %strOptionsMouseControl%, 0
}

return
;------------------------------------------------------------


;------------------------------------------------------------
OptionsMouseChanged:
;------------------------------------------------------------

strOptionsMouseControl := A_GuiControl ; hotkey var name
GuiControlGet, strOptionsMouseValue, , %strOptionsMouseControl%

if (strOptionsMouseValue = lOptionsMouseNone) ; this is the translated "None"
{
	GuiControl, , blnOptionsShift, 0
	GuiControl, , blnOptionsCtrl, 0
	GuiControl, , blnOptionsAlt, 0
	GuiControl, , blnOptionsWin, 0
}

if (intHotkeyType = 3) ; both keyboard and mouse options are available
{
	StringReplace, strOptionsHotkeyControl, strOptionsMouseControl, Mouse, Key ; get the hotkey var

	; we have a mouse button, empty the hotkey control
	GuiControl, , %strOptionsHotkeyControl%, None
}

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonChangeHotkeySave1:
ButtonChangeHotkeySave2:
ButtonChangeHotkeySave3:
ButtonChangeHotkeySave4:
ButtonChangeHotkeySave5:
ButtonChangeHotkeySave6:
;------------------------------------------------------------
Gui, 3:Submit

StringReplace, intIndex, A_ThisLabel, ButtonChangeHotkeySave

StringSplit, arrIniVarNames, strIniKeyNames, |

if StrLen(strOptionsMouse)
	strOptionsMouse := GetMouseButton4Text(strOptionsMouse) ; get mouse button system name from dropdown localized text
else
	strMouseButton%intIndex% := "" ;  empty mouse button text

strHotkey := Trim(strOptionsKey . strOptionsMouse)

if StrLen(strHotkey)
	if (strHotkey = "None") ; do not compare with lOptionsMouseNone because it is translated
	{
		Hotkey, % "$" . arrHotkeyVarNames%intIndex%, Off, UseErrorLevel ; UseErrorLevel in case we had an invalid key name in the ini file
		IniWrite, None, %strIniFile%, Global, % arrIniVarNames%intIndex% ; do not write lOptionsMouseNone because it is translated
	}
	else
	{
		if (blnOptionsWin)
			strHotkey := "#" . strHotkey
		if (blnOptionsAlt)
			strHotkey := "!" . strHotkey
		if (blnOptionsShift)
			strHotkey := "+" . strHotkey
		if (blnOptionsCtrl)
			strHotkey := "^" . strHotkey

		if (strHotkey = "LButton")
		{
			Oops(lOptionsMouseCheckLButton)
			Gosub, 3GuiClose
			return
		}

		Hotkey, % "$" . arrHotkeyVarNames%intIndex%, Off, UseErrorLevel ; UseErrorLevel in case we had an invalid key name in the ini file
		IniWrite, %strHotkey%, %strIniFile%, Global, % arrIniVarNames%intIndex%
		
	}
else
	Oops(L(lOptionsNoKeyOrMouseSpecified, arrIniVarNames%intIndex%))

Gosub, LoadIniHotkeys ; reload ini variables and reset hotkeys

GuiControl, 2:, lblHotkeyText%intIndex%, % Hotkey2Text(strModifiers%intIndex%, strMouseButton%intIndex%, strOptionsKey%intIndex%)

Gosub, 3GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonResetHotkey1:
ButtonResetHotkey2:
ButtonResetHotkey3:
ButtonResetHotkey4:
ButtonResetHotkey5:
ButtonResetHotkey6:
;------------------------------------------------------------

StringReplace, intIndex, A_ThisLabel, ButtonResetHotkey

StringSplit, arrIniVarNames, strIniKeyNames, |
IniWrite, % arrHotkeyDefaults%intIndex%, %strIniFile%, Global, % arrIniVarNames%intIndex%

MsgBox, 52, %strAppName%, % L(lReloadPromptDefaultHotkey, strAppName)
IfMsgBox, Yes
{
	Gosub, RestoreBackupMenuObjects
	Reload
}

Gosub, 3GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonChangeHotkeyCancel:
;------------------------------------------------------------
Gosub, 3GuiClose

return
;------------------------------------------------------------



;========================================================================================================================
; ABOUT-DONATE-HELP
;========================================================================================================================


;------------------------------------------------------------
GuiAbout:
;------------------------------------------------------------

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide
Gui, 2:New, , % L(lAboutTitle, strAppName, strAppVersion)
Gui, 2:Color, %strGuiWindowColor%
Gui, 2:+Owner1
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Link, y10 w350 vlblAboutText1, % L(lAboutText1, strAppName, strAppVersion, str32or64)
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Link, , % L(lAboutText2, strAppName)
Gui, 2:Add, Link, , % L(lAboutText3, chr(169))
Gui, 2:Font, s10 w400, Verdana
Gui, 2:Add, Link, , % L(lAboutText4)
Gui, 2:Font, s8 w400, Verdana

Gui, 2:Add, Button, y+20 vbtnAboutDonate gGuiDonate, %lDonateButton%
Gui, 2:Add, Button, yp vbtnAboutClose g2GuiClose vbtnAboutClose, %lGui2Close%
GuiCenterButtons(L(lAboutTitle, strAppName, strAppVersion), 10, 5, 20, , "btnAboutDonate", "btnAboutClose")

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
Gui, 2:Color, %strGuiWindowColor%
Gui, 2:+Owner1
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Link, y10 w420, % L(lDonateText1, strAppName)
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Link, x175 w185 y+10, % L(lDonateText2, "http://code.jeanlalonde.ca/support-freeware/")
loop, 2
{
	Gui, 2:Add, Button, % (A_Index = 1 ? "y+10 Default vbtnDonateDefault " : "") . " xm w150 gButtonDonate" . A_Index, % lDonatePlatformName%A_Index%
	Gui, 2:Add, Link, x+10 w235 yp, % lDonatePlatformComment%A_Index%
}

Gui, 2:Font, s10 w700, Verdana
Gui, 2:Add, Link, xm y+20 w420, %lDonateText3%
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Link, xm y+10 w420 Section, % L(lDonateText4, strAppName)

strDonateReviewUrlLeft1 := "http://download.cnet.com/FoldersPopup/3000-2344_4-76062382.html"
strDonateReviewUrlLeft2 := "http://www.portablefreeware.com/index.php?id=2557"
strDonateReviewUrlLeft3 := "http://www.softpedia.com/get/System/OS-Enhancements/FoldersPopup.shtml"
strDonateReviewUrlRight1 := "http://fileforum.betanews.com/detail/Folders-Popup/1385175626/1"
strDonateReviewUrlRight2 := "http://www.filecluster.com/System-Utilities/Other-Utilities/Download-FoldersPopup.html"
strDonateReviewUrlRight3 := "http://freewares-tutos.blogspot.fr/2013/12/folders-popup-un-logiciel-portable-pour.html"

loop, 3
	Gui, 2:Add, Link, % (A_Index = 1 ? "ys+20" : "y+5") . " x25 w150", % "<a href=""" . strDonateReviewUrlLeft%A_Index% . """>" . lDonateReviewNameLeft%A_Index% . "</a>"

loop, 3
	Gui, 2:Add, Link, % (A_Index = 1 ? "ys+20" : "y+5") . " x175 w150", % "<a href=""" . strDonateReviewUrlRight%A_Index% . """>" . lDonateReviewNameRight%A_Index% . "</a>"

Gui, 2:Add, Link, y+10 x130, <a href="http://code.jeanlalonde.ca/support-freeware/">%lDonateText5%</a>

Gui, 2:Font, s8 w400, Verdana
Gui, 2:Add, Button, x175 y+20 g2GuiClose vbtnDonateClose, %lGui2Close%
GuiCenterButtons(L(lDonateTitle, strAppName, strAppVersion), 10, 5, 20, , "btnDonateClose")

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
Gui, 2:Color, %strGuiWindowColor%
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

Gui, 2:Add, Button, x180 y+20 vbtnHelpDonate gGuiDonate, %lDonateButton%
Gui, 2:Add, Button, x+80 yp g2GuiClose vbtnHelpClose, %lGui2Close%
GuiCenterButtons(L(lHelpTitle, strAppName, strAppVersion), 10, 5, 20, , "btnHelpDonate", "btnHelpClose")

GuiControl, Focus, btnHelpClose
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------



;========================================================================================================================
; GUI CLOSE-CANCEL
;========================================================================================================================


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
3GuiClose:
3GuiEscape:
;------------------------------------------------------------

Gui, 2:-Disabled
Gui, 3:Destroy
WinActivate, ahk_id %intGui2WinID%

return
;------------------------------------------------------------



;========================================================================================================================
; THIRD-PARTY
;========================================================================================================================


;------------------------------------------------------------
RunDOpusRt(strCommand, strLocation := "", strParam := "")
; put A_Space at the beginning of strParam if required - some param (like ",paths") must have no space 
;------------------------------------------------------------
{
	global strDirectoryOpusRtPath
	global strDirectoryOpusPath

	if FileExist(strDirectoryOpusRtPath)
		Run, % """" . strDirectoryOpusRtPath . """ " . strCommand . " """ . strLocation . """" . strParam
	else
		Oops(lOopsWrongThirdPartyPath, strAppName, "Directory Opus", "DirectoryOpus", strDirectoryOpusPath)
}
;------------------------------------------------------------


;------------------------------------------------------------
CheckDirectoryOpus:
;------------------------------------------------------------

strCheckDirectoryOpusPath := A_ProgramFiles . "\GPSoftware\Directory Opus\dopus.exe"

if FileExist(strCheckDirectoryOpusPath)
{
	MsgBox, 52, %strAppName%, % L(lDialogThirdPartyDetected, strAppName, "Directory Opus")
	IfMsgBox, No
		strDirectoryOpusPath := "NO"
	else
	{
		strDirectoryOpusPath := strCheckDirectoryOpusPath
		Gosub, SetDOpusRt
	}
	blnUseDirectoryOpus := (strDirectoryOpusPath <> "NO")
	IniWrite, %strDirectoryOpusPath%, %strIniFile%, Global, DirectoryOpusPath
	blnDirectoryOpusUseTabs := 1
	IniWrite, %blnDirectoryOpusUseTabs%, %strIniFile%, Global, DirectoryOpusUseTabs
	; strDirectoryOpusNewTabOrWindow will contain "NEWTAB" to open in a new tab if DirectoryOpusUseTabs is 1 (default) or "NEW" to open in a new lister
	strDirectoryOpusNewTabOrWindow := "NEWTAB"
}

return
;------------------------------------------------------------


;------------------------------------------------------------
SetDOpusRt:
;------------------------------------------------------------

IniRead, blnDirectoryOpusUseTabs, %strIniFile%, Global, DirectoryOpusUseTabs, 1 ; should be intialized here but true by default for safety
if (blnDirectoryOpusUseTabs)
	strDirectoryOpusNewTabOrWindow := "NEWTAB" ; open new folder in a new tab
else
	strDirectoryOpusNewTabOrWindow := "NEW" ; open new folder in a new lister

strDOpusTempFilePath := strTempDir . "\dopus-list.txt"
StringReplace, strDirectoryOpusRtPath, strDirectoryOpusPath, \dopus.exe, \dopusrt.exe

; additional icon for Directory Opus
objIconsFile["DirectoryOpus"] := strDirectoryOpusPath
objIconsIndex["DirectoryOpus"] := 1

return
;------------------------------------------------------------


;------------------------------------------------------------
CheckTotalCommander:
;------------------------------------------------------------

strCheckTotalCommanderPath := GetTotalCommanderPath()

if FileExist(strCheckTotalCommanderPath)
{
	MsgBox, 52, %strAppName%, % L(lDialogThirdPartyDetected, strAppName, "Total Commander")
	IfMsgBox, No
		strTotalCommanderPath := "NO"
	else
	{
		strTotalCommanderPath := strCheckTotalCommanderPath
		Gosub, SetTCCommand
		
		; disable Group menu for TC users until the tabs issue is resolved
		blnDisplayGroupMenu := 0
		IniWrite, %blnDisplayGroupMenu%, %strIniFile%, Global, DisplaySwitchMenu
	}
	blnUseTotalCommander := (strTotalCommanderPath <> "NO")
	IniWrite, %strTotalCommanderPath%, %strIniFile%, Global, TotalCommanderPath
	blnTotalCommanderUseTabs := 1
	IniWrite, %blnTotalCommanderUseTabs%, %strIniFile%, Global, TotalCommanderUseTabs
	; strTotalCommanderNewTabOrWindow will contain "/O /T" to open in a new tab if TotalCommanderUseTabs is 1 (default) or "/N" to open in a new file list
	strTotalCommanderNewTabOrWindow := "/O /T"
}

return
;------------------------------------------------------------


;------------------------------------------------------------
GetTotalCommanderPath()
;------------------------------------------------------------
{
	RegRead, strPath, HKEY_CURRENT_USER, Software\Ghisler\Total Commander\, InstallDir
	If !StrLen(strPath)
		RegRead, strPath, HKEY_LOCAL_MACHINE, Software\Ghisler\Total Commander\, InstallDir

	if FileExist(strPath . "\TOTALCMD64.EXE")
		strPath := strPath . "\TOTALCMD64.EXE"
	else
		strPath := strPath . "\TOTALCMD.EXE"
	return strPath
}
;------------------------------------------------------------


;------------------------------------------------------------
SetTCCommand:
;------------------------------------------------------------

IniRead, blnTotalCommanderUseTabs, %strIniFile%, Global, TotalCommanderUseTabs, 1 ; should be intialized here but true by default for safety
if (blnTotalCommanderUseTabs)
	strTotalCommanderNewTabOrWindow := "/O /T" ; open new folder in a new tab
else
	strTotalCommanderNewTabOrWindow := "/N" ; open new folder in a new window (TC instance)

; additional icon for TotalCommander
objIconsFile["TotalCommander"] := strTotalCommanderPath
objIconsIndex["TotalCommander"] := 1

return
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



;========================================================================================================================
; GENERAL
;========================================================================================================================


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
GuiCenterButtons(strWindow, intInsideHorizontalMargin := 10, intInsideVerticalMargin := 0, intDistanceBetweenButtons := 20, intLeftRightOffset := 0, arrControls*)
; This is a variadic function. See: http://ahkscript.org/docs/Functions.htm#Variadic
;------------------------------------------------------------
{
	DetectHiddenWindows, On
	Gui, Show, Hide
	WinGetPos, , , intWidth, , %strWindow%

	intWidth := intWidth + intLeftRightOffset
	intMaxControlWidth := 0
	intMaxControlHeight := 0
	for intIndex, strControl in arrControls
	{
		GuiControlGet, arrControlPos, Pos, %strControl%
		if (arrControlPosW > intMaxControlWidth)
			intMaxControlWidth := arrControlPosW
		if (arrControlPosH > intMaxControlHeight)
			intMaxControlHeight := arrControlPosH
	}
	
	intMaxControlWidth := intMaxControlWidth + intInsideHorizontalMargin
	intButtonsWidth := (arrControls.MaxIndex() * intMaxControlWidth) + ((arrControls.MaxIndex()  - 1) * intDistanceBetweenButtons)
	intLeftMargin := (intWidth - intButtonsWidth) // 2

	for intIndex, strControl in arrControls
		GuiControl, Move, %strControl%
			, % "x" . intLeftMargin + ((intIndex - 1) * intMaxControlWidth) + ((intIndex - 1) * intDistanceBetweenButtons)
			. " w" . intMaxControlWidth
			. " h" . intMaxControlHeight + intInsideVerticalMargin
}
;------------------------------------------------------------


;------------------------------------------------------------
Hotkey2Text(strModifiers, strMouseButton, strOptionKey)
;------------------------------------------------------------
{
	if (strMouseButton = "None") ; do not compare with lOptionsMouseNone because it is translated
		str := lOptionsMouseNone ; use lOptionsMouseNone because this is displayed
	else
	{
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
GetMouseButton4Text(strSource)
; Returns the string in arrMouseButtons at the same position of strSource in arrMouseButtonsText
;------------------------------------------------------------
{
	global
	
	loop, %arrMouseButtonsText0%
		if (strSource = arrMouseButtonsText%A_Index%)
			return arrMouseButtons%A_Index%
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


;------------------------------------------------------------
GetDeepestFolderName(strLocation)
;------------------------------------------------------------
{
	SplitPath, strLocation, , , , strDeepestName, strDrive
	if !StrLen(strDeepestName) ; we are probably at the root of a drive
		return strDrive
	else
		return strDeepestName
}
;------------------------------------------------------------


;------------------------------------------------------------
LocationIsDocument(strLocation)
;------------------------------------------------------------
{
    FileGetAttrib, strAttributes, %strLocation%
    return !InStr(strAttributes, "D") ; not a folder
}
;------------------------------------------------------------


;------------------------------------------------------------
Url2Var(strUrl)
;------------------------------------------------------------
{
	objWebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	/*
	if (A_LastError)
		; an error occurred during ComObjCreate (A_LastError probably is E_UNEXPECTED = -2147418113 #0x8000FFFFL)
		BUT DO NOT ABORT because the following commands will be executed even if an error occurred in ComObjCreate (!)
	*/
	objWebRequest.Open("GET", strUrl)
	objWebRequest.Send()

	Return objWebRequest.ResponseText()
}
;------------------------------------------------------------


;------------------------------------------------------------
OSVersionIsWorkstation()
;------------------------------------------------------------
{
	If (OSVersion := GetOSVersionInfo())
		return (OSVersion.ProductType = 1)
	else
		return false
}
;------------------------------------------------------------


;------------------------------------------------------------
GetOSVersionInfo()
; by shajul (http://www.autohotkey.com/board/topic/54639-getosversion/?p=414249)
; reference: http://msdn.microsoft.com/en-ca/library/windows/desktop/ms724833(v=vs.85).aspx
;------------------------------------------------------------
{
	static Ver
	If !Ver
	{
		VarSetCapacity(OSVer, 284, 0)
		NumPut(284, OSVer, 0, "UInt")
		If !DllCall("GetVersionExW", "Ptr", &OSVer)
		   return 0 ; GetSysErrorText(A_LastError)
		Ver := Object()
		Ver.MajorVersion      := NumGet(OSVer, 4, "UInt")
		Ver.MinorVersion      := NumGet(OSVer, 8, "UInt")
		Ver.BuildNumber       := NumGet(OSVer, 12, "UInt")
		Ver.PlatformId        := NumGet(OSVer, 16, "UInt")
		Ver.ServicePackString := StrGet(&OSVer+20, 128, "UTF-16")
		Ver.ServicePackMajor  := NumGet(OSVer, 276, "UShort")
		Ver.ServicePackMinor  := NumGet(OSVer, 278, "UShort")
		Ver.SuiteMask         := NumGet(OSVer, 280, "UShort")
		Ver.ProductType       := NumGet(OSVer, 282, "UChar") ; 1 = VER_NT_WORKSTATION, 2 = VER_NT_DOMAIN_CONTROLLER, 3 = VER_NT_SERVER
		Ver.EasyVersion       := Ver.MajorVersion . "." . Ver.MinorVersion . "." . Ver.BuildNumber
	}
	return Ver
}
;------------------------------------------------------------


