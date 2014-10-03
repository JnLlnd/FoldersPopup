# FoldersPopup - Read me

Popup menu to jump from folders to folders. Freeware.

Written using AHKScript v1.1.09.03+ (http://www.ahkscript.org)  
By JnLlnd on [AHKScript forum](http://ahkscript.org/boards/memberlist.php?mode=viewprofile&u=66), based on DirMenu v2 by Robert Ryan 2012 (rbrtryn on [AutoHotkey.com forum](http://www.autohotkey.com/board/user/15020-rbrtryn/))  
Thanks to LearningOne for sharing his code and to others who helped in [this thread on AHKScript.org forum](http://ahkscript.org/boards/viewtopic.php?f=5&t=526).

## Links

* [Application home](http://code.jeanlalonde.ca/folderspopup/)
* [Download 32-bits / 64-bits](http://code.jeanlalonde.ca/ahk/folderspopup/folderspopup.zip)

## History

### Version: 3.2.2 (2014-10-02)
* fix layout in options gui
* remove support for MS Office 2003/2007 file dialog boxes
* German language update

### Version: 3.2.1 (2014-09-20)
* When Explorer replacement activated in DOpus, ghost Explorer in the Switch Explorer menu skipped
* Removed Flattr from donation platforms
* Remove Switch Explorer support for DOpus listers containing an FTP folder (until issue resolved - https://github.com/JnLlnd/FoldersPopup/issues/84)
* Addition of the korean language - thanks to Om Il-Sung (Dollnamul)

### Version: 3.2 (2014-09-16)
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

### Version: 3.1.3 (2014-09-07)
* bug fix: make all special folders menu items work when popup menu is activated from the tray icon
* improve handling of the hash (aka Sharp / "#") bug in Shell.Application (see v1.2.6)
* fix bug when navigating in a CMD window with path including AHK reserved chars

### Version: 3.1.2 (2014-09-03)
* Menu icons now supporting Windows Vista
* Stop building recent folders menu at startup (unnecessary since this menu is refreshed on demand)

### Version: 3.1 (2014-08-29)
* First public release of Folders Popup v3
* Fix a bug in Switch in dialog box menu
	
### Version: 3.0.12 BETA (2014-08-27)
* German and Dutch translation update (Thanks to Edgar "Fast Edi" Hoffmann and Pieter Dejonghe)
* Left click on Tray icon to show favorites menu
	
### Version: 3.0.11 BETA (2014-08-24)
* fix an icon error under WinXP
	
### Version: 3.0.10 BETA (2014-08-23)
* fix bug when selecting a mouse hotkey after None was selected for that hotkey
* in Change Hotkey, unselect modifiers when None is selected as mouse trigger
* additional text to clarify triggers in Settings, Options
* new menu icon for submenus

### Version: 3.0.9 BETA (2014-08-22)
* replaces Send command with SendInput
* fix bug when navigating to network folder in DOpus
* add popup menu and color to tray menu
	
### Version: 3.0.8 BETA (2014-08-20)
* add type of favorites for links, display default browser icon for link favorites and open links in default browser
* fix bug with DOpus when path includes AHK reserved chars
* better support of DOpus when in dual listers
	
### Version: 3.0.7 BETA (2014-08-18)
* make display icons optional, refactor Add Menu commands in a centralized function
* allow to select no mouse trigger for popup menu, add None to the dropdown list in Change hotkey window
* add mouse or keyboard hotkey to open the recent folders list
* fix error when icon location contains %1
* fix error when assigning color to an empty submenu
* fix a v2 bug with shortcuts numbers increment in Switch menus
	
### Version: 3.0.6 BETA (2014-07-26)
* Redesign of buttons in Settings
* Addition to ini file of themes with colors for dialog boxes and menu
* Implementation of colors to menus and dialog boxes
* Add option in Settings/Options to select theme
	
### Version: 3.0.5 BETA (2014-07-23)
* fix a v2 bug allowing editing in Settings with no item selected
* fix a v3.0.2 bug when adding an item to a menu other than the current menu in Settings
* change cursor to hand for all buttons in Settings
* refactor (merge) Add and Edit favorites GUI and Save commands (no change visible to users)
	
### Version: 3.0.4 BETA (2014-07-21)
* fix a bug when adding a menu and numeric shortcuts are active
* lighter tray tip message after menu is updated in settings
* fix a bug when retrieving icons for documents
* change cursor for an hand for all buttons in Gui
* support icons for document being executable files
	
### Version: 3.0.3 BETA (2014-07-20)
* remove "supported dialog boxes" management
* in gui remove listview, add/edit/remove buttons, reposition other buttons
* remove add dialog box menu, save dialog box, dialog is supported function
	
### Version: 3.0.2 BETA (2014-07-19)
* add favorite type "F" folder, "D" document or "S" submenu and refactor all
* remove or add ... to main menu items
* manage icons resource at init, supporting XP and Win7+
* include parent menu dropdown list when add favorite
* fix old 2.0 bug not detecting name already used when adding from add this folder
* menu icon size default size to 16 for XP and 24 for other OS
	
### Version: 3.0.1 BETA (2014-07-15)
* do not check if network favorites exist
* error icon when local favorite does not exist (removed feature)
* error message when unavailable local favorite is selected in popup menu
* traytip status when refreshing menus
	
### Version: 3.0.0 BETA (2014-07-14)
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
	
### 2014-07-11 2.2.1
* fix bug when adding a folder to a submenu using drag and drop
* add an incentive message about drag and drop at the bottom of Settings window
* ignore submenu change in Settings when user select the current menu

### 2014-07-06 2.2
* support drag and drop to add favorite
* make the cursor change to a hand when the mouse pointer is over buttons or clickable text in Settings dialog box (tried to also implement tooltips but even with a timer, it flickers too much)
* Recent folders menu now shown in a detached menu, at the calling popup menu location, refreshed each time it is opened, with tooltip while refreshing
* fix a bug with number of Recent folders hide/display in Settings, Options
* fix layout bug in edit folder dialog box
* fix bug with Switch to Explorer opening a new window
* replace PCAstuces review URL with Freewares & Tutos

### 2014-06-25 2.1.1
* complete translation of mouse button names
* fix bug when changing Settings shortcut
* fix PCAstuces missing URL

### 2014-06-17 2.1.0
* when adding this folder, select in which menu to add the new folder
* new button when edit menu entry to open this menu
* in edit folder dialog box, set focus to and select folder name
* on-demand recent folders update to keep the popup menu snappy regardless of the number of recent items to parse or the performance of the PC
* option in settings to choose the number of recent folders in popup menu, now default to 10
* refactor (code merge) of GuiAddFolderSave and GuiEditFolderSave
* allow to add this folder from a network folder starting with "\\"
* fix bug with up arrow to go to parent menu
* addition of Dutch translation (thanks to Pieter Dejonghe!)
* fix missing translations

### 2014-06-06 2.0.3
* fix bugs with switch folders and recent folders options
* update german translation

### 2014-06-03 2.0.2
* improve performance of Recent Folders menu building, process only recent folders in recent items
* fix bug when a recent folder is not available (only XP?)
* fix header bug in diagnostic mode

### 2014-06-01 2.0.1
* complete german translation
* fix language typos

### 2014-05-28 2.0.0
* see all additions from v1.5 ALPHA to v1.9 BETA

### 2014-05-27 v1.9 BETA (not to be released)
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

### 2014-05-04 v1.8 ALPHA (not to be released)
* add switch in dialog box to other explorer windows already opened
* lMenuReservedShortcuts management with translations
* sort folders button
* folder up button
* translated help to French
* support freeware to popup menu
* blnMenuReady before popup

### 2014-04-27 v1.7 ALPHA (not to be released)
* new settings dialog box layout with icons to add, edit or remove folders or dialog boxes
* icons to open help, about and settings dialog boxes
* dropdown to select the submenu to edit
* left arrow to go back to edit the menu(s) previously displayed
* double-click to edit folders or supported dialog boxes
* adjustments to dialog boxes for German and French translation
* updated about and help dialog boxes
* solved a bug when Add this folder in some type of dialog boxes

### 2014-04-19 v1.6 ALPHA (not to be released)
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

### 2014-03-22 v1.5 ALPHA (not to be released)
* add recent folders sub-menu
* add ini variable RecentFolders
* when blnDisplayMenuShortcuts reserve shortcut chars for app's items in menus
* add GetDeepestFolderName as function
* add ValueIsInObject function
* add language dropdown
* display full folder names in recent folders
* add swith submenu to activate any other open Explorer
* add DisplayRecentFolders and DisplaySwitchMenu options in Options dialog box and ini file

### 2014-04-25 v1.2.7
* Workaround to make the "Run" command work on some system
* Fix end-of-line at end of version number bug

### 2014-04-24 v1.2.6
* Workaround for the hash (aka Sharp or #) bug in Shell.Application that occurs only when navigating in the current Explorer window to a subfolder including “#” in its parent path (e.g.: C:\C#\Project)
* Windows XP only: fix a bug when navigating to the special folder “My Pictures” in dialog boxes

### 2014-04-19 v1.2.5
* Support for FreeCommander XE
* Compatible with Clover (opens the folder in a new tab)
* Fix wrong error message issue #28

### 2014-04-17 v1.2.4
* Fix shortcut (hotkey) assignments error (not a valid key name error) on Windows system with keyboard regional settings supporting Cyrillic letters (Russian and others)

### 2014-02-25 v1.2.3
* For Windows XP users only, revert to v1.2.1 state due to unwanted behavior of FP v1.2.2 in Windows Explorer XP

### 2014-02-20 v1.2.2
* Opens new Explorer windows complying with the Explorer navigation pane setting

### 2014-02-01 v1.2.1
* Fix a bug that added separator lines at the bottom of Tray Menu (one line added at each display of the popu menu)
* Improve diagnostic data collection (always at the user's discretion)

### 2013-01-26 v1.2
* Add an option to add numeric keyboard shortcuts to launch folders in popup menu
* Add an option to display the popup menu at a fix position
* Add a diagnostic mode to collect support info (add DiagMode=1 under [Global] section in ini file)
* Redesign of the Options dialog box

### 2013-12-24 v1.01
* Bug fix: mouse and keyboard triggers were disabled in non-explorer windows

### 2013-12-23 v1.0 (First Official Release)
* Configurable mouse button and keyboard triggers in a new "Options" dialog box
* New keyboard triggers (by default, Windows-K and Shift-Windows-K) in addition to mouse button triggers (by default, Middle mouse and Shift-Middle mouse buttons)
* Add "Run at startup" checkbox to "Options" dialog box to launch Folders Popup automatically at Windows startup
* Add "Display the startup tray tip" checkbox to "Options" dialog box to display or hide the Folders popup's tray tip
* Add "Display Special Folders" checkbox to "Options" dialog box to enable/disable navigation to special folders (My Computer, Network, Recycle bion, etc.) in popup menu
* Better formated startup help tray tip
* Close "Settings" dialog box with Escape key

### 2013-11-11 v0.9 (beta version)
* Implemented startup option in tray and check4update, standardize dialog box titles, various text fixes
* Renamed the app FoldersPopup, removed debugging code, prepare for compiler, removed external pictures

### 2013-11-11 v0.5 (last alpha version)
* Implemented GuiAbout and GuiHelp, added About and Help to tray menu, tray tip displayed only 5 times
* Removed file:/// protocol prefix, added support for ExploreWClass, implemented try/catch to Explore shell method, offer to add manually when add folder failed

### 2013-11-11 v0.4
* Add settings hotkey to ini file (default Ctrl-Windows-F)
* Enable Add This Window menu only in WIN_7/WIN_8 dialog boxes (not working in WIN_XP) and in Explorer windows (all versions)
* Add GuiAddFolder and GuiAddDialog buttons commands, add AddThisFolder and AddThisDialog menus commands
	
### 2013-11-10 v0.3
* Add NavigateConsole for console support (command prompt CMD)
* Change .ini filename to new app name
	
### 2013-11-09 v0.2

* Create language file, build gui, tray menu, popup folder menu and skeleton for front end buttons and commands
* Create CanOpenFavorite condition for MButton with WindowIsAnExplorer, WindowIsDesktop and DialogIsSupported, AddThisDialog menu
* Add SpecialFolders menu, OpenFavorite for Explorer and Desktop, NavigateExplorer
* Add support for MS Office dialog boxes on WinXP (bosa_sdm_), open special folders in Explorer and Desktop
* Implement NavigateDialog, add Desktop/Document/Pictures special folders, open it in dialog boxes, enable appropriate menus in popup menu
* Renamed the app PopupFolders, isolate text into language variables, update GitHub and ReadMe.md with v0.2

### 2013-11-05 v0.1

* First release of ALPHA version
* Initialize script, read ini file and create arrays for folders menu and supported dialog boxes


## <a name="copyright"></a>Copyright

This software is provided 'as-is', without any express or implied warranty.  In no event will the authors be held liable for any damages arising from the use of this software.  
  
Permission is granted to anyone to use this software for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:  
  
1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation would be appreciated but is not required.  
2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.  
3. This notice may not be removed or altered from any source distribution.  
  
Jean Lalonde, <A HREF="mailto:ahk@jeanlalonde.ca">ahk@jeanlalonde.ca</A>


