# FoldersPopup - Read me

Move like a breeze between your frequently used folders and documents! Freeware.

Written using AHKScript v1.1.09.03+ (http://www.ahkscript.org)  
By JnLlnd on [AHKScript forum](http://ahkscript.org/boards/memberlist.php?mode=viewprofile&u=66)
Thanks to LearningOne for sharing his code and to others who helped in [this thread on AHKScript.org forum](http://ahkscript.org/boards/viewtopic.php?f=5&t=526).

## Links

* [Application home](http://code.jeanlalonde.ca/folderspopup/)
* [Download 32-bits / 64-bits](http://code.jeanlalonde.ca/ahk/folderspopup/folderspopup.zip)

## History

### Version: 5.1 (2015-05-06)
* See beta versions v5.0.9 to 5.0.9.0
	
### Version: 5.0.9.9 (2015-05-06)
* Enable keyboard shortcuts even if Current folders, Groups of Folders and Clipboard menus are disabled
* Dutch language update

### Version: 5.0.9.8 (2015-05-02)
* fix bug causing error when trying to show icon in Clipboard menu when icons are not allowed
* fix bug with None in change hotkey dialog box
* updates of language files
	
### Version: 5.0.9.7 (2015-04-27)
* fix a bug with relative paths being combined wrongly when the location in an URL
* in change hotkey dialog box, make the selection of no hotkey (None) more obvious
* preserve standard order of modifiers in hotkey labels when changing hotkey
* updates of language files
	
### Version: 5.0.9.5/6 (2015-04-24)
* simplified implementation of the copy location to clipboard feature; English language adapted
* addition of the Brazilian Portuguese language !
* update to Spanish language file
	
### Version: 5.0.9.4 (2015-04-23)
* fix bug when adding folders using the drag-and-drop technique, these favorites being considered as application favorites
* re-wording of the language around the "Paste Location" feature to "Copy location"
* expand the relative path in favorite location, based on the current working directory, making the change folder support relative paths
* adjustments to all translation language files
	
### Version: 5.0.9.3 (2015-04-15)
* sort URLs in Clipboard menu
* Spanish, Dutch and French text updates
* help text updates, addition of help text about Clipboard and Paste Favorite's Location menus
* display current customized shortcuts in Help text (English only for now)
* support comments starting with ";" in language files
* support comments at end of lines after ";" in language files

### Version: 5.0.9.2 (2015-04-11)
* addition of spanish language
* fix Change hotkey description text for hotkeys 3 to 6
* Add English description text in Change hotkey for hotkeys 7 to 10
* fix buttons centering bug in Option GUI
* fix text layout in Change hotkey GUI
	
### Version: 5.0.9.1 (2015-04-10)
* merge changes in v5.0.1
	
### Version: 5.0.9 (2015-04-05)
* new Paste Favorite's Location in the main menu
* new shortcut (default Shift-Windows-V) to open the Paste Location menu
* new radio button options for paste location destination in Options tab 1
* paste favorite's location to keyboard or clipboard, according to destination selected in Options
* new checkbox option in Options tab 1 to display or not the Paste Favorite's Location menu
* add tray tip when showing the Paste Favorite's Location
* disable Groups, Settings, Add this folder and Support freeware menus when showing Paste menu

### Version: 5.0.1 (2015-04-10)
* change default hotkeys for Current Folders (+^f), Groups of Folders (+^g), Recent Folders (+^r), Clipboard (+^c) and Settings (+^s) for Windows 8.1 compatibility
* fix bug with special folders Pictures and Favorites (Internet) when user change these folders default location

### Version: 5.0 (2015-04-05)
(see history for v4.9.1 to 4.9.9.1)
	
### Version: 4.9.9.1 (2015-04-03)
* removed menu shortcuts in main menu to let user select menu item by their name first letter
* German, Dutch, Italian and Korean language update
	
### Version: 4.9.9 (2015-03-30)
* keep current position of add favorite window when changing favorite type
* fix bug making the exit routine running twice
* moving OnExit before InitFileInstall to ensure deletion of temporary files
* unused language variable removed
	
### Version: 4.9.8 (2015-03-28)
* add an option in third tab to display or not special menu shortcut
* adjust layout ot Options tab 3
* save and load display special menu shortcut option to ini file
* function to build main menu with special menu shortcuts text
* shorten button names in Hotkey2Text
* fix bug in group load for special folders
* Italian and Swedish Options language updates
	
### Version: 4.9.7 (2015-03-25)
* fix a bug with "New window" (Shift-MMB or Shift-Win-A) not opening in a new Explorer when mouse over an Explorer
* improve target window identification when special menu are called using their shortcuts (if target window can open favorite, then navigate, if not new window)
* sets menu position correctly when special menu are called using their shortcuts
* fix bug in WindowIsFPconnect when target window id or class is unknown
* review of English text in Menu hotkeys Options tab and improve Menu hotkeys tab layout
* Italian, Swedish, French and Korean language update
	
### Version: 4.9.6.2 (2015-03-21)
* fix a bug in OpenFavorite (and OpenClipboard) in situations where the target window could not be detected
	
### Version: 4.9.6.1 (2015-03-20)
* addition of debugging code around OpenClipboard
* fix a bug introduced in v4.9.2 breaking the creation of default menu at first run
	
### Version: 4.9.6 (2015-03-19) (no v4.9.5)
* change default hotkleys for Settings (+#s), Current Folders (+#f) and Clipboard (+#s)
* review hotkeys array naming
* add hotkey reminders in special menu labels in main menus
	
### Version: 4.9.4 (2015-03-18)
* add a hotkey to open directly the Current folders menu (by default Ctrl-Win-C)
* add a hotkey to open directly the Groups menu (by default Ctrl-Win-G)
* add a hotkey to open directly the Clipboard menu (by default Ctrl-Win-V)
* redesign the Options dialog box splitting hotkeys in two tabs: one for popup menu hotkeys and one for other hotkeys
* review hotkeys language in Options
* fix a bug from v4.2 when opening a special folder (Libraries, My Computer, etc.) from the Current Folders menu

### Version: 4.9.3 (2015-03-14)
* add URL parsing in Clipboard submenu
* keep only URLs shorter than 260 chars
* fix icon bug inside Clipboard menu (using only Folder and URL icons for now)
* filter out illegal characters in paths / ? : * " > < | (in addition to space, tab and line-feed) from the beginning and the end of each clipboard line
* fix two bugs in OpenClipboard making folders always opening in new window

### Version: 4.9.2 (2015-03-12)
* check for beta versions updates
* enabled only for users who ran a beta version previously and who enabled the Check for update option
	
### Version: 4.9.1 (2015-03-10)
* add the favorite type Application
* add Arguments and Working directory fields to Application favorites
* execute the Application favorites passing properly the arguments and setting the working directory
* make room in the Add Favorite window for additional property fields
* support default and custom icons for Application favorites
* add the Clipboard menu item in the main menu and add to the submenu folders, documents or applications paths found in the Clipboard
* if no path is found in the current Clipboard, the previous submenu content is preserved
* add an option to determine if the Clipboard menu is shown (default true)
* disable clipboard submenu if empty
* add clipboard icon to Clipboard menu
* remove arguments double quotes when there is no argument
* process environment vars for app favorites and clipboard paths

### Version: 4.3 (2015-02-22)
* make the Settings window resizable
* save Settings Gui size state to ini file on quit
* restore Settings Gui size on load
* when saved maximized, restore at default size and center
* prevent minimizing the settings window to avoid user to forget to save settings
* adjust hand mouse pointer when hover clickable controls

### Version: 4.2.2 (2015-01-31)
* fix a bug with environment variables not being expanded when checking if target file exist
* fix bug under XP during group load when an Explorer already contains the target folder, the existing Explorer is now activated and resized (consequence: a group cannot contain the same folder twice)
* fix bug with check for update URL on some browsers
* adding diag code to Check4Update command
* stop incrementing usage counter when checking for update manually

### Version: 4.2.1 (2015-01-18)
* make FP compliant with Windows themes by adding a FP theme named "Windows" that keeps Windows theme's colors (making FP display OK when user selects a dark Windows theme)
* making the FP theme "Windows" selected by default for new users
* because of a side-effect in XL 2010, revert a patch in v4.2 to prevent double-click up/down buttons in Settings to overwrite the clipboard with the image URL (a Windows "undesired feature")

### Version: 4.2 (2015-01-15)
* See all changes from v4.1.8 to v4.1.9.6 BETA

### Version: 4.1.9.1 to 4.1.9.6 BETA (2015-01-10 to 2015-01-15)
* German, Dutch, Swedish and Italian language updates
* Minimized language variable added
* change landing page URL in FP code for a redirect page easier to manage on the website
* remove timeout from msgbox in check4update
* revert ampersand (&) handling in menu as it was in 4.1, one & for shortcut, && to display an ampersand
* Added the FP ico file to the portable package
* re-enabled the Special Folders menu in Windows XP

### Version: 4.1.8.7 BETA (2015-01-08)
* add the new customizable My Special Folders menu as last item in the user's main menu (except for XP)
* add a bln value in ini file to track that the new My Special Folders was created (except for XP)
* protection if user already has a My Special Folders menu before FP creates it
* stop building the old Special Folders menu (except for XP)
* remove option to display special folders menu (except for XP)
* in Settings, Ctrl-Left is now as clicking on on the left arrow (instead of the up arrow) beside the menu dropdown list
* in Settings, remember the last menu position when returning to a previously displayed menu
* refactor of code around navigation to previous menu (arrows left of the Menu to edit in Settings)
* remove & in special folders menu names in language files
* fix bug when moving up/down or removing favorite, the items list in add favorite is now updated
* fix error message bug when moving folder under itself
* swedish translation update

### Version: 4.1.8.6 BETA (2015-01-06)
* improve performance when moving large number of favorite from one submenu to another
* fix bug & not being kept in menu names
* fix bug in add favorite when changing favorite type, default icon not being properly set and location not being properly reset
* fix bug when moving out all favorite from a submenu, menu item is now grayed out
* Korean language updates

### Version: 4.1.8.5 BETA (2015-01-04)
* prevent double-click on Up/Down arrows buttons to overwrite the clipboard
* fix a bug when moving multiple favorites with Up/Down or Ctrl-Up/Ctrl-Down, selection is now kept

### Version: 4.1.8.4 BETA (2015-01-03)
* add hotkeys to Gui to move favorites (Ctrl-Down/Up), edit favorite (Enter), open submenu (Ctrl-Right), return to parent menu (Ctrl-Left), Select All (Ctrl-A), Add new (Ctrl-N) and Remove (Del)
* add shortcuts help in main Gui, new layout for drag & drop help
* allow multiple select of favorites to move or delete them; adapt gui Edit and Remove buttons if multiple selection
* add separator and column break not allowed if multiple favorites selected
* looped uses of adapted guiaddfavoritesave to move favorites
* add moved favorites at end of destination menu
* arrows move multiple favorites
* in add/edit favorite, save default button 
* special folder Performance Information and Tool only on Win7 (not available on Win8 and more)
* add support for six special folders in TC with use of :: instead of shell:::
* fix bug special folder Images with Total Commander
* fix bug in manage groups, Select a group was sorted with list of groups

### Version: 4.1.8.3 BETA (2014-12-31)
* fix bug default position in menu not correct after last items in menu was removed
* fix bug when change to submenmu using edit button, the delete button in new submenu deleted items in the previous menu

### Version: 4.1.8.2 BETA (2014-12-31)
* complete refactor of special folders using CLSID, Shell commands, Shell constants, AHK constants, DOpus alias or TC commands, and supporting NavigateExplorer, NewExplorer, Dialog, Console, DOpus, TC and FPc
* adaptation of OpenFavorite and navigate/new window functions to the refactored special folders
* error message when a special folder cuold not be open
* add support for FPconnect TargetPath filename
* support environment variables in FPc paths
* fix bug when moving a favorite to another submenu

### Version: 4.1.8.1 BETA (2014-12-27)
* removed support for FreeCommander XE (now available via FPconnect)
* add version and os info to check4update request

### Version: 4.1.8 BETA (2014-12-26)
* add dropdown list in Add Favorite dialog box to select the position of the new favorite in the menu
* for Windows 7 and more, refactor InitSpecialFolders with ClassID and exceptions for unavailable ClassID (the Special Folders submenu is maintained but users could replace it by creating their own Special Folders in any menus)
* add icons and translateble default names for exceptions
* extended support for FPconnect (universal file manager connecteor from Roland Toth) with auto-detection of FPconnect, open in current tab/window or new tab/window
* fix bug not showing icons for system menus in main menu under Win_XP
* fix delay in group load for slow drives
* the FP menu can now be open over the FP Settings window with middle-mouse click (or Win-A)
* language files updates

### Version: 4.1 (2014-12-20)
* addition of Italian language, thanks to Riccardo Leone
* redesign the Help and Options windows into three tabs to save height on small screens
* change mouse cursor to hand only in Settings window
* change delays in group load
* add diagnostic info for clipboard in group load
* solve icon issue with multi column menus under Win XP, show icons only in first columns
* change default to "add to existing windows" when creating a new group of folders
* add BETA support for file manager connector FPConnect (from Roland Toth)
	
### Version: 4.0.4 (2014-12-13)
* add a button to select or deselect all Explorer windows in Group Save
* support column click in Group Save to sort on column content
* fix bug in Explorer collection causing the Save Group button and menu to be disabled

### Version: 4.0.3 (2014-12-13)
* more robust group load and window move and resize
* fix a bug in Explorer collections in case ComObjCreate returns an invalid handle
* remove forgotten testing code in DOpus group load
* when close before restoring group stop closing IE windows
* stop closing TC windows before restoring

### Version: 4.0.2 (2014-12-12)
* bug making language (other than English) in setup not being taken into account at FP first run

### Version: 4.0.1 (2014-12-09)
* fix bug with Recent shortcut opening in a new Explorer window instead of navigating in the correct window
* fix bug properly exit group load loop when an error occurs within an iteration

### Version: 4.0.0 (2014-12-07)
* See all changes from v3.9.1 to v3.9.9 BETA

### Version: 3.9.9 BETA (2014-12-04)
* detect if app is started in program files folder and set working dir to appdata
* create a backup of the ini file at launch time
* Dutch and Korean language updates

### Version: 3.9.8 BETA (2014-12-01)
* fix lOptionsDisplayFoldersInExplorerMenu label.
* add column break and system variable in default menu
* fix bug when edit and save a submenu under the same name (from v3.3.2)
* add double-quotes to Run command parameters
* sort groups list in manage groups and edit group

### Version: 3.9.7 BETA (2014-11-25)
* add an item in the right-click Tray menu to open the FoldersPopup.ini file
* add an option to disable check for update at startup
* add Downloads folder to Special Folders menu, support for DOpus and TC, not available on Win_XP
* fix a bug making custom icons not following when favorites were moved up or down in the menu
* merge and refactor GuiMoveFavoriteUp and GuiMoveFavoriteDown commands
* fix a bug visible only to Total Commander users occuring when you left-click the tray icon button or when left-click on the tray icon was in the overflow area
* add location URL of folders in groups saved to the ini file

### Version: 3.9.6 BETA (2014-11-21)
* refactor BuildGroupMenu into BuildFoldersInExplorerMenu and stripped BuildGroupMenu
* add numeric shortcuts to groups menu
* exclude DOpus collection windows of Current folders menu
* merged OpenRecentFolder and OpenFolderInExplorer with OpenFavorite
* merged 2-in-1 command PopupMenuMouse + PopupMenuKeyboard and 2-in-1 command PopupMenuNewWindowMouse + PopupMenuNewWindowKeyboard, into 4-in-1 command
* support for system environment variables in favorite location (e.g.: APPDATA, LOCALAPPDATA, ProgramData, PUBLIC, TEMP, TMP, USERPROFILE)
* make the vertical bar (or pipe "|") a reserved character in submenu or favorite name
* fix bug clicking the correct pane in DOpus when popup in new window
* fix bug with document favorite custom icons
* fix a bug occurring in some situation when a favorite location contains a comma (from v3.3.1)

### Version: 3.9.5 BETA (2014-11-15)
* display and select icon for folders, url and documents in add/edit favorite and in menu
* better error management around menu icon assignment, fix the *.msc bug
* use shell32.dll icon #1 for unknown icon
* fix a bug in Group menu for network locations

### Version: 3.9.4 BETA (2014-11-09)
* Swedish, German and Korean translations for new features in v3.9.1 and v3.9.2
* Custom icons for submenus (custom icons for folders in next release)
* Add hidden column in Settings listview for icon resource
* Add field in Folders section of ini file for icon resource
* Add icon selector to add/edit favorite dialog box (for submenu only in this release)
* New special menu Folders in Explorer to open in Explorer or a dialog box a folder already open in another Explorer
* Merge open folder in Explorer with Open recent folders
* Add option to display or not the Folders in Explorer menu
* Regroup Display options in Options dialog box
* In Options, add the size 48 pixels to the choie of icon size

### Version: 3.9.3 BETA (2014-11-08)
* retrieve language from ini file created by setup program and use when creating the FP ini file
* accept space in Change hotkey dialog box to allow combinations with spacebar as a potential hotkey
* detect TreeView folder select dialog box and exclude them because of a Windows limitation (Edit1 control not handling the Enter)
* add the option OpenMenuOnTaskbar to open or not the popup menu over the taskbar (class Shell_TrayWnd)
* add column breaks in menu
* improve reliability and performance of group load with Explorer and DOpus
* fix bug with windows move/resize when group load
* fix bug with minimize/maximize Explorers when group load
* fix wrong web link when an beta version update is available

### Version: 3.9.2 BETA (2014-11-05)
* Addition of German language to Setup program
* Add the possibility to overwrite an existing group of folders in the save group dialog box
* allow to edit a group from the manage groups gui
* Delete the startup shortcut when uninstall with Inno Setup
* After installation with Inno Setup, copy an existing FoldersPopup.ini file if one exist in a previous protable installation (findable only if a shortcut to the portable installation exists)

### Version: 3.9.1 BETA (2014-11-02)
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

### Version: 3.3.2 (2014-12-01)
* fix a bug occurring when editing a submenu and saving it under the same name

### Version: 3.3.1 (2014-11-17)
* fix a bug occurring in some situation when a favorite location contains a comma

### Version: 3.3 (2014-10-24)
	
TotalCommander integration
* automatic detection for Total Commander support
* Total Commander configuration in Options
* ini configuration for TotalCommander
* add a checkbox in options to let Total Commander users choose to open new folders (Shift+Middle-Mouse) in a new tab or in a new window
* new TotalCommanderUseTabs and TotalCommanderNewTabOrWindow switches in ini file
* show popup menu in TotalCommander windows
* add this folder from Total Commander window
* navigate regular and special folder in TotalCommander existing window
* open regular and special folder in new TotalCommander window or tab according to TotalCommanderNewTabOrWindow
* disable Switch menu first time TC is enabled until the tabs issue is resolved in TC

Other changes
* addition of Swedish language, thanks to Åke Engelbrektson
* fix a bug when user select a hotkey replacement for Middle-mouse button that involves a modifier key (e.g. Shift+Right-click)
* fix bug with icons on Windows Server (disable icons)
* fix bug making new folders opening in Explorer instead of Total Commander or Directory Opus when called from the Tray left-click menu
* change DOpus command to open a new lister to Go with NEW parameter

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


