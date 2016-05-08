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


	Version: 5.2.3 (2016-05-08)
	* display clearer message when prompting for upgrade from FoldersPopup to Quick Access Popup (prompt displayed only for new QAP releases with new features)
	* add the auto-detection of .ahk and .vbs extensions when user add a favorite using drag-and-drop to the Settings window
	* stop launching Directory Opus when refreshing the list of open folders in listers if Directory Opus is not running
	* Sweeden and German language files update
	* new runtime v1.1.23.5 from AHK

	Version: 5.2.2 (2015-11-29)
	* Italian and French language updates
	* Fix an error in link for setup file in v5.2.1
	
	Version: 5.2.1 (2015-11-15)
	* update special folders initialization for Windows 10
	* adjust menu icons to Windows 10 icon files
	* shorten application description in executable file for Windows 10 display
	* shorten notification tray tip texts for better display on Windows 10
	* add a short delay after tray tip in notification zone for Windows 10 compatibility
	* add option to disable sound on some tray tips
	* fix bug with favorite application parameters, letting user enclose parameters with double-quotes only if required
	* disabled collecting group load diagnostic data
	* check for updates prompt for download v6+ only if OS is Win7+, if v6+ prompt for ugrade to Quick Access Popup
	* German and French language updates

	Version: 5.1.2 (2015-08-28)
	* fix description label errors when changing a hotkey in "Options, Other hotkeys"
	* when the Explorer extension Clover is installed, support folder navigation in the current tab instead of opening a new tab
	
	Version: 5.1.1 (2015-07-21)
	* fix a bug for FPconnect users preventing the middle-mouse-button click to be recognized by FPconnected file manager (see: http://blog.rolandtoth.hu/post/106133423662/fpconnect-for-folderspopup-windows)
	* improve group load error handling
	
	Version: 5.1 (2015-05-06)
	* See beta versions v5.0.9 to 5.0.9.0
	
	Version: 5.0.9.9 (2015-05-06)
	* Enable keyboard shortcuts even if Current folders, Groups of Folders and Clipboard menus are disabled
	* Dutch language update

	Version: 5.0.9.8 (2015-05-02)
	* fix bug causing error when trying to show icon in Clipboard menu when icons are not allowed
	* fix bug with None in change hotkey dialog box
	* updates of language files
	
	Version: 5.0.9.7 (2015-04-27)
	* fix a bug with relative paths being combined wrongly when the location in an URL
	* in change hotkey dialog box, make the selection of no hotkey (None) more obvious
	* preserve standard order of modifiers in hotkey labels when changing hotkey
	* updates of language files
	
	Version: 5.0.9.5/6 (2015-04-24)
	* simplified implementation of the copy location to clipboard feature; English language adapted
	* addition of the Brazilian Portuguese language !
	* update to Spanish language file
	
	Version: 5.0.9.4 (2015-04-23)
	* fix bug when adding folders using the drag-and-drop technique, these favorites being considered as application favorites
	* re-wording of the language around the "Paste Location" feature to "Copy location"
	* expand the relative path in favorite location, based on the current working directory, making the change folder support relative paths
	* adjustments to all translation language files
	
	Version: 5.0.9.3 (2015-04-15)
	* sort URLs in Clipboard menu
	* Spanish, Dutch and French text updates
	* help text updates, addition of help text about Clipboard and Paste Favorite's Location menus
	* display current customized shortcuts in Help text (English only for now)
	* support comments starting with ";" in language files
	* support comments at end of lines after ";" in language files

	Version: 5.0.9.2 (2015-04-11)
	* addition of spanish language
	* fix Change hotkey description text for hotkeys 3 to 6
	* Add English description text in Change hotkey for hotkeys 7 to 10
	* fix buttons centering bug in Option GUI
	* fix text layout in Change hotkey GUI
	
	Version: 5.0.9.1 (2015-04-10)
	* merge changes in v5.0.1
	
	Version: 5.0.9 (2015-04-05)
	* new Paste Favorite's Location in the main menu
	* new shortcut (default Shift-Windows-V) to open the Paste Location menu
	* new radio button options for paste location destination in Options tab 1
	* paste favorite's location to keyboard or clipboard, according to destination selected in Options
	* new checkbox option in Options tab 1 to display or not the Paste Favorite's Location menu
	* add tray tip when showing the Paste Favorite's Location
	* disable Groups, Settings, Add this folder and Support freeware menus when showing Paste menu

	Version: 5.0.1 (2015-04-10)
	* change default hotkeys for Current Folders (+^f), Groups of Folders (+^g), Recent Folders (+^r), Clipboard (+^c) and Settings (+^s) for Windows 8.1 compatibility
	* fix bug with special folders Pictures and Favorites (Internet) when user change these folders default location
	
	Version: 5.0 (2015-04-05)
	(see history for v4.9.1 to 4.9.9)
	
	Version: 4.9.9.1 (2015-04-04)
	* removed menu shortcuts in main menu to let user select menu item by their name first letter
	* German, Dutch, Italian and Korean language update
	
	Version: 4.9.9 (2015-03-30)
	* keep current position of add favorite window when changing favorite type
	* fix bug making the exit routine running twice
	* moving OnExit before InitFileInstall to ensure deletion of temporary files
	* unused language variable removed
	
	Version: 4.9.8 (2015-03-28)
	* add an option in third tab to display or not special menu shortcut
	* adjust layout ot Options tab 3
	* save and load display special menu shortcut option to ini file
	* function to build main menu with special menu shortcuts text
	* shorten button names in Hotkey2Text
	* fix bug in group load for special folders
	* Italian and Swedish Options language updates
	
	Version: 4.9.7 (2015-03-25)
	* fix a bug with "New window" (Shift-MMB or Shift-Win-A) not opening in a new Explorer when mouse over an Explorer
	* improve target window identification when special menu are called using their shortcuts (if target window can open favorite, then navigate, if not new window)
	* sets menu position correctly when special menu are called using their shortcuts
	* fix bug in WindowIsFPconnect when target window id or class is unknown
	* review of English text in Menu hotkeys Options tab and improve Menu hotkeys tab layout
	* Italian, Swedish, French and Korean language update
	
	Version: 4.9.6.2 (2015-03-21)
	* fix a bug in OpenFavorite (and OpenClipboard) in situations where the target window could not be detected
	
	Version: 4.9.6.1 (2015-03-20)
	* addition of debugging code around OpenClipboard
	* fix a bug introduced in v4.9.2 breaking the creation of default menu at first run
	
	Version: 4.9.6 (2015-03-19) (no v4.9.5)
	* change default hotkleys for Settings (+#s), Current Folders (+#f) and Clipboard (+#s)
	* review hotkeys array naming
	* add hotkey reminders in special menu labels in main menus
	
	Version: 4.9.4 (2015-03-18)
	* add a hotkey to open directly the Current folders menu (by default Ctrl-Win-C)
	* add a hotkey to open directly the Groups menu (by default Ctrl-Win-G)
	* add a hotkey to open directly the Clipboard menu (by default Ctrl-Win-V)
	* redesign the Options dialog box splitting hotkeys in two tabs: one for popup menu hotkeys and one for other hotkeys
	* review hotkeys language in Options
	* fix a bug from v4.2 when opening a special folder (Libraries, My Computer, etc.) from the Current Folders menu

	Version: 4.9.3 (2015-03-14)
	* add URL parsing in Clipboard submenu
	* keep only URLs shorter than 260 chars
	* fix icon bug inside Clipboard menu (using only Folder and URL icons for now)
	* filter out illegal characters in paths / ? : * " > < | (in addition to space, tab and line-feed) from the beginning and the end of each clipboard line
	* fix two bugs in OpenClipboard making folders always opening in new window

	Version: 4.9.2 (2015-03-12)
	* check for beta versions updates
	* enabled only for users who ran a beta version previously and who enabled the Check for update option
	
	Version: 4.9.1 (2015-03-10)
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
	
	Version: 4.3 (2015-02-22)
	* make the Settings window resizable
	* adjust hand mouse pointer when hover clickable controls
	* save Settings Gui size state to ini file on quit
	* restore Settings Gui size on load
	* when saved maximized, restore at default size and center
	* prevent minimizing the settings window to avoid user to forget to save settings
	
	Version: 4.2.4 (2015-02-08)
	* fix a version number in v4.2.3 causing an error in update checking
	* fix a bug with expanded environment variables in favorites of type Special folders
	
	Version: 4.2.2 (2015-01-31)
	* fix a bug with environment variables not being expanded when checking if target file exist
	* fix bug under XP during group load when an Explorer already contains the target folder, the existing Explorer is now activated and resized (consequence: a group cannot contain the same folder twice)
	* fix bug with check for update URL on some browsers
	* adding diag code to Check4Update command
	* stop incrementing usage counter when checking for update manually
	
	Version: 4.2.1 (2015-01-18)
	* make FP compliant with Windows themes by adding a FP theme named "Windows" that keeps Windows theme's colors (making FP display OK when user selects a dark Windows theme)
	* making the FP theme "Windows" selected by default for new users
	* because of a side-effect in XL 2010, revert a patch in v4.2 to prevent double-click up/down buttons in Settings to overwrite the clipboard with the image URL (a Windows "undesired feature")

	Version: 4.2.0 (2015-01-15)
	* see changes in beta version 4.1.8 to 4.1.9.6
	
	Version: 4.1.9.6 BETA (2015-01-15)
	* italian and german translation fixes
	
	Version: 4.1.9.5 BETA (2015-01-14)
	* italian languag fixes
	* Minimized language variable added
	
	Version: 4.1.9.4 BETA (2015-01-14)
	* change beta landing page URL in FP code for a redirect page easier to manage on the website
	* remove timeout from msgbox in check4update
	* German, Dutch and Italian language updates
	
	Version: 4.1.9.3 BETA (2015-01-11)
	* revert ampersand (&) handling in menu as it was in 4.1, one & for shortcut, && to display an ampersand

	Version: 4.1.9.2 BETA (2015-01-10)
	* Italian translation update
	* Korean translation update
	* Swedish translation fix
	* Added the FP ico file to the portable package
	
	Version: 4.1.9.1 BETA (2015-01-10)
	* re-enabled the Special Folders menu in Windows XP
	
	Version: 4.1.8.7 BETA (2015-01-08)
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
	
	Version: 4.1.8.6 BETA (2015-01-06)
	* improve performance when moving large number of favorite from one submenu to another
	* fix bug & not being kept in menu names
	* fix bug in add favorite when changing favorite type, default icon not being properly set and location not being properly reset
	* fix bug when moving out all favorite from a submenu, menu item is now grayed out
	* Korean language updates
	
	Version: 4.1.8.5 BETA (2015-01-04)
	* prevent double-click on Up/Down arrows buttons to overwrite the clipboard (note: feature reverted in v4.2.1 because of a side effect in XL 2010)
	* fix a bug when moving multiple favorites with Up/Down or Ctrl-Up/Ctrl-Down, selection is now kept
	
	Version: 4.1.8.4 BETA (2015-01-03)
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

	Version: 4.1.8.3 BETA (2014-12-31)
	* fix bug default position in menu not correct after last items in menu was removed
	* fix bug when change to submenmu using edit button, the delete button in new submenu deleted items in the previous menu

	Version: 4.1.8.2 BETA (2014-12-31)
	* complete refactor of special folders using CLSID, Shell commands, Shell constants, AHK constants, DOpus alias or TC commands, and supporting NavigateExplorer, NewExplorer, Dialog, Console, DOpus, TC and FPc
	* adaptation of OpenFavorite and navigate/new window functions to the refactored special folders
	* error message when a special folder cuold not be open
	* add support for FPconnect TargetPath filename
	* support environment variables in FPc paths
	* fix bug when moving a favorite to another submenu
	
	Version: 4.1.8.1 BETA (2014-12-27)
	* removed support for FreeCommander XE (now available via FPconnect)
	* add version and os info to check4update request
	
	Version: 4.1.8 BETA (2014-12-26)
	* add dropdown list in Add Favorite dialog box to select the position of the new favorite in the menu
	* for Windows 7 and more, refactor InitSpecialFolders with ClassID and exceptions for unavailable ClassID (the Special Folders submenu is maintained but users could replace it by creating their own Special Folders in any menus)
	* add icons and translateble default names for exceptions
	* extended support for FPconnect (universal file manager connecteor from Roland Toth) with auto-detection of FPconnect, open in current tab/window or new tab/window
	* fix bug not showing icons for system menus in main menu under Win_XP
	* fix delay in group load for slow drives
	* the FP menu can now be open over the FP Settings window with middle-mouse click (or Win-A)
	* language files updates

	Version: 4.1 (2014-12-20)
	* addition of Italian language, thanks to Riccardo Leone
	* redesign the Help and Options windows into three tabs to save height on small screens
	* change mouse cursor to hand only in Settings window
	* change delays in group load
	* add diagnostic info for clipboard in group load
	* solve icon issue with multi column menus under Win XP, show icons only in first columns
	* change default to "add to existing windows" when creating a new group of folders
	* add BETA support for file manager connector FPconnect (from Roland Toth)
	
	Version: 4.0.4 (2014-12-13)
	* add a button to select or deselect all Explorer windows in Group Save
	* support column click in Group Save to sort on column content
	* fix bug in Explorer collection causing the Save Group button and menu to be disabled

	Version: 4.0.3 (2014-12-13)
	* more robust group load and window move and resize
	* fix a bug in Explorer collections in case ComObjCreate returns an invalid handle
	* remove forgotten testing code in DOpus group load
	* when close before restoring group stop closing IE windows
	* stop closing TC windows before restoring

	Version: 4.0.2 (2014-12-12)
	* fix bug making language (other than English) in setup not being taken into account oinly at FP first run

	Version: 4.0.1 (2014-12-09)
	* fix bug with Recent shortcut opening in a new Explorer window instead of navigating in the correct window
	* fix bug properly exit group load loop when an error occurs within an iteration

	Version: 4.0.0 (2014-12-07)
	* See all changes from v3.9.1 to v3.9.9 BETA
	* Korean language update

	Version: 3.9.9 BETA (2014-12-04)
	* detect if app is started in program files folder and set working dir to appdata
	* create a backup of the ini file at launch time
	* Dutch and Korean language updates
	
	Version: 3.9.8 BETA (2014-12-01)
	* fix lOptionsDisplayFoldersInExplorerMenu label.
	* add column break and system variable in default menu
	* fix bug when edit and save a submenu under the same name (from v3.3.2)
	* add double-quotes to Run command parameters
	* sort groups list in manage groups and edit group

	Version: 3.9.7 BETA (2014-11-25)
	* add an item in the right-click Tray menu to open the FoldersPopup.ini file
	* add an option to disable check for update at startup
	* add Downloads folder to Special Folders menu, support for DOpus and TC, not available on Win_XP
	* fix a bug making custom icons not following when favorites were moved up or down in the menu
	* merge and refactor GuiMoveFavoriteUp and GuiMoveFavoriteDown commands
	* fix a bug visible only to Total Commander users occuring when you left-click the tray icon button or when left-click on the tray icon was in the overflow area
	* add location URL of folders in groups saved to the ini file
	
	Version: 3.9.6 BETA (2014-11-21)
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
	
	Version: 3.9.5 BETA (2014-11-15)
	* display and select icon for folders, url and documents in add/edit favorite and in menu
	* better error management around menu icon assignment, fix the *.msc bug
	* use shell32.dll icon #1 for unknown icon
	* fix a bug in Group menu for network locations
	
	Version: 3.9.4 BETA (2014-11-09)
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
	
	Version: 3.9.3 BETA (2014-11-08)
	* retrieve language from ini file created by setup program and use when creating the FP ini file
	* accept space in Change hotkey dialog box to allow combinations with spacebar as a potential hotkey
	* detect TreeView folder select dialog box and exclude them because of a Windows limitation (Edit1 control not handling the Enter)
	* add the option OpenMenuOnTaskbar to open or not the popup menu over the taskbar (class Shell_TrayWnd)
	* add column breaks in menu
	* improve reliability and performance of group load with Explorer and DOpus
	* fix bug with windows move/resize when group load
	* fix bug with minimize/maximize Explorers when group load
	* fix wrong web link when an beta version update is available
	
	Version: 3.9.2 BETA (2014-11-05)
	* Addition of German language to Setup program
	* Add the possibility to overwrite an existing group of folders in the save group dialog box
	* allow to edit a group from the manage groups gui
	* Delete the startup shortcut when uninstall with Inno Setup
	* After installation with Inno Setup, copy an existing FoldersPopup.ini file if one exist in a previous protable installation (findable only if a shortcut to the portable installation exists)

	Version: 3.9.1 BETA (2014-11-02)
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


	Version: 3.3.2 (2014-12-01)
	* fix a bug occurring when editing a submenu and saving it under the same name
	
	Version: 3.3.1 (2014-11-17)
	* fix a bug occurring in some situation when a favorite location contains a comma
	
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
;@Ahk2Exe-SetDescription Folders Popup (freeware)
;@Ahk2Exe-SetVersion 5.2.3
;@Ahk2Exe-SetOrigFilename FoldersPopup.exe


;========================================================================================================================
; INITIALIZATION
;========================================================================================================================

#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off
DetectHiddenWindows, On
ComObjError(False) ; we will do our own error handling

; avoid error message when shortcut destination is missing
; see http://ahkscript.org/boards/viewtopic.php?f=5&t=4477&p=25239#p25236
DllCall("SetErrorMode", "uint", SEM_FAILCRITICALERRORS := 1)

; By default, the A_WorkingDir is A_ScriptDir.
; When the shortcut is created by Inno Setup, the working is set to the folder under {userappdata}.
; In portable mode, the user can set the working directory in his own Windows shortcut.
; If user enable "Run at startup", the "Start in:" shortcut option is set to the current A_WorkingDir.

; If A_WorkingDir equals A_ScriptDir and the file _do_not_remove_or_rename.txt is found in A_WorkingDir
; it means that FP has been installed with the setup program but that it was launched directly in the
; Program Files directory instead of using the Start menu or Startup shortcuts. In this situation, we
; know that the working directory has not been set properly. The following lines will fix it.
if (A_WorkingDir = A_ScriptDir) and FileExist(A_WorkingDir . "\_do_not_remove_or_rename.txt")
	SetWorkingDir, %A_AppData%\FoldersPopup

; Force A_WorkingDir to A_ScriptDir if uncomplied (development phase)
;@Ahk2Exe-IgnoreBegin
; Piece of code for development phase only - won't be compiled
; see http://fincs.ahk4.net/Ahk2ExeDirectives.htm
SetWorkingDir, %A_ScriptDir%
; to test user data directory: SetWorkingDir, %A_AppData%\FoldersPopup

ListLines, On
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd

OnExit, CleanUpBeforeExit ; must be positioned before InitFileInstall to ensure deletion of temporary files

Gosub, InitFileInstall
Gosub, InitLanguageVariables

global strAppName := "FoldersPopup"
global strCurrentVersion := "5.2.3" ; "major.minor.bugs" or "major.minor.beta.release"
global strCurrentBranch := "prod" ; "prod" or "beta", always lowercase for filename
global strAppVersion := "v" . strCurrentVersion . (strCurrentBranch = "beta" ? " " . strCurrentBranch : "")

global str32or64 := A_PtrSize * 8
global blnDiagMode := False
global strDiagFile := A_WorkingDir . "\" . strAppName . "-DIAG.txt"
global strIniFile := A_WorkingDir . "\" . strAppName . ".ini"
global strIniBackupFile := A_WorkingDir . "\" . strAppName . "-backup.ini"
global blnMenuReady := false
global arrSubmenuStack := Object()
global arrSubmenuStackPosition := Object()
global objIconsFile := Object()
global objIconsIndex := Object()
global strHotkeyNoneModifiers := ">^!+#" ; right-control/atl/shift/windows impossible keys combination
global strHotkeyNoneKey := "9"
global strColumnBreakIndicator := "==="

if InStr(A_ScriptDir, A_Temp) ; must be positioned after strAppName is created
; if the app runs from a zip file, the script directory is created under the system Temp folder
{
	Oops(lOopsZipFileError, strAppName)
	ExitApp
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
Gosub, InitGuiControls
Gosub, LoadIniFile
; must be after LoadIniFile
IniWrite, %strCurrentVersion%, %strIniFile%, Global, % "LastVersionUsed" . (strCurrentBranch = "beta" ? "Beta" : "Prod")

if (blnDiagMode)
	Gosub, InitDiagMode
if (blnUseColors)
	Gosub, LoadTheme

; build even if blnDisplaySpecialFolders or blnDisplaySwitchMenu are false because they could become true
; no need to build Recent folders menu at startup since this menu is refreshed on demand
if (A_OSVersion = "WIN_XP")
	Gosub, BuildSpecialFoldersMenu
Gosub, BuildFoldersInExplorerMenu ; need to be initialized here - will be updated at each call to popup menu
Gosub, BuildGroupMenu
Gosub, BuildClipboardMenu
Gosub, BuildMainMenu
Gosub, BuildGui
Gosub, BuildTrayMenu

if (blnCheck4Update)
	Gosub, Check4Update

IfExist, %A_Startup%\%strAppName%.lnk ; update the shortcut in case the exe filename changed
{
	FileDelete, %A_Startup%\%strAppName%.lnk
	FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\%strAppName%.lnk, %A_WorkingDir%
	Menu, Tray, Check, %lMenuRunAtStartup%
}

if (blnDisplayTrayTip)
{
	TrayTip, % L(lTrayTipInstalledTitle, strAppName)
		, % L(lTrayTipInstalledDetail, strAppName
			, Hotkey2Text(strModifiers1, strMouseButton1, strOptionsKey1)
			, Hotkey2Text(strModifiers3, strMouseButton3, strOptionsKey3)
			, Hotkey2Text(strModifiers2, strMouseButton2, strOptionsKey2)
			, Hotkey2Text(strModifiers4, strMouseButton4, strOptionsKey4))
			; ~4~ and ~5~ not used in new EN text but keep them for compatibility with untranslated texts
		, , 17 ; 1 info icon + 16 no sound)
	Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)
}

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
; Gosub, BuildFoldersInExplorerMenu
; Gosub, BuildGroupMenu
; Gosub, BuildClipboardMenu
; Gosub, GuiGroupSaveFromMenu
; Gosub, GuiGroupsManage
; Gosub, FoldersInExplorerMenuShortcut
; Gosub, PopupMenuCopyLocation

return


/*
REMOVED IN v4.2.1 BECAUSE OF A SIDE EFFECT IN XL 2010
; prevent double-click on some static control to overwrite the clipboard with the image URL (a windows "undesired feature")
; see http://www.autohotkey.com/board/topic/94962-doubleclick-on-gui-pictures-puts-their-path-in-your-clipboard/
OnClipboardChange:
If A_EventInfo
  ClipboardAllBK := ClipboardAll
return
*/


;========================================================================================================================
; END OF INITIALIZATION
;========================================================================================================================


;========================================================================================================================
; HOTKEYS
;========================================================================================================================

; Gui Hotkeys
#If WinActive("ahk_id " . strAppHwnd)

^Up::
if (LV_GetCount("Selected") > 1)
	Gosub, GuiMoveMultipleFavoritesUp
else
	Gosub, GuiMoveFavoriteUp
return

^Down::
if (LV_GetCount("Selected") > 1)
	Gosub, GuiMoveMultipleFavoritesDown
else
	Gosub, GuiMoveFavoriteDown
return

^Right::
Gosub, HotkeyChangeMenu
return

^Left::
GuiControlGet, blnUpMenuVisible, Visible, picUpMenu
if (blnUpMenuVisible)
	Gosub, GuiGotoPreviousMenu
return

^A::
LV_Modify(0, "Select")
return

^N::
Gosub, GuiAddFavorite
return

Enter::
if (LV_GetCount("Selected") > 1)
	Gosub, GuiMoveMultipleFavorites
else
	Gosub, GuiEditFavorite
return

Del::
if (LV_GetCount("Selected") > 1)
	Gosub, GuiRemoveMultipleFavorites
else
	Gosub, GuiRemoveFavorite
return

#If
; End of Gui Hotkeys


;========================================================================================================================
; END OF INITIALIZATION
;========================================================================================================================


;-----------------------------------------------------------
InitLanguageVariables:
;-----------------------------------------------------------

#Include %A_ScriptDir%\FoldersPopup_LANG.ahk

return
;-----------------------------------------------------------


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
FileInstall, FileInstall\FoldersPopup_LANG_IT.txt, %strTempDir%\FoldersPopup_LANG_IT.txt, 1
FileInstall, FileInstall\FoldersPopup_LANG_ES.txt, %strTempDir%\FoldersPopup_LANG_ES.txt, 1
FileInstall, FileInstall\FoldersPopup_LANG_PT-BR.txt, %strTempDir%\FoldersPopup_LANG_PT-BR.txt, 1

FileInstall, FileInstall\default_browser_icon.html, %strTempDir%\default_browser_icon.html, 1

FileInstall, FileInstall\about-32.png, %strTempDir%\about-32.png
FileInstall, FileInstall\add_property-48.png, %strTempDir%\add_property-48.png
FileInstall, FileInstall\delete_property-48.png, %strTempDir%\delete_property-48.png
FileInstall, FileInstall\channel_mosaic-48.png, %strTempDir%\channel_mosaic-48.png
FileInstall, FileInstall\separator-26.png, %strTempDir%\separator-26.png
FileInstall, FileInstall\column-26.png, %strTempDir%\column-26.png
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
strIniKeyNames := "PopupHotkeyMouse|PopupHotkeyNewMouse|PopupHotkeyKeyboard|PopupHotkeyNewKeyboard|SettingsHotkey|FoldersInExplorerHotkey|GroupsHotkey|RecentsHotkey|ClipboardHotkey|CopyLocationHotkey"
StringSplit, arrIniKeyNames, strIniKeyNames, |
strHotkeyDefaults := "MButton|+MButton|#a|+#a|+^s|+^f|+^g|+^r|+^c|+^v"
StringSplit, arrHotkeyDefaults, strHotkeyDefaults, |
strHotkeyLabels := "PopupMenuMouse|PopupMenuNewWindowMouse|PopupMenuKeyboard|PopupMenuNewWindowKeyboard|GuiShow|FoldersInExplorerMenuShortcut|GroupsMenuShortcut|RecentFoldersShortcut|ClipboardMenuShortcut|PopupMenuCopyLocation"
StringSplit, arrHotkeyLabels, strHotkeyLabels, |

strMouseButtons := "None|LButton|MButton|RButton|XButton1|XButton2|WheelUp|WheelDown|WheelLeft|WheelRight|"
; leave last | to enable default value on the last item
StringSplit, arrMouseButtons, strMouseButtons, |

strIconsMenus := "lMenuDesktop|lMenuDocuments|lMenuPictures|lMenuMyComputer|lMenuNetworkNeighborhood|lMenuControlPanel|lMenuRecycleBin"
	. "|menuRecentFolders|menuGroupDialog|menuGroupExplorer|lMenuSpecialFolders|lMenuGroup|lMenuFoldersInExplorer"
	. "|lMenuRecentFolders|lMenuSettings|lMenuAddThisFolder|lDonateMenu|Submenu|Network|UnknownDocument|Folder"
	. "|menuGroupSave|menuGroupLoad|lMenuDownloads|Templates|MyMusic|MyVideo|History|Favorites|Temporary|Winver"
	. "|Fonts|Application|Clipboard"

if (A_OSVersion = "WIN_XP")
{
	strIconsFile := "shell32|shell32|shell32|shell32|shell32|shell32|shell32"
				. "|shell32|shell32|shell32|shell32|shell32|shell32"
				. "|shell32|shell32|shell32|shell32|shell32|shell32|shell32|shell32"
				. "|shell32|shell32|shell32|shell32|shell32|shell32|shell32|shell32|shell32|winver"
				. "|shell32|shell32|shell32"
	strIconsIndex := "35|127|118|16|19|22|33"
				. "|4|147|4|4|147|147"
				. "|214|166|111|161|85|10|1|4"
				. "|7|7|156|55|129|130|21|209|132|1"
				. "|39|97|145"
}
else if (A_OSVersion = "WIN_VISTA")
{
	strIconsFile := "imageres|imageres|imageres|imageres|imageres|imageres|imageres"
				. "|imageres|imageres|imageres|imageres|shell32|imageres"
				. "|imageres|imageres|shell32|shell32|shell32|imageres|shell32|shell32"
				. "|shell32|shell32|shell32|shell32|imageres|imageres|shell32|shell32|shell32|winver"
				. "|shell32|shell32|shell32"
	strIconsIndex := "105|85|67|104|114|22|49"
				. "|112|174|3|3|251|174"
				. "|112|109|88|161|85|28|1|3"
				. "|259|259|123|55|103|177|240|87|153|1"
				. "|39|97|261"
}
else if (A_OSVersion = "WIN_7")
{
	strIconsFile := "imageres|imageres|imageres|imageres|imageres|imageres|imageres"
				. "|imageres|imageres|imageres|imageres|shell32|imageres"
				. "|imageres|imageres|imageres|imageres|shell32|imageres|shell32|shell32"
				. "|shell32|shell32|imageres|shell32|imageres|imageres|shell32|shell32|shell32|winver"
				. "|shell32|shell32|shell32"
	strIconsIndex := "106|189|68|105|115|23|50"
				. "|113|176|203|203|99|176"
				. "|113|110|217|208|298|29|1|4"
				. "|297|46|176|55|104|179|240|87|153|1"
				. "|39|304|261"
}
else ; Windows 8.1 / 10 (tested and more (not tested)
{
	strIconsFile := "imageres|imageres|imageres|imageres|imageres|imageres|imageres"
				. "|imageres|imageres|imageres|imageres|shell32|imageres"
				. "|imageres|imageres|imageres|imageres|shell32|imageres|shell32|shell32"
				. "|shell32|shell32|imageres|shell32|imageres|imageres|shell32|shell32|shell32|winver"
				. "|shell32|shell32|shell32"
	strIconsIndex := "106|189|68|105|115|23|50"
				. "|113|176|204|204|99|176"
				. "|113|110|307|209|300|29|1|4"
				. "|299|46|176|55|104|179|240|87|153|1"
				. "|39|324|261"
}

StringSplit, arrIconsFile, strIconsFile, |
StringSplit, arrIconsIndex, strIconsIndex, |

Loop, Parse, strIconsMenus, |
{
	objIconsFile[A_LoopField] := A_WinDir . "\System32\" . arrIconsFile%A_Index% . (arrIconsFile%A_Index% = "winver" ? ".exe" : ".dll")
	objIconsIndex[A_LoopField] := arrIconsIndex%A_Index%
}
; example: objIconsFile["lMenuPictures"] and objIconsIndex["lMenuPictures"]

return
;-----------------------------------------------------------


;-----------------------------------------------------------
LoadIniFile:
;-----------------------------------------------------------

; create a backup of the ini file before loading
FileCopy, %strIniFile%, %strIniBackupFile%, 1

; reinit after Settings save if already exist
Global arrMenus := Object() ; list of menus (Main and submenus), non-hierachical
arrMainMenu := Object() ; array of favorites of the Main menu
arrMenus.Insert(lMainMenuName, arrMainMenu) ; lMainMenuName is used in the objects but not saved/restores in the ini file

IfNotExist, %strIniFile%
{
	strPopupHotkeyMouseDefault := arrHotkeyDefaults1 ; "MButton"
	strPopupHotkeyMouseNewDefault := arrHotkeyDefaults2 ; "+MButton"
	strPopupHotkeyKeyboardDefault := arrHotkeyDefaults3 ; "#a"
	strPopupHotkeyKeyboardNewDefault := arrHotkeyDefaults4 ; "+#a"
	strSettingsHotkeyDefault := arrHotkeyDefaults5 ; "+^s"
	strFoldersInExplorerHotkeyDefault := arrHotkeyDefaults6 ; "+^f"
	strGroupsHotkeyDefault := arrHotkeyDefaults7 ; "+^g"
	strRecentsHotkeyDefault := arrHotkeyDefaults8 ; "+^r"
	strClipboardHotkeyDefault := arrHotkeyDefaults9 ; "+^c"
	
	intIconSize := (A_OSVersion = "WIN_XP" ? 16 : 24)
	
	FileAppend,
		(LTrim Join`r`n
			[Global]
			PopupHotkeyMouse=%strPopupHotkeyMouseDefault%
			PopupHotkeyNewMouse=%strPopupHotkeyMouseNewDefault%
			PopupHotkeyKeyboard=%strPopupHotkeyKeyboardDefault%
			PopupHotkeyNewKeyboard=%strPopupHotkeyKeyboardNewDefault%
			SettingsHotkey=%strSettingsHotkeyDefault%
			FoldersInExplorerHotkey=%strFoldersInExplorerHotkeyDefault%
			GroupsHotkey=%strGroupsHotkeyDefault%
			RecentsHotkey=%strRecentsHotkeyDefault%
			ClipboardHotkey=%strClipboardHotkeyDefault%
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
			LanguageCode=%strLanguageCode%
			DirectoryOpusPath=
			IconSize=%intIconSize%
			OpenMenuOnTaskbar=1
			[Folders]
			Folder1=C:\|C:\
			Folder2=Windows|%A_WinDir%
			Folder3=Program Files|%A_ProgramFiles%
			Folder4=User Profile|`%USERPROFILE`%


)
		, %strIniFile%
}

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
IniRead, blnDisplaySpecialMenusShortcuts, %strIniFile%, Global, DisplaySpecialMenusShortcuts, 1
IniRead, blnDisplaySpecialFolders, %strIniFile%, Global, DisplaySpecialFolders, 1
IniRead, blnDisplayRecentFolders, %strIniFile%, Global, DisplayRecentFolders, 1
IniRead, blnDisplayFoldersInExplorerMenu, %strIniFile%, Global, DisplayFoldersInExplorerMenu, 1
IniRead, blnDisplayGroupMenu, %strIniFile%, Global, DisplaySwitchMenu, 1 ; keep "Switch" in label instead of "Group" for backward compatibility
IniRead, blnDisplayClipboardMenu, %strIniFile%, Global, DisplayClipboardMenu, 1
IniRead, blnDisplayCopyLocationMenu, %strIniFile%, Global, DisplayCopyLocationMenu, 1
IniRead, intPopupMenuPosition, %strIniFile%, Global, PopupMenuPosition, 1
IniRead, strPopupFixPosition, %strIniFile%, Global, PopupFixPosition, 20,20
StringSplit, arrPopupFixPosition, strPopupFixPosition, `,
IniRead, blnDisplayMenuShortcuts, %strIniFile%, Global, DisplayMenuShortcuts, 0
IniRead, strLatestSkippedProd, %strIniFile%, Global, LatestVersionSkipped, 0.0 ; do not add "Prod" to ini variable for backward compatibility
IniRead, strLatestSkippedBeta, %strIniFile%, Global, LatestVersionSkippedBeta, 0.0
IniRead, strLatestUsedProd, %strIniFile%, Global, LastVersionUsedProd, 0.0
IniRead, strLatestUsedBeta, %strIniFile%, Global, LastVersionUsedBeta, 0.0
IniRead, blnDiagMode, %strIniFile%, Global, DiagMode, 0
IniRead, intStartups, %strIniFile%, Global, Startups, 1
IniRead, intRecentFolders, %strIniFile%, Global, RecentFolders, 10
IniRead, intIconSize, %strIniFile%, Global, IconSize, 24
IniRead, strGroups, %strIniFile%, Global, Groups, %A_Space% ; empty string if not found
IniRead, blnCheck4Update, %strIniFile%, Global, Check4Update, 1
IniRead, blnOpenMenuOnTaskbar, %strIniFile%, Global, OpenMenuOnTaskbar, 1
IniRead, blnRememberSettingsPosition, %strIniFile%, Global, RememberSettingsPosition, 1

IniRead, strSettingsPosition, %strIniFile%, Global, SettingsPosition, -1 ; center at minimal size
StringSplit, arrSettingsPosition, strSettingsPosition, |

IniRead, blnDonor, %strIniFile%, Global, Donor, 0 ; Please, be fair. Don't cheat with this.

IniRead, strDirectoryOpusPath, %strIniFile%, Global, DirectoryOpusPath, %A_Space% ; empty string if not found
IniRead, blnDirectoryOpusUseTabs, %strIniFile%, Global, DirectoryOpusUseTabs, 1 ; use tabs by default
if StrLen(strDirectoryOpusPath)
{
	blnUseDirectoryOpus := FileExist(strDirectoryOpusPath)
	if (blnUseDirectoryOpus)
		Gosub, SetDOpusRt
	else
		if (strDirectoryOpusPath <> "NO")
			Oops(lOopsWrongThirdPartyPath, "Directory Opus", strDirectoryOpusPath, lOptionsThirdParty)
}
else
	if (strDirectoryOpusPath <> "NO")
		Gosub, CheckDirectoryOpus

IniRead, strTotalCommanderPath, %strIniFile%, Global, TotalCommanderPath, %A_Space% ; empty string if not found
IniRead, blnTotalCommanderUseTabs, %strIniFile%, Global, TotalCommanderUseTabs, 1 ; use tabs by default
if StrLen(strTotalCommanderPath)
{
	blnUseTotalCommander := FileExist(strTotalCommanderPath)
	if (blnUseTotalCommander)
		Gosub, SetTCCommand
	else
		if (strTotalCommanderPath <> "NO")
			Oops(lOopsWrongThirdPartyPath, "Total Commander", strTotalCommanderPath, lOptionsThirdParty)
}
else
	if (strTotalCommanderPath <> "NO")
		Gosub, CheckTotalCommander

IniRead, strFPconnectPath, %strIniFile%, Global, FPconnectPath, %A_Space% ; empty string if not found
if StrLen(strFPconnectPath)
{
	blnUseFPconnect := FileExist(strFPconnectPath)
	if (blnUseFPconnect)
		Gosub, SetFPconnect
	else
		if (strFPconnectPath <> "NO")
			Oops(lOopsWrongThirdPartyPath, "FPconnect", strFPconnectPath, lOptionsThirdParty)
}
else
	if (strFPconnectPath <> "NO")
		Gosub, CheckFPconnect

IniRead, strTheme, %strIniFile%, Global, Theme
if (strTheme = "ERROR") ; if Theme not found, we have a v1 or v2 ini file - add the themes to the ini file
{
	strTheme := "Windows"
	strAvailableThemes := "Windows|Grey|Light Blue|Light Green|Light Red|Yellow"
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
{
	IniRead, strAvailableThemes, %strIniFile%, Global, AvailableThemes
	if !InStr(strAvailableThemes, "Windows|")
	{
		strAvailableThemes := "Windows|" . strAvailableThemes
		IniWrite, %strAvailableThemes%, %strIniFile%, Global, AvailableThemes
	}
}
blnUseColors := (strTheme <> "Windows")
	
IniRead, blnMySystemFoldersBuilt, %strIniFile%, Global, MySystemFoldersBuilt, 0 ; default false
if !(blnMySystemFoldersBuilt) and (A_OSVersion <> "WIN_XP")
	Gosub, AddToIniMySystemFoldersMenu ; modify the ini file Folders section before reading it

Loop
{
	IniRead, strIniLine, %strIniFile%, Folders, Folder%A_Index% ; keep "Folders" label instead of "Favorite" for backward compatibility
	if (strIniLine = "ERROR")
		Break
	strIniLine := strIniLine . "|||" ; additional "|" to make sure we have all empty items
	; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
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

	objFavorite.IconResource := arrThisFavorite6 ; icon resource in format "iconfile,iconindex"
	objFavorite.AppArguments := arrThisFavorite7 ; application arguments
	objFavorite.AppWorkingDir := arrThisFavorite8 ; application working directory

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
	ExitApp
}

return
;------------------------------------------------------------


;------------------------------------------------------------
AddToIniMySystemFoldersMenu:
;------------------------------------------------------------

strInstance := ""
Loop
{
	IniRead, strIniLine, %strIniFile%, Folders, Folder%A_Index% ; keep "Folders" label instead of "Favorite" for backward compatibility
	if InStr(strIniLine, lMenuMySystemMenu . strInstance)
		strInstance := strInstance . "+"
	if (strIniLine = "ERROR")
	{
		intNextFolderNumber := A_Index
		Break
	}
}
strMySystemMenu := lMenuMySystemMenu . strInstance

AddToIniOneSystemFolderMenu(intNextFolderNumber + 0, lMenuSeparator, lMenuSeparator . lMenuSeparator, , , "F")
AddToIniOneSystemFolderMenu(intNextFolderNumber + 1, strMySystemMenu, lGuiSubmenuSeparator, , lGuiSubmenuSeparator . strMySystemMenu, "S")
AddToIniOneSystemFolderMenu(intNextFolderNumber + 2, lMenuDesktop, A_Desktop, lGuiSubmenuSeparator . strMySystemMenu)
AddToIniOneSystemFolderMenu(intNextFolderNumber + 3, , "{450D8FBA-AD25-11D0-98A8-0800361B1103}", lGuiSubmenuSeparator . strMySystemMenu)
AddToIniOneSystemFolderMenu(intNextFolderNumber + 4, , strMyPicturesPath, lGuiSubmenuSeparator . strMySystemMenu)
AddToIniOneSystemFolderMenu(intNextFolderNumber + 5, , strDownloadPath, lGuiSubmenuSeparator . strMySystemMenu)
AddToIniOneSystemFolderMenu(intNextFolderNumber + 6, lMenuSeparator, lMenuSeparator . lMenuSeparator, lGuiSubmenuSeparator . strMySystemMenu, , "F")
AddToIniOneSystemFolderMenu(intNextFolderNumber + 7, , "{20D04FE0-3AEA-1069-A2D8-08002B30309D}", lGuiSubmenuSeparator . strMySystemMenu)
AddToIniOneSystemFolderMenu(intNextFolderNumber + 8, , "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", lGuiSubmenuSeparator . strMySystemMenu)
AddToIniOneSystemFolderMenu(intNextFolderNumber + 9, lMenuSeparator, lMenuSeparator . lMenuSeparator, lGuiSubmenuSeparator . strMySystemMenu, , "F")
AddToIniOneSystemFolderMenu(intNextFolderNumber + 10, , "{21EC2020-3AEA-1069-A2DD-08002B30309D}", lGuiSubmenuSeparator . strMySystemMenu)
AddToIniOneSystemFolderMenu(intNextFolderNumber + 11, , "{645FF040-5081-101B-9F08-00AA002F954E}", lGuiSubmenuSeparator . strMySystemMenu)

IniWrite, 1, %strIniFile%, Global, MySystemFoldersBuilt

return
;------------------------------------------------------------


;------------------------------------------------------------
AddToIniOneSystemFolderMenu(intNumber, strSpecialFolderName := "", strSpecialFolderLocation := "", strMenuName := "", strSubmenuFullName := "", strFavoriteType := "P")
; 0 Folder number, 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir)
;------------------------------------------------------------
{
	if (strFavoriteType = "S")
		strIconResource := objIconsFile["lMenuSpecialFolders"] . "," . objIconsIndex["lMenuSpecialFolders"]
	else
		strIconResource := objSpecialFolders[strSpecialFolderLocation].DefaultIcon
	if !StrLen(strSpecialFolderName)
		strSpecialFolderName := objSpecialFolders[strSpecialFolderLocation].DefaultName

	strNewIniLine := strSpecialFolderName . "|" . strSpecialFolderLocation . "|" . strMenuName . "|" . strSubmenuFullName . "|" . strFavoriteType . "|" . strIconResource
	
	IniWrite, %strNewIniLine%, %strIniFile%, Folders, Folder%intNumber% ; keep "Folders" label instead of "Favorite" for backward compatibility
}
;------------------------------------------------------------


;-----------------------------------------------------------
LoadIniHotkeys:
;-----------------------------------------------------------
; Read the values and set hotkey shortcuts
loop, % arrIniKeyNames%0%
{
	; Prepare global arrays used by SplitHotkey function
	IniRead, arrHotkeys%A_Index%, %strIniFile%, Global, % arrIniKeyNames%A_Index%, % arrHotkeyDefaults%A_Index%
	SplitHotkey(arrHotkeys%A_Index%, strMouseButtons
		, strModifiers%A_Index%, strOptionsKey%A_Index%, strMouseButton%A_Index%, strMouseButtonsWithDefault%A_Index%)
	; example: Hotkey, $MButton, PopupMenuMouse
	if (arrHotkeys%A_Index% = "None") ; do not compare with lOptionsMouseNone because it is translated
		Hotkey, % "$" . strHotkeyNoneModifiers . strHotkeyNoneKey, % arrHotkeyLabels%A_Index%, On UseErrorLevel
	else
		Hotkey, % "$" . arrHotkeys%A_Index%, % arrHotkeyLabels%A_Index%, On UseErrorLevel
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

IfNotExist, %strIniFile%
	; read language code from ini file created by the Inno Setup script in the user data folder
	IniRead, strLanguageCode, % A_WorkingDir . "\" . strAppName . "-setup.ini", Global , LanguageCode, EN
else
	IniRead, strLanguageCode, %strIniFile%, Global, LanguageCode, EN

strLanguageFile := strTempDir . "\" . strAppName . "_LANG_" . strLanguageCode . ".txt"

if FileExist(strLanguageFile)
{
	FileRead, strLanguageStrings, %strLanguageFile%
	Loop, Parse, strLanguageStrings, `n, `r
	{
		if (SubStr(A_LoopField, 1, 1) <> ";")
		{
			StringSplit, arrLanguageBit, A_LoopField, `t
			if SubStr(arrLanguageBit1, 1, 1) = "l"
				%arrLanguageBit1% := arrLanguageBit2
			StringReplace, %arrLanguageBit1%, %arrLanguageBit1%, ``n, `n, All
			
			if InStr(%arrLanguageBit1%, ";;")
				StringReplace, %arrLanguageBit1%, %arrLanguageBit1%, % ";;", % "!r4nd0mt3xt!"
			if InStr(%arrLanguageBit1%, ";")
				%arrLanguageBit1% := Trim(SubStr(%arrLanguageBit1%, 1, InStr(%arrLanguageBit1%, ";") - 1)) ; trim from ; and trim spaces and tabs
			if InStr(%arrLanguageBit1%, "!r4nd0mt3xt!")
				StringReplace, %arrLanguageBit1%, %arrLanguageBit1%, % "!r4nd0mt3xt!", % ";"
		}
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
strOptionsLanguageCodes := "EN|FR|DE|NL|KO|SV|IT|ES|PT-BR"
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

; Shell numeric Constants
; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx

; Shell Commands:
; http://www.sevenforums.com/tutorials/4941-shell-command.html
; http://www.eightforums.com/tutorials/6050-shell-commands-windows-8-a.html

; Environment system variables
; http://en.wikipedia.org/wiki/Environment_variable#Windows

global objClassIdOrPathByDefaultName := Object() ; also used by CollectExplorers
global objSpecialFolders := Object()
global strSpecialFoldersList

; InitSpecialFolderObject(strClassIdOrPath, strShellConstant, intShellConstant, strAHKConstant, strDOpusAlias, strTCCommand
;	, strDefaultName, strDefaultIcon
;	, strUse4NavigateExplorer, strUse4NewExplorer, strUse4Dialog, strUse4Console, strUse4DOpus, strUse4TC, strUse4FPc)

; 		CLS: Class ID
;		SCT: Shell Constant Text
;		SCN: Shell Constant Numeric
;		DOA: Directory Opus Alias
;		TCC: Total Commander Commands
;		NEW: Open in new Explorer anyway
/*
NOTES
- Total Commander commands: cm_OpenDesktop (2121), cm_OpenDrives (2122), cm_OpenControls (2123), cm_OpenFonts (2124), cm_OpenNetwork (2125), cm_OpenPrinters (2126), cm_OpenRecycled (2127)
- DOpus see http://resource.dopus.com/viewtopic.php?f=3&t=23691
*/

;---------------------
; CLSID giving localized name and icon, with valid Shell Command

InitSpecialFolderObject("{D20EA4E1-3957-11d2-A40B-0C5020524153}", "Common Administrative Tools", -1, "", "commonadmintools", ""
	, "Administrative Tools", "" ; Outils d’administration
	, "CLS", "CLS", "NEW", "NEW", "DOA", "NEW", "NEW")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "MyComputerFolder", 17, "", "mycomputer", 2122
	, "Computer", "" ; Ordinateur
	, "SCT", "SCT", "SCT", "NEW", "DOA", "TCC", "NEW") ; for 1,2,3 CLS works, 7 OK for FPc but CLS does not work with DoubleCommander
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{21EC2020-3AEA-1069-A2DD-08002B30309D}", "ControlPanelFolder", 3, "", "controls", 2123
	, "Control Panel (Icons view)", "" ; Tous les Panneaux de configuration
	, "SCT", "SCT", "NEW", "NEW", "DOA", "CLS", "NEW")
	; OK     OK      OK     OK    OK  NO-use NEW
InitSpecialFolderObject("{450D8FBA-AD25-11D0-98A8-0800361B1103}", "Personal", 5, "A_MyDocuments", "mydocuments", ""
	, "Documents", "" ; Mes documents
	, "SCT", "SCT", "AHK", "AHK", "DOA", "AHK", "AHK")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{ED228FDF-9EA8-4870-83b1-96b02CFE0D52}", "Games", -1, "", "", ""
	, "Games Explorer", "" ; Jeux
	, "SCT", "SCT", "NEW", "NEW", "NEW", "CLS", "NEW")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{B4FB3F98-C1EA-428d-A78A-D1F5659CBA93}", "HomeGroupFolder", -1, "", "", ""
	, "HomeGroup", "" ; Groupe résidentiel
	, "SCT", "SCT", "SCT", "NEW", "NEW", "CLS", "NEW")
	; OK     OK      OK     OK    OK     OK
InitSpecialFolderObject("{031E4825-7B94-4dc3-B131-E946B44C8DD5}", "Libraries", -1, "", "libraries", ""
	, "Libraries", "" ; Bibliothèque
	, "SCT", "SCT", "SCT", "NEW", "DOA", "CLS", "NEW")
	; OK     OK      OK     OK     OK      OK
InitSpecialFolderObject("{7007ACC7-3202-11D1-AAD2-00805FC1270E}", "ConnectionsFolder", -1, "", "", ""
	, "Network Connections", "" ; Connexions réseau
	, "SCT", "SCT", "NEW", "NEW", "NEW", "CLS", "NEW")
	; OK     OK      OK     OK     OK      OK
InitSpecialFolderObject("{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", "NetworkPlacesFolder", 18, "", "network", 2125
	, "Network", "" ; Réseau
	, "SCT", "SCT", "SCT", "NEW", "DOA", "TCC", "NEW")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{2227A280-3AEA-1069-A2DE-08002B30309D}", "PrintersFolder", -1, "", "printers", 2126
	, "Printers and Faxes", "" ; Imprimantes
	, "SCT", "SCT", "NEW", "NEW", "DOA", "TCC", "NEW")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{645FF040-5081-101B-9F08-00AA002F954E}", "RecycleBinFolder", 0, "", "trash", 2127
	, "Recycle Bin", "" ; Corbeille
	, "SCT", "SCT", "NEW", "NEW", "DOA", "TCC", "NEW")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{59031a47-3f72-44a7-89c5-5595fe6b30ee}", "Profile", -1, "", "profile", ""
	, lMenuUserFolder, "" ; Dossier de l'utilisateur
	, "SCT", "SCT", "SCT", "NEW", "DOA", "CLS", "NEW")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{1f3427c8-5c10-4210-aa03-2ee45287d668}", "User Pinned", -1, "", "", ""
	, lMenuUserPinned, "" ; Epinglé par l'utilisateur
	, "SCT", "SCT", "SCT", "NEW", "NEW", "NEW", "NEW")
	; OK     OK      OK     OK    OK      OK
InitSpecialFolderObject("{BD84B380-8CA2-1069-AB1D-08000948534}", "Fonts", -1, "", "fonts", 2124
	, lMenuFonts, "Fonts"
	, "SCT", "SCT", "NEW", "NEW", "DOA", "TCC", "NEW")
	; OK     OK      OK     OK    OK      OK

;---------------------
; CLSID giving localized name and icon, no valid Shell Command, must be open in a new Explorer using CLSID - to be tested with DOpus, TC and FPc

InitSpecialFolderObject("{B98A2BEA-7D42-4558-8BD1-832F41BAC6FD}", "", -1, "", "", ""
	, "Backup and Restore", "" ; Sauvegarder et restaurer
	, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
InitSpecialFolderObject("{ED7BA470-8E54-465E-825C-99712043E01C}", "", -1, "", "", ""
	, "Control Panel (All Tasks)", "" ; Toutes les tâches
	, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
InitSpecialFolderObject("{323CA680-C24D-4099-B94D-446DD2D7249E}", "", -1, "", "favorites", ""
	, "Favorites", "" ; Favoris (<> Favorites (Internet))
	, "CLS", "CLS", "CLS", "NEW", "DOA", "NEW", "NEW")
if (GetOsVersion() <> "WIN_10")
	InitSpecialFolderObject("{3080F90E-D7AD-11D9-BD98-0000947B0257}", "", -1, "", "", ""
		, "Flip 3D", "" ; Pas de traduction
		, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
InitSpecialFolderObject("{6DFD7C5C-2451-11d3-A299-00C04F8EF6AF}", "", -1, "", "", ""
	, "Folder Options", "" ; Options des dossiers
	, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
if (A_OSVersion = "WIN_7") ; Performance Information and Tool not available on Win8+
	InitSpecialFolderObject("{78F3955E-3B90-4184-BD14-5397C15F1EFC}", "", -1, "", "", ""
		, "Performance Information and Tools", "" ; Informations et outils de performance
		, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
InitSpecialFolderObject("{35786D3C-B075-49b9-88DD-029876E11C01}", "", -1, "", "", ""
	, "Portable Devices", "" ; Appareils mobiles
	, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
InitSpecialFolderObject(A_Programs, "", -1, "A_Programs", "programs", "" ; CLSID was "{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}" but not working under Win 10
	, lMenuProgramsFolderStartMenu, "" ; Menu Démarrer / Programmes (Menu Start/Programs)
	, "CLS", "CLS", "CLS", "CLS", "DOA", "AHK", "AHK")
InitSpecialFolderObject("{22877a6d-37a1-461a-91b0-dbda5aaebc99}", "", -1, "", "", ""
	, "Recent Places", "" ; Emplacements récents
	, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
InitSpecialFolderObject("{3080F90D-D7AD-11D9-BD98-0000947B0257}", "", -1, "", "", ""
	, "Show Desktop", "" ; Afficher le Bureau
	, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")
InitSpecialFolderObject("{BB06C0E4-D293-4f75-8A90-CB05B6477EEE}", "", -1, "", "", ""
	, "System", "" ; Système
	, "CLS", "CLS", "NEW", "NEW", "NEW", "NEW", "NEW")

;---------------------
; Path from registry (no CLSID), localized name and icon provided, no Shell Command - to be tested with DOpus, TC and FPc

RegRead, strDownloadPath, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, {374DE290-123F-4565-9164-39C4925E467B}
InitSpecialFolderObject(strDownloadPath, "", -1, "", "downloads", ""
	, lMenuDownloads, "lMenuDownloads"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
RegRead, strException, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Music
InitSpecialFolderObject(strException, "", -1, "", "mymusic", ""
	, lMenuMyMusic, "MyMusic"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
RegRead, strException, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Video
InitSpecialFolderObject(strException, "", -1, "", "myvideos", ""
	, lMenuMyVideo, "MyVideo"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
RegRead, strException, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, Templates
InitSpecialFolderObject(strException, "", -1, "", "templates", ""
	, lMenuTemplates, "Templates"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
RegRead, strMyPicturesPath, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Pictures
InitSpecialFolderObject(strMyPicturesPath, "", 39, "", "mypictures", ""
	, lMenuPictures, "lMenuPictures"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
RegRead, strException, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, Favorites ; Favorites (Internet)
InitSpecialFolderObject(strException, "", -1, "", "", ""
	, lMenuFavoritesInternet, "Favorites"
	, "CLS", "CLS", "CLS", "CLS", "CLS", "CLS", "CLS")

;---------------------
; Path under %APPDATA% (no CLSID), localized name and icon provided, no Shell Command - to be tested with DOpus, TC and FPc

InitSpecialFolderObject("%APPDATA%\Microsoft\Windows\Start Menu", "", -1, "A_StartMenu", "start", ""
	, lMenuStartMenu, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup", "", -1, "A_Startup", "startup", ""
	, lMenuStartup, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%APPDATA%", "", -1, "A_AppData", "appdata", ""
	, lMenuAppData, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%APPDATA%\Microsoft\Windows\Recent", "", -1, "", "recent", ""
	, lMenuRecentItems, "menuRecentFolders"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
if (GetOsVersion() = "WIN_10")
	InitSpecialFolderObject("%LocalAppData%\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\Cookies", "", -1, "", "cookies", ""
		, lMenuCookies, "Folder"
		, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
else
	InitSpecialFolderObject("%APPDATA%\Microsoft\Windows\Cookies", "", -1, "", "cookies", ""
		, lMenuCookies, "Folder"
		, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%APPDATA%\Microsoft\Internet Explorer\Quick Launch", "", -1, "", "", ""
	, lMenuQuickLaunch, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "CLS", "CLS", "CLS")
InitSpecialFolderObject("%APPDATA%\Microsoft\SystemCertificates", "", -1, "", "", ""
	, lMenuSystemCertificates, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "CLS", "CLS", "CLS")

;---------------------
; Path under other environment variables (no CLSID), localized name and icon provided, no Shell Command - to be tested with DOpus, TC and FPc

InitSpecialFolderObject("%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu", "", -1, "A_StartMenuCommon", "commonstartmenu", ""
	, lMenuCommonStartMenu, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Startup", "", -1, "A_StartupCommon", "commonstartup", ""
	, lMenuCommonStartupMenu, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%ALLUSERSPROFILE%", "", -1, "A_AppDataCommon", "commonappdata", ""
	, lMenuCommonAppData, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%LOCALAPPDATA%\Microsoft\Windows\Temporary Internet Files", "", -1, "", "", ""
	, lMenuCache, "Temporary"
	, "CLS", "CLS", "CLS", "CLS", "CLS", "CLS", "CLS")
InitSpecialFolderObject("%LOCALAPPDATA%\Microsoft\Windows\History", "", -1, "", "history", ""
	, lMenuHistory, "History"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%ProgramFiles%", "", -1, "A_ProgramFiles", "programfiles", ""
	, lMenuProgramFiles, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
if (A_Is64bitOS)
	InitSpecialFolderObject("%ProgramFiles(x86)%", "", -1, "", "programfilesx86", ""
		, lMenuProgramFiles . " (x86)", "Folder"
		, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject("%PUBLIC%\Libraries", "", -1, "", "", ""
	, lMenuPublicLibraries, "Folder"
	, "CLS", "CLS", "CLS", "CLS", "CLS", "CLS", "CLS")

;---------------------
; Path under the Users folder (no CLSID, localized name and icon provided), no Shell Command

StringReplace, strPathUsername, A_AppData, \AppData\Roaming
StringReplace, strPathUsers, strPathUsername, \%A_UserName%

InitSpecialFolderObject(strPathUsers . "\Public", "Public", -1, "", "common", ""
	, "Public Folder", "" ; Public
	, "SCT", "SCT", "SCT", "CLS", "DOA", "CLS", "CLS")
	; OK     OK      OK     OK    OK      OK

;---------------------
; Path using AHK constants (no CLSID), localized name and icon provided, no Shell Command - to be tested with DOpus, TC and FPc

InitSpecialFolderObject(A_Desktop, "", 0, "A_Desktop", "desktop", 2121
	, lMenuDesktop, "lMenuDesktop"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "TCC", "CLS")
InitSpecialFolderObject(A_DesktopCommon, "", -1, "A_DesktopCommon", "commondesktopdir", ""
	, lMenuCommonDesktop, "lMenuDesktop"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject(A_Temp, "", -1, "A_Temp", "temp", ""
	, lMenuTemporaryFiles, "Temporary"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")
InitSpecialFolderObject(A_WinDir, "", -1, "A_WinDir", "windows", ""
	, "Windows", "Winver"
	, "CLS", "CLS", "CLS", "CLS", "DOA", "CLS", "CLS")

;------------------------------------------------------------
; Build folders list for dropdown

strSpecialFoldersList := ""
for strSpecialFolderName in objClassIdOrPathByDefaultName
	strSpecialFoldersList := strSpecialFoldersList . strSpecialFolderName . "|"
StringTrimRight, strSpecialFoldersList, strSpecialFoldersList, 1

return
;------------------------------------------------------------


;------------------------------------------------------------
InitSpecialFolderObject(strClassIdOrPath, strShellConstantText, intShellConstantNumeric, strAHKConstant, strDOpusAlias, strTCCommand
	, strDefaultName, strDefaultIcon
	, strUse4NavigateExplorer, strUse4NewExplorer, strUse4Dialog, strUse4Console, strUse4DOpus, strUse4TC, strUse4FPc)

; strClassIdOrPath: CLSID or Path, used as key to access objSpecialFolder objects
;		CLSID Win_7: http://www.sevenforums.com/tutorials/110919-clsid-key-list-windows-7-a.html
;		CLSID Win_8: http://www.eightforums.com/tutorials/13591-clsid-key-guid-shortcuts-list-windows-8-a.html
; 		Environment system variables: http://en.wikipedia.org/wiki/Environment_variable#Windows
;		HKEY_CLASSES_ROOT Key: http://msdn.microsoft.com/en-us/library/windows/desktop/ms724475(v=vs.85).aspx
; 		NOTES How to call in Explorer...
;		... CLSID: shell:::{{20D04FE0-3AEA-1069-A2D8-08002B30309D}}
;		... ShellConstant: shell:MyComputerFolder

; strShellConstantText: text constant used to navigate using Explorer or Dialog box? What with DOpus and TC?
;		http://www.sevenforums.com/tutorials/4941-shell-command.html
;		http://www.eightforums.com/tutorials/6050-shell-commands-windows-8-a.html

; intShellConstantNumeric: numeric ShellSpecialFolderConstants constant 
;		http://msdn.microsoft.com/en-us/library/windows/desktop/bb774096%28v=vs.85%29.aspx

; CLSID, strShellConstantText (by version XP!, Vista, 7) and intShellConstantNumeric: http://docs.rainmeter.net/tips/launching-windows-special-folders

; strAHKConstant: AutoHotkey constant

; strDOpusAlias: Directory Opus constant

; strTCCommand: Total Commander constant

; strDefaultName: name for menu if path is used, fallback name if CLSID is used to access localized name

; strDefaultIcon: icon in strIconsMenus if path is used, fallback icon (?) if CLSID returns no icon resource

; Constants for "use" flags:
; 		CLS: Class ID
;		SCT: Shell Constant Text
;		SCN: Shell Constant Numeric
;		DOA: Directory Opus Alias
;		TCC: Total Commander Commands

; Usage flags:
; 		strUse4NavigateExplorer
; 		strUse4NewExplorer
; 		strUse4Dialog
; 		strUse4Console
; 		strUse4DOpus
; 		strUse4TC
;		strUse4FPc

; Special Folder Object definition:
;		ClassIdOrPath: key to access one Special Folder object (example: objSpecialFolders[strClassIdOrPath]
;		objSpecialFolder.ShellConstantText: text constant used to navigate using Explorer or Dialog box? What with DOpus and TC?
;		objSpecialFolder.ShellConstantNumeric: numeric ShellSpecialFolderConstants constant 
;		objSpecialFolder.AHKConstant: AutoHotkey constant
;		objSpecialFolder.DOpusAlias: Directory Opus constant
;		objSpecialFolder.TCCommand: Total Commander constant
;		objSpecialFolder.DefaultName:
;		objSpecialFolder.DefaultIcon: icon resource name in the format "file,index"
;		objSpecialFolder.Use4NavigateExplorer:
;		objSpecialFolder.Use4NewExplorer:
;		objSpecialFolder.Use4Dialog:
;		objSpecialFolder.Use4Console:
;		objSpecialFolder.Use4DOpus:
;		objSpecialFolder.Use4TC:
;		objSpecialFolder.Use4FPc:

;------------------------------------------------------------
{
	objOneSpecialFolder := Object()
	
	blnIsClsId := (SubStr(strClassIdOrPath, 1, 1) = "{")

	if (blnIsClsId)
		strThisDefaultName := GetLocalizedNameForClassId(strClassIdOrPath)
	If !StrLen(strThisDefaultName)
		strThisDefaultName := strDefaultName
    objClassIdOrPathByDefaultName.Insert(strThisDefaultName, strClassIdOrPath)
	objOneSpecialFolder.DefaultName := strThisDefaultName
	
	if (blnIsClsId)
		strThisDefaultIcon := GetIconForClassId(strClassIdOrPath)
	if !StrLen(strThisDefaultIcon) and StrLen(objIconsFile[strDefaultIcon]) and StrLen(objIconsIndex[strDefaultIcon])
		strThisDefaultIcon := objIconsFile[strDefaultIcon] . "," . objIconsIndex[strDefaultIcon]
	if !StrLen(strThisDefaultIcon)
		strThisDefaultIcon := "%SystemRoot%\System32\shell32.dll,4"
	objOneSpecialFolder.DefaultIcon := strThisDefaultIcon

	objOneSpecialFolder.ShellConstantText := strShellConstantText
	objOneSpecialFolder.ShellConstantNumeric := intShellConstantNumeric
	objOneSpecialFolder.AHKConstant := strAHKConstant
	objOneSpecialFolder.DOpusAlias := strDOpusAlias
	objOneSpecialFolder.TCCommand := strTCCommand
	
	objOneSpecialFolder.Use4NavigateExplorer := strUse4NavigateExplorer
	objOneSpecialFolder.Use4NewExplorer := strUse4NewExplorer
	objOneSpecialFolder.Use4Dialog := strUse4Dialog
	objOneSpecialFolder.Use4Console := strUse4Console
	objOneSpecialFolder.Use4DOpus := strUse4DOpus
	objOneSpecialFolder.Use4TC := strUse4TC
	objOneSpecialFolder.Use4FPc := strUse4FPc
	
	objSpecialFolders.Insert(strClassIdOrPath, objOneSpecialFolder)
}
;------------------------------------------------------------


;------------------------------------------------------------
InitClassId(strClassId, strFallbackName, strShellConstant)
;------------------------------------------------------------
{
    strLocalizedName := GetLocalizedNameForClassId(strClassId)
	If !StrLen(strLocalizedName)
			strLocalizedName := strFallbackName
			
    objClassIdByLocalizedName.Insert(strLocalizedName, strClassId)
    objLocalizedNameByClassId.Insert(strClassId, strLocalizedName)
	objShellConstantByClassId.Insert(strClassId, strShellConstant)

	objIconByClassId.Insert(strClassId, GetIconForClassId(strClassId))
	if !StrLen(objIconByClassId[strClassId])
		objIconByClassId[strClassId] := "%SystemRoot%\System32\shell32.dll,4"
}
;------------------------------------------------------------


;------------------------------------------------------------
InitClassIdException(strPath, strName, strIconMenu)
; Exceptions to CLSID:
; how to detect exception: first char of ClassId is not "{"
; instead of CLSID we use the path
; instead of the localized name, we use the localized name from the FP language files
; instead of the icon resource, we use the strIconsMenus name (see InitSystemArrays:)
;------------------------------------------------------------
{
    objClassIdByLocalizedName.Insert(strName, strPath)
    objLocalizedNameByClassId.Insert(strPath, strName)
	objIconByClassId.Insert(strPath, strIconMenu)
	objShellConstantByClassId.Insert(strPath, strShellConstant)
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
GetIconForClassId(strClassId)
;------------------------------------------------------------
{
	RegRead, strDefaultIcon, HKEY_CLASSES_ROOT, CLSID\%strClassId%\DefaultIcon
    return strDefaultIcon
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
	Diag("GetOsVersion()", GetOsVersion())
	Diag("A_Is64bitOS", A_Is64bitOS)
	Diag("A_Language", A_Language)
	Diag("A_IsAdmin", A_IsAdmin)
}

FileRead, strDiag, %strIniFile%
StringReplace, strDiag, strDiag, `", `"`"
Diag("IniFile", """" . strDiag . """")
FileAppend, `n, %strDiagFile% ; required when the last line of the existing file ends with "

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
	Global lGuiFullTitle

	WinGetTitle, strCurrentWindow, A
	if (strCurrentWindow <> lGuiFullTitle)
		return

	MouseGetPos, , , , strControl ; Static1, StaticN, Button1, ButtonN
	if InStr(strControl, "Static")
	{
		StringReplace, intControl, strControl, Static
		; 3-23, 25-26
		if (intControl < 3) or (intControl = 24) or (intControl > 26)
			return
	}
	else if !InStr(strControl, "Button")
		return

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
	global blnClickOnTrayIcon
	
	if (lParam = 0x202) ; WM_LBUTTONUP
	{
		blnClickOnTrayIcon := 1
		SetTimer, PopupMenuNewWindowMouse, -1
		return 0
	}
} 
;------------------------------------------------------------



;========================================================================================================================
; EXIT
;========================================================================================================================


;-----------------------------------------------------------
TrayMenuExitApp:
;-----------------------------------------------------------

ExitApp
;-----------------------------------------------------------


;-----------------------------------------------------------
CleanUpBeforeExit:
;-----------------------------------------------------------

strSettingsPosition := "-1" ; center at minimal size
if (blnRememberSettingsPosition)
{
	WinGet, intMinMax, MinMax, ahk_id %strAppHwnd%
	if (intMinMax <> 1) ; if window is maximized, we keep the default positionand size (center at minimal size)
	{
		WinGetPos, intX, intY, intW, intH, ahk_id %strAppHwnd%
		strSettingsPosition := intX . "|" . intY . "|" . intW . "|" . intH
	}
}
IniWrite, %strSettingsPosition%, %strIniFile%, Global, SettingsPosition

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
Menu, Tray, Add, %  BuildSpecialMenuItemName(5, L(lMenuSettings, strAppName)), GuiShow
Menu, Tray, Add, % strAppName . ".ini", ShowIniFile
Menu, Tray, Add
Menu, Tray, Add, %lMenuRunAtStartup%, RunAtStartup
Menu, Tray, Add, %lMenuSuspendHotkeys%, SuspendHotkeys
Menu, Tray, Add
Menu, Tray, Add, %lMenuUpdate%, Check4Update
Menu, Tray, Add, %lMenuHelp%, GuiHelp
Menu, Tray, Add, %lMenuAbout%, GuiAbout
Menu, Tray, Add, %lDonateMenu%, GuiDonate
Menu, Tray, Add
Menu, Tray, Add, % L(lMenuExitFoldersPopup, strAppName), TrayMenuExitApp
Menu, Tray, Default, %  BuildSpecialMenuItemName(5, L(lMenuSettings, strAppName))
if (blnUseColors)
	Menu, Tray, Color, %strMenuBackgroundColor%
Menu, Tray, Tip, % strAppName . " " . strAppVersion . " (" . str32or64 . "-bit)`n" . (blnDonor ? lDonateThankyou : lDonateButton)

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildSpecialFoldersMenu:
; Windows XP only
;------------------------------------------------------------

Menu, menuSpecialFolders, Add
Menu, menuSpecialFolders, DeleteAll ; had problem with DeleteAll making the Special menu to disappear 1/2 times - now OK
if (blnUseColors)
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
BuildFoldersInExplorerMenu:
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

objFoldersInExplorers := Object()

intExplorersIndex := 0 ; used in PopupMenu and SaveGroup to check if we disable menu or button when empty

if (blnUseDirectoryOpus)
	for intIndex, objLister in objDOpusListers
	{
		; if we have no path or or DOpus collection, skip it
		if !StrLen(objLister.LocationURL) or InStr(objLister.LocationURL, "coll://")
			continue
		
		if NameIsInObject(objLister.LocationURL, objFoldersInExplorers)
			continue
		
		intExplorersIndex := intExplorersIndex + 1
			
		objFolderInExplorer := Object()
		objFolderInExplorer.LocationURL := objLister.LocationURL
		objFolderInExplorer.Name := objLister.LocationURL
		
		; used for DOpus windows to discriminatre different listers
		objFolderInExplorer.WindowId := objLister.Lister
		
		; info used to create groups
		objFolderInExplorer.TabId := objLister.Tab
		objFolderInExplorer.Position := objLister.Position
		objFolderInExplorer.MinMax := objLister.MinMax
		objFolderInExplorer.Pane := (objLister.Pane = 0 ? 1 : objLister.Pane) ; consider pane 0 as pane 1
		objFolderInExplorer.WindowType := "DO"
		
		objFoldersInExplorers.Insert(intExplorersIndex, objFolderInExplorer)
	}

/* when TC tabs resolved
if (blnUseTotalCommander)
	for intIndex, objTCList in objTCLists
	{
		if !NameIsInObject(objTCList.LocationURL, objFoldersInExplorers)
		{
			intExplorersIndex := intExplorersIndex + 1
			
			objGroupMenuExplorer := Object()
			objGroupMenuExplorer.LocationURL := objTCList.LocationURL
			objGroupMenuExplorer.Name := objTCList.LocationURL
			objGroupMenuExplorer.WindowId := objTCList.WindowId
			objGroupMenuExplorer.Position := objTCList.Position
			objGroupMenuExplorer.MinMax := objTCList.MinMax
			objGroupMenuExplorer.Pane := objTCList.Pane
			objGroupMenuExplorer.WindowType := "TC"
			
			objFoldersInExplorers.Insert(intExplorersIndex, objGroupMenuExplorer)
		}
	}
*/

for intIndex, objFolder in objExplorersWindows
{
	; if we have no path, skip it
	if !StrLen(objFolder.LocationURL)
		continue
		
	if NameIsInObject(objFolder.LocationName, objFoldersInExplorers)
		continue
	
	intExplorersIndex := intExplorersIndex + 1
	
	objFolderInExplorer := Object()
	objFolderInExplorer.LocationURL := objFolder.LocationURL
	objFolderInExplorer.Name := objFolder.LocationName
	objFolderInExplorer.IsSpecialFolder := objFolder.IsSpecialFolder
	
	; not used for Explorer windows, but keep it
	objFolderInExplorer.WindowId := objFolder.WindowId

	; info used to create groups
	objFolderInExplorer.Position := objFolder.Position
	objFolderInExplorer.MinMax := objFolder.MinMax
	objFolderInExplorer.WindowType := "EX"

	objFoldersInExplorers.Insert(intExplorersIndex, objFolderInExplorer)
}

Menu, menuFoldersInExplorer, Add
Menu, menuFoldersInExplorer, DeleteAll
if (blnUseColors)
	Menu, menuFoldersInExplorer, Color, %strMenuBackgroundColor%

intShortcutFoldersInExplorer := 0
objLocationUrlByName := Object()

for intIndex, objFolderInExplorer in objFoldersInExplorers
{
	strMenuName := (blnDisplayMenuShortcuts and (intShortcutFoldersInExplorer <= 35) ? "&" . NextMenuShortcut(intShortcutFoldersInExplorer) . " " : "") . objFolderInExplorer.Name
	objLocationUrlByName.Insert(strMenuName, objFolderInExplorer.LocationURL) ; can include the numeric shortcut
	AddMenuIcon("menuFoldersInExplorer", strMenuName, "OpenFolderInExplorer", "Folder")
}

objDOpusListers := 
objExplorersWindows :=

return
;------------------------------------------------------------


;------------------------------------------------------------
CollectExplorers(objExplorers, pExplorers)
;------------------------------------------------------------
{
	intExplorers := 0
	
	For pExplorer in pExplorers
	; see http://msdn.microsoft.com/en-us/library/windows/desktop/aa752084(v=vs.85).aspx
	{
		/* in v.3.9.8: stop interupting Explorer collection if an error occurs - just check for content and continue
		if (A_LastError)
			; an error occurred during ComObjCreate (A_LastError probably is E_UNEXPECTED = -2147418113 #0x8000FFFFL)
			break
		*/

		strType := ""
		try strType := pExplorer.Type ; Gets the type name of the contained document object. "Document HTML" for IE windows. Should be empty for file Explorer windows.
		strWindowID := ""
		try strWindowID := pExplorer.HWND ; Try to get the handle of the window. Some ghost Explorer in the ComObjCreate may return an empty handle
		
		if !StrLen(strType) ; must be empty
			and StrLen(strWindowID) ; must not be empty
		{
			intExplorers := intExplorers + 1
			objExplorer := Object()
			objExplorer.Position := pExplorer.Left . "|" . pExplorer.Top . "|" . pExplorer.Width . "|" . pExplorer.Height

			objExplorer.IsSpecialFolder := !StrLen(pExplorer.LocationURL) ; empty for special folders like Recycle bin
			
			if (objExplorer.IsSpecialFolder)
			{
				objExplorer.LocationURL := pExplorer.Document.Folder.Self.Path
				objExplorer.LocationName := pExplorer.LocationName ; see http://msdn.microsoft.com/en-us/library/aa752084#properties
			}
			else
			{
				objExplorer.LocationURL := pExplorer.LocationURL
				strLocationName :=  UriDecode(pExplorer.LocationURL)
				StringReplace, strLocationName, strLocationName, file:///
				StringReplace, strLocationName, strLocationName, /, \, A
				objExplorer.LocationName := strLocationName
			}
			
			objExplorer.WindowId := pExplorer.HWND ; not used for Explorer windows, but keep it
			WinGet, intMinMax, MinMax, % "ahk_id " . pExplorer.HWND
			objExplorer.MinMax := intMinMax
			
			objExplorers.Insert(intExplorers, objExplorer) ; I was checking if StrLen(pExplorer.HWND) - any reason?
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
			objTCList.LocationURL := ClipBoard
			objTCList.WindowId := strWindowId
			objTCList.Position := intX . "|" . intY . "|" . intW . "|" . intH
			objTCList.MinMax := intMinMax
			objTCList.Pane := 1
			objLists.Insert(intTCWindows, objTCList)
			ClipBoard := ""
			sleep, 50
		}
		SendMessage, 0x433, %cm_CopyTrgPathToClip%, , , ahk_id %strWindowId%
		ClipWait, 2, 1 ; wait up to 2 seconds for text (1) in the Clipboard
		sleep, 10 ; wait 10 additional seconds to improve reliability
		if StrLen(ClipBoard)
		{
			objTCList := Object()
			objTCList.LocationURL := ClipBoard
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
			objLister.LocationURL := SubStr(strSubStr, InStr(strSubStr, ">") + 1)

			WinGetPos, intX, intY, intW, intH, % "ahk_id " . objLister.lister
			objLister.Position := intX . "|" . intY . "|" . intW . "|" . intH
			WinGet, intMinMax, MinMax, % "ahk_id " . objLister.lister
			objLister.MinMax := intMinMax
			objLister.Pane := objLister.Side
			
			if !InStr(objLister.LocationURL, "ftp://")
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
NameIsInObject(strName, obj)
;------------------------------------------------------------
{
	loop, % obj.MaxIndex()
		if (strName = obj[A_Index].Name)
			return true
		
	return false
}
;------------------------------------------------------------


;------------------------------------------------------------
FoldersInExplorerMenuShortcut:
;------------------------------------------------------------

blnCopyLocation := false

blnNewWindow := !CanOpenFavorite("", strTargetWinId, strTargetClass, strTargetControl)
Gosub, SetMenuPosition ; sets strTargetWinId or activate the window strTargetWinId set by CanOpenFavorite

Gosub, BuildFoldersInExplorerMenu
if (intExplorersIndex) ; there are Folders in Explorer menu
{
	CoordMode, Menu, % (intPopupMenuPosition = 2 ? "Window" : "Screen")
	Menu, menuFoldersInExplorer, Show, %intMenuPosX%, %intMenuPosY%
}
else
{
	TrayTip, % L(lTrayTipNoFoldersInExplorerMenuTitle, strAppName), %lTrayTipNoFoldersInExplorerMenuDetail%, , 18
	Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildGroupMenu:
;------------------------------------------------------------

Menu, menuGroups, Add
Menu, menuGroups, DeleteAll
if (blnUseColors)
	Menu, menuGroups, Color, %strMenuBackgroundColor%

intShortcut := 0
Loop, Parse, strGroups, |
{
	IniRead, blnReplaceWhenRestoringThisGroup, %striniFile%, Group-%A_LoopField%, ReplaceWhenRestoringThisGroup, %A_Space% ; empty if not found
	strMenuName := (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut) . " " : "")
		. A_LoopField . " (" . (blnReplaceWhenRestoringThisGroup ? lMenuGroupReplace : lMenuGroupAdd) . ")"
	AddMenuIcon("menuGroups", strMenuName, "GroupLoad", "lMenuGroup")
}

if StrLen(strGroups)
	Menu, menuGroups, Add
AddMenuIcon("menuGroups", lMenuGroupSave, "GuiGroupSaveFromMenu", "menuGroupSave")

return
;------------------------------------------------------------


;------------------------------------------------------------
GroupsMenuShortcut:
;------------------------------------------------------------

blnCopyLocation := false

Gosub, SetMenuPosition

Menu, menuGroups
	, % (!intExplorersIndex ? "Disable" : "Enable") ; disable Save group menu if no Explorer
	, %lMenuGroupSave%
CoordMode, Menu, % (intPopupMenuPosition = 2 ? "Window" : "Screen")
Menu, menuGroups, Show, %intMenuPosX%, %intMenuPosY%

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildRecentFoldersMenu:
;------------------------------------------------------------

Menu, menuRecentFolders, Add
Menu, menuRecentFolders, DeleteAll ; had problem with DeleteAll making the Special menu to disappear 1/2 times - now OK
if (blnUseColors)
	Menu, menuRecentFolders, Color, %strMenuBackgroundColor%

objRecentFolders := Object()
intRecentFoldersIndex := 0 ; used in PopupMenu... to check if we disable the menu when empty

RegRead, strRecentsFolder, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, Recent

/*
Alternative to collect recent files *** NOT WORKING with XP
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
	
	strMenuName := (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut) . " " : "") . strOutTarget
	AddMenuIcon("menuRecentFolders", strMenuName, "OpenRecentFolder", "Folder")

	if (intRecentFoldersIndex >= intRecentFolders)
		break
}

return
;------------------------------------------------------------


;------------------------------------------------------------
RefreshRecentFolders:
RecentFoldersShortcut:
;------------------------------------------------------------

blnCopyLocation := false

blnNewWindow := !CanOpenFavorite("", strTargetWinId, strTargetClass, strTargetControl)
Gosub, SetMenuPosition ; sets strTargetWinId or activate the window strTargetWinId set by CanOpenFavorite

ToolTip, %lMenuRefreshRecent%...
Gosub, BuildRecentFoldersMenu
ToolTip
CoordMode, Menu, % (intPopupMenuPosition = 2 ? "Window" : "Screen")
Menu, menuRecentFolders, Show, %intMenuPosX%, %intMenuPosY%

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildClipboardMenu:
;------------------------------------------------------------

Menu, menuClipboard, Add
Menu, menuClipboard, DeleteAll
if (blnUseColors)
	Menu, menuClipboard, Color, %strMenuBackgroundColor%

blnClipboardMenuEnable := 0

Gosub, RefreshClipboardMenu

return
;------------------------------------------------------------


;------------------------------------------------------------
RefreshClipboardMenu:
;------------------------------------------------------------

blnPreviousClipboardMenuDeleted := 0
intShortcutClipboardMenu := 0
strURLsInClipboard := ""

; Parse Clipboard for folder, document or application filenames (filenames alone on one line)
Loop, parse, Clipboard, `n, `r%A_Space%%A_Tab%/?:*`"><|
{
    strClipboardLine = %A_LoopField%
	strClipboardLineExpanded := EnvVars(strClipboardLine) ; only to test if file exist - will not be displayed in menu

	if StrLen(FileExist(strClipboardLineExpanded))
	{
		if !(blnPreviousClipboardMenuDeleted)
		{
			Menu, menuClipboard, Add
			Menu, menuClipboard, DeleteAll
			blnPreviousClipboardMenuDeleted := 1
		}
		blnClipboardMenuEnable := 1

		strMenuName := (blnDisplayMenuShortcuts and (intShortcutFoldersInExplorer <= 35) ? "&" . NextMenuShortcut(intShortcutClipboardMenu) . " " : "") . strClipboardLine
		if (blnDisplayIcons)
			if LocationIsDocument(strClipboardLineExpanded)
			{
				GetIcon4Location(strClipboardLineExpanded, strThisIconFile, intThisIconIndex)
				strIconValue := strThisIconFile . "," . intThisIconIndex
			}
			else
				strIconValue := "Folder"
		AddMenuIcon("menuClipboard", strMenuName, "OpenClipboard", strIconValue)
	}

	; Parse Clipboard line for URLs (anywhere on the line)
	strURLSearchString := strClipboardLine
	Gosub, GetURLsInClipboardLine
}

Sort, strURLsInClipboard

Loop, parse, strURLsInClipboard, `n
{
	if !StrLen(A_LoopField)
		break
	
	; if we get here, we have at least one URL, check if we need to delete previous menu
	if !(blnPreviousClipboardMenuDeleted)
	{
		Menu, menuClipboard, Add
		Menu, menuClipboard, DeleteAll
		blnPreviousClipboardMenuDeleted := 1
	}
	blnClipboardMenuEnable := 1

	strMenuName := (blnDisplayMenuShortcuts and (intShortcutFoldersInExplorer <= 35) ? "&" . NextMenuShortcut(intShortcutClipboardMenu) . " " : "") . A_LoopField
	if StrLen(strMenuName) < 260 ; skip too long URLs
	{
		Menu, menuClipboard, Add, %strMenuName%, OpenClipboard
		if (blnDisplayIcon)
			Menu, menuClipboard, Icon, %strMenuName%, %strThisIconFile%, %intThisIconIndex%, %intIconSize%
	}
}

return
;------------------------------------------------------------


;------------------------------------------------------------
GetURLsInClipboardLine:
;------------------------------------------------------------
; Adapted from AHK help file: http://ahkscript.org/docs/commands/LoopReadFile.htm
; It's done this particular way because some URLs have other URLs embedded inside them:
StringGetPos, intURLStart1, strURLSearchString, http://
StringGetPos, intURLStart2, strURLSearchString, https://
StringGetPos, intURLStart3, strURLSearchString, www.

; Find the left-most starting position:
intURLStart := intURLStart1 ; Set starting default.
Loop
{
	; It helps performance (at least in a script with many variables) to resolve
	; "intURLStart%A_Index%" only once:
	intArrayElement := intURLStart%A_Index%
	if (intArrayElement = "") ; End of the array has been reached.
		break
	if (intArrayElement = -1) ; This element is disqualified.
		continue
	if (intURLStart = -1)
		intURLStart := intArrayElement
	else ; intURLStart has a valid position in it, so compare it with intArrayElement.
	{
		if (intArrayElement <> -1)
			if (intArrayElement < intURLStart)
				intURLStart := intArrayElement
	}
}

if (intURLStart = -1) ; No URLs exist in strURLSearchString.
	return ; (exit loop)

; Otherwise, extract this strURL:
StringTrimLeft, strURL, strURLSearchString, %intURLStart% ; Omit the beginning/irrelevant part.
Loop, parse, strURL, %A_Tab%%A_Space%<> ; Find the first space, tab, or angle (if any).
{
	strURL := A_LoopField
	break ; i.e. perform only one loop iteration to fetch the first "field".
}
; If the above loop had zero iterations because there were no ending characters found,
; leave the contents of the strURL var untouched.

; If the strURL ends in a double quote, remove it.  For now, StringReplace is used, but
; note that it seems that double quotes can legitimately exist inside URLs, so this
; might damage them:
StringReplace, strURLCleansed, strURL, ",, All


; See if there are any other URLs in this line:
StringLen, intCharactersToOmit, strURL
intCharactersToOmit := intCharactersToOmit + intURLStart
StringTrimLeft, strURLSearchString, strURLSearchString, %intCharactersToOmit%

Gosub, GetURLsInClipboardLine ; Recursive call to self (end of loop)

strURLsInClipboard := strURLsInClipboard . strURLCleansed . "`n"

return

;------------------------------------------------------------


;------------------------------------------------------------
ClipboardMenuShortcut:
;------------------------------------------------------------

blnCopyLocation := false

blnNewWindow := !CanOpenFavorite("", strTargetWinId, strTargetClass, strTargetControl)
Gosub, SetMenuPosition ; sets strTargetWinId or activate the window strTargetWinId set by CanOpenFavorite

Gosub, RefreshClipboardMenu
if (blnClipboardMenuEnable)
{
    CoordMode, Menu, % (intPopupMenuPosition = 2 ? "Window" : "Screen")
    Menu, menuClipboard, Show, %intMenuPosX%, %intMenuPosY%
}
else
{
    TrayTip, % L(lTrayTipNoClipboardMenuTitle, strAppName), %lTrayTipNoClipboardMenuDetail%, , 18
	Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildMainMenu:
BuildMainMenuWithStatus:
;------------------------------------------------------------

if (A_ThisLabel = "BuildMainMenuWithStatus")
{
	TrayTip, % L(lTrayTipWorkingTitle, strAppName)
		, %lTrayTipWorkingDetail%, , 17
	Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)
}

objMenuColumnBreaks := Object()
blnMainIsFirstColumn := True

Menu, %lMainMenuName%, Add
Menu, %lMainMenuName%, DeleteAll
if (blnUseColors)
	Menu, %lMainMenuName%, Color, %strMenuBackgroundColor%

BuildOneMenu(lMainMenuName) ; and recurse for submenus

if !IsColumnBreak(arrMenus[lMainMenuName][arrMenus[lMainMenuName].MaxIndex()].FavoriteName)
; column break not allowed if first item is a separator
	Menu, %lMainMenuName%, Add

if (blnDisplaySpecialFolders) and (A_OSVersion = "WIN_XP")
	AddMenuIcon(lMainMenuName, lMenuSpecialFolders, ":menuSpecialFolders", "lMenuSpecialFolders")

if (blnDisplayFoldersInExplorerMenu)
{
	AddMenuIcon(lMainMenuName, BuildSpecialMenuItemName(6, lMenuFoldersInExplorer), ":menuFoldersInExplorer", "lMenuFoldersInExplorer")
	if (blnUseColors)
		Menu, menuFoldersInExplorer, Color, %strMenuBackgroundColor%
}

if (blnDisplayGroupMenu)
{
	AddMenuIcon(lMainMenuName, BuildSpecialMenuItemName(7, lMenuGroup), ":menuGroups", "lMenuGroup")
	if (blnUseColors)
		Menu, menuGroups, Color, %strMenuBackgroundColor%
}

if (blnDisplayRecentFolders)
	AddMenuIcon(lMainMenuName, BuildSpecialMenuItemName(8, lMenuRecentFolders), "RefreshRecentFolders", "lMenuRecentFolders")

if (blnDisplayClipboardMenu)
	AddMenuIcon(lMainMenuName, BuildSpecialMenuItemName(9, lMenuClipboard), ":menuClipboard", "Clipboard")

if ((blnDisplaySpecialFolders and A_OSVersion = "WIN_XP") or blnDisplayRecentFolders or blnDisplayFoldersInExplorerMenu or blnDisplayGroupMenu or blnDisplayClipboardMenu)
	Menu, %lMainMenuName%, Add

AddMenuIcon(lMainMenuName, BuildSpecialMenuItemName(5, L(lMenuSettings, strAppName) . "..."), "GuiShow", "lMenuSettings")
Menu, %lMainMenuName%, Default, %  BuildSpecialMenuItemName(5, L(lMenuSettings, strAppName) . "...")
AddMenuIcon(lMainMenuName, lMenuAddThisFolder . "...", "AddThisFolder", "lMenuAddThisFolder")

if (blnDisplayCopyLocationMenu)
	AddMenuIcon(lMainMenuName, lMenuCopyLocation . "...", "PopupMenuCopyLocation", "Clipboard")

if !(blnDonor)
{
	Menu, %lMainMenuName%, Add
	AddMenuIcon(lMainMenuName, lDonateMenu . "...", "GuiDonate", "lDonateMenu")
}

if (A_ThisLabel = "BuildMainMenuWithStatus")
{
	TrayTip, % L(lTrayTipInstalledTitle, strAppName)
		, %lTrayTipWorkingDetailFinished%, , 17
	Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)
}

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
	global objMenuColumnBreaks
	global blnMainIsFirstColumn
	
	intShortcut := 0
	
	; try because at first execution strMenu does not exist and produces an error,
	; but DeleteAll is required later for menu updates
	try Menu, %strMenu%, DeleteAll
	
	arrThisMenu := arrMenus[strMenu]
	intMenuItemsCount := 0
	intMenuArrayItemsCount := 0
	blnIsFirstColumn := True
	
	Loop, % arrThisMenu.MaxIndex()
	{	
		intMenuItemsCount := intMenuItemsCount + 1
		intMenuArrayItemsCount := intMenuArrayItemsCount + 1
		
		if (arrThisMenu[A_Index].FavoriteType = "S") ; this is a submenu
		{
			strSubMenuFullName := arrThisMenu[A_Index].SubmenuFullName
			strSubMenuDisplayName := arrThisMenu[A_Index].FavoriteName
			strSubMenuParent := arrThisMenu[A_Index].MenuName
			
			BuildOneMenu(strSubMenuFullName) ; recursive call
			if (blnUseColors)
				Try Menu, %strSubMenuParent%, Color, %strMenuBackgroundColor% ; Try because this can fail if submenu is empty
			
			strMenuName := (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut) . " " : "") . strSubMenuDisplayName
			Try Menu, %strSubMenuParent%, Add, %strMenuName%, % ":" . strSubMenuFullName
			catch e ; when menu is empty
			{
				Menu, % arrThisMenu[A_Index].MenuName, Add, %strMenuName%, OpenFavorite ; will never be called because disabled
				Menu, % arrThisMenu[A_Index].MenuName, Disable, %strMenuName%
			}
			Menu, % arrThisMenu[A_Index].MenuName, % (arrMenus[strSubMenuFullName].MaxIndex() ? "Enable" : "Disable"), %strMenuName% ; disable menu if empty
			if (blnDisplayIcons and (A_OSVersion <> "WIN_XP" or blnIsFirstColumn))
			{
				ParseIconResource(arrThisMenu[A_Index].IconResource, strThisIconFile, intThisIconIndex, "Submenu")

				Menu, % arrThisMenu[A_Index].MenuName, UseErrorLevel, on
				Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%
					, %strThisIconFile%, %intThisIconIndex% , %intIconSize%
				if (ErrorLevel)
					Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%
						, % objIconsFile["UnknownDocument"], % objIconsIndex["UnknownDocument"], %intIconSize%
				Menu, % arrThisMenu[A_Index].MenuName, UseErrorLevel, off
			}
		}
		
		else if (arrThisMenu[A_Index].FavoriteName = lMenuSeparator) ; this is a separator
			
			if IsColumnBreak(arrThisMenu[A_Index - 1].FavoriteName)
				intMenuItemsCount := intMenuItemsCount - 1 ; separator not allowed as first item is a column, skip it
			else
				Menu, % arrThisMenu[A_Index].MenuName, Add
			
		else if IsColumnBreak(arrThisMenu[A_Index].FavoriteName)
		{
			blnIsFirstColumn := False
			if (strMenu = lMainMenuName)
				blnMainIsFirstColumn := False
			intMenuItemsCount := intMenuItemsCount - 1
			objMenuColumnBreak := Object()
			objMenuColumnBreak.MenuName := strMenu
			objMenuColumnBreak.MenuPosition := intMenuItemsCount
			objMenuColumnBreak.MenuArrayPosition := intMenuArrayItemsCount
			objMenuColumnBreaks.Insert(objMenuColumnBreak)
		}
		else ; this is a favorite (folder, document, application or URL)
		{
			strSubMenuDisplayName := arrThisMenu[A_Index].FavoriteName
			strMenuName := (blnDisplayMenuShortcuts and (intShortcut <= 35) ? "&" . NextMenuShortcut(intShortcut) . " " : "")
				. strSubMenuDisplayName
			Menu, % arrThisMenu[A_Index].MenuName, Add, %strMenuName%, OpenFavorite

			if (blnDisplayIcons and (A_OSVersion <> "WIN_XP" or blnIsFirstColumn))
			{
				Menu, % arrThisMenu[A_Index].MenuName, UseErrorLevel, on
				if (arrThisMenu[A_Index].FavoriteType = "F") ; this is a folder
					ParseIconResource(arrThisMenu[A_Index].IconResource, strThisIconFile, intThisIconIndex, "Folder")
				else if (arrThisMenu[A_Index].FavoriteType = "U") ; this is an URL
					if StrLen(arrThisMenu[A_Index].IconResource)
						ParseIconResource(arrThisMenu[A_Index].IconResource, strThisIconFile, intThisIconIndex)
					else
						GetIcon4Location(strTempDir . "\default_browser_icon.html", strThisIconFile, intThisIconIndex)
						; not sure it is required to have a physical file with .html extension - but keep it as is by safety
				else ; this is a document
					if StrLen(arrThisMenu[A_Index].IconResource)
						ParseIconResource(arrThisMenu[A_Index].IconResource, strThisIconFile, intThisIconIndex)
					else
						GetIcon4Location(arrThisMenu[A_Index].FavoriteLocation, strThisIconFile, intThisIconIndex)
					
				ErrorLevel := 0 ; for safety clear in case Menu is not called in next if
				if StrLen(strThisIconFile)
					Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%, %strThisIconFile%, %intThisIconIndex%, %intIconSize%
				if (!StrLen(strThisIconFile) or ErrorLevel)
					Menu, % arrThisMenu[A_Index].MenuName, Icon, %strMenuName%
						, % objIconsFile["UnknownDocument"], % objIconsIndex["UnknownDocument"], %intIconSize%
						
				Menu, % arrThisMenu[A_Index].MenuName, UseErrorLevel, off
			}
		}
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
NextMenuShortcut(ByRef intShortcut)
;------------------------------------------------------------
{
	if (intShortcut < 10)
		strShortcut := intShortcut ; 0 .. 9
	else
		strShortcut := Chr(intShortcut + 55) ; Chr(10 + 55) = "A" .. Chr(35 + 55) = "Z"
	
	intShortcut := intShortcut + 1
	return strShortcut
}
;------------------------------------------------------------


;------------------------------------------------------------
AddMenuIcon(strMenuName, ByRef strMenuItemName, strLabel, strIconValue)
; strIconValue can be an item from strIconsMenus (eg: "Folder") or a "file,index" combo (eg: "imageres.dll,33")
;------------------------------------------------------------
{
	global intIconSize
	global blnDisplayIcons
	global objIconsFile
	global objIconsIndex
	global blnMainIsFirstColumn

	if !StrLen(strMenuItemName)
		return
	
	; The names of menus and menu items can be up to 260 characters long.
	if StrLen(strMenuItemName) > 260
		strMenuItemName := SubStr(strMenuItemName, 1, 256) . "..." ; minus one for the luck ;-)
	
	Menu, %strMenuName%, Add, %strMenuItemName%, %strLabel%
	if (blnDisplayIcons) and ((A_OSVersion <> "WIN_XP") or blnMainIsFirstColumn or (strMenuName <> lMainMenuName))
		; under Win_XP, display icons in main menu only when in first column (for other menus, this fuction is not called)
	{
		Menu, %strMenuName%, UseErrorLevel, on
		if InStr(strIconValue, ",")
			ParseIconResource(strIconValue, strIconFile, intIconIndex)
		else
		{
			strIconFile := objIconsFile[strIconValue]
			intIconIndex := objIconsIndex[strIconValue]
		}
		
		Menu, %strMenuName%, Icon, %strMenuItemName%, %strIconFile%, %intIconIndex%, %intIconSize%
		if (ErrorLevel)
			Menu, %strMenuName%, Icon, %strMenuItemName%
				, % objIconsFile["UnknownDocument"], % objIconsIndex["UnknownDocument"], %intIconSize%
		Menu, %strMenuName%, UseErrorLevel, off
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
InsertColumnBreaks:
; Based on Lexikos
; http://www.autohotkey.com/board/topic/69553-menu-with-columns-problem-with-adding-column-separator/#entry440866
;------------------------------------------------------------

VarSetCapacity(mii, cb:=16+8*A_PtrSize, 0) ; A_PtrSize is used for 64-bit compatibility.
NumPut(cb, mii, "uint")
NumPut(0x100, mii, 4, "uint") ; fMask = MIIM_FTYPE
NumPut(0x20, mii, 8, "uint") ; fType = MFT_MENUBARBREAK

for intIndex, objMenuColumnBreak in objMenuColumnBreaks
{
	pMenuHandle := GetMenuHandle(objMenuColumnBreak.MenuName) 
	DllCall("SetMenuItemInfo", "ptr", pMenuHandle, "uint", objMenuColumnBreak.MenuPosition, "int", 1, "ptr", &mii)
}

return
;------------------------------------------------------------


;------------------------------------------------------------
GetMenuHandle(strMenuName)
; from MenuIcons v2 by Lexikos
; http://www.autohotkey.com/board/topic/20253-menu-icons-v2/
;------------------------------------------------------------
{
	static pMenuDummy
	
	; v2.2: Check for !pMenuDummy instead of pMenuDummy="" in case init failed last time.
	If !pMenuDummy
	{
		Menu, menuDummy, Add
		Menu, menuDummy, DeleteAll
		
		Gui, 99:Menu, menuDummy
		; v2.2: Use LastFound method instead of window title. [Thanks animeaime.]
		Gui, 99:+LastFound
		
		pMenuDummy := DllCall("GetMenu", "uint", WinExist())
		
		Gui, 99:Menu
		Gui, 99:Destroy
		
		; v2.2: Return only after cleaning up. [Thanks animeaime.]
		if !pMenuDummy
			return 0
	}

	Menu, menuDummy, Add, :%strMenuName%
	pMenu := DllCall( "GetSubMenu", "uint", pMenuDummy, "int", 0 )
	DllCall( "RemoveMenu", "uint", pMenuDummy, "uint", 0, "uint", 0x400 )
	Menu, menuDummy, Delete, :%strMenuName%

	return pMenu
}
;------------------------------------------------------------



;========================================================================================================================
; POPUP MENU
;========================================================================================================================


;------------------------------------------------------------
PopupMenuMouse: ; default MButton
PopupMenuKeyboard: ; default #A
PopupMenuNewWindowMouse: ; default +MButton::
PopupMenuNewWindowKeyboard: ; default +#A
PopupMenuCopyLocation: ; default +#V
;------------------------------------------------------------

if !(blnMenuReady)
	return

blnCopyLocation := (A_ThisLabel = "PopupMenuCopyLocation")
if (blnCopyLocation)
{
	TrayTip, %strAppName%, %lPopupMenuCopyLocationTrayTip%, , 17
	Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)
}

blnMouse := InStr(A_ThisLabel, "Mouse")
blnNewWindow := InStr(A_ThisLabel, "New") ; used in SetMenuPosition and BuildFoldersInExplorerMenu

if !(blnNewWindow) and !(blnCopyLocation)
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

Gosub, SetMenuPosition ; sets strTargetWinId or activate the window strTargetWinId set by CanOpenFavorite

if (blnNewWindow) ; must be placed after SetMenuPosition to have a fresh strTargetWinId
	WinGetClass strTargetClass, % "ahk_id " . strTargetWinId

if (blnMouse) and (WindowIsDirectoryOpus(strTargetClass) or WindowIsTotalCommander(strTargetClass))
{
	Click ; to make sure the DOpus lister or TC pane under the mouse become active
	Sleep, 20
}

if (blnDiagMode)
{
	Diag("A_ThisLabel", A_ThisLabel)
	WinGetTitle strDiag, % "ahk_id " . strTargetWinId
	Diag("WinTitle", strDiag)
	Diag("WinId", strTargetWinId)
	Diag("Class", strTargetClass)
	Diag("ShowMenu", "Favorites Menu " . (WindowIsAnExplorer(strTargetClass) ; removed for FPconnect: or WindowIsFreeCommander(strTargetClass) 
	or WindowIsTotalCommander(strTargetClass) or WindowIsDirectoryOpus(strTargetClass)
		or (WindowIsDialog(strTargetClass, strTargetWinId) and WindowsIsVersion7OrMore()) ? "WITH" : "WITHOUT") . " Add this folder")
}

; Enable when adapted to use ::ClassID location

if (blnDisplaySpecialFolders) and (A_OSVersion = "WIN_XP")
	if (blnNewWindow)
	{
		; In case it was disabled while in a dialog box
		Menu, menuSpecialFolders, Enable, %lMenuMyComputer%
		Menu, menuSpecialFolders, Enable, %lMenuNetworkNeighborhood%
		Menu, menuSpecialFolders, Enable, %lMenuControlPanel%
		Menu, menuSpecialFolders, Enable, %lMenuRecycleBin%
	}
	else
	{
		Menu, menuSpecialFolders
			, % WindowIsConsole(strTargetClass) ; removed for FPconnect: or WindowIsFreeCommander(strTargetClass)
			or WindowIsDialog(strTargetClass, strTargetWinId) ? "Disable" : "Enable"
			, %lMenuMyComputer%
		Menu, menuSpecialFolders
			, % WindowIsConsole(strTargetClass)  ; removed for FPconnect: or WindowIsFreeCommander(strTargetClass)
			or WindowIsDialog(strTargetClass, strTargetWinId) ? "Disable" : "Enable"
			, %lMenuNetworkNeighborhood%
	
		; There is no point to navigate a dialog box or console to Control Panel or Recycle Bin
		Menu, menuSpecialFolders
			, % WindowIsConsole(strTargetClass)  ; removed for FPconnect: or WindowIsFreeCommander(strTargetClass)
			or WindowIsDialog(strTargetClass, strTargetWinId) ? "Disable" : "Enable"
			, %lMenuControlPanel%
		Menu, menuSpecialFolders
			, % WindowIsConsole(strTargetClass)  ; removed for FPconnect: or WindowIsFreeCommander(strTargetClass)
			or WindowIsDialog(strTargetClass, strTargetWinId) ? "Disable" : "Enable"
			, %lMenuRecycleBin%
	}

Gosub, BuildFoldersInExplorerMenu ; build anyway for keyboard shortcut

if (blnDisplayFoldersInExplorerMenu)
{
	Menu, %lMainMenuName%
		, % (!intExplorersIndex ? "Disable" : "Enable") ; disable Folders in Explorer menu if no Explorer
		, % BuildSpecialMenuItemName(6, lMenuFoldersInExplorer)
}

if (blnDisplayGroupMenu)
{
	Menu, menuGroups
		, % (!intExplorersIndex ? "Disable" : "Enable") ; disable Save group menu if no Explorer
		, %lMenuGroupSave%
	Menu, %lMainMenuName%
		, % (blnCopyLocation ? "Disable" : "Enable") ; disable if in Copy location menu
		, % BuildSpecialMenuItemName(7, lMenuGroup)
}

Gosub, RefreshClipboardMenu ; refresh anyway for keyboard shortcut

if (blnDisplayClipboardMenu)
{
	Menu, %lMainMenuName%
		, % (blnClipboardMenuEnable ? "Enable" : "Disable")
		, %  BuildSpecialMenuItemName(9, lMenuClipboard)
}

Menu, %lMainMenuName%
	, % (blnCopyLocation ? "Disable" : "Enable") ; disable if in Copy location menu
	, % BuildSpecialMenuItemName(5, L(lMenuSettings, strAppName) . "...")

; Enable "Add This Folder" only if the target window is an Explorer, TotalCommander,
; Directory Opus or a dialog box under WIN_7 (does not work under WIN_XP). Tested on WIN_XP and WIN_7.
; Other tests shown that WIN_8 behaves like WIN_7.
; Disable if blnCopyLocation
if (blnDiagMode)
	Diag("ShowMenu", "Favorites Menu " 
		. (WindowIsAnExplorer(strTargetClass) ; removed for FPconnect: or WindowIsFreeCommander(strTargetClass)
		or WindowIsTotalCommander(strTargetClass) or WindowIsDirectoryOpus(strTargetClass)
		or (WindowIsDialog(strTargetClass, strTargetWinId) and WindowsIsVersion7OrMore() and !(blnCopyLocation)) ? "WITH" : "WITHOUT")
		. " Add this folder")

Menu, %lMainMenuName%
	, % WindowIsAnExplorer(strTargetClass) ; removed for FPconnect: or WindowIsFreeCommander(strTargetClass)
	or WindowIsTotalCommander(strTargetClass) or WindowIsDirectoryOpus(strTargetClass)
	or (WindowIsDialog(strTargetClass, strTargetWinId) and WindowsIsVersion7OrMore() and !(blnCopyLocation)) ? "Enable" : "Disable"
	, %lMenuAddThisFolder%...

if (blnDisplayCopyLocationMenu)
	Menu, %lMainMenuName%
		, % (blnCopyLocation ? "Disable" : "Enable") ; disable if in Copy location menu
		, % lMenuCopyLocation . "..."

if !(blnDonor)
	Menu, %lMainMenuName%
		, % (blnCopyLocation ? "Disable" : "Enable") ; disable if in Copy location menu
		, % lDonateMenu . "..."

Gosub, InsertColumnBreaks

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
else ; (keyboard)
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
	global blnUseFPconnect
	
	if (strMouseOrKeyboard = "PopupMenuMouse")
		MouseGetPos, , , strWinId, strControlID
	else
	{
		strWinId := WinExist("A")
		strControlID := ""
	}
	WinGetClass strClass, % "ahk_id " . strWinId

	blnCanOpenFavorite := WindowIsAnExplorer(strClass) or WindowIsDesktop(strClass) or WindowIsConsole(strClass)
		or WindowIsDialog(strClass, strWinId) ; removed for FPconnect: or WindowIsFreeCommander(strClass)
		or (blnUseDirectoryOpus and WindowIsDirectoryOpus(strClass))
		or (blnUseTotalCommander and WindowIsTotalCommander(strClass))
		or (blnUseFPconnect and WindowIsFPconnect(strWinId))
		or WindowIsFoldersPopup(strClass)

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
	global blnOpenMenuOnTaskbar
	global blnClickOnTrayIcon
	
	blnWindowIsDesktop := (strClass = "ProgMan")
		or (strClass = "WorkerW")
		or (strClass = "Shell_TrayWnd" and (blnOpenMenuOnTaskbar or blnClickOnTrayIcon))
		or (strClass = "NotifyIconOverflowWindow")

	blnClickOnTrayIcon := 0
	; blnClickOnTrayIcon was turned on by AHK_NOTIFYICON
	; turn it off to avoid further clicks on taskbar to be accepted if blnOpenMenuOnTaskbar is off

	return blnWindowIsDesktop
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
WindowIsDialog(strClass, strWinId)
;------------------------------------------------------------
{
	return (strClass = "#32770") and !WindowIsTreeview(strWinId)
	
	; or InStr(strClass, "bosa_sdm_")
	; Removed 2014-09-27  (see http://code.jeanlalonde.ca/folderspopup/#comment-7912)
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsTreeview(strWinId)
; Disable popup menu in folder select dialog boxes (like those displayed by FileSelectFolder)
; because their Edit1 control does not react as expected in NavigateDialog.
; Signature: contains both SysTreeView321 and SHBrowseForFolder controls (tested on Win7 only)
; but NOT 100% sure this is a unique signature...
;------------------------------------------------------------
{
	WinGet, strControlsList, ControlList, ahk_id %strWinId%
	blnIsTreeView := InStr(strControlsList, "SysTreeView321") and InStr(strControlsList, "SHBrowseForFolder")
	if (blnIsTreeView)
	{
		TrayTip, %lWindowIsTreeviewTitle%, % L(lWindowIsTreeviewText, strAppName), , 18
		Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)
	}
	
	return blnIsTreeView
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsDirectoryOpus(strClass)
;------------------------------------------------------------
{
	return InStr(strClass, "dopus")
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
WindowIsFPconnect(strWinId)
;------------------------------------------------------------
{
	global strFPconnectAppPathFilename
	global strFPconnectTargetPathFilename

	if (strTargetWinId = 0)
		return false
	
    intPID := 0
    DllCall("GetWindowThreadProcessId", "UInt", strWinId, "UInt *", intPID)
    hProcess := DllCall("OpenProcess", "UInt", 0x400 | 0x10, "Int", False, "UInt", intPID)
    intPathLength = 260*2
    VarSetCapacity(strFCAppFile, intPathLength, 0)
    DllCall("Psapi.dll\GetModuleFileNameExW", "UInt", hProcess, "Int", 0, "Str", strFCAppFile, "UInt", intPathLength)
    DllCall("CloseHandle", "UInt", hProcess)
	
	SplitPath, strFCAppFile, strFCAppFile
	return (strFCAppFile = strFPconnectAppPathFilename) or (strFCAppFile = strFPconnectTargetPathFilename)
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowIsFoldersPopup(strClass)
;------------------------------------------------------------
{
	return (strClass = "JeanLalonde.ca")
}
;------------------------------------------------------------



;========================================================================================================================
; MENU ACTIONS
;========================================================================================================================


;------------------------------------------------------------
OpenFavorite:
OpenRecentFolder:
OpenFolderInExplorer:
OpenClipboard:
;------------------------------------------------------------

if (A_ThisLabel = "OpenRecentFolder")
{
	strLocation := objRecentFolders[A_ThisMenuItemPos]
	strFavoriteType := "F" ; folder
}
else if (A_ThisLabel = "OpenFolderInExplorer")
{
	If (InStr(objLocationUrlByName[A_ThisMenuItem], "::") = 1) ; A_ThisMenuItem can include the numeric shortcut
	{
		strLocation := SubStr(objLocationUrlByName[A_ThisMenuItem], 3) ; remove "::" from beginning
		strFavoriteType := "P" ; sPecial
	}
	else
	{
		if (blnDisplayMenuShortcuts)
			StringTrimLeft, strLocation, A_ThisMenuItem, 3 ; remove "&1 " from menu item
		else
			strLocation :=  A_ThisMenuItem
		strFavoriteType := "F" ; folder
	}
}
else if (A_ThisLabel = "OpenClipboard")
{
	if (blnDisplayMenuShortcuts)
		StringTrimLeft, strLocation, A_ThisMenuItem, 3 ; remove "&1 " from menu item
	else
		strLocation :=  A_ThisMenuItem

	if InStr(strLocation, "http://") = 1 or InStr(strLocation, "https://") = 1 or InStr(strLocation, "www.") = 1
		strFavoriteType := "U"
	else
	{
		strLocation :=  EnvVars(strLocation)
		SplitPath, strLocation, , , strExtension
		if StrLen(strExtension) and InStr("exe.com.bat.vbs.ahk", strExtension)
		{
			strFavoriteType := "A" ; application
			strAppArguments := "" ; make sure it is empty from previous calls
			strAppWorkingDir := "" ; make sure it is empty from previous calls
		}
		else
			strFavoriteType := (LocationIsDocument(strLocation) ? "D" : "F") ; folder or document
	}
}
else ; this is a favorite
{
	if (blnDisplayMenuShortcuts)
		StringTrimLeft, strThisMenu, A_ThisMenuItem, 3 ; remove "&1 " from menu item
	else
		strThisMenu := A_ThisMenuItem
	GetFavoriteProperties(A_ThisMenu, strThisMenu, strLocation, strFavoriteType, strAppArguments, strAppWorkingDir)
}

if (blnDiagMode)
{
	Diag("A_ThisHotkey", A_ThisHotkey)
	Diag("Label", A_ThisLabel)
	Diag("Location", strLocation)
	Diag("FavoriteType", strFavoriteType)
	Diag("TargetWinId", strTargetWinId)
	Diag("TargetClass", strTargetClass)
	Diag("blnCopyLocation", blnCopyLocation)
}

objThisSpecialFolder := objSpecialFolders[strLocation] ; save objThisSpecialFolder before expanding EnvVars
strLocation := EnvVars(strLocation) ; expand the environment variables inside location

if (blnCopyLocation) ; before or after expanding EnvVars?
{
	gosub, CopyLocation
	blnCopyLocation := false
	
	return
}

if (blnDiagMode)
	Diag("EnvVars(Path)", strLocation)

if !((strFavoriteType = "U") or (strFavoriteType = "P")) ; not URL or Special Folder, this strLocation should exist
{
	; make the location absolute based on the current working directory
	strLocation := PathCombine(A_WorkingDir, strLocation) ; expand the relative path, based on the current working directory

	if !FileExist(strLocation)
		; this favorite does not exist and it is not a special folder or an URL
		{
			Gui, 1:+OwnDialogs
			MsgBox, 0, % L(lDialogFavoriteDoesNotExistTitle, strAppName)
				, % L(lDialogFavoriteDoesNotExistPrompt, strLocation)
				, 30
			return
		}
}

if (strFavoriteType = "D" or strFavoriteType = "U") ; this is a document or an URL
{
	Run, %strLocation%
	return
}

if (strFavoriteType = "A") ; this is an application
{
	; in 5.2 double-quotes were added arount parameters - bad practice
	; was: Run, % strLocation . (StrLen(strAppArguments) ? " """ . EnvVars(strAppArguments) . """" : ""), % EnvVars(strAppWorkingDir) ; double-quotes required around strAppArguments, no dbl-quotes for strAppWorkingDir

	; No double-quotes around strAppArguments (let user add double-quotes if/when required), no dbl-quotes for strAppWorkingDir
	; From AHK doc: To pass parameters, add them immediately after the program or document name. If a parameter contains spaces, it is safest to enclose it in double quotes (even though it may work without them in some cases).
	; Also: Double-quotes are required if there are more than one parameter (ex.: [ "/a" "/b" ] is OK instead of [ "/a /b" ] or [ /a /b ]
	Run, % strLocation . (StrLen(strAppArguments) ? " " . EnvVars(strAppArguments) : ""), % EnvVars(strAppWorkingDir)
	return
}

; else this is a folder

if !StrLen(strTargetClass) or (strTargetWinId = 0) ; for situations where the target window could not be detected
	or InStr(GetIniName4Hotkey(A_ThisHotkey), "New") or WindowIsDesktop(strTargetClass)
	or ((strFavoriteType = "P") and objThisSpecialFolder.Use4NavigateExplorer = "NEW")
	
	Gosub, OpenFavoriteInNewWindow

else if WindowIsAnExplorer(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateExplorer")
	
	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4NavigateExplorer = "CLS")
			if (SubStr(strLocation, 1, 1) = "{")
				strThisLocation := "shell:::" . strLocation
			else
				strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4NavigateExplorer = "SCT")
			strThisLocation := "shell:" . objThisSpecialFolder.ShellConstantText
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "Explorer", strLocation)
			return
		}
	else
		strThisLocation := strLocation
	
	NavigateExplorer(strThisLocation, strTargetWinId)
}
else if WindowIsConsole(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateConsole")
	
	if (strFavoriteType = "P")
		if ((objThisSpecialFolder.Use4Console = "CLS") and SubStr(strLocation, 1, 1) <> "{")
			strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4Console = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4Console = "NEW")
		{
			Gosub, OpenFavoriteInNewWindow
			return
		}
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "Console (CMD)", strLocation)
			return
		}
	else
		strThisLocation := strLocation

	NavigateConsole(strThisLocation, strTargetWinId)
}
else if WindowIsFPconnect(strTargetWinId) ; must be before other third-party file managers
{
	if (blnDiagMode)
	{
		Diag("Navigate", "FPconnect")
		Diag("TargetWinId", strTargetWinId)
	}
	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4FPc = "CLS")
			if (SubStr(strLocation, 1, 1) = "{")
				strThisLocation := "shell:::" . strLocation
			else
				strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4FPc = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4FPc = "NEW")
		{
			Gosub, OpenFavoriteInNewWindow
			return
		}
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "FPconnect", strLocation)
			return
		}
	else
		strThisLocation := strLocation
	
	NavigateFPconnect(strThisLocation, strTargetWinId, strTargetClass)
}
else if WindowIsDirectoryOpus(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "DirectoryOpus")

	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4DOpus = "DOA")
			strThisLocation := "/" . objThisSpecialFolder.DOpusAlias
		else if (objThisSpecialFolder.Use4DOpus = "CLS")
			if (SubStr(strLocation, 1, 1) = "{")
				strThisLocation := "shell:::" . strLocation
			else
				strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4DOpus = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4DOpus = "NEW")
		{
			Gosub, OpenFavoriteInNewWindow
			return
		}
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "Directory Opus", strLocation)
			return
		}
	else
		strThisLocation := strLocation
	
	NavigateDirectoryOpus(strThisLocation, strTargetWinId)
}
/* removed for FPconnect: 
else if WindowIsFreeCommander(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateFreeCommander")
	
	NavigateFreeCommander(strLocation, strTargetWinId, strTargetControl)
}
*/
else if WindowIsTotalCommander(strTargetClass)
{
	if (blnDiagMode)
		Diag("Navigate", "NavigateTotalCommander")
	
	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4TC = "TCC")
			strThisLocation := objThisSpecialFolder.TCCommand
		else if (objThisSpecialFolder.Use4TC = "CLS")
			if (SubStr(strLocation, 1, 1) = "{")
				strThisLocation := "::" . strLocation
			else
				strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4TC = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4TC = "NEW")
		{
			Gosub, OpenFavoriteInNewWindow
			return
		}			
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "Total Commander", strLocation)
			return
		}
	else
		strThisLocation := strLocation
	
	if strThisLocation is not integer 
		NavigateTotalCommander(strThisLocation, strTargetWinId, strTargetControl)
	else
		NavigateTotalCommanderCommand(strThisLocation, strTargetWinId, strTargetControl)
}
else if WindowIsDialog(strTargetClass, strTargetWinId)
{
	if (blnDiagMode)
	{
		Diag("Navigate", "NavigateDialog")
		Diag("TargetClass", strTargetClass)
	}

	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4Dialog = "CLS")
			if (SubStr(strLocation, 1, 1) = "{")
				strThisLocation := "shell:::" . strLocation
			else
				strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4Dialog = "SCT")
			strThisLocation := "shell:" . objThisSpecialFolder.ShellConstantText
		else if (objThisSpecialFolder.Use4Dialog = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4Dialog = "NEW")
		{
			Gosub, OpenFavoriteInNewWindow
			return
		}
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, lOopsDialogBox, strLocation)
			return
		}
	else
		strThisLocation := strLocation
	
	NavigateDialog(strThisLocation, strTargetWinId, strTargetClass)
}
else ; we open the folder in a new window
	Gosub, OpenFavoriteInNewWindow

if (blnDiagMode)
	Diag("NavigateResult", ErrorLevel)

return
;------------------------------------------------------------


;------------------------------------------------------------
CopyLocation:
;------------------------------------------------------------

Clipboard := strLocation
TrayTip, %strAppName%, %lCopyLocationCopiedToClipboard%, , 17
Sleep, 20 ; tip from Lexikos for Windows 10 "Just sleep for any amount of time after each call to TrayTip" (http://ahkscript.org/boards/viewtopic.php?p=50389&sid=29b33964c05f6a937794f88b6ac924c0#p50389)

return
;------------------------------------------------------------


;------------------------------------------------------------
OpenFavoriteInNewWindow:
;------------------------------------------------------------

strThisLocation := ""
if (strFavoriteType = "P")
	if (objThisSpecialFolder.Use4NewExplorer = "SCT")
		strThisLocation := "shell:" . objThisSpecialFolder.ShellConstantText
	else if (SubStr(strLocation, 1, 1) = "{")
		strThisLocation := "shell:::" . strLocation
if !StrLen(strThisLocation)
	strThisLocation := strLocation ; already expanded by EnvVars() in OpenFavorite

if (blnDiagMode)
{
	Diag("OpenInNewWindowWinID", strTargetWinId)
	Diag("OpenInNewWindowClass", strTargetControl)
}

if (blnUseFPconnect)
{	
	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4FPc = "NEW")
		{
			Run, Explorer "%strThisLocation%"
			return
		}
		else if (objThisSpecialFolder.Use4FPc = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4FPc <> "CLS")
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "FPconnect", strLocation)
			return
		}

	NewFPconnect(strThisLocation, strTargetWinId, strTargetControl)
}	
else if (blnUseDirectoryOpus)
{
	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4DOpus = "DOA")
			strThisLocation := "/" . objThisSpecialFolder.DOpusAlias
		else if (objThisSpecialFolder.Use4DOpus = "CLS")
			if (SubStr(strLocation, 1, 1) = "{")
				strThisLocation := "shell:::" . strLocation
			else
				strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4DOpus = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4DOpus = "NEW")
		{
			Run, Explorer "%strThisLocation%"
			return
		}		
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "Directory Opus", strLocation)
			return
		}
	else
		strThisLocation := strLocation
	
	RunDOpusRt("/acmd Go ", strThisLocation, " " . strDirectoryOpusNewTabOrWindow) ; open in a new lister or tab
	WinActivate, ahk_class dopus.lister
}
else if (blnUseTotalCommander)
{	
	if (strFavoriteType = "P")
		if (objThisSpecialFolder.Use4TC = "TCC")
			strThisLocation := objThisSpecialFolder.TCCommand
		else if (objThisSpecialFolder.Use4TC = "CLS")
			if (SubStr(strLocation, 1, 1) = "{")
				strThisLocation := "::" . strLocation
			else
				strThisLocation := strLocation
		else if (objThisSpecialFolder.Use4TC = "AHK")
		{
			strThisLocation := objThisSpecialFolder.AHKConstant
			strThisLocation := %strThisLocation%
		}
		else if (objThisSpecialFolder.Use4TC = "NEW")
		{
			Run, Explorer "%strThisLocation%"
			return
		}		
		else
		{
			Oops(lOopsCouldNotOpenSpecialFolder, "Total Commander", strLocation)
			return
		}
	else
		strThisLocation := strLocation
	
	if strThisLocation is not integer 
		NewTotalCommander(strThisLocation, strTargetWinId, strTargetControl)
	else
		NavigateTotalCommanderCommand(strThisLocation, strTargetWinId, strTargetControl, true)
}	
else
	if (A_OSVersion = "WIN_XP")
		ComObjCreate("Shell.Application").Explore(strThisLocation)
	else
		Run, Explorer "%strThisLocation%" ; there was a bug prior to v3.3.1 because the lack of double-quotes

return
;------------------------------------------------------------


;------------------------------------------------------------
GetFavoriteProperties(strMenu, strName, ByRef strLocation, ByRef strFavoriteType, ByRef strAppArguments, ByRef strAppWorkingDir)
;------------------------------------------------------------
{
	global arrMenus
	
	Loop, % arrMenus[strMenu].MaxIndex()
		if (strName = arrMenus[strMenu][A_Index].FavoriteName)
		{
			strLocation := arrMenus[strMenu][A_Index].FavoriteLocation
			strFavoriteType := arrMenus[strMenu][A_Index].FavoriteType
			strAppArguments := arrMenus[strMenu][A_Index].AppArguments
			strAppWorkingDir := arrMenus[strMenu][A_Index].AppWorkingDir
			break
		}
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
	else if (A_ThisMenuItem = lMenuDownloads)
		strDOpusAlias := "downloads"
	
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
	else if (A_ThisMenuItem = lMenuDownloads)
		RegRead, strLocation, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, {374DE290-123F-4565-9164-39C4925E467B}
	
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
else ; this is Explorer, Console or a dialog box
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
	else if (A_ThisMenuItem = lMenuDownloads)
		; intSpecialFolder contains a sting here - this is awkward, I know, but it works and I will rework all OpenSpecialFolder very soon...
		RegRead, intSpecialFolder, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, {374DE290-123F-4565-9164-39C4925E467B}

	if (blnNewSpecialWindow)
		ComObjCreate("Shell.Application").Explore(intSpecialFolder)
		; http://msdn.microsoft.com/en-us/library/windows/desktop/bb774073%28v=vs.85%29.aspx
	else
	{
		if WindowIsAnExplorer(strTargetClass)
			NavigateExplorer(intSpecialFolder, strTargetWinId)
		else ; this is Console or a dialog box
		{
			if (intSpecialFolder = 0)
				strLocation := A_Desktop
			else if (intSpecialFolder = 5)
				strLocation := A_MyDocuments
			else if (intSpecialFolder = 39)
				RegRead, strLocation, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders, My Pictures
			else if (A_ThisMenuItem = lMenuDownloads)
				strLocation := intSpecialFolder ; still awkward, of course...
			else
				; We cannot open the special folders lMenuMyComputer, lMenuNetworkNeighborhood, lMenuControlPanel and lMenuRecycleBin
				; in the current window. Need to open an Explorer with ComObjCreate
				blnNewSpecialWindow := true
			
			if (blnNewSpecialWindow)
				ComObjCreate("Shell.Application").Explore(intSpecialFolder)
			else if WindowIsConsole(strTargetClass)
				NavigateConsole(strLocation, strTargetWinId)
			/* removed for FPconnect: 
			else if WindowIsFreeCommander(strTargetClass)
				NavigateFreeCommander(strLocation, strTargetWinId, strTargetControl)
			*/
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
GuiGroupSaveFromMenu:
GuiGroupSaveFromManage:
GuiGroupEditFromManage:
;------------------------------------------------------------

intGui2WinID := WinExist("A")

strGuiGroupSaveEditTitle := L(lGuiGroupSaveEditTitle, (A_ThisLabel = "GuiGroupEditFromManage" ? lDialogEdit : lDialogSave), strAppName, strAppVersion)
Gui, 3:New, , %strGuiGroupSaveEditTitle%
Gui, 3:Margin, 10, 10
if (blnUseColors)
	Gui, 3:Color, %strGuiWindowColor%
Gui, 3:+OwnDialogs

Gui, 3:Font, s10 w700, Verdana
Gui, 3:Add, Text, x10 y10 w670 center, % L(lGuiGroupSaveEditPrompt, (A_ThisLabel = "GuiGroupEditFromManage" ? lDialogEdit : lDialogSave))
Gui, 3:Font

Gui, 3:Add, Text, x10 y+15 w670 center, %lGuiGroupSaveSelect%
Gui, 3:Add, Link, x10 yp w100 vlblGroupSelect gGroupSelectClicked, <a>%lGuiGroupSaveDeselectAll%</a>

; Gui, 3:Add, ListView
;	, xm w680 h200 Checked Count32 -Multi LV0x10 c%strGuiListviewTextColor% Background%strGuiListviewBackgroundColor% gGuiGroupListViewEvents
;	, %lGuiGroupSaveLvHeader%|Hidden: LocationURL|IsSpecialFolder|WindowId|Position|TabId
Gui, 3:Add, ListView
	, % "xm w680 h200 Checked Count32 -Multi LV0x10 " . (blnUseColors ? "c" . strGuiListviewTextColor . " Background" . strGuiListviewBackgroundColor : "") . " gGuiGroupListViewEvents"
	, %lGuiGroupSaveLvHeader%|Hidden: LocationURL|IsSpecialFolder|WindowId|Position|TabId
Loop, 4
	LV_ModifyCol(A_Index + 3, "Right")

if (A_ThisLabel = "GuiGroupEditFromManage")
{
	strGroup := "Group-" . strGroupToEdit
	
	objIniExplorersInGroup := Object()
	Gosub, Group2Object

	for intIndex, objIniExplorerInGroup in objIniExplorersInGroup
		LV_Add("Check", objIniExplorerInGroup.Name
			, (objIniExplorerInGroup.WindowType = "DO" ? "Directory Opus" : (objIniExplorerInGroup.WindowType = "TC" ? "Total Commander" : "Windows Explorer"))
			, (objIniExplorerInGroup.MinMax = -1 ? lDialogMinimized : (objIniExplorerInGroup.MinMax = "1" ? lDialogMaximized : lDialogNormal))
			, (objIniExplorerInGroup.MinMax = 0 ? objIniExplorerInGroup.Left : "-")
			, (objIniExplorerInGroup.MinMax = 0 ? objIniExplorerInGroup.Top : "-")
			, (objIniExplorerInGroup.MinMax = 0 ? objIniExplorerInGroup.Width : "-")
			, (objIniExplorerInGroup.MinMax = 0 ? objIniExplorerInGroup.Height : "-")
			, (objIniExplorerInGroup.Pane ? objIniExplorerInGroup.Pane : "-")
			, objIniExplorerInGroup.LocationURL, objIniExplorerInGroup.IsSpecialFolder, objIniExplorerInGroup.WindowId, objIniExplorerInGroup.Position, objIniExplorerInGroup.TabId)
	IniRead, blnReplaceWhenRestoringThisGroup, %strIniFile%, %strGroup%, ReplaceWhenRestoringThisGroup, 0 ; false if not found
}
else
{
	for intIndex, objGroupMenuExplorer in objFoldersInExplorers
	{
		arrExplorerPosition := objGroupMenuExplorer.Position
		StringSplit, arrExplorerPosition, arrExplorerPosition, "|"
		LV_Add("Check", objGroupMenuExplorer.Name
			, (objGroupMenuExplorer.WindowType = "DO" ? "Directory Opus" : (objGroupMenuExplorer.WindowType = "TC" ? "Total Commander" : "Windows Explorer"))
			, (objGroupMenuExplorer.MinMax = -1 ? lDialogMinimized : (objGroupMenuExplorer.MinMax = "1" ? "Maximized" : "Normal"))
			, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition1 : "-")
			, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition2 : "-")
			, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition3 : "-")
			, (objGroupMenuExplorer.MinMax = 0 ? arrExplorerPosition4 : "-")
			, (objGroupMenuExplorer.Pane ? objGroupMenuExplorer.Pane : "-")
			, objGroupMenuExplorer.LocationURL, objGroupMenuExplorer.IsSpecialFolder, objGroupMenuExplorer.WindowId, objGroupMenuExplorer.Position, objGroupMenuExplorer.TabId)
	}
	blnReplaceWhenRestoringThisGroup := False ; default for new group
}

LV_Modify(1, "Select Focus")
Loop, 12
	LV_ModifyCol(A_Index, "AutoHdr")
Loop, % (blnUseDirectoryOpus ? 5 : 6)
	LV_ModifyCol(A_Index + (blnUseDirectoryOpus ? 8 : 7), 0) ; hide 8th-12th columns

Gui, 3:Add, Text, x10 y+10 w250, %lGuiGroupSaveShortName%
Gui, 3:Add, Edit, x10 y+5 w250 Limit248 vstrGroupSaveName gGuiGroupSaveNameChanged section, % (A_ThisLabel = "GuiGroupEditFromManage" ? strGroupToEdit : "") ; maximum length of ini section name is 255 ("Group-" is prepended)

Gui, 3:Add, Text, x10 y+5 w250, %lGuiGroupSaveSelectExisting%
Gui, 3:Add, DropDownList, x10 y+5 w250 vdrpGroupsOverwriteList gGuiGroupSaveOverwriteListChanged, %lGuiGroupSaveNewGroup%|%strGroups%
GuiControl, ChooseString, drpGroupsOverwriteList, %lGuiGroupSaveNewGroup%

Gui, 3:Add, Button, y+20 vbtnGroupSave gButtonGroupSave, %lGuiSave%
Gui, 3:Add, Button, yp vbtnGroupSaveCancel gButtonGroupSaveCancel, %lGuiCancel%

Gui, 3:Add, Text, ys x300, %lGuiGroupSaveRestoreOption%
Gui, 3:Add, Radio, % "x+10 yp section vblnRadioAddWindows " . (blnReplaceWhenRestoringThisGroup ? "" : "checked"), %lGuiGroupSaveAddWindowsLabel%
Gui, 3:Add, Radio, % "xs y+5 vblnRadioReplaceWindows " . (blnReplaceWhenRestoringThisGroup ? "checked" : ""), %lGuiGroupSaveReplaceWindowsLabel%

GuiCenterButtons(strGuiGroupSaveEditTitle, 10, 5, 20, "btnGroupSave", "btnGroupSaveCancel")

GuiControl, Focus, strGroupSaveName
if (A_ThisLabel = "GuiGroupEditFromManage")
	SendInput, ^a
GuiControl, +Default, btnGroupSave
Gui, 3:Show, AutoSize Center

if InStr(A_ThisLabel, "FromManage")
	Gui, 2:+Disabled

objIniExplorersInGroup :=

return
;------------------------------------------------------------


;------------------------------------------------------------
GroupSelectClicked:
;------------------------------------------------------------

GuiControlGet, strGroupSelect, , lblGroupSelect
Loop, % LV_GetCount()
	LV_Modify(A_Index, InStr(strGroupSelect, lGuiGroupSaveDeselectAll) ? "-Check" : "Check")

GuiControl, , lblGroupSelect, % "<a>" . (InStr(strGroupSelect, lGuiGroupSaveDeselectAll) ? lGuiGroupSaveSelectAll : lGuiGroupSaveDeselectAll) . "</a>"

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupListViewEvents:
;------------------------------------------------------------
if (A_GuiEvent = "ColClick")
	LV_ModifyCol(A_EventInfo, "Sort")

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupSaveOverwriteListChanged:
;------------------------------------------------------------
Gui, 3:Submit, NoHide

if (drpGroupsOverwriteList <> lGuiGroupSaveNewGroup) and StrLen(strGroupSaveName)
	GuiControl, , strGroupSaveName

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupSaveNameChanged:
;------------------------------------------------------------
Gui, 3:Submit, NoHide

if (drpGroupsOverwriteList <> lGuiGroupSaveNewGroup) and StrLen(strGroupSaveName)
	GuiControl, ChooseString, drpGroupsOverwriteList, %lGuiGroupSaveNewGroup%

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonGroupSave:
;------------------------------------------------------------
Gui, 3:Submit, NoHide
Gui, 3:+OwnDialogs

if (drpGroupsOverwriteList <> lGuiGroupSaveNewGroup)
	strGroupSaveName := drpGroupsOverwriteList

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
	Sort, strGroups, D|
	IniWrite, %strGroups%, %strIniFile%, Global, Groups
}

IniWrite, %blnRadioReplaceWindows%, %strIniFile%, Group-%strGroupSaveName%, ReplaceWhenRestoringThisGroup
intRow := 0
Loop
{
	intRow := LV_GetNext(intRow, "Checked")
    if !(intRow)
        break
	IniWrite, % objFoldersInExplorers[intRow].Name
		. "|" . objFoldersInExplorers[intRow].WindowType
		. "|" . objFoldersInExplorers[intRow].MinMax
		. "|" . objFoldersInExplorers[intRow].Position
		. "|" . objFoldersInExplorers[intRow].Pane
		. "|" . objFoldersInExplorers[intRow].WindowId
		. "|" . objFoldersInExplorers[intRow].IsSpecialFolder
		. "|" . objFoldersInExplorers[intRow].LocationURL
		, %strIniFile%, Group-%strGroupSaveName%, Explorer%A_Index%
}

StringReplace, strUseMenuLoad, lMenuGroup, &, , All
MsgBox, 0, % L(lGuiGroupSaveTitle, lMenuGroupSave, strAppName)
	, % L(lGuiGroupSaved, strGroupSaveName, strUseMenuLoad)
	, 30

strGroupSaveName := ""

Gosub, BuildGroupMenu
Gosub, ButtonGroupSaveCancel

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonGroupSaveCancel:
;------------------------------------------------------------

Gosub, 3GuiClose

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

blnGroupLoadExplorerError := False
strGroupLoadDiag := "DateTime`tType`tData`n"

DiagGroupLoad("DIAGNOSTIC FILE", "Copy this in an email and send it to ahk@jeanlalonde.ca with any detail that could help make a diagnostic of the Group Load issue. Thanks.")
DiagGroupLoad("AppName", strAppName)
DiagGroupLoad("AppVersion", strAppVersion)
DiagGroupLoad("A_ScriptFullPath", A_ScriptFullPath)
DiagGroupLoad("A_WorkingDir", A_WorkingDir)
DiagGroupLoad("A_AhkVersion", A_AhkVersion)
DiagGroupLoad("A_OSVersion", A_OSVersion)
DiagGroupLoad("GetOsVersion()", GetOsVersion())
DiagGroupLoad("A_Is64bitOS", A_Is64bitOS)
DiagGroupLoad("A_Language", A_Language)
DiagGroupLoad("A_IsAdmin", A_IsAdmin)

FileRead, strDiag, %strIniFile%
StringReplace, strDiag, strDiag, `", `"`"
DiagGroupLoad("IniFile", """" . strDiag . """`n")

if (A_ThisLabel = "GroupLoad")
{
	if (blnDisplayMenuShortcuts)
		StringTrimLeft, strGroup, A_ThisMenuItem, 3 ; remove "&1 " from menu item
	else
		strGroup :=  A_ThisMenuItem
	strGroup := "Group-" . strGroup
	StringLeft, strGroup, strGroup, % InStr(strGroup, "(") - 2 ; remove " (add)" or " (replace)"
}
else
	strGroup := "Group-" . strSelectedGroup
DiagGroupLoad("strGroup", strGroup)

IniRead, blnReplaceWhenRestoringThisGroup, %strIniFile%, %strGroup%, ReplaceWhenRestoringThisGroup, 0 ; false if not found
DiagGroupLoad("ReplaceWhenRestoringThisGroup", ReplaceWhenRestoringThisGroup)

if (blnReplaceWhenRestoringThisGroup)
	Gosub, CloseBeforeRestoringGroup

objIniExplorersInGroup := Object()
Gosub, Group2Object
intTotalWindowsInIni := objIniExplorersInGroup.MaxIndex()
intActualWindowInIni := 1

while, intExplorer := WindowOfType("EX") ; returns the index of the first Explorer saved window in the group
{
	blnGroupLoadError := False
	if !StrLen(objIniExplorersInGroup[intExplorer].LocationURL) ; for compatibility before v3.9.7
		objIniExplorersInGroup[intExplorer].LocationURL := objIniExplorersInGroup[intExplorer].Name
	DiagGroupLoad("objIniExplorersInGroup[intExplorer].LocationURL", objIniExplorersInGroup[intExplorer].LocationURL)
	
	Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni%
	DiagGroupLoad("objIniExplorersInGroup", intActualWindowInIni . " " . lGuiGroupOf . " " . intTotalWindowsInIni)
	
	if (objIniExplorersInGroup[intExplorer].IsSpecialFolder)
		strExplorerLocationOrClassId := "shell:::" . objClassIdOrPathByDefaultName[objIniExplorersInGroup[intExplorer].Name]
	else
		strExplorerLocationOrClassId := objIniExplorersInGroup[intExplorer].Name
	DiagGroupLoad("strExplorerLocationOrClassId", strExplorerLocationOrClassId)
	DiagGroupLoad("objIniExplorersInGroup[intExplorer].LocationURL", objIniExplorersInGroup[intExplorer].LocationURL)
	DiagGroupLoad("objIniExplorersInGroup[intExplorer].MinMax", objIniExplorersInGroup[intExplorer].MinMax)
	
	strNewWindowId := GetExplorerIdIfContains(objIniExplorersInGroup[intExplorer].LocationURL) ; returns the ID of the Explorer containing LocationURL or an empty string if no Explorer contain LocationURL

	if StrLen(strNewWindowId) ; then we activate the existing Explorer
		WinActivate, ahk_id %strNewWindowId%
	else ; then we create a new Explorer with this LocationURL
	{
		Loop
		{
			if (A_Index > 1)
				Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni% (take %A_Index%)
			if (A_Index > 3)
			{
				blnGroupLoadError := True
				blnGroupLoadExplorerError := True
				Break
			}
			strExplorerIDsBefore := GetExplorersIDs() ;  get a list of existing Explorer windows before launching this new Explorer
			DiagGroupLoad("strExplorerIDsBefore", strExplorerIDsBefore)
			DiagGroupLoad("GroupLoadRun Take", A_Index)
			
			Run, % "explorer.exe """ . strExplorerLocationOrClassId . """",
				, % (objIniExplorersInGroup[intExplorer].MinMax = -1 ? "Min" : (objIniExplorersInGroup[intExplorer].MinMax = 1 ? "Max" : ""))
				
			Loop
			{
				if (A_Index > 10)
				{
					Oops(lDialogGroupLoadErrorLoading, strExplorerLocationOrClassId)
					blnGroupLoadError := True
					blnGroupLoadExplorerError := True
					Break
				}
				Sleep, 500 ; was 200 before v5.1.1
				strExplorerIDsAfter := GetExplorersIDs() ;  get an updated list of existing Explorer windows
				DiagGroupLoad("strExplorerIDsAfter Take " . A_Index, strExplorerIDsAfter)
				strNewWindowId := GetNewExplorer(strExplorerIDsBefore, strExplorerIDsAfter) ; check if we have a new Explorer window
				DiagGroupLoad("strNewWindowId", strNewWindowId)
				blnGroupLoadWindowIdFound := StrLen(strNewWindowId)
				if (blnGroupLoadWindowIdFound)
					Break ; if we have a new window, exit loops 1, 2 and WinMove it
			}
			if (blnGroupLoadWindowIdFound)
				Break ; if we have a new window, exit loop 2 and WinMove it
		}
	}
	
	if !(blnGroupLoadError)
	{
		WinWait, ahk_id %strNewWindowId%, , 5
		Sleep, 300
		
		if !(objIniExplorersInGroup[intExplorer].MinMax) ; window normal, not min nor max
			WinMove, ahk_id %strNewWindowId%,
				, % objIniExplorersInGroup[intExplorer].Left
				, % objIniExplorersInGroup[intExplorer].Top
				, % objIniExplorersInGroup[intExplorer].Width
				, % objIniExplorersInGroup[intExplorer].Height
		else
			if (objIniExplorersInGroup[intExplorer].MinMax = -1)
				WinMinimize, ahk_id %strNewWindowId%
			else
				WinMaximize, ahk_id %strNewWindowId%
	}
	
	objIniExplorersInGroup.Remove(intExplorer) ; remove the first Explorer saved window from the group
	intActualWindowInIni := intActualWindowInIni + 1
}

while, intDOWindow := WindowOfType("DO") ; returns the index of the first DOpus saved window in the group
{
	blnGroupLoadError := False
	strFirstWindowId := objIniExplorersInGroup[intDOWindow].WindowId
	strNewWindowId := ""
	
	while, intDOIndex := DOpusWindowOfId(strFirstWindowId) ; returns the paths of the first group of DOpus windows (grouped by WindowsId)
	{
		if !StrLen(strNewWindowId) ; the window does not exist, create it with the first path in the pane 1
		{
			Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni%
			
			intWinIdBeforeRun := WinExist("A")
			RunDOpusRt("/acmd Go ", objIniExplorersInGroup[intDOWindow].Name, " NEW=nodual") ; open in a new lister
			Loop
			{
				if (A_Index > 25)
				{
					Oops(lDialogGroupLoadErrorLoading, objIniExplorersInGroup[intDOWindow].Name)
					blnGroupLoadError := True
					Break
				}
				Sleep, 200
				strNewWindowId := WinExist("A")		
			} until (intWinIdBeforeRun <> strNewWindowId)

			if !(blnGroupLoadError)
			{
				WinWait, ahk_id %strNewWindowId%, , 5
				Sleep, 300
				
				if !(objIniExplorersInGroup[intDOWindow].MinMax) ; window normal, not min nor max
					WinMove, ahk_id %strNewWindowId%,
						, % objIniExplorersInGroup[intDOWindow].Left
						, % objIniExplorersInGroup[intDOWindow].Top
						, % objIniExplorersInGroup[intDOWindow].Width
						, % objIniExplorersInGroup[intDOWindow].Height
				else
					if (objIniExplorersInGroup[intDOWindow].MinMax = -1)
						WinMinimize, ahk_id %strNewWindowId%
					else
						WinMaximize, ahk_id %strNewWindowId%
			}

			objIniExplorersInGroup.Remove(intDOWindow) ; remove the first DOpus saved window from the group
			intActualWindowInIni := intActualWindowInIni + 1
		}
		else ; the window exist, add other path of pane tabs or pane 2
		{
			intSleepAfterDOpusRtCommand := 200
			
			while, intDOIndexPane := DOpusWindowOfPane(1, strFirstWindowId) ; returns the paths of the windows of the left pane for this group of windows
			{
				Tooltip, %intActualWindowInIni% %lGuiGroupOf% %intTotalWindowsInIni%
				RunDOpusRt("/acmd Go ", objIniExplorersInGroup[intDOIndexPane].Name, " NEWTAB") ; open in a new tab of pane 1
				Sleep, %intSleepAfterDOpusRtCommand%
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
					Sleep, %intSleepAfterDOpusRtCommand%
					blnNewPane2 := false
				}
				else
				{
					RunDOpusRt("/acmd Go ", objIniExplorersInGroup[intDOIndexPane].Name, " OPENINRIGHT NEWTAB") ; open in a new tab of pane 2
					Sleep, %intSleepAfterDOpusRtCommand%
				}
				objIniExplorersInGroup.Remove(intDOIndexPane) ; remove the first DOpus saved window from the group
				intActualWindowInIni := intActualWindowInIni + 1
			}
		}
	}
}
Tooltip ; clear tooltip

objIniExplorersInGroup :=

/*
if (blnGroupLoadExplorerError)
{
	MsgBox, 20, %strAppName% Group Load Diagnostic, Following the error you encountered, do you want to copy DIAGNOSTIC info to your clipboard and send it to HELP the developer?
	IfMsgBox, Yes
	{
		Clipboard := strGroupLoadDiag
		MsgBox, 0, %strAppName% Group Load Diagnostic, Paste your CLIPBOARD in an EMAIL and send it to the address indicated in the first line of the info.`n`nThank you!
	}
}
*/
strGroupLoadDiag := ""

return
;------------------------------------------------------------


;------------------------------------------------------------
CloseBeforeRestoringGroup:
;------------------------------------------------------------

intSleepTime := 67 ; for visual effect only...
Tooltip, %lGuiGroupClosing%

strWindowsId := ""
for objExplorer in ComObjCreate("Shell.Application").Windows
{
	; do not close in this loop as it mess up the handlers
	strType := ""
	try strType := objExplorer.Type ; Gets the type name of the contained document object. "Document HTML" for IE windows. Should be empty for file Explorer windows.
	strWindowID := ""
	try strWindowID := objExplorer.HWND ; Try to get the handle of the window. Some ghost Explorer in the ComObjCreate may return an empty handle
	if !StrLen(strType) and StrLen(strWindowID) ; strType must be empty and strWindowID must not be empty
		strWindowsId := strWindowsId . objExplorer.HWND . "|"
}
StringTrimRight, strWindowsId, strWindowsId, 1
Loop, Parse, strWindowsId, |
{
	WinClose, ahk_id %A_LoopField%
	Sleep, %intSleepTime%
}

/*
if (blnUseTotalCommander)
{
	WinGet, arrIDs, List, ahk_class TTOTAL_CMD
	Loop, %arrIDs%
	{
		WinClose, % "ahk_id " . arrIDs%A_Index%
		Sleep, %intSleepTime%
	}
}
*/

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

	;  1	objFoldersInExplorers[intRow].Name
	;  2	objFoldersInExplorers[intRow].WindowType
	;  3	objFoldersInExplorers[intRow].MinMax
	;  4	objFoldersInExplorers[intRow].Left
	;  5	objFoldersInExplorers[intRow].Top
	;  6	objFoldersInExplorers[intRow].Width
	;  7	objFoldersInExplorers[intRow].Height
	;  8	objFoldersInExplorers[intRow].Pane
	;  9	objFoldersInExplorers[intRow].WindowId
	; 10	objFoldersInExplorers[intRow].IsSpecialFolder
	; 11	objFoldersInExplorers[intRow].LocationURL
	
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
	objIniEntry.LocationURL := arrThisExplorer11
	
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


;------------------------------------------------------------
GetExplorersIDs()
;------------------------------------------------------------
{
	strExplorerIDs := ""
	objShell := ComObjCreate("Shell.Application")
	for objExplorer in objShell.Windows
	{
		strType := ""
		try strType := objExplorer.Type ; Gets the type name of the contained document object. "Document HTML" for IE windows. Should be empty for file Explorer windows.
		strWindowID := ""
		try strWindowID := objExplorer.HWND ; Try to get the handle of the window. Some ghost Explorer in the ComObjCreate may return an empty handle
		if !StrLen(strType) and StrLen(strWindowID) ; strType must be empty and strWindowID must not be empty
			strExplorerIDs := strExplorerIDs . objExplorer.HWND . "|"
	}
	objExplorer := "" ; release object
	objShell := "" ; release object
	return strExplorerIDs
}
;------------------------------------------------------------


;------------------------------------------------------------
GetExplorerIdIfContains(strLocationURL)
;------------------------------------------------------------
{
	strExplorerID := ""
	for objExplorer in ComObjCreate("Shell.Application").Windows
	{
		strType := ""
		try strType := objExplorer.Type ; Gets the type name of the contained document object. "Document HTML" for IE windows. Should be empty for file Explorer windows.
		strWindowID := ""
		try strWindowID := objExplorer.HWND ; Try to get the handle of the window. Some ghost Explorer in the ComObjCreate may return an empty handle
		if !StrLen(strType) and StrLen(strWindowID) ; strType must be empty and strWindowID must not be empty
		{
			if !StrLen(objExplorer.LocationURL) ; empty for special folders like Recycle bin
			{
				if (objExplorer.Document.Folder.Self.Path = strLocationURL)
				{
					strExplorerID := strWindowID
					DiagGroupLoad("Special folder LocationURL Is", objExplorer.Document.Folder.Self.Path)
					break
				}
			}
			else if (objExplorer.LocationURL = strLocationURL)
			{
				strExplorerID := strWindowID
				DiagGroupLoad("LocationURL Is", objExplorer.LocationURL)
				break
			}
			else
				DiagGroupLoad("LocationURL Is Not", objExplorer.LocationURL)
		}
	}
	return strExplorerID
}
;------------------------------------------------------------


;------------------------------------------------------------
GetNewExplorer(strIDsBefore, strIDsAfter)
;------------------------------------------------------------
{
	Loop, Parse, strIDsAfter, |
		if !InStr(strIDsBefore, A_LoopField . "|")
			return A_LoopField
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
				SendInput, {F4}{Esc}{Raw}%varPath%`n ; was Run, Explorer "%varPath%"
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


/*
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
*/


;------------------------------------------------------------
NavigateTotalCommander(strLocation, strWinId, strControl)
;------------------------------------------------------------
{
	global strTotalCommanderPath
	
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
	
	; strTotalCommanderNewTabOrWindow in ini file should contain "/O /T" to open in an new tab of the existing file list (default), or "/N" to open in a new file list
	Run, %strTotalCommanderPath% %strTotalCommanderNewTabOrWindow% /S "/L=%strLocation%" ; /L= left pane of the new window
}
;------------------------------------------------------------



;------------------------------------------------------------
NavigateTotalCommanderCommand(strLocation, strWinId, strControl, blnNewSpecialWindow := false)
;------------------------------------------------------------
{
	global strTotalCommanderPath
	global strTotalCommanderNewTabOrWindow
	
	intTCCommand := strLocation
	
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
		Sleep, 100 ; wait to improve SendMessage reliability
	}
	SendMessage, 0x433, %intTCCommand%, , , ahk_class TTOTAL_CMD
	Sleep, 100 ; wait to improve SendMessage reliability
	WinActivate, ahk_class TTOTAL_CMD
}
;------------------------------------------------------------


;------------------------------------------------------------
NavigateFPconnect(strLocation, strWinId, strControl)
;------------------------------------------------------------
{
	global strFPconnectPath
	
	if (WinExist("A") <> strWinId) ; in case that some window just popped out, and initialy active window lost focus
	{
		WinActivate, ahk_id %strWinId% ; we'll activate initialy active window
		Sleep, 200
	}
	Run, %strFPconnectPath% %strLocation%
}
;------------------------------------------------------------


;------------------------------------------------------------
NewFPconnect(strLocation, strWinId, strControl)
;------------------------------------------------------------
{
	global strFPconnectPath

	Run, %strFPconnectPath% %strLocation% /new
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
			Run, Explorer "%strLocation%"
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
ShowIniFile:
;------------------------------------------------------------

Run, %strIniFile%

return
;------------------------------------------------------------


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
Check4Update:
;------------------------------------------------------------

strUrlCheck4Update := "http://code.jeanlalonde.ca/ahk/folderspopup/latest-version/latest-version-3.php"
strAppLandingPage := "http://code.jeanlalonde.ca/folderspopup/"
strBetaLandingPage := "http://code.jeanlalonde.ca/ahk/folderspopup/check4update-beta-redirect.html"

if (blnDiagMode)
{
	Diag("Check4Update strAppLandingPage", strAppLandingPage)
	Diag("Check4Update strBetaLandingPage", strBetaLandingPage)
}

Gui, 1:+OwnDialogs

if (A_ThisMenuItem <> lMenuUpdate)
{
	if Time2Donate(intStartups, blnDonor)
	{
		MsgBox, 36, % L(lDonateCheckTitle, intStartups, strAppName), % L(lDonateCheckPrompt, strAppName, intStartups)
		IfMsgBox, Yes
			Gosub, GuiDonate
	}
	IniWrite, % (intStartups + 1), %strIniFile%, Global, Startups
}

blnSetup := (FileExist(A_ScriptDir . "\_do_not_remove_or_rename.txt") = "" ? 0 : 1)

strLatestVersions := Url2Var(strUrlCheck4Update
	. "?v=" . strCurrentVersion
	. "&os=" . GetOsVersion() ; A_OSVersion
	. "&is64=" . A_Is64bitOS
    . "&setup=" . (blnSetup)
				+ (2 * (blnDonor ? 1 : 0))
				+ (4 * (blnUseDirectoryOpus ? 1 : 0))
				+ (8 * (blnUseTotalCommander ? 1 : 0))
				+ (16 * (blnUseFPconnect ? 1 : 0))
    . "&lsys=" . A_Language
    . "&lfp=" . strLanguageCode)
if !StrLen(strLatestVersions)
	if (A_ThisMenuItem = lMenuUpdate)
	{
		Oops(lUpdateError)
		return ; an error occured during ComObjCreate
	}

strLatestVersions := SubStr(strLatestVersions, InStr(strLatestVersions, "[[") + 2) 
strLatestVersions := SubStr(strLatestVersions, 1, InStr(strLatestVersions, "]]") - 1) 
strLatestVersions := Trim(strLatestVersions, "`n`l") ; remove en-of-line if present
Loop, Parse, strLatestVersions, , 0123456789.| ; strLatestVersions should only contain digits, dots and one pipe (|) between prod and beta versions
	; if we get here, the content returned by the URL above is wrong
	if (A_ThisMenuItem <> lMenuUpdate)
		return ; return silently
	else
	{
		Oops(lUpdateError) ; return with an error message
		return
	}

StringSplit, arrLatestVersions, strLatestVersions, |
strLatestVersionProd := arrLatestVersions1
strLatestVersionBeta := arrLatestVersions2

if (blnDiagMode)
{
	Diag("Check4Update strCurrentVersion", strCurrentVersion)
	Diag("Check4Update strLatestVersionProd", strLatestVersionProd)
	Diag("Check4Update strLatestVersionBeta", strLatestVersionBeta)
	Diag("Check4Update strLatestSkippedProd", strLatestSkippedProd)
	Diag("Check4Update strLatestSkippedBeta", strLatestSkippedBeta)
	Diag("Check4Update strLatestUsedProd", strLatestUsedProd)
	Diag("Check4Update strLatestUsedBeta", strLatestUsedBeta)
}

Gui, 1:+OwnDialogs

if (strLatestUsedBeta <> "0.0")
{
	if (FirstVsSecondIs(strLatestVersionBeta, strCurrentVersion) = 1) ; latest beta more recent than current
		and ((FirstVsSecondIs("6.0.0", strLatestVersionBeta) = 1) or !InStr("WIN_VISTA|WIN_2003|WIN_XP|WIN_2000", A_OSVersion)) ; latest beta under 6.0.0 or OS Win_7+
	{
		SetTimer, Check4UpdateChangeButtonNames, 50

		MsgBox, 3, % L(lUpdateTitle, strAppName) ; do not add BETA to keep buttons rename working
			, % L(lUpdatePromptBeta, strAppName, strCurrentVersion, strLatestVersionBeta)
		IfMsgBox, Yes
			Run, %strBetaLandingPage%
		IfMsgBox, Cancel ; Remind me
			IniWrite, 0.0, %strIniFile%, Global, LatestVersionSkippedBeta
		IfMsgBox, No
		{
			IniWrite, %strLatestVersionBeta%, %strIniFile%, Global, LatestVersionSkippedBeta
			MsgBox, 4, % L(lUpdateTitle, strAppName . " BETA"), %lUpdatePromptBetaContinue%
			IfMsgBox, No
				IniWrite, 0.0, %strIniFile%, Global, LastVersionUsedBeta
		}
	}
}

if (FirstVsSecondIs(strLatestSkippedProd, strLatestVersionProd) >= 0 and (A_ThisMenuItem <> lMenuUpdate))
	return

if (FirstVsSecondIs(strLatestVersionProd, strCurrentVersion) = 1) ; latest prod more recent than current
	and ((FirstVsSecondIs("6.0.0", strLatestVersionProd) = 1) or !InStr("WIN_VISTA|WIN_2003|WIN_XP|WIN_2000", A_OSVersion)) ; latest prod under 6.0.0 or OS Win7+
{
	if FirstVsSecondIs("6.0.0", strLatestVersionProd) = 1 ; this is FoldersPopup update
		strThisUpdatePrompt := L(lUpdatePrompt, strAppName, strCurrentVersion, strLatestVersionProd)
	else ; this is a Quick Access Popup update
		if (A_OSVersion = "WIN_XP")
			return
		else
			strThisUpdatePrompt := L(lUpdatePromptQAP, strAppName, strCurrentVersion, strLatestVersionProd, "Quick Access Popup")
	
	SetTimer, Check4UpdateChangeButtonNames, 50
	MsgBox, 3, % L(lUpdateTitle, strAppName), %strThisUpdatePrompt%
	
	IfMsgBox, Yes
		Run, %strAppLandingPage%
	IfMsgBox, No
		IniWrite, %strLatestVersionProd%, %strIniFile%, Global, LatestVersionSkipped ; do not add "Prod" to ini variable for backward compatibility
	IfMsgBox, Cancel ; Remind me
		IniWrite, 0.0, %strIniFile%, Global, LatestVersionSkipped ; do not add "Prod" to ini variable for backward compatibility
}
else if (A_ThisMenuItem = lMenuUpdate)
{
	strThisAppName := strAppName ; requires to keep buttons rename working 
	MsgBox, 4, % L(lUpdateTitle, strThisAppName), % L(lUpdateYouHaveLatest, strAppVersion, strThisAppName)
	IfMsgBox, Yes
	{
		if (blnDiagMode)
		{
			Diag("Check4Update lMenuUpdate strCurrentBranch", strCurrentBranch)
			Diag("Check4Update lMenuUpdate strAppLandingPage", strAppLandingPage)
			Diag("Check4Update lMenuUpdate strBetaLandingPage", strBetaLandingPage)
		}
		Run, %strAppLandingPage%
	}
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
Check4UpdateChangeButtonNames:
;------------------------------------------------------------

IfWinNotExist, % L(lUpdateTitle, strAppName)
    return  ; Keep waiting.
SetTimer, Check4UpdateChangeButtonNames, Off
WinActivate 
ControlSetText, Button3, %lUpdateButtonRemind%

return
;------------------------------------------------------------


;------------------------------------------------------------
Time2Donate(intStartups, blnDonor)
;------------------------------------------------------------
{
	return !Mod(intStartups, 20) and (intStartups > 40) and !(blnDonor)
}
;------------------------------------------------------------



;========================================================================================================================
; MAIN GUI
;========================================================================================================================

;------------------------------------------------------------
InitGuiControls:
;------------------------------------------------------------

objGuiControls := Object()

; Order of controls important to avoid drawgins gliches when resizing

InsertGuiControlPos("lnkGuiDropHelpClicked",	 -88, -130)
InsertGuiControlPos("lnkGuiHotkeysHelpClicked",	  40, -130)

InsertGuiControlPos("picGuiOptions",			 -44,   10, true) ; true = center
InsertGuiControlPos("picGuiAddFavorite",		 -44,  122, true)
InsertGuiControlPos("picGuiEditFavorite",		 -44,  199, true)
InsertGuiControlPos("picGuiRemoveFavorite",		 -44,  274, true)
InsertGuiControlPos("picGuiGroupsManage",		 -44, -150, true, true) ; true = center, true = draw
InsertGuiControlPos("picGuiDonate",				  50,  -62, true, true)
InsertGuiControlPos("picGuiHelp",				 -44,  -62, true, true)
InsertGuiControlPos("picGuiAbout",				-104,  -62, true, true)

InsertGuiControlPos("picAddColumnBreak",		  10,  230)
InsertGuiControlPos("picAddSeparator",			  10,  200)
InsertGuiControlPos("picMoveFavoriteDown",		  10,  170)
InsertGuiControlPos("picMoveFavoriteUp",		  10,  140)
InsertGuiControlPos("picPreviousMenu",			  10,   84)
InsertGuiControlPos("picSortFavorites",			  10, -165)
InsertGuiControlPos("picUpMenu",				  25,   84)

InsertGuiControlPos("btnGuiSave",				   0,  -90, , true)				
InsertGuiControlPos("btnGuiCancel",				   0,  -90, , true)

InsertGuiControlPos("drpMenusList",				  40,   84)

InsertGuiControlPos("lblGuiDonate",				  50,  -20, true)
InsertGuiControlPos("lblGuiAbout",				-104,  -20, true)
InsertGuiControlPos("lblGuiHelp",				 -44,  -20, true)
InsertGuiControlPos("lblAppName",				  10,   10)
InsertGuiControlPos("lblAppTagLine",			  10,   42)
InsertGuiControlPos("lblGuiAddFavorite",		 -44,  172, true)
InsertGuiControlPos("lblGuiEditFavorite",		 -44,  249, true)
InsertGuiControlPos("lblGuiOptions",			 -44,   45, true)
InsertGuiControlPos("lblGuiRemoveFavorite",		 -44,  324, true)
InsertGuiControlPos("lblSubmenuDropdownLabel",	  40,   66)
InsertGuiControlPos("lblGuiGroupsManage",		 -44,  -95, true)

InsertGuiControlPos("lvFavoritesList",			  40,  115)

return
;------------------------------------------------------------


;------------------------------------------------------------
InsertGuiControlPos(strControlName, intX, intY, blnCenter := false, blnDraw := false)
;------------------------------------------------------------
{
	global objGuiControls

	objGuiControl := Object()
	objGuiControl.Name := strControlName
	objGuiControl.X := intX
	objGuiControl.Y := intY
	objGuiControl.Center := blnCenter
	objGuiControl.Draw := blnDraw
	
	objGuiControls.Insert(objGuiControl)
}
;------------------------------------------------------------


;------------------------------------------------------------
GuiSize:
;------------------------------------------------------------

if (A_EventInfo = 1)  ; The window has been minimized.  No action needed.
    return

intListW := A_GuiWidth - 40 - 88
intListH := A_GuiHeight - 115 - 132

intButtonSpacing := (intListW - (100 * 2)) // 3

for intIndex, objGuiControl in objGuiControls
{
	intX := objGuiControl.X
	intY := objGuiControl.Y

	if (intX < 0)
		intX:= A_GuiWidth + intX
	if (intY < 0)
		intY := A_GuiHeight + intY

	if (objGuiControl.Center)
	{
		GuiControlGet, arrPos, Pos, % objGuiControl.Name
		intX := intX - (arrPosW // 2) ; Floor divide
	}

	if (objGuiControl.Name = "lnkGuiDropHelpClicked")
	{
		GuiControlGet, arrPos, Pos, lnkGuiDropHelpClicked
		intX := intX - arrPosW
	}
	else if (objGuiControl.Name = "btnGuiSave")
		intX := 40 + intButtonSpacing
	else if (objGuiControl.Name = "btnGuiCancel")
		intX := 40 + (2 * intButtonSpacing) + 100
		
	GuiControl, % "1:Move" . (objGuiControl.Draw ? "Draw" : ""), % objGuiControl.Name, % "x" . intX	.  " y" . intY
		
}

GuiControl, 1:Move, drpMenusList, w%intListW%
GuiControl, 1:Move, lvFavoritesList, w%intListW% h%intListH%

Gosub, AjustColumnWidth

return
;------------------------------------------------------------


;------------------------------------------------------------
AjustColumnWidth:
;------------------------------------------------------------

LV_ModifyCol(1, "Auto") ; adjust column width

; See http://www.autohotkey.com/board/topic/6073-get-listview-column-width-with-sendmessage/
intCol1 := 0 ; column index, zero-based
SendMessage, 0x1000+29, %intCol1%, 0, SysListView321, ahk_id %strAppHwnd%
intCol1 := ErrorLevel ; column width
LV_ModifyCol(2, intListW - intCol1 - 21) ; adjust column width (-21 is for vertical scroll bar width)

return
;------------------------------------------------------------


;------------------------------------------------------------
BuildGui:
;------------------------------------------------------------

lGuiFullTitle := L(lGuiTitle, strAppName, strAppVersion)
Gui, 1:New, +Resize -MinimizeBox +MinSize636x538, %lGuiFullTitle%

Gui, +LastFound
strAppHwnd := WinExist()

if (blnUseColors)
	Gui, 1:Color, %strGuiWindowColor%

; Order of controls important to avoid drawgins gliches when resizing

Gui, 1:Font, % "s12 w700 " . (blnUseColors ? "c" . strTextColor : ""), Verdana
Gui, 1:Add, Text, vlblAppName x0 y0, %strAppName% %strAppVersion%
Gui, 1:Font, s9 w400, Verdana
Gui, 1:Add, Text, vlblAppTagLine, %lAppTagline%

Gui, 1:Add, Picture, vpicGuiAddFavorite gGuiAddFavorite, %strTempDir%\add_property-48.png ; Static3
Gui, 1:Add, Picture, vpicGuiEditFavorite gGuiEditFavorite x+1 yp, %strTempDir%\edit_property-48.png ; Static4
Gui, 1:Add, Picture, vpicGuiRemoveFavorite gGuiRemoveFavorite x+1 yp, %strTempDir%\delete_property-48.png ; Static5
Gui, 1:Add, Picture, vpicGuiGroupsManage gGuiGroupsManage x+1 yp, %strTempDir%\channel_mosaic-48.png ; Static6
Gui, 1:Add, Picture, vpicGuiOptions gGuiOptions x+1 yp, %strTempDir%\settings-32.png ; Static7
Gui, 1:Add, Picture, vpicPreviousMenu gGuiGotoPreviousMenu hidden x+1 yp, %strTempDir%\left-12.png ; Static8
Gui, 1:Add, Picture, vpicUpMenu gGuiGotoUpMenu hidden x+1 yp, %strTempDir%\up-12.png ; Static9
Gui, 1:Add, Picture, vpicMoveFavoriteUp gGuiMoveFavoriteUp x+1 yp, %strTempDir%\up_circular-26.png ; Static10
Gui, 1:Add, Picture, vpicMoveFavoriteDown gGuiMoveFavoriteDown x+1 yp, %strTempDir%\down_circular-26.png ; Static11
Gui, 1:Add, Picture, vpicAddSeparator gGuiAddSeparator x+1 yp, %strTempDir%\separator-26.png ; Static12
Gui, 1:Add, Picture, vpicAddColumnBreak gGuiAddColumnBreak x+1 yp, %strTempDir%\column-26.png ; Static13
Gui, 1:Add, Picture, vpicSortFavorites gGuiSortFavorites x+1 yp, %strTempDir%\generic_sorting2-26-grey.png ; Static14
Gui, 1:Add, Picture, vpicGuiAbout gGuiAbout x+1 yp, %strTempDir%\about-32.png ; Static15
Gui, 1:Add, Picture, vpicGuiHelp gGuiHelp x+1 yp, %strTempDir%\help-32.png ; Static16

Gui, 1:Font, s8 w400, Arial ; button legend
Gui, 1:Add, Text, vlblGuiOptions gGuiOptions x0 y+20, %lGuiOptions% ; Static17
Gui, 1:Add, Text, vlblGuiAddFavorite center gGuiAddFavorite x+1 yp, %lGuiAddFavorite% ; Static18
Gui, 1:Add, Text, vlblGuiEditFavorite center gGuiEditFavorite x+1 yp w88, %lGuiEditFavorite% ; Static19, w88 to make room fot when multiple favorites are selected
Gui, 1:Add, Text, vlblGuiRemoveFavorite center gGuiRemoveFavorite x+1 yp, %lGuiRemoveFavorite% ; Static20
Gui, 1:Add, Text, vlblGuiGroupsManage center gGuiGroupsManage x+1 yp, %lDialogGroups% ; Static21
Gui, 1:Add, Text, vlblGuiAbout center gGuiAbout x+1 yp, %lGuiAbout% ; Static22
Gui, 1:Add, Text, vlblGuiHelp center gGuiHelp x+1 yp, %lGuiHelp% ; Static23

Gui, 1:Font, s8 w400 italic, Verdana
Gui, 1:Add, Link, vlnkGuiHotkeysHelpClicked gGuiHotkeysHelpClicked x0 y+1, <a>%lGuiHotkeysHelp%</a> ; center option not working SysLink1
Gui, 1:Add, Link, vlnkGuiDropHelpClicked gGuiDropFilesHelpClicked right x+1 yp, <a>%lGuiDropFilesHelp%</a> ; SysLink2

Gui, 1:Font, s8 w400 normal, Verdana
Gui, 1:Add, Text, vlblSubmenuDropdownLabel x+1 yp, %lGuiSubmenuDropdownLabel%
Gui, 1:Add, DropDownList, vdrpMenusList gGuiMenusListChanged x0 y+1

; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
Gui, 1:Add, ListView
	, % "vlvFavoritesList Count32 AltSubmit NoSortHdr LV0x10 " . (blnUseColors ? "c" . strGuiListviewTextColor . " Background" . strGuiListviewBackgroundColor : "") . " gGuiFavoritesListEvents x+1 yp"
	, %lGuiLvFavoritesHeader%|Hidden Menu|Hidden Submenu|Hidden FavoriteType|Hidden IconResource|Hidden AppArguments|Hidden AppWorkingDir
Loop, 6
 	LV_ModifyCol(A_Index + 2, 0) ; hide 3rd-8th columns

Gui, 1:Font, s9 w600, Verdana
Gui, 1:Add, Button, vbtnGuiSave Disabled Default gGuiSave x200 y400 w100 h50, %lGuiSave% ; Button1
Gui, 1:Add, Button, vbtnGuiCancel gGuiCancel x350 yp w100 h50, %lGuiClose% ; Close until changes occur - Button2

if !(blnDonor)
{
	strDonateButtons := "thumbs_up|solutions|handshake|conference|gift"
	StringSplit, arrDonateButtons, strDonateButtons, |
	Random, intDonateButton, 1, 5

	Gui, 1:Add, Picture, vpicGuiDonate gGuiDonate x0 y+1, % strTempDir . "\" . arrDonateButtons%intDonateButton% . "-32.png" ; Static25
	Gui, 1:Font, s8 w400, Arial ; button legend
	Gui, 1:Add, Text, vlblGuiDonate center gGuiDonate x0 y+1, %lGuiDonate% ; Static26
}

Gui, 1:Show, % "Hide "
	. (arrSettingsPosition1 = -1 or arrSettingsPosition1 = "" or arrSettingsPosition2 = ""
	? "center w636 h538"
	: "x" . arrSettingsPosition1 . " y" . arrSettingsPosition2)
sleep, 100
if (arrSettingsPosition1 <> -1)
	WinMove, ahk_id %strAppHwnd%, , , , %arrSettingsPosition3%, %arrSettingsPosition4%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiHotkeysHelpClicked:
;------------------------------------------------------------
Gui, 1:+OwnDialogs

MsgBox, 0, %strAppName% - %lGuiHotkeysHelp%
	, %lGuiHotkeysHelpText%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiDropFilesHelpClicked:
;------------------------------------------------------------
Gui, 1:+OwnDialogs

MsgBox, 0, %strAppName% - %lGuiDropFilesHelp%
	, % L(lGuiDropFilesIncentive, strAppName, lDialogFolderLabel, lDialogFileLabel)

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

; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
Loop, % arrMenus[strCurrentMenu].MaxIndex()
	if (arrMenus[strCurrentMenu][A_Index].FavoriteType = "S") ; this is a submenu
		LV_Add(, arrMenus[strCurrentMenu][A_Index].FavoriteName, lGuiSubmenuLocation, arrMenus[strCurrentMenu][A_Index].MenuName
			, arrMenus[strCurrentMenu][A_Index].SubmenuFullName, arrMenus[strCurrentMenu][A_Index].FavoriteType
			, arrMenus[strCurrentMenu][A_Index].IconResource)
	else ; this is a folder, document, URL or application
		LV_Add(, arrMenus[strCurrentMenu][A_Index].FavoriteName, arrMenus[strCurrentMenu][A_Index].FavoriteLocation, arrMenus[strCurrentMenu][A_Index].MenuName
			, "", arrMenus[strCurrentMenu][A_Index].FavoriteType, arrMenus[strCurrentMenu][A_Index].IconResource
			, arrMenus[strCurrentMenu][A_Index].AppArguments, arrMenus[strCurrentMenu][A_Index].AppWorkingDir)
LV_Modify(1, "Select Focus")
Gosub, AjustColumnWidth

GuiControl, , drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"

return
;------------------------------------------------------------


;------------------------------------------------------------
SaveCurrentListviewToMenuObject:
;------------------------------------------------------------

arrMenus[strCurrentMenu] := Object() ; re-init current menu array
Gui, 1:ListView, lvFavoritesList

; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
Loop % LV_GetCount()
{
	LV_GetText(strName, A_Index, 1)
	LV_GetText(strLocation, A_Index, 2)
	LV_GetText(strMenu, A_Index, 3)
	LV_GetText(strSubmenu, A_Index, 4)
	LV_GetText(strFavoriteType, A_Index, 5)
	LV_GetText(strIconResource, A_Index, 6)
	LV_GetText(strAppArguments, A_Index, 7)
	LV_GetText(strAppWorkingDir, A_Index, 8)

	objFavorite := Object() ; new menu item
	objFavorite.FavoriteName := strName ; display name of this menu item
	objFavorite.FavoriteLocation := strLocation ; path for this menu item
	objFavorite.MenuName := strCurrentMenu ; parent menu of this menu item - do not use strMenu because lack of lMainMenuName
	objFavorite.SubmenuFullName := strSubmenu ; if menu, full name of the submenu
	objFavorite.FavoriteType := strFavoriteType ; "F" folder, "D" document, "U" url, "P" sPecial, "A" application or "S" submenu
	objFavorite.IconResource := strIconResource ; icon resource in format "iconfile,iconindex"
	objFavorite.AppArguments := strAppArguments ; application arguments
	objFavorite.AppWorkingDir := strAppWorkingDir ; application working directory
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
GuiFavoritesListEvents:
;------------------------------------------------------------
Gui, 1:ListView, lvFavoritesList

if (A_GuiEvent = "DoubleClick")
	gosub, GuiEditFavorite
else if (A_GuiEvent = "I")
{
	intFavoriteSelected := LV_GetCount("Selected")
	if (intFavoriteSelected > 1)
	{
		GuiControl, , lblGuiEditFavorite, % lGuiMove . " (" . intFavoriteSelected . ")"
		GuiControl, +gGuiMoveMultipleFavorites, lblGuiEditFavorite
		GuiControl, +gGuiMoveMultipleFavorites, picGuiEditFavorite
		GuiControl, , lblGuiRemoveFavorite, % lGuiRemoveFavorite . " (" . intFavoriteSelected . ")"
		GuiControl, +gGuiRemoveMultipleFavorites, lblGuiRemoveFavorite
		GuiControl, +gGuiRemoveMultipleFavorites, picGuiRemoveFavorite
		GuiControl, +gGuiMoveMultipleFavoritesUp, picMoveFavoriteUp
		GuiControl, +gGuiMoveMultipleFavoritesDown, picMoveFavoriteDown
	}
	else
	{
		GuiControl, , lblGuiEditFavorite, %lGuiEditFavorite%
		GuiControl, +gGuiEditFavorite, lblGuiEditFavorite
		GuiControl, +gGuiEditFavorite, picGuiEditFavorite
		GuiControl, , lblGuiRemoveFavorite, %lGuiRemoveFavorite%
		GuiControl, +gGuiRemoveFavorite, lblGuiRemoveFavorite
		GuiControl, +gGuiRemoveFavorite, picGuiRemoveFavorite
		GuiControl, +gGuiMoveFavoriteUp, picMoveFavoriteUp
		GuiControl, +gGuiMoveFavoriteDown, picMoveFavoriteDown
	}
}

return
;------------------------------------------------------------


;------------------------------------------------------------
HotkeyChangeMenu:
;------------------------------------------------------------
Gui, 1:ListView, lvFavoritesList

intRowToEdit := LV_GetNext()
LV_GetText(strCurrentName, intRowToEdit, 1)
LV_GetText(strCurrentSubmenuFullName, intRowToEdit, 4)
LV_GetText(strFavoriteType, intRowToEdit, 5)

if (strFavoriteType = "S")
	Gosub, OpenMenuFromGuiHotkey

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMenusListChanged:
GuiGotoUpMenu:
GuiGotoPreviousMenu:
OpenMenuFromEditForm:
OpenMenuFromGuiHotkey:
;------------------------------------------------------------

intCurrentLastPosition := 0

if (A_ThisLabel = "GuiMenusListChanged")
{
	GuiControlGet, strNewDropdownMenu, , drpMenusList
	if (strNewDropdownMenu = strCurrentMenu) ; user selected the current menu in the dropdown
		return
}

Gosub, SaveCurrentListviewToMenuObject ; save current LV

if (A_ThisLabel = "GuiGotoPreviousMenu")
{
	strCurrentMenu := arrSubmenuStack[1] ; pull the top menu from the left arrow stack
	arrSubmenuStack.Remove(1) ; remove the top menu from the left arrow stack

	intCurrentLastPosition := arrSubmenuStackPosition[1] ; pull the focus position in top menu from the left arrow stack
	arrSubmenuStackPosition.Remove(1) ; remove the top position from the left arrow stack
}
else
{
	arrSubmenuStack.Insert(1, strCurrentMenu) ; push the current menu to the left arrow stack
	
	if (A_ThisLabel = "GuiMenusListChanged")
		strCurrentMenu := strNewDropdownMenu
	else if (A_ThisLabel = "GuiGotoUpMenu")
		strCurrentMenu := SubStr(strCurrentMenu, 1, InStr(strCurrentMenu, lGuiSubmenuSeparator, , 0) - 1) 
	else if (A_ThisLabel = "OpenMenuFromEditForm") or (A_ThisLabel = "OpenMenuFromGuiHotkey")
		strCurrentMenu := strCurrentSubmenuFullName

	arrSubmenuStackPosition.Insert(1, LV_GetNext("Focused"))
}

GuiControl, % (arrSubmenuStack.MaxIndex() ? "Show" : "Hide"), picPreviousMenu
GuiControl, % (strCurrentMenu <> lMainMenuName ? "Show" : "Hide"), picUpMenu

Gosub, LoadOneMenuToGui

if (intCurrentLastPosition) ; we went to a previous menu
{
	LV_Modify(0, "-Select")
	LV_Modify(intCurrentLastPosition, "Select Focus Vis")
}

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupsManage:
;------------------------------------------------------------

intWidth := 350

Gosub, BuildFoldersInExplorerMenu ; refresh explorers object and intExplorersIndex counter

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide

Gui, 2:New, , % L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:+OwnDialogs
if (blnUseColors)
	Gui, 2:Color, %strGuiWindowColor%

Gui, 2:Font, w600 
Gui, 2:Add, Text, x10 y10, %lDialogGroupManageAbout%
Gui, 2:Font

Gui, 2:Add, Text, x10 y+10 w%intWidth%, %lDialogGroupManageIntro%

Gui, 2:Font, w600 
Gui, 2:Add, Text, x10 y+20, %lDialogGroupManageCreatingTitle%
Gui, 2:Font 

strUseMenuSave := lMenuGroup . " > " . lMenuGroupSave
Gui, 2:Add, Text, x10 y+10 w%intWidth%, % L(lDialogGroupManageCreatingPrompt, lDialogGroupNew, strUseMenuSave)
Gui, 2:Add, Button, x10 y+10 vbtnGroupManageNew gGuiGroupManageNew, %lDialogGroupNew%
GuiControl, % (!intExplorersIndex ? "Disable" : "Enable") ; disable Save group menu if no Explorer
	, btnGroupManageNew
GuiCenterButtons(L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion), , , , "btnGroupManageNew")
if !(intExplorersIndex)
	Gui, 2:Add, Text, x10 y+10 w%intWidth%, %lDialogGroupManageCannotSave%

Gui, 2:Font, w600 
Gui, 2:Add, Text, x10 y+20, %lDialogGroupManageManagingTitle%
Gui, 2:Font

Gui, 2:Add, DropDownList, x10 y+10 w%intWidth% vdrpGroupsList, %lDialogGroupSelect%||%strGroups%

Gui, 2:Add, Button, x10 y+10 vbtnGroupManageLoad  gGuiGroupManageLoad, %lDialogGroupLoad%
Gui, 2:Add, Button, x10 yp vbtnGroupManageEdit gGuiGroupManageEdit, %lDialogGroupEdit%
Gui, 2:Add, Button, x10 yp vbtnGroupManageDelete gGuiGroupManageDelete, %lDialogGroupDelete%
GuiCenterButtons(L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion), , , , "btnGroupManageLoad", "btnGroupManageEdit", "btnGroupManageDelete")

Gui, 2:Add, Button, x+10 y+30 vbtnGroupManageClose g2GuiClose h33, %lGui2Close%
GuiCenterButtons(L(lDialogGroupManageGroupsTitle, strAppName, strAppVersion), , , , "btnGroupManageClose")
Gui, 2:Add, Text, x10, %A_Space%

Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiGroupManageEdit:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

if !StrLen(drpGroupsList) or (drpGroupsList = lDialogGroupSelect)
{
	Oops(lDialogGroupSelectError, lDialogGroupEditError)
	return
}

strGroupToEdit := drpGroupsList
Gosub, GuiGroupEditFromManage
GuiControl, 2:, drpGroupsList, |%lDialogGroupSelect%||%strGroups%

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

Gosub, BuildGroupMenu

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

Gosub, GuiGroupSaveFromManage
GuiControl, 2:, drpGroupsList, |%lDialogGroupSelect%||%strGroups%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddSeparator:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList

if (LV_GetCount("Selected") > 1)
	return

intInsertPosition := LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1
LV_Modify(0, "-Select")
LV_Insert(intInsertPosition, "Select Focus", lMenuSeparator, lMenuSeparator . lMenuSeparator, strCurrentMenu, "", "", "", "", "")
LV_Modify(LV_GetNext(), "Vis")
Gosub, AjustColumnWidth

GuiControl, Enable, btnGuiSave
GuiControl, , btnGuiCancel, %lGuiCancel%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddColumnBreak:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList

if (LV_GetCount("Selected") > 1)
	return

intInsertPosition := LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1
LV_Modify(0, "-Select")
; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir

LV_Insert(intInsertPosition, "Select Focus"
	, strColumnBreakIndicator . " " . lMenuColumnBreak . " " . strColumnBreakIndicator
	, strColumnBreakIndicator . " " . lMenuColumnBreak . " " . strColumnBreakIndicator
	, strCurrentMenu, "", "", "", "", "")
LV_Modify(LV_GetNext(), "Vis")
Gosub, AjustColumnWidth

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
GuiMoveMultipleFavoritesUp:
GuiMoveMultipleFavoritesDown:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList

blnAbortGroupMove := false
strSelected := ""
intRowToProcess := 0
loop
{
	intRowToProcess := LV_GetNext(intRowToProcess)
	strSelected := strSelected . intRowToProcess . "|"
}
until !LV_GetNext(intRowToProcess)
StringTrimRight, strSelected, strSelected, 1

Loop
{
	if (A_ThisLabel = "GuiMoveMultipleFavoritesUp")
		Gosub, GetFirstSelected ; will re-init intRowToProcess
	else
		Gosub, GetLastSelected ; will re-init intRowToProcess
	
	if (!intRowToProcess) or (blnAbortGroupMove)
		break
	
	if (A_ThisLabel = "GuiMoveMultipleFavoritesUp")
	{
		intSelectedRow := intRowToProcess
		Gosub, GuiMoveOneFavoriteUp
	}
	else
	{
		intSelectedRow := intRowToProcess
		Gosub, GuiMoveOneFavoriteDown
	}
}

if (!blnAbortGroupMove)
	Loop, Parse, strSelected, |
		LV_Modify(A_LoopField  + (A_ThisLabel = "GuiMoveMultipleFavoritesUp" ? -1 : 1), "Select")
LV_Modify(LV_GetNext(0), "Focus") ; give focus to the first selected row


return
;------------------------------------------------------------


;------------------------------------------------------------
GetFirstSelected:
GetLastSelected:
;------------------------------------------------------------

intRowToProcess := 0
if (A_ThisLabel = "GetFirstSelected")
	intRowToProcess := LV_GetNext(intRowToProcess) ; start from first selected
else
	loop
		intRowToProcess := LV_GetNext(intRowToProcess) ; start with last selected
	until !LV_GetNext(intRowToProcess)

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveFavoriteUp:
GuiMoveFavoriteDown:
GuiMoveOneFavoriteUp:
GuiMoveOneFavoriteDown:
;------------------------------------------------------------

; prevent double-click on some static control to overwrite the clipboard with the image URL (a windows "undesired feature")
; see http://www.autohotkey.com/board/topic/94962-doubleclick-on-gui-pictures-puts-their-path-in-your-clipboard/
If (A_GuiEvent="DoubleClick")
	; would be used to restore clipboard's  previous content if there was not a side effect in XL 2010
	; (see: https://github.com/JnLlnd/FoldersPopup/issues/128)
	; Clipboard := ClipboardAllBK
	Clipboard := "" ; better than nothing, empty the clipboard because we could not restore its previous content

if !InStr(A_ThisLabel, "One")
{
	GuiControl, Focus, lvFavoritesList
	Gui, 1:ListView, lvFavoritesList
	intSelectedRow := LV_GetNext()
}
if (intSelectedRow = (InStr(A_ThisLabel, "Up") ? 1 : LV_GetCount()))
{
	if InStr(A_ThisLabel, "One")
		blnAbortGroupMove := true
	return
}

Loop, 8
	LV_GetText(arrThis%A_Index%, intSelectedRow, A_Index)

Loop, 8
	LV_GetText(arrOther%A_Index%, intSelectedRow + (InStr(A_ThisLabel, "Up") ? -1 : 1), A_Index)

LV_Modify(intSelectedRow, "-Select")
LV_Modify(intSelectedRow, "", arrOther1, arrOther2, arrOther3, arrOther4, arrOther5, arrOther6, arrOther7, arrOther8)
LV_Modify(intSelectedRow + (InStr(A_ThisLabel, "Up") ? -1 : 1), , arrThis1, arrThis2, arrThis3, arrThis4, arrThis5, arrThis6, arrThis7, arrThis8)

if !InStr(A_ThisLabel, "One")
	LV_Modify(intSelectedRow + (InStr(A_ThisLabel, "Up") ? -1 : 1), "Select Focus Vis")

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
Gosub, BuildMainMenuWithStatus
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
		; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
		strValue := arrMenus[strMenu][A_Index].FavoriteName
			. "|" . arrMenus[strMenu][A_Index].FavoriteLocation
			. "|" . SubStr(arrMenus[strMenu][A_Index].MenuName, StrLen(lMainMenuName) + 1) ; remove main menu name for ini file
			. "|" . SubStr(arrMenus[strMenu][A_Index].SubmenuFullName, StrLen(lMainMenuName) + 1) ; remove main menu name for ini file
			. "|" . arrMenus[strMenu][A_Index].FavoriteType
			. "|" . arrMenus[strMenu][A_Index].IconResource
			. "|" . arrMenus[strMenu][A_Index].AppArguments
			. "|" . arrMenus[strMenu][A_Index].AppWorkingDir
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
Gui, 1:Show

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
		if (A_OSVersion = "WIN_XP")
			Gosub, BuildSpecialFoldersMenu
		Gosub, BuildFoldersInExplorerMenu
		Gosub, BuildMainMenu ; need to be initialized here - will be updated at each call to popup menu

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
		; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
		objFavoriteBK := Object()
		objFavoriteBK.MenuName := objFavorite.MenuName
		objFavoriteBK.FavoriteName := objFavorite.FavoriteName
		objFavoriteBK.FavoriteLocation := objFavorite.FavoriteLocation
		objFavoriteBK.SubmenuFullName := objFavorite.SubmenuFullName
		objFavoriteBK.FavoriteType := objFavorite.FavoriteType
		objFavoriteBK.IconResource := objFavorite.IconResource
		objFavoriteBK.AppArguments := objFavorite.AppArguments
		objFavoriteBK.AppWorkingDir := objFavorite.AppWorkingDir
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
		; 1 FavoriteName, 2 FavoriteLocation, 3 MenuName, 4 SubmenuFullName, 5 FavoriteType, 6 IconResource, 7 AppArguments, 8 AppWorkingDir
		objFavorite := Object()
		objFavorite.MenuName := objFavoriteBK.MenuName
		objFavorite.FavoriteName := objFavoriteBK.FavoriteName
		objFavorite.FavoriteLocation := objFavoriteBK.FavoriteLocation
		objFavorite.SubmenuFullName := objFavoriteBK.SubmenuFullName
		objFavorite.FavoriteType := objFavoriteBK.FavoriteType
		objFavorite.IconResource := objFavoriteBK.IconResource
		objFavorite.AppArguments := objFavoriteBK.AppArguments
		objFavorite.AppWorkingDir := objFavoriteBK.AppWorkingDir
		arrMenu.Insert(objFavorite)
	}
	arrMenus.Insert(strMenuName, arrMenu) ; add this submenu to the array of menus
	objFavoriteBK := 
}
arrMenusBK :=

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
			strCurrentLocation := objLister.LocationURL
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
	/* removed for FPconnect: 
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
	*/
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

objDOpusListers :=

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFavorite:
GuiAddFromPopup:
GuiAddFromDropFiles:
GuiEditFavorite:
;------------------------------------------------------------

; Icon resource in the format "iconfile,index", examnple "shell32.dll,2"
; strDefaultIconResource -> default icon for the current type of favorite (F, P, D, U, A or S)
; strSavedIconResource -> actual value in the ListView or in the menu object
; strCurrentIconResource -> icon currently displayed in the Add/Edit dialog box
; strCurrentIconResource is always equal to strDefaultIconResource or strSavedIconResource

if (A_ThisLabel = "GuiEditFavorite")
{
	Gui, 1:ListView, lvFavoritesList
	intRowToEdit := LV_GetNext()
	LV_GetText(strCurrentName, intRowToEdit, 1)
	
	if (strCurrentName = lMenuSeparator) or IsColumnBreak(strCurrentName)
		return

	if !StrLen(strCurrentName) or (intRowToEdit = 0)
	{
		Oops(lDialogSelectItemToEdit)
		return
	}
	LV_GetText(strCurrentLocation, intRowToEdit, 2)
	LV_GetText(strCurrentSubmenuFullName, intRowToEdit, 4)
	LV_GetText(strFavoriteType, intRowToEdit, 5)
	LV_GetText(strSavedIconResource, intRowToEdit, 6)
	strCurrentIconResource := strSavedIconResource
	LV_GetText(strAppArguments, intRowToEdit, 7)
	LV_GetText(strAppWorkingDir, intRowToEdit, 8)

	blnRadioFolder := (strFavoriteType = "F")
	blnRadioSpecial := (strFavoriteType = "P")
	blnRadioFile := (strFavoriteType = "D")
	blnRadioURL := (strFavoriteType = "U")
	blnRadioApplication := (strFavoriteType = "A")
	blnRadioSubmenu := (strFavoriteType = "S")
}
else
{
	Gosub, SaveCurrentListviewToMenuObject ; update menu object from LV, for items dropdown list
	
	intRowToEdit := 0 ;  used when saving to flag to insert a new row
	strCurrentName := "" ; make sure it is empty
	strCurrentSubmenuFullName := "" ;  make sure it is empty
	strFavoriteType := "" ;  make sure it is empty
	strSavedIconResource := "" ;  make sure it is empty
	strDefaultIconResource := "" ;  make sure it is empty
	strCurrentIconResource := "" ;  make sure it is empty
	strAppArguments := "" ;  make sure it is empty
	strAppWorkingDir := "" ;  make sure it is empty
	
	if (A_ThisLabel = "GuiAddFromPopup" or A_ThisLabel = "GuiAddFromDropFiles")
		; strCurrentLocation is received from AddThisFolder or GuiDropFiles
		strFavoriteShortName := GetDeepestFolderName(strCurrentLocation)
	else
	{
		;  make sure these variables are empty
		strCurrentLocation := ""
		strFavoriteShortName := ""
	}

	blnRadioFolder := true
	blnRadioSpecial := false
	blnRadioFile := false
	blnRadioURL := false
	blnRadioApplication := false
	blnRadioSubmenu := false
	
	if (A_ThisLabel = "GuiAddFromPopup")
		blnRadioFolder := true
	else if (A_ThisLabel = "GuiAddFromDropFiles")
	{
		SplitPath, strCurrentLocation, , , strExtension
		if StrLen(strExtension) and InStr("exe|com|bat|vbs|ahk", strExtension)
		{
			blnRadioApplication := true
			blnRadioFile := false
			blnRadioFolder := false
		}
		else
		{
			blnRadioApplication := false
			blnRadioFile := LocationIsDocument(strCurrentLocation)
			blnRadioFolder := not blnRadioFile
		}
	}
}

if InStr("GuiAddFavorite|GuiAddFromDropFiles", A_ThisLabel)
	intItemPosition := (LV_GetCount() ? (LV_GetNext() ? LV_GetNext() : 0xFFFF) : 1)
else
	intItemPosition := 0 ; will display lDialogEndOfMenu

intGui1WinID := WinExist("A")
Gui, 1:Submit, NoHide

Gui, 2:New, , % L(lDialogAddEditFavoriteTitle, (A_ThisLabel = "GuiEditFavorite" ? lDialogEdit : lDialogAdd), strAppName, strAppVersion)
Gui, 2:+Owner1
Gui, 2:+OwnDialogs
if (blnUseColors)
	Gui, 2:Color, %strGuiWindowColor%

Gui, 2:Add, Text, % x10 y10 vlblFavoriteParentMenu, % (blnRadioSubmenu ? lDialogSubmenuParentMenu : lDialogFavoriteParentMenu)
Gui, 2:Add, DropDownList, x10 w300 vdrpParentMenu gDropdownParentMenuChanged, % BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu, strCurrentSubmenuFullName) . "|"
Gui, 2:Add, Text, yp x+10 section
Gui, 2:Add, Text, xs y10 w64 center vlblIcon gGuiPickIconDialog, %lDialogIcon%
Gui, Add, Picture, % "xs+" . ((64-32)/2) . " y+5 w32 h32 vpicIcon gGuiPickIconDialog"
Gui, Add, Text, x+5 yp vlblRemoveIcon gGuiRemoveIcon, X

If (A_ThisLabel <> "GuiEditFavorite")
{
	Gui, 2:Add, Text, x20 ys+25 vlblFavoriteParentMenuPosition, %lDialogFavoriteMenuPosition%
	Gui, 2:Add, DropDownList, x20 w290 vdrpParentMenuItems AltSubmit
}

if (A_ThisLabel = "GuiAddFavorite")
{
	Gui, 2:Add, Text, x10, %lDialogAdd%:
	Gui, 2:Add, Radio, x+10 yp vblnRadioFolder checked gRadioButtonsChanged section, %lDialogFolderLabel%
	If WindowsIsVersion7OrMore()
		Gui, 2:Add, Radio, xs vblnRadioSpecial gRadioButtonsChanged, %lDialogSpecialLabel%
	Gui, 2:Add, Radio, xs vblnRadioFile gRadioButtonsChanged, %lDialogFileLabel%
	Gui, 2:Add, Radio, xs vblnRadioApplication gRadioButtonsChanged, %lDialogApplicationLabel%
	Gui, 2:Add, Radio, xs vblnRadioURL gRadioButtonsChanged, %lDialogURLLabel%
	Gui, 2:Add, Radio, xs vblnRadioSubmenu gRadioButtonsChanged, %lDialogSubmenuLabel%
}

Gui, 2:Add, Text, x10 y+10 w300 vlblShortName
	, % (blnRadioSubmenu ? lDialogSubmenuShortName
	: (blnRadioFile ? lDialogFileShortName 
	: (blnRadioURL ? lDialogURLShortName 
	: (blnRadioSpecial ? lDialogSpecialLabel 
	: (blnRadioApplication ? lDialogApplicationLabel 
	: lDialogFolderShortName)))))
Gui, 2:Add, Edit, x10 w300 Limit250 vstrFavoriteShortName, % (A_ThisLabel = "GuiEditFavorite" ? strCurrentName : strFavoriteShortName)

if (blnRadioSubmenu)
	Gui, 2:Add, Button, x+10 yp vbnlEditFolderOpenMenu gGuiOpenThisMenu, %lDialogOpenThisMenu%
else
{
	Gui, 2:Add, Text, x10 w300 vlblFolder section, % (blnRadioFile ? lDialogFileLabel : (blnRadioURL ? lDialogURLLabel : (blnRadioApplication ? lDialogApplicationLabel : lDialogFolderLabel)))
	
	Gui, 2:Add, DropDownList, xs ys w300 vdrpSpecialFolder gDropdownSpecialFolderChanged hidden, %strSpecialFoldersList%
	if (A_ThisLabel = "GuiEditFavorite")
		GuiControl, ChooseString, drpSpecialFolder, %strCurrentName%

	Gui, 2:Add, Edit, ys+20 x10 w300 h20 vstrFavoriteLocation gEditFolderLocationChanged, %strCurrentLocation%
	if !(blnRadioURL)
		Gui, 2:Add, Button, x+10 yp vbtnSelectFolderLocation gButtonSelectFolderLocation, %lDialogBrowseButton%
}

Gui, 2:Add, Text, x10 w300 vlblArguments, %lDialogArgumentsLabel%
Gui, 2:Add, Edit, x10 w300 Limit250 vstrAppArguments, %strAppArguments%
Gui, 2:Add, Text, x10 w300 y+1 vlblArgumentsHelp, %lDialogArgumentsLabelHelp%
Gui, 2:Add, Text, x10 w300 vlblWorkingDir, %lDialogWorkingDirLabel%
Gui, 2:Add, Edit, x10 w300 Limit250 vstrAppWorkingDir, %strAppWorkingDir%
Gui, 2:Add, Button, x+10 yp vbtnSelectWorkingDir gButtonSelectWorkingDir, %lDialogBrowseButton%

GuiControlGet, arrPos, Pos, strAppArguments
intMinButtonY := arrPosY

if (A_ThisLabel = "GuiEditFavorite")
{
	Gui, 2:Add, Button, y+20 vbtnEditFolderSave gGuiEditFavoriteSave default, %lDialogSave%
	Gui, 2:Add, Button, yp vbtnEditFolderCancel gGuiEditFavoriteCancel, %lGuiCancel%
	
	if (!blnRadioApplication)
	{
		GuiControl, Move, btnEditFolderSave, y%intMinButtonY%
		GuiControl, Move, btnEditFolderCancel, y%intMinButtonY%
	}
	
	GuiCenterButtons(L(lDialogAddEditFavoriteTitle, lDialogEdit, strAppName, strAppVersion), 10, 5, 20, "btnEditFolderSave", "btnEditFolderCancel")
}
else
{
	Gui, 2:Add, Button, y+20 vbtnAddFolderAdd gGuiAddFavoriteSave default, %lDialogAdd%
	Gui, 2:Add, Button, yp vbtnAddFolderCancel gGuiAddFavoriteCancel, %lGuiCancel%
	
	GuiControlGet, arrPos, Pos, btnAddFolderCancel
	intMaxButtonY := arrPosY

	if !(A_ThisLabel = "GuiAddFromDropFiles" and blnRadioApplication)
	{
		GuiControl, Move, btnAddFolderAdd, y%intMinButtonY%
		GuiControl, Move, btnAddFolderCancel, y%intMinButtonY%
	}
	
	GuiCenterButtons(L(lDialogAddEditFavoriteTitle, lDialogAdd, strAppName, strAppVersion), 10, 5, 20, "btnAddFolderAdd", "btnAddFolderCancel")
}

Gosub, GuiFavoriteIconDefault
Gosub, GuiFavoriteIconDisplay
Gosub, DropdownParentMenuChanged ; to init the content of menu items
Gosub, RadioButtonsChanged ; to hide unused control when edit a special folder

if (blnRadioSpecial)
	GuiControl, 2:Focus, drpSpecialFolder
else
	GuiControl, 2:Focus, strFavoriteShortName
if (A_ThisLabel = "GuiEditFavorite") and (!blnRadioSpecial)
	SendInput, ^a
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveMultipleFavorites:
;------------------------------------------------------------

Gui, 2:New, , % L(lDialogMoveFavoritesTitle, strAppName, strAppVersion)
Gui, 2:Add, Text, % x10 y10 vlblFavoriteParentMenu, % L(lDialogFavoritesParentMenuMove, intFavoriteSelected)
Gui, 2:Add, DropDownList, x10 w300 vdrpParentMenu, % BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"

Gui, 2:Add, Button, y+20 vbtnMoveFavoritesSave gGuiMoveMultipleFavoritesSave, %lGuiMove%
Gui, 2:Add, Button, yp vbtnMoveFavoritesCancel gGuiEditFavoriteCancel, %lGuiCancel%
GuiCenterButtons(L(lDialogMoveFavoritesTitle, strAppName, strAppVersion), 10, 5, 20, "btnMoveFavoritesSave", "btnMoveFavoritesCancel")

GuiControl, 2:Focus, drpParentMenu
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
DropdownParentMenuChanged:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

strDropdownParentMenuItems := ""

Loop, % arrMenus[drpParentMenu].MaxIndex()
	strDropdownParentMenuItems := strDropdownParentMenuItems . arrMenus[drpParentMenu][A_Index].FavoriteName . "|"

GuiControl, , drpParentMenuItems, % "|" . strDropdownParentMenuItems . strColumnBreakIndicator . " " . lDialogEndOfMenu . " " . strColumnBreakIndicator
if (intItemPosition)
	GuiControl, Choose, drpParentMenuItems, %intItemPosition%
else
	GuiControl, ChooseString, drpParentMenuItems, % strColumnBreakIndicator . " " . lDialogEndOfMenu . " " . strColumnBreakIndicator

intItemPosition := 0 ; if called again for a new partent menu, will display lDialogEndOfMenu

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
RadioButtonsChanged:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

GuiControlGet, blnAddMode, Visible, blnRadioFolder ; if radio buttons visible, we are in add mode
if (ErrorLevel) ; if control does not exist, we are in edit mode
	blnAddMode := 0

GuiControl, % "2:" . (blnRadioSpecial ? "Show" : "Hide"), drpSpecialFolder
GuiControl, % "2:" . (blnRadioSubmenu or blnRadioSpecial ? "Hide" : "Show"), lblFolder
GuiControl, % "2:" . (blnRadioSubmenu or blnRadioSpecial ? "Hide" : "Show"), strFavoriteLocation
GuiControl, % "2:" . (blnRadioSubmenu or blnRadioSpecial or blnRadioURL ? "Hide" : "Show"), btnSelectFolderLocation
GuiControl, % "2:" . (blnRadioApplication ? "Show" : "Hide"), lblArguments
GuiControl, % "2:" . (blnRadioApplication ? "Show" : "Hide"), strAppArguments
GuiControl, % "2:" . (blnRadioApplication ? "Show" : "Hide"), lblArgumentsHelp
GuiControl, % "2:" . (blnRadioApplication ? "Show" : "Hide"), lblWorkingDir
GuiControl, % "2:" . (blnRadioApplication ? "Show" : "Hide"), strAppWorkingDir
GuiControl, % "2:" . (blnRadioApplication ? "Show" : "Hide"), btnSelectWorkingDir

if (blnAddMode) and StrLen(strFavoriteLocation) ; in add mode keep only folder name if we have a file name, else empty it
{
	SplitPath, strFavoriteLocation, , strFolderNameOnly, strFilenameExtension
	if StrLen(strFilenameExtension)
		GuiControl, 2:, strFavoriteLocation, %strFolderNameOnly%
	else
		GuiControl, 2:, strFavoriteLocation
}

if (blnRadioFolder)
{
	GuiControl, 2:, lblShortName, %lDialogFolderShortName%
	GuiControl, 2:, lblFolder, %lDialogFolderLabel%
}
else if (blnRadioSpecial)
{
	GuiControl, 2:, lblShortName, %lDialogSpecialLabel%
	ControlClick, ComboBox3 ; open drpSpecialFolder dropdown
}
else if (blnRadioFile)
{
	GuiControl, 2:, lblShortName, %lDialogFileShortName%
	GuiControl, 2:, lblFolder, %lDialogFileLabel%
}
else if (blnRadioURL)
{
	GuiControl, 2:, lblShortName, %lDialogURLShortName%
	GuiControl, 2:, lblFolder, %lDialogURLLabel%
}
else if (blnRadioApplication)
{
	GuiControl, 2:, lblShortName, %lDialogApplicationShortName%
	GuiControl, 2:, lblFolder, %lDialogApplicationLabel%
}
else ; blnRadioSubmenu
{
	GuiControl, 2:, lblShortName, %lDialogSubmenuShortName%
}

if (blnAddMode) ; move buttons considering app properties fields
{
	GuiControl, Move, btnAddFolderAdd, % "y" . (blnRadioApplication ? intMaxButtonY : intMinButtonY)
	GuiControl, Move, btnAddFolderCancel, % "y" . (blnRadioApplication ? intMaxButtonY : intMinButtonY)
}
Gui, Show, AutoSize ; resize window considering app properties fields, but do not Center

Gosub, GuiFavoriteIconDefault
if (blnAddMode)
	strCurrentIconResource := strDefaultIconResource
Gosub, GuiFavoriteIconDisplay

if (blnRadioFolder or blnRadioFile or blnRadioApplication)
	GuiControl, 2:+Default, btnSelectFolderLocation
else
	GuiControl, 2:+Default, btnAddFolderAdd
GuiControl, 2:Focus, % (blnRadioSpecial ? "drpSpecialFolder" : "strFavoriteShortName")

return
;------------------------------------------------------------


;------------------------------------------------------------
DropdownSpecialFolderChanged:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

strThisFavoriteShortName := drpSpecialFolder
GuiControl, , strFavoriteShortName, %strThisFavoriteShortName% ; also assign values to gui control

strThisFavoriteLocation := objClassIdOrPathByDefaultName[strThisFavoriteShortName]
GuiControl, , strFavoriteLocation, %strThisFavoriteLocation% ; also assign values to gui control

strCurrentIconResource := objSpecialFolders[strThisFavoriteLocation].DefaultIcon

Gosub, GuiFavoriteIconDisplay

return
;------------------------------------------------------------


;------------------------------------------------------------
EditFolderLocationChanged:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

if (blnRadioFile or blnRadioApplication)
{
	strCurrentIconResource := ""
	Gosub, GuiFavoriteIconDefault
	Gosub, GuiFavoriteIconDisplay
}

return
;------------------------------------------------------------


;------------------------------------------------------------
ButtonSelectFolderLocation:
ButtonSelectWorkingDir:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

if (blnRadioFolder) or (A_ThisLabel = "ButtonSelectWorkingDir")
	FileSelectFolder, strNewLocation, *%strCurrentLocation%, 3, %lDialogAddFolderSelect%
else
	FileSelectFile, strNewLocation, S3, %strCurrentLocation%, %lDialogAddFileSelect%

if !(StrLen(strNewLocation))
	return

if (A_ThisLabel = "ButtonSelectFolderLocation")
{
	GuiControl, 2:, strFavoriteLocation, %strNewLocation%
	if !StrLen(strFavoriteShortName)
		GuiControl, 2:, strFavoriteShortName, % GetDeepestFolderName(strNewLocation)
}
else
	GuiControl, 2:, strAppWorkingDir, %strNewLocation%

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiPickIconDialog:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

if ((blnRadioFile or blnRadioApplication) and !StrLen(strFavoriteLocation))
{
	Oops(lPickIconNoLocation)
	return
}
if InStr(strCurrentIconResource, "%")
	strCurrentIconResource := EnvVars(strCurrentIconResource)

; Source: http://ahkscript.org/boards/viewtopic.php?f=5&t=5108#p29970
VarSetCapacity(strThisIconFile, 1024) ; must be placed before strNewIconFile is initialized because VarSetCapacity erase its content
ParseIconResource(strCurrentIconResource, strThisIconFile, intThisIconIndex)

WinGet, hWnd, ID, A
if (intThisIconIndex >= 0) ; adjust index for positive index only (not for negative index)
	intThisIconIndex := intThisIconIndex - 1
DllCall("shell32\PickIconDlg", "Uint", hWnd, "str", strThisIconFile, "Uint", 260, "intP", intThisIconIndex)
if (intThisIconIndex >= 0) ; adjust index for positive index only (not for negative index)
	intThisIconIndex := intThisIconIndex + 1

if StrLen(strThisIconFile)
	strCurrentIconResource := strThisIconFile . "," . intThisIconIndex
Gosub, GuiFavoriteIconDisplay

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveIcon:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

if (blnRadioFile or blnRadioApplication)
{
	GetIcon4Location(strFavoriteLocation, strThisIconFile, intThisIconIndex)
	strCurrentIconResource := strThisIconFile . "," . intThisIconIndex
}
else
	strCurrentIconResource := strDefaultIconResource

Gosub, GuiFavoriteIconDisplay

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiFavoriteIconDefault:
;------------------------------------------------------------
Gui, 2:Submit, NoHide

if (blnRadioSubmenu)
	; default submenu icon
	strDefaultIconResource := objIconsFile["Submenu"] . "," . objIconsIndex["Submenu"]
else if (blnRadioURL)
{
	; default browser icon
	GetIcon4Location(strTempDir . "\default_browser_icon.html", strThisIconFile, intThisIconIndex)
	strDefaultIconResource := strThisIconFile . "," . intThisIconIndex
}
else if (blnRadioFile or blnRadioApplication)
{
	; default icon for the selected file in add/edit favorite
	GetIcon4Location(strFavoriteLocation, strThisIconFile, intThisIconIndex)
	strDefaultIconResource := strThisIconFile . "," . intThisIconIndex
}
else ; blnRadioFolder
	strDefaultIconResource := objIconsFile["Folder"] . "," . objIconsIndex["Folder"]

if !StrLen(strCurrentIconResource)
	strCurrentIconResource := strDefaultIconResource

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiFavoriteIconDisplay:
;------------------------------------------------------------

ParseIconResource(strCurrentIconResource, strThisIconFile, intThisIconIndex)
GuiControl, , picIcon, *icon%intThisIconIndex% %strThisIconFile%
GuiControl, % (EnvVars(strCurrentIconResource) <> (strDefaultIconResource) ? "Show" : "Hide"), lblRemoveIcon

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiMoveMultipleFavoritesSave:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

GuiControlGet, strParentMenu, , drpParentMenu
if (strParentMenu = strCurrentMenu)
	return

Gui, 1:Default
Gui, ListView, lvFavoritesList
intRowToEdit := 0

Loop
{
	intRowToEdit := LV_GetNext(intRowToEdit)
	if (!intRowToEdit)
        break
	intNewItemPos := arrMenus[strParentMenu].MaxIndex() + 1 ; add favorite at end of destination menu
	LV_GetText(strFavoriteShortName, intRowToEdit, 1)
	LV_GetText(strFavoriteLocation, intRowToEdit, 2)
	LV_GetText(strCurrentSubmenuFullName, intRowToEdit, 4)
	LV_GetText(strFavoriteType, intRowToEdit, 5)
	LV_GetText(strCurrentIconResource, intRowToEdit, 6)
	LV_GetText(strAppArguments, intRowToEdit, 7)
	LV_GetText(strAppWorkingDir, intRowToEdit, 8)

	blnRadioFolder := (strFavoriteType = "F")
	blnRadioSpecial := (strFavoriteType = "P")
	blnRadioFile := (strFavoriteType = "D")
	blnRadioURL := (strFavoriteType = "U")
	blnRadioApplication := (strFavoriteType = "A")
	blnRadioSubmenu := (strFavoriteType = "S")

	Gosub, GuiMoveOneFavoriteSave
	intRowToEdit := intRowToEdit - 1 ; because we deleted the previous item
}

Gosub, BuildMainMenuWithStatus ; update menus
Gosub, GuiEditFavoriteCancel

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiAddFavoriteSave:
GuiEditFavoriteSave:
GuiMoveOneFavoriteSave:
;------------------------------------------------------------
Gui, 2:Submit, NoHide
Gui, 2:+OwnDialogs

if (A_ThisLabel <> "GuiMoveOneFavoriteSave")
{
	GuiControlGet, strParentMenu, , drpParentMenu
	GuiControlGet, intNewItemPos, , drpParentMenuItems
}
if !(intNewItemPos)
	intNewItemPos := 1

if !StrLen(strFavoriteShortName)
{
	Oops(blnRadioSubmenu ? lDialogSubmenuNameEmpty : lDialogFavoriteNameEmpty)
	return
}

if InStr(strFavoriteShortName, "|")
{
	Oops(lDialogFavoriteNameNoPipe)
	return
}

if IsColumnBreak(strFavoriteShortName)
{
	Oops(L(lDialogFavoriteNameNoColumnBreak, strColumnBreakIndicator))
	return
}

if  ((blnRadioFolder or blnRadioFile or blnRadioApplication or blnRadioURL) and !StrLen(strFavoriteLocation))
{
	Oops(lDialogFavoriteLocationEmpty)
	return
}

if !FolderNameIsNew(strFavoriteShortName, (strParentMenu = strCurrentMenu ? "" : strParentMenu))
	and (strFavoriteShortName <> lMenuSeparator) and (!IsColumnBreak(strFavoriteShortName)) ; same name OK for separators
	if ((strParentMenu <> strCurrentMenu) or (strFavoriteShortName <> strCurrentName)) ; same name is OK in current menu
	{
		Oops(lDialogFavoriteNameNotNew, strFavoriteShortName)
		if (A_ThisLabel = "GuiMoveOneFavoriteSave")
			intRowToEdit := intRowToEdit + 1
		return
	}

if (blnRadioSubmenu)
{
	if InStr(strFavoriteShortName, lGuiSubmenuSeparator)
	{
		Oops(L(lDialogFavoriteNameNoSeparator, lGuiSubmenuSeparator))
		return
	}
	if ((A_ThisLabel = "GuiMoveOneFavoriteSave") and InStr(strParentMenu, strCurrentMenu . lGuiSubmenuSeparator . strFavoriteShortName) <> 0)
	{
		Oops(lDialogMenuNotMoveUnderItself, strFavoriteShortName)
		intRowToEdit := intRowToEdit + 1
		return
	}
	
	strNewSubmenuFullName := strParentMenu . lGuiSubmenuSeparator . strFavoriteShortName
	
	if (A_ThisLabel = "GuiAddFavoriteSave")
	{
		strFavoriteLocation := lGuiSubmenuLocation
		
		arrNewMenu := Object() ; array of folders of the new menu
		arrMenus.Insert(strNewSubmenuFullName, arrNewMenu)
	}
	else ; GuiEditFavoriteSave or GuiMoveOneFavoriteSave
	{
		strFavoriteLocation := lGuiSubmenuLocation
		UpdateMenuNameInSubmenus(strCurrentSubmenuFullName, strNewSubmenuFullName) ; change names in arrMenus and arrMenu objects
	}
}
else
	strNewSubmenuFullName := ""

strFavoriteType := (blnRadioSubmenu ? "S" : (blnRadioURL ? "U" : (blnRadioFile ? "D" : (blnRadioURL ? "U" : (blnRadioSpecial ? "P" : (blnRadioApplication ? "A" : "F")))))) ; submenu, document, URL, sPecial, application or else folder

Gosub, 2GuiClose

Gui, 1:Default
GuiControl, 1:Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList

if (strParentMenu = strCurrentMenu)
{
	if (A_ThisLabel = "GuiAddFavoriteSave")
	{
		LV_Modify(0, "-Select")
		LV_Insert(intNewItemPos, "Select Focus"
			, strFavoriteShortName, strFavoriteLocation, strCurrentMenu, strNewSubmenuFullName, strFavoriteType, strCurrentIconResource, strAppArguments, strAppWorkingDir)
		LV_Modify(LV_GetNext(), "Vis")
	
		Gosub, SaveCurrentListviewToMenuObject ; save current LV tbefore update the dropdown menu
		GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
	}
	else ; GuiEditFavoriteSave or GuiMoveOneFavoriteSave
	{
		LV_Modify(intRowToEdit, "Select Focus"
			, strFavoriteShortName, strFavoriteLocation, strCurrentMenu, strNewSubmenuFullName, strFavoriteType, strCurrentIconResource, strAppArguments, strAppWorkingDir)
	}
}
else ; add menu item to selected menu object
{
	objFavorite := Object() ; new menu item
	objFavorite.MenuName := strParentMenu ; parent menu of this menu item
	objFavorite.FavoriteName := strFavoriteShortName ; display name of this menu item
	objFavorite.FavoriteLocation := strFavoriteLocation ; path for this menu item
	objFavorite.SubmenuFullName := strNewSubmenuFullName ; full name of the submenu
	objFavorite.FavoriteType := strFavoriteType ; "F" folder, "P" sPecial, "D" document, "U" for URL or "S" submenu
	objFavorite.IconResource := strCurrentIconResource ; icon resource in format "iconfile,iconindex"
	objFavorite.AppArguments := strAppArguments ; application arguments
	objFavorite.AppWorkingDir := strAppWorkingDir ; application working directory
	arrMenus[objFavorite.MenuName].Insert(intNewItemPos, objFavorite) ; add this menu item to the new menu

	if (A_ThisLabel = "GuiEditFavoriteSave") or (A_ThisLabel = "GuiMoveOneFavoriteSave")
		LV_Delete(intRowToEdit)
	if (A_ThisLabel = "GuiEditFavoriteSave")
		LV_Modify(intRowToEdit, "Select Focus")
}

GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"

Gosub, AjustColumnWidth

if (A_ThisLabel = "GuiEditFavoriteSave") or (A_ThisLabel = "GuiMoveOneFavoriteSave")
{
	Gosub, SaveCurrentListviewToMenuObject ; save current LV tbefore update the dropdown menu
	GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
}

if (A_ThisLabel <> "GuiMoveOneFavoriteSave")
	Gosub, BuildMainMenuWithStatus ; update menus

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
	if (strOldMenu = strNewMenu) ; do not continue if both menus have same name: this destroys the menu
		return
	
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
GuiAddFavoriteCancel:
GuiEditFavoriteCancel:
;------------------------------------------------------------
Gosub, 2GuiClose

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveMultipleFavorites:
;------------------------------------------------------------

GuiControl, Focus, lvFavoritesList
Gui, 1:ListView, lvFavoritesList

MsgBox, 52, %strAppName%, % L(lDialogRemoveMultipleFavorites, LV_GetCount("Selected"))
IfMsgBox, No
	return

Loop
	Gosub, GuiRemoveOneFavorite
until !LV_GetNext()

return
;------------------------------------------------------------


;------------------------------------------------------------
GuiRemoveFavorite:
GuiRemoveOneFavorite:
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
if (A_ThisLabel = "GuiRemoveFavorite")
{
	LV_Modify(intItemToRemove, "Select Focus")
	if !LV_GetNext() ; if last item was deleted, select the new last item
		LV_Modify(LV_GetCount(), "Select Focus")
}
Gosub, AjustColumnWidth

if (strFavoriteType = "S")
{
	Gosub, SaveCurrentListviewToMenuObject ; save current LV before update the dropdown menu
	GuiControl, 1:, drpMenusList, % "|" . BuildMenuTreeDropDown(lMainMenuName, strCurrentMenu) . "|"
}

Gosub, SaveCurrentListviewToMenuObject

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
if (blnUseColors)
	Gui, 2:Color, %strGuiWindowColor%
Gui, 2:+Owner1
Gui, 2:Font, s10 w700, Verdana
Gui, 2:Add, Text, x10 y10 w595 center, % L(lOptionsGuiTitle, strAppName)

Gui, 2:Font, s8 w600, Verdana
Gui, 2:Add, Tab2, vintOptionsTab w620 h400 AltSubmit, %A_Space%%lOptionsOtherOptions% | %lOptionsMouseAndKeyboard% | %lOptionsHotkeys% | %lOptionsThirdParty%%A_Space%

;---------------------------------------
; Tab 1: General options

Gui, 2:Tab, 1

Gui, 2:Font
Gui, 2:Add, Text, x10 y+10 w595 center, % L(lOptionsTabOtherOptionsIntro, strAppName)

; column 1
Gui, 2:Add, Text, y+10 x15 Section, %lOptionsLanguage%
Gui, 2:Add, DropDownList, ys x+10 w120 vdrpLanguage Sort, %lOptionsLanguageLabels%
GuiControl, ChooseString, drpLanguage, %strLanguageLabel%

Gui, 2:Add, CheckBox, y+10 xs w220 vblnOptionsRunAtStartup, %lOptionsRunAtStartup%
GuiControl, , blnOptionsRunAtStartup, % FileExist(A_Startup . "\" . strAppName . ".lnk") ? 1 : 0

Gui, 2:Add, CheckBox, y+10 xs w220 vblnDisplayMenuShortcuts, %lOptionsDisplayMenuShortcuts%
GuiControl, , blnDisplayMenuShortcuts, %blnDisplayMenuShortcuts%

Gui, 2:Add, CheckBox, y+10 xs w220 vblnDisplayTrayTip, %lOptionsTrayTip%
GuiControl, , blnDisplayTrayTip, %blnDisplayTrayTip%

Gui, 2:Add, CheckBox, y+10 xs w220 vblnDisplayIcons, %lOptionsDisplayIcons%
GuiControl, , blnDisplayIcons, %blnDisplayIcons%
if !OSVersionIsWorkstation()
	GuiControl, Disable, blnDisplayIcons

Gui, 2:Add, CheckBox, y+10 xs w220 vblnCheck4Update, %lOptionsCheck4Update%
GuiControl, , blnCheck4Update, %blnCheck4Update%

Gui, 2:Add, CheckBox, y+10 xs w220 vblnOpenMenuOnTaskbar, %lOptionsOpenMenuOnTaskbar%
GuiControl, , blnOpenMenuOnTaskbar, %blnOpenMenuOnTaskbar%

; column 2
Gui, 2:Add, Text, ys x240 Section, %lOptionsIconSize%
Gui, 2:Add, DropDownList, ys x+10 w40 vdrpIconSize Sort, 16|24|32|48|64
GuiControl, ChooseString, drpIconSize, %intIconSize%

Gui, 2:Add, Text, y+7 x240 w200, %lOptionsDisplayMenus%

if (A_OSVersion = "WIN_XP")
{
	Gui, 2:Add, CheckBox, y+10 xs w180 vblnDisplaySpecialFolders, %lOptionsDisplaySpecialFolders%
	GuiControl, , blnDisplaySpecialFolders, %blnDisplaySpecialFolders%
}

Gui, 2:Add, CheckBox, y+10 xs w180 vblnDisplayFoldersInExplorerMenu, %lOptionsDisplayFoldersInExplorerMenu%
GuiControl, , blnDisplayFoldersInExplorerMenu, %blnDisplayFoldersInExplorerMenu%

Gui, 2:Add, CheckBox, y+10 xs w180 vblnDisplayGroupMenu, %lOptionsDisplayGroupMenu%
GuiControl, , blnDisplayGroupMenu, %blnDisplayGroupMenu%

Gui, 2:Add, CheckBox, y+10 xs w180 vblnDisplayClipboardMenu, %lOptionsDisplayClipboardMenu%
GuiControl, , blnDisplayClipboardMenu, %blnDisplayClipboardMenu%

Gui, 2:Add, CheckBox, y+10 xs w180 vblnDisplayCopyLocationMenu, %lOptionsDisplayCopyLocationMenu%
GuiControl, , blnDisplayCopyLocationMenu, %blnDisplayCopyLocationMenu%

Gui, 2:Add, CheckBox, y+10 xs w180 vblnDisplayRecentFolders gDisplayRecentFoldersClicked, %lOptionsDisplayRecentFolders%
GuiControl, , blnDisplayRecentFolders, %blnDisplayRecentFolders%

Gui, 2:Add, Edit, % "y+5 xs+15 w36 h17 vintRecentFolders center " . (blnDisplayRecentFolders ? "" : "hidden"), %intRecentFolders%
Gui, 2:Add, Text, % "yp x+10 vlblRecentFolders " . (blnDisplayRecentFolders ? "" : "hidden"), %lOptionsRecentFolders%

; column 3

Gui, 2:Add, Text, ys x430 Section, %lOptionsTheme%
Gui, 2:Add, DropDownList, ys x+10 w120 vdrpTheme, %strAvailableThemes%
GuiControl, ChooseString, drpTheme, %strTheme%

Gui, 2:Add, CheckBox, y+10 xs w190 vblnRememberSettingsPosition, %lOptionsRememberSettingsPosition%
GuiControl, , blnRememberSettingsPosition, %blnRememberSettingsPosition%

Gui, 2:Add, Text, y+12 xs w190 Section, %lOptionsMenuPositionPrompt%

Gui, 2:Add, Radio, % "y+5 xs w190 vradPopupMenuPosition1 gPopupMenuPositionClicked Group " . (intPopupMenuPosition = 1 ? "Checked" : ""), %lOptionsMenuNearMouse%
Gui, 2:Add, Radio, % "y+5 xs w190 vradPopupMenuPosition2 gPopupMenuPositionClicked " . (intPopupMenuPosition = 2 ? "Checked" : ""), %lOptionsMenuActiveWindow%
Gui, 2:Add, Radio, % "y+5 xs w190 vradPopupMenuPosition3 gPopupMenuPositionClicked " . (intPopupMenuPosition = 3 ? "Checked" : ""), %lOptionsMenuFixPosition%

Gui, 2:Add, Text, % "y+5 xs+18 vlblPopupFixPositionX " . (intPopupMenuPosition = 3 ? "" : "hidden"), %lOptionsPopupFixPositionX%
Gui, 2:Add, Edit, % "yp x+5 w36 h17 vstrPopupFixPositionX center " . (intPopupMenuPosition = 3 ? "" : "hidden"), %arrPopupFixPosition1%
Gui, 2:Add, Text, % "yp x+5 vlblPopupFixPositionY " . (intPopupMenuPosition = 3 ? "" : "hidden"), %lOptionsPopupFixPositionY%
Gui, 2:Add, Edit, % "yp x+5 w36 h17 vstrPopupFixPositionY center " . (intPopupMenuPosition = 3 ? "" : "hidden"), %arrPopupFixPosition2%

;---------------------------------------
; Tab 2: Popup menu hotkeys

Gui, 2:Tab, 2

Gui, 2:Font
Gui, 2:Add, Text, x10 y+10 w595 center, % L(lOptionsTabMouseAndKeyboardIntro, strAppName)

loop, 4
{
	Gui, 2:Font, s8 w700
	Gui, 2:Add, Text, x15 y+20 w610, % arrOptionsTitles%A_Index%
	Gui, 2:Font, s9 w500, Courier New
	Gui, 2:Add, Text, Section x260 y+5 w280 h23 center 0x1000 vlblHotkeyText%A_Index% gButtonOptionsChangeHotkey%A_Index%, % Hotkey2Text(strModifiers%A_Index%, strMouseButton%A_Index%, strOptionsKey%A_Index%)
	Gui, 2:Font
	Gui, 2:Add, Button, yp x555 vbtnChangeHotkey%A_Index% gButtonOptionsChangeHotkey%A_Index%, %lOptionsChangeHotkey%
	Gui, 2:Font, s8 w500
	Gui, 2:Add, Text, x15 ys w240, % arrOptionsTitlesSub%A_Index%
}

;---------------------------------------
; Tab 3: Other hotkeys

Gui, 2:Tab, 3

Gui, 2:Font
Gui, 2:Add, Text, x10 y+10 w595 center, %lOptionsTabHotkeysIntro%

loop, 6
{
	intIndex := A_Index + 4
	Gui, 2:Font, s8 w700
	Gui, 2:Add, Text, Section x15 y+10, % arrOptionsTitles%intIndex%
	Gui, 2:Font, s9 w500, Courier New
	Gui, 2:Add, Text, x260 ys+5 w280 h23 center 0x1000 vlblHotkeyText%intIndex% gButtonOptionsChangeHotkey%intIndex%, % Hotkey2Text(strModifiers%intIndex%, strMouseButton%intIndex%, strOptionsKey%intIndex%)
	Gui, 2:Font
	Gui, 2:Add, Button, yp x555 vbtnChangeHotkey%intIndex% gButtonOptionsChangeHotkey%intIndex%, %lOptionsChangeHotkey%
	Gui, 2:Font, s8 w500
	Gui, 2:Add, Text, x15 ys+15 w240, % arrOptionsTitlesSub%intIndex%
}

Gui, 2:Add, CheckBox, y+35 x20 vblnDisplaySpecialMenusShortcuts, %lOptionsDisplaySpecialMenusShortcuts%
GuiControl, , blnDisplaySpecialMenusShortcuts, %blnDisplaySpecialMenusShortcuts%


;---------------------------------------
; Tab 4: File Managers

Gui, 2:Tab, 4

Gui, 2:Add, Text, x10 y+10 w595 center, %lOptionsTabFileManagersIntro%

Gui, 2:Font, s8 w700
Gui, 2:Add, Link, y+15 x15, % L(lOptionsThirdPartyTitle, "Directory Opus") . " (<a href=""http://code.jeanlalonde.ca/using-folderspopup-with-directory-opus/"">" . lGuiHelp . "</a>)"
Gui, 2:Font
Gui, 2:Add, Text, y+5 x15, % L(lOptionsThirdPartyDetail, "Directory Opus")
Gui, 2:Add, Text, y+10 x15, %lOptionsThirdPartyPrompt%
Gui, 2:Add, Edit, x+10 yp w300 h20 vstrDirectoryOpusPath, %strDirectoryOpusPath%
Gui, 2:Add, Button, x+10 yp vbtnSelectDOpusPath gButtonSelectDOpusPath, %lDialogBrowseButton%
Gui, 2:Add, Checkbox, x+10 yp vblnDirectoryOpusUseTabs, %lOptionsDirectoryOpusUseTabs%
GuiControl, , blnDirectoryOpusUseTabs, %blnDirectoryOpusUseTabs%

Gui, 2:Font, s8 w700
Gui, 2:Add, Link, y+25 x15, % L(lOptionsThirdPartyTitle, "Total Commander") . " (<a href=""http://code.jeanlalonde.ca/using-folderspopup-with-total-commander/"">" . lGuiHelp . "</a>)"
Gui, 2:Font
Gui, 2:Add, Text, y+5 x15, % L(lOptionsThirdPartyDetail, "Total Commander")
Gui, 2:Add, Text, y+10 x15, %lOptionsThirdPartyPrompt%
Gui, 2:Add, Edit, x+10 yp w300 h20 vstrTotalCommanderPath, %strTotalCommanderPath%
Gui, 2:Add, Button, x+10 yp vbtnSelectTCPath gButtonSelectTCPath, %lDialogBrowseButton%
Gui, 2:Add, Checkbox, x+10 yp vblnTotalCommanderUseTabs, %lOptionsTotalCommanderUseTabs%
GuiControl, , blnTotalCommanderUseTabs, %blnTotalCommanderUseTabs%

Gui, 2:Font, s8 w700
Gui, 2:Add, Link, y+25 x15, %lOptionsThirdPartyTitleFPconnect% (<a href="https://github.com/rolandtoth/FPconnect">%lGuiHelp%</a>)
Gui, 2:Font
Gui, 2:Add, Text, y+5 x15, %lOptionsThirdPartyDetailFPconnect%
Gui, 2:Add, Text, y+10 x15, %lOptionsThirdPartyPrompt%
Gui, 2:Add, Edit, x+10 yp w300 h20 vstrFPconnectPath, %strFPconnectPath%
Gui, 2:Add, Button, x+10 yp vbtnSelectFPcPath gButtonSelectFPcPath, %lDialogBrowseButton%

;---------------------------------------
; Build Gui footer

Gui, 2:Tab

GuiControlGet, arrTabPos, Pos, intOptionsTab

Gui, 2:Add, Button, % "y" . arrTabPosY + arrTabPosH + 10 . " x10 vbtnOptionsSave gButtonOptionsSave Default", %lGuiSave%
Gui, 2:Add, Button, yp vbtnOptionsCancel gButtonOptionsCancel, %lGuiCancel%
Gui, 2:Add, Button, yp vbtnOptionsDonate gGuiDonate, %lDonateButton%
GuiCenterButtons(L(lOptionsGuiTitle, strAppName, strAppVersion), 10, 5, 20, "btnOptionsSave", "btnOptionsCancel", "btnOptionsDonate")

Gui, 2:Add, Text
GuiControl, Focus, btnOptionsSave

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
ButtonSelectFPcPath:
;------------------------------------------------------------

if StrLen(strFPconnectPath) and (strFPconnectPath <> "NO")
	strCurrentFPcLocation := strFPconnectPath
else
	strCurrentFPcLocation := A_ScriptDir . "\FPconnect\FPconnect.exe"

FileSelectFile, strNewFPcLocation, 3, %strCurrentFPcLocation%, %lDialogAddFolderSelect%

if !(StrLen(strNewFPcLocation))
	return

GuiControl, 2:, strFPconnectPath, %strNewFPcLocation%

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
if (A_OSVersion = "WIN_XP")
	IniWrite, %blnDisplaySpecialFolders%, %strIniFile%, Global, DisplaySpecialFolders
IniWrite, %blnDisplaySpecialMenusShortcuts%, %strIniFile%, Global, DisplaySpecialMenusShortcuts
IniWrite, %blnDisplayRecentFolders%, %strIniFile%, Global, DisplayRecentFolders
IniWrite, %intRecentFolders%, %strIniFile%, Global, RecentFolders
IniWrite, %blnDisplayFoldersInExplorerMenu%, %strIniFile%, Global, DisplayFoldersInExplorerMenu
IniWrite, %blnDisplayGroupMenu%, %strIniFile%, Global, DisplaySwitchMenu ; keep "Switch" for backward compatibility
IniWrite, %blnDisplayClipboardMenu%, %strIniFile%, Global, DisplayClipboardMenu
IniWrite, %blnDisplayCopyLocationMenu%, %strIniFile%, Global, DisplayCopyLocationMenu
IniWrite, %blnDisplayMenuShortcuts%, %strIniFile%, Global, DisplayMenuShortcuts
IniWrite, %blnCheck4Update%, %strIniFile%, Global, Check4Update
IniWrite, %blnOpenMenuOnTaskbar%, %strIniFile%, Global, OpenMenuOnTaskbar
IniWrite, %blnRememberSettingsPosition%, %strIniFile%, Global, RememberSettingsPosition

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

IniWrite, %strFPconnectPath%, %strIniFile%, Global, FPconnectPath
blnUseFPconnect := StrLen(strFPconnectPath)
if (blnUseFPconnect)
{
	blnUseFPconnect := FileExist(strFPconnectPath)
	if (blnUseFPconnect)
		Gosub, SetFPconnect
}

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
if (A_OSVersion = "WIN_XP")
	Gosub, BuildSpecialFoldersMenu
Gosub, BuildFoldersInExplorerMenu
Gosub, BuildGroupMenu

; and rebuild Folders menus w/ or w/o optional folders and shortcuts
for strMenuName, arrMenu in arrMenus
{
	Menu, %strMenuName%, Add
	Menu, %strMenuName%, DeleteAll
	arrMenu := ; free object's memory
}
Gosub, BuildMainMenuWithStatus

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
ButtonOptionsChangeHotkey7:
ButtonOptionsChangeHotkey8:
ButtonOptionsChangeHotkey9:
ButtonOptionsChangeHotkey10:
;------------------------------------------------------------

StringReplace, intIndex, A_ThisLabel, ButtonOptionsChangeHotkey

intGui2WinID := WinExist("A")
Gui, 2:Submit, NoHide

Gui, 3:New, , % L(lOptionsChangeHotkeyTitle, strAppName, strAppVersion)
if (blnUseColors)
	Gui, 3:Color, %strGuiWindowColor%
Gui, 3:+Owner2
Gui, 3:Font, s10 w700, Verdana
Gui, 3:Add, Text, x10 y10 w350 center, % L(lOptionsChangeHotkeyTitle, strAppName)
Gui, 3:Font

intHotkeyType := 3 ; Folders in Explorer, Groups, Recent folders, Clipboard and Settings
if InStr(arrIniKeyNames%intIndex%, "Mouse")
	intHotkeyType := 1
if InStr(arrIniKeyNames%intIndex%, "Keyboard")
	intHotkeyType := 2

Gui, 3:Add, Text, y+15 x10 , %lOptionsTriggerFor%
Gui, 3:Font, s8 w700
Gui, 3:Add, Text, x+5 yp w300, % arrOptionsTitles%intIndex%
Gui, 3:Font

Gui, 3:Add, Text, x10 y+5 w350, % lOptionsArrDescriptions%intIndex%

Gui, 3:Add, CheckBox, y+20 x50 vblnOptionsShift, %lOptionsShift%
GuiControl, , blnOptionsShift, % InStr(strModifiers%intIndex%, "+") ? 1 : 0
GuiControlGet, arrTop, Pos, blnOptionsShift
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
	Gui, 3:Add, DropDownList, % "y" . arrTopY . " x150 w200 vstrOptionsMouse gOptionsMouseChanged", % strMouseButtonsWithDefault%intIndex%
if (intHotkeyType <> 1)
{
	Gui, 3:Add, Text, % "y" . arrTopY . " x150 w60", %lOptionsKeyboard%
	Gui, 3:Add, Hotkey, yp x+10 w130 vstrOptionsKey gOptionsHotkeyChanged
	GuiControl, , strOptionsKey, % strOptionsKey%intIndex%
}
if (intHotkeyType = 3)
{
	Gui, 3:Add, Text, % "y" . arrTopY + 30 . " x150 w60", %lOptionsMouse%
	Gui, 3:Add, DropDownList, yp x+10 w130 vstrOptionsMouse gOptionsMouseChanged, % strMouseButtonsWithDefault%intIndex%
}
if (intHotkeyType <> 1)
{
	Gui, 3:Add, Link, y+10 x150 gOptionsSelectNoneHotkeyClicked, <a>%lOptionsMouseNone%</a>
	Gui, 3:Add, Link, yp x+10 w130 gOptionsHotkeySpaceClicked, <a>%lOptionsSpacebar%</a>
}

Gui, 3:Add, Button, % "x10 y" . arrTopY + 100 . " vbtnResetHotkey gButtonResetHotkey" . intIndex, %lGuiResetDefault%
GuiCenterButtons(L(lOptionsChangeHotkeyTitle, strAppName, strAppVersion), 10, 5, 20, "btnResetHotkey")
Gui, 3:Add, Button, y+20 x10 vbtnChangeHotkeySave gButtonChangeHotkeySave%intIndex%, %lGuiSave%
Gui, 3:Add, Button, yp x+20 vbtnChangeHotkeyCancel gButtonChangeHotkeyCancel, %lGuiCancel%
GuiCenterButtons(L(lOptionsChangeHotkeyTitle, strAppName, strAppVersion), 10, 5, 20, "btnChangeHotkeySave", "btnChangeHotkeyCancel")

Gui, 3:Add, Text
GuiControl, Focus, btnChangeHotkeySave
Gui, 3:Show, AutoSize Center
Gui, 2:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
OptionsSelectNoneHotkeyClicked:
;------------------------------------------------------------

GuiControl, , strOptionsKey, %lOptionsMouseNone%
GuiControl, Choose, strOptionsMouse, %lOptionsMouseNone%
GuiControl, , blnOptionsShift, 0
GuiControl, , blnOptionsCtrl, 0
GuiControl, , blnOptionsAlt, 0
GuiControl, , blnOptionsWin, 0

strModifiers%intIndex% := "" ; intIndex comes from N in ButtonOptionsChangeHotkeyN label

return
;------------------------------------------------------------


;------------------------------------------------------------
OptionsHotkeySpaceClicked:
;------------------------------------------------------------

GuiControl, , strOptionsKey, %A_Space%
GuiControl, Choose, strOptionsMouse, 0

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
ButtonChangeHotkeySave7:
ButtonChangeHotkeySave8:
ButtonChangeHotkeySave9:
ButtonChangeHotkeySave10:
;------------------------------------------------------------
Gui, 3:Submit

StringReplace, intIndex, A_ThisLabel, ButtonChangeHotkeySave

StringSplit, arrIniVarNames, strIniKeyNames, |

if StrLen(strOptionsMouse)
	strOptionsMouse := GetMouseButton4Text(strOptionsMouse) ; get mouse button system name from dropdown localized text
else
	strMouseButton%intIndex% := "" ;  empty mouse button text

strHotkey := Trim(strOptionsKey . strOptionsMouse)
if !StrLen(strHotkey)
	strHotkey := "None"

if (strHotkey = "None") ; do not compare with lOptionsMouseNone because it is translated
{
	Hotkey, % "$" . arrHotkeys%intIndex%, Off, UseErrorLevel ; UseErrorLevel in case we had an invalid key name in the ini file
	IniWrite, None, %strIniFile%, Global, % arrIniVarNames%intIndex% ; do not write lOptionsMouseNone because it is translated
}
else
{
	; Order of modifiers important to keep modifiers labels in correct order
	if (blnOptionsWin)
		strHotkey := "#" . strHotkey
	if (blnOptionsAlt)
		strHotkey := "!" . strHotkey
	if (blnOptionsCtrl)
		strHotkey := "^" . strHotkey
	if (blnOptionsShift)
		strHotkey := "+" . strHotkey

	if (strHotkey = "LButton")
	{
		Oops(lOptionsMouseCheckLButton)
		Gosub, 3GuiClose
		return
	}

	Hotkey, % "$" . arrHotkeys%intIndex%, Off, UseErrorLevel ; UseErrorLevel in case we had an invalid key name in the ini file
	IniWrite, %strHotkey%, %strIniFile%, Global, % arrIniVarNames%intIndex%
}

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
ButtonResetHotkey7:
ButtonResetHotkey8:
ButtonResetHotkey9:
ButtonResetHotkey10:
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
if (blnUseColors)
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
GuiCenterButtons(L(lAboutTitle, strAppName, strAppVersion), 10, 5, 20, "btnAboutDonate", "btnAboutClose")

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
if (blnUseColors)
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
GuiCenterButtons(L(lDonateTitle, strAppName, strAppVersion), 10, 5, 20, "btnDonateClose")

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
if (blnUseColors)
	Gui, 2:Color, %strGuiWindowColor%
Gui, 2:+Owner1
intWidth := 600
Gui, 2:Font, s12 w700, Verdana
Gui, 2:Add, Text, x10 y10, %strAppName%
Gui, 2:Font, s10 w400, Verdana
Gui, 2:Add, Link, x10 w%intWidth%, %lHelpTextLead%

Gui, 2:Font, s8 w600, Verdana
Gui, 2:Add, Tab2, vintHelpTab w640 h400 AltSubmit, %A_Space%%lHelpTabGettingStarted% | %lHelpTabAddingFavorite% | %lHelpTabTitlesTipsAndTricks%%A_Space%

; Hotkeys: 1) PopupHotkeyMouse 2) PopupHotkeyNewMouse 3) PopupHotkeyKeyboard 4) PopupHotkeyNewKeyboard
; 5) SettingsHotkey 6) FoldersInExplorerHotkey 7) GroupsHotkey 8) RecentsHotkey 9) ClipboardHotkey 10) CopyLocationHotkey
Gui, 2:Font, s8 w400, Verdana
Gui, 2:Tab, 1
Gui, 2:Add, Link, w%intWidth%, % L(lHelpText1, Hotkey2Text(strModifiers1, strMouseButton1, strOptionsKey1), Hotkey2Text(strModifiers3, strMouseButton3, strOptionsKey3))
Gui, 2:Add, Link, w%intWidth%, % lHelpText2
Gui, 2:Add, Button, vbtnNext1 gNextHelpButtonClicked, %lDialogTabNext%
GuiCenterButtons(L(lHelpTitle, strAppName, strAppVersion), 10, 5, 20, "btnNext1")

Gui, 2:Tab, 2
Gui, 2:Add, Link, w%intWidth%, % L(lHelpText3, Hotkey2Text(strModifiers1, strMouseButton1, strOptionsKey1), Hotkey2Text(strModifiers3, strMouseButton3, strOptionsKey3))
Gui, 2:Add, Link, w%intWidth%, % L(lHelpText4, Hotkey2Text(strModifiers5, strMouseButton5, strOptionsKey5))
Gui, 2:Add, Button, vbtnNext2 gNextHelpButtonClicked, %lDialogTabNext%
GuiCenterButtons(L(lHelpTitle, strAppName, strAppVersion), 10, 5, 20, "btnNext2")

Gui, 2:Tab, 3
Gui, 2:Add, Link, w%intWidth%, % L(lHelpText5
	, Hotkey2Text(strModifiers2, strMouseButton2, strOptionsKey2)
	, Hotkey2Text(strModifiers4, strMouseButton4, strOptionsKey4)
	, Hotkey2Text(strModifiers8, strMouseButton8, strOptionsKey8)
	, Hotkey2Text(strModifiers6, strMouseButton6, strOptionsKey6)
	, Hotkey2Text(strModifiers7, strMouseButton7, strOptionsKey7)
	, Hotkey2Text(strModifiers9, strMouseButton9, strOptionsKey9)
	, Hotkey2Text(strModifiers10, strMouseButton10, strOptionsKey10))
Gui, 2:Add, Link, w%intWidth%, % lHelpText6

Gui, 2:Tab

GuiControlGet, arrTabPos, Pos, intHelpTab
Gui, 2:Add, Button, % "x180 y" . arrTabPosY + arrTabPosH + 10. " vbtnHelpDonate gGuiDonate", %lDonateButton%
Gui, 2:Add, Button, x+80 yp g2GuiClose vbtnHelpClose, %lGui2Close%
GuiCenterButtons(L(lHelpTitle, strAppName, strAppVersion), 10, 5, 20, "btnHelpDonate", "btnHelpClose")

GuiControl, Focus, btnHelpClose
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled

return
;------------------------------------------------------------


;------------------------------------------------------------
NextHelpButtonClicked:
;------------------------------------------------------------

Gui, 2:Submit, NoHide

GuiControl, Choose, intHelpTab, % intHelpTab + 1

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

	if (strCommand = "/info")
	{
		Process, Exist, dopus.exe
		; abort if DOpus.exe is not running
		if !(ErrorLevel)
			return
	}
	if FileExist(strDirectoryOpusRtPath)
		Run, % """" . strDirectoryOpusRtPath . """ " . strCommand . " """ . strLocation . """" . strParam
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
		
		; disable Folders in Explorer and Group menus for TC users until the tabs issue is resolved
		blnDisplayFoldersInExplorerMenu := 0
		IniWrite, %blnDisplayFoldersInExplorerMenu%, %strIniFile%, Global, DisplayFoldersInExplorerMenu
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
CheckFPconnect:
;------------------------------------------------------------

strCheckFPconnectPath := A_ScriptDir . "\FPconnect\FPconnect.exe"

if FileExist(strCheckFPconnectPath)
{
	MsgBox, 52, %strAppName%, % L(lDialogThirdPartyDetected, strAppName, "FPconnect")
	IfMsgBox, No
		strFPconnectPath := "NO"
	else
	{
		strFPconnectPath := strCheckFPconnectPath
		Gosub, SetFPconnect
		
		; disable Folders in Explorer and Group menus for FPconnect users
		blnDisplayFoldersInExplorerMenu := 0
		IniWrite, %blnDisplayFoldersInExplorerMenu%, %strIniFile%, Global, DisplayFoldersInExplorerMenu
		blnDisplayGroupMenu := 0
		IniWrite, %blnDisplayGroupMenu%, %strIniFile%, Global, DisplaySwitchMenu
	}
	blnUseFPconnect := (strFPconnectPath <> "NO")
	IniWrite, %strFPconnectPath%, %strIniFile%, Global, FPconnectPath
}

return
;------------------------------------------------------------


;------------------------------------------------------------
SetFPconnect:
;------------------------------------------------------------

StringTrimRight, strFPconnectIniPath, strFPconnectPath, 4
strFPconnectIniPath := strFPconnectIniPath . ".ini"

IniRead, strFPconnectAppPathFilename, %strFPconnectIniPath%, Options, AppPath, %A_Space% ; empty by default
blnUseFPconnect := FileExist(EnvVars(strFPconnectAppPathFilename))

IniRead, strFPconnectTargetPathFilename, %strFPconnectIniPath%, Options, TargetPath, %A_Space% ; empty by default

if (blnUseFPconnect)
{
	strFPconnectAppPathFilename := EnvVars(strFPconnectAppPathFilename)
	SplitPath, strFPconnectAppPathFilename, strFPconnectAppPathFilename
	strFPconnectTargetPathFilename := EnvVars(strFPconnectTargetPathFilename)
	SplitPath, strFPconnectTargetPathFilename, strFPconnectTargetPathFilename
}
else
	Oops(lOopsWrongFPconnectAppPathFilename, strFPconnectPath, strFPconnectIniPath)

return
;------------------------------------------------------------



;========================================================================================================================
; VARIOUS FUNCTIONS
;========================================================================================================================


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


;------------------------------------------------
DiagGroupLoad(strName, strData)
;------------------------------------------------
{
	global strGroupLoadDiag
	
	FormatTime, strNow, %A_Now%, yyyyMMdd@HH:mm:ss
	strGroupLoadDiag := strGroupLoadDiag . strNow . A_MSec . "`t" . strName . "`t" . strData . "`n"
}
;------------------------------------------------


;------------------------------------------------------------
GuiCenterButtons(strWindow, intInsideHorizontalMargin := 10, intInsideVerticalMargin := 0, intDistanceBetweenButtons := 20, arrControls*)
; This is a variadic function. See: http://ahkscript.org/docs/Functions.htm#Variadic
;------------------------------------------------------------
{
	Gui, Show, Hide ; why?
	WinGetPos, , , intWidth, , %strWindow%

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
BuildSpecialMenuItemName(intHotkeyIndex, strMenuItemName)
;------------------------------------------------------------
{
	global
	
	if (blnDisplaySpecialMenusShortcuts)
	{
		strHotkey := Hotkey2Text(strModifiers%intHotkeyIndex%, strMouseButton%intHotkeyIndex%, strOptionsKey%intHotkeyIndex%, true)
		if (strHotkey <> lOptionsMouseNone)
			strMenuItemName := strMenuItemName . " (" . strHotkey . ")"
	}

	return strMenuItemName
}
;------------------------------------------------------------


;------------------------------------------------------------
Hotkey2Text(strModifiers, strMouseButton, strOptionKey, blnShort := false)
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
				str := str . (blnShort ? lOptionsCtrlShort : lOptionsCtrl) . "+"
			if (A_LoopField = "+")
				str := str . lOptionsShift . "+"
			if (A_LoopField = "#")
				str := str . (blnShort ? lOptionsWinShort : lOptionsWin) . "+"
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
; Returns the string in arrIniKeyNames at the same position of strSource in arrHotkeys array
;------------------------------------------------------------
{
	global
	
	loop, %arrIniKeyNames0%
		if (strSource = "$" . arrHotkeys%A_Index%)
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
GetOSVersion()
;------------------------------------------------------------
{
	if (GetOSVersionInfo().MajorVersion = 10)
		return "WIN_10"
	else
		return A_OSVersion
}
;------------------------------------------------------------


;------------------------------------------------------------
OSVersionIsWorkstation()
;------------------------------------------------------------
{
	return (GetOSVersionInfo() and (GetOSVersionInfo().ProductType = 1))
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


;------------------------------------------------------------
IsColumnBreak(strMenuName)
;------------------------------------------------------------
{
	return (SubStr(strMenuName, 1, StrLen(strColumnBreakIndicator)) = strColumnBreakIndicator)
}
;------------------------------------------------------------


;------------------------------------------------------------
GetIcon4Location(strLocation, ByRef strDefaultIcon, ByRef intDefaultIcon, strDefaultType := "")
; get icon, extract from kiu http://www.autohotkey.com/board/topic/8616-kiu-icons-manager-quickly-change-icon-files/
;------------------------------------------------------------
{
	global blnDiagMode
	global objIconsFile
	global objIconsIndex
	global blnRadioApplication
	
	if !StrLen(strLocation)
	{
		if (blnRadioApplication)
		{
			strDefaultIcon := objIconsFile["Application"]
			intDefaultIcon := objIconsIndex["Application"]
		}
		else
		{
			strDefaultIcon := objIconsFile["UnknownDocument"]
			intDefaultIcon := objIconsIndex["UnknownDocument"]
		}
		return
	}
	
	SplitPath, strLocation, , , strExtension
	RegRead, strHKeyClassRoot, HKEY_CLASSES_ROOT, .%strExtension%
	RegRead, strRegistryIconResource, HKEY_CLASSES_ROOT, %strHKeyClassRoot%\DefaultIcon
	if (blnDiagMode)
	{
		Diag("BuildOneMenuIcon", strLocation)
		Diag("strHKeyClassRoot", strHKeyClassRoot)
		Diag("strRegistryIconResource", strRegistryIconResource)
	}
	
	if (strRegistryIconResource = "%1") ; use the file itself (for executable)
	{
		strDefaultIcon := strLocation
		intDefaultIcon := 1
		return
	}
	
	ParseIconResource(strRegistryIconResource, strDefaultIcon, intDefaultIcon, strDefaultType)

	if (blnDiagMode)
	{
		Diag("strDefaultIcon", strDefaultIcon)
		Diag("intDefaultIcon", intDefaultIcon)
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
ParseIconResource(strIconResource, ByRef strIconFile, ByRef intIconIndex, strDefaultType := "")
;------------------------------------------------------------
{
	global objIconsFile
	global objIconsIndex
	
	if !StrLen(strDefaultType)
		strDefaultType := "UnknownDocument"
	
	if StrLen(strIconResource)
		If InStr(strIconResource, ",") ; this is icongroup files
		{
			intIconResourceCommaPosition := InStr(strIconResource, ",", , 0) ; search reverse
			StringLeft, strIconFile, strIconResource, % intIconResourceCommaPosition - 1
			StringRight, intIconIndex, strIconResource, % StrLen(strIconResource) - intIconResourceCommaPosition
		}
		else
		{
			strIconFile := strIconResource
			intIconIndex := 1
		}
	else
	{
		strIconFile := objIconsFile[strDefaultType]
		intIconIndex := objIconsIndex[strDefaultType]
	}
}
;------------------------------------------------------------


;------------------------------------------------------------
EnvVars(str)
; from Lexikos http://www.autohotkey.com/board/topic/40115-func-envvars-replace-environment-variables-in-text/#entry310601
;------------------------------------------------------------
{
    if sz:=DllCall("ExpandEnvironmentStrings", "uint", &str
                    , "uint", 0, "uint", 0)
    {
        VarSetCapacity(dst, A_IsUnicode ? sz*2:sz)
        if DllCall("ExpandEnvironmentStrings", "uint", &str
                    , "str", dst, "uint", sz)
            return dst
    }
    return src
}
;------------------------------------------------------------


;------------------------------------------------------------
WindowsIsVersion7OrMore()
;------------------------------------------------------------
{
	return !InStr("WIN_VISTA|WIN_2003|WIN_XP|WIN_2000", A_OSVersion)
}
;------------------------------------------------------------


;------------------------------------------------------------
PathCombine(strAbsolutePath, strRelativePath)
; see http://www.autohotkey.com/board/topic/17922-func-relativepath-absolutepath/page-3#entry117355
; and http://stackoverflow.com/questions/29783202/combine-absolute-path-with-a-relative-path-with-ahk/
;------------------------------------------------------------
{
    VarSetCapacity(strCombined, (A_IsUnicode ? 2 : 1) * 260, 1) ; MAX_PATH
    DllCall("Shlwapi.dll\PathCombine", "UInt", &strCombined, "UInt", &strAbsolutePath, "UInt", &strRelativePath)
    Return, strCombined
}
;------------------------------------------------------------


