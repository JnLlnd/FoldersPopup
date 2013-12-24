# FoldersPopup - Read me

Popup menu to jump from folders to folders. Freeware.

Written using AHKScript v1.1.09.03+ (http://www.ahkscript.org)  
By JnLlnd on [AHKScript forum](http://ahkscript.org/boards/memberlist.php?mode=viewprofile&u=66), based on DirMenu v2 by Robert Ryan 2012 (rbrtryn on [AutoHotkey.com forum](http://www.autohotkey.com/board/user/15020-rbrtryn/))  
Thanks to LearningOne for sharing his code and to others who helped in [this thread on AHKScript.org forum](http://ahkscript.org/boards/viewtopic.php?f=5&t=526).

## Links

* [Application home](http://code.jeanlalonde.ca/folderspopup/)
* [Download 32-bits / 64-bits](http://code.jeanlalonde.ca/ahk/folderspopup/folderspopup.zip)

## History

### v1.01
* bug fix: mouse and keyboard triggers were disabled in non-explorer windows

### v1.0 (First Official Release)
* configurable mouse button and keyboard triggers in a new "Options" dialog box
* new keyboard triggers (by default, Windows-K and Shift-Windows-K) in addition to mouse button triggers (by default, Middle mouse and Shift-Middle mouse buttons)
* add "Run at startup" checkbox to "Options" dialog box to launch Folders Popup automatically at Windows startup
* add "Display the startup tray tip" checkbox to "Options" dialog box to display or hide the Folders popup's tray tip
* add "Display Special Folders" checkbox to "Options" dialog box to enable/disable navigation to special folders (My Computer, Network, Recycle bion, etc.) in popup menu
* better formated startup help tray tip
* close "Settings" dialog box with Escape key

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


