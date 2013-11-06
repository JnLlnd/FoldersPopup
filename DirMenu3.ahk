;===============================================
/*
DirMenu3
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By Jean Lalonde (JnLlnd on AHKScript.org forum), based on DirMenu v2 by Robert Ryan (rbrtryn on AutoHotkey.com forum)

	Version 0.1 ALPHA
	- init skeleton, read ini file and create arrays for folders menu and supported dialog boxes

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

#NoEnv
#SingleInstance force
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

LoadIniFile()
BuildGUI()
BuildTrayMenu()
LoadSettings()

return



LoadIniFile()
{
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
return
}



BuildGui()
{
}



BuildTrayMenu()
{
}


LoadSettings()
{
}


