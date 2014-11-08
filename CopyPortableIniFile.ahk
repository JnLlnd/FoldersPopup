#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off

strStartupShortcut := A_Startup . "\FoldersPopup.lnk"
strUserDataFolder := A_AppData . "\FoldersPopup"
strUserDataIniFile := strUserDataFolder . "\FoldersPopup.ini"

if !FileExist(strUserDataFolder)
	return ; strUserDataFolder  does not exist - the setup did not complete - abort

if FileExist(strUserDataIniFile)
	return ; strUserDataIniFile already exists - do not overwrite - abort

if !FileExist(strStartupShortcut)
	return ; strStartupShortcut does not exist - no portable shortcut to read to get the portable folder - abort

FileGetShortcut, %strStartupShortcut%, , strStartupWorkingDir
if InStr(strStartupWorkingDir, A_AppData)
	return ; the data folder is already under the user data folder - do not risk to overwrite - abort

; we have an existing portable FoldersPopup.ini file to place as an initial ini file in the brand new user data folder - copy!
FileCopy, %strStartupWorkingDir%\FoldersPopup.ini, %strUserDataIniFile%, 0 ; do not overwrite
