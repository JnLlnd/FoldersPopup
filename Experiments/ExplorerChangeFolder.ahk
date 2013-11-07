#NoEnv
#SingleInstance force

Global strEnv := A_AhkVersion . " | " . A_OSVersion . " | " . A_Is64bitOS . " | " . A_Language
Global strLog := "Test_Name | Test_Result | A_AhkVersion | A_OSVersion | A_Is64bitOS | A_Language`r`n"
StringReplace, strLogFile, A_ScriptName, .ahk
strLogFile := A_Temp . "\" . strLogFile . ".log"

Info("The script will open ""Windows Explorer"".")
run, Explorer
Info("The script will now perform the same task (changing folders in Explorer) using 4 different methods.")

if (A_OSVersion = "WIN_XP")
{
	; WinXP ControlSend Method
	strFolder := "C:\Windows"
	Info("Method 1)`nIn the next step, the folder in Explorer should change to your """ . strFolder . """ folder.")
	if not ControlEdit1Exist()
	{
		PostMessage, 0x111, 41477, 0, , A ; Show Address Bar
		while not ControlEdit1Exist()
			Sleep 0
		PostMessage, 0x111, 41477, 0, , A ; Hide Address Bar
	}
	ControlFocus, Edit1, A
	ControlSetText, Edit1, %strFolder%, A
	ControlSend, Edit1, {Enter}, A
	CheckResult("Explorer_CtrlSendXP", "Is your Explorer folder changed to """ . strFolder . """?")
}
else
	Info("Method 1)`nThis method is reserved for Windows XP users. This test is skipped because you are running on " . A_OSVersion . ".")

; F4 Method
SplitPath, A_AhkPath, , strFolder
Info("Method 2)`nIn the next step, the folder in Explorer should change to your """ . strFolder . """ folder.")
Send, {F4}{Esc}
Sleep, 500 ; long delay for safety
Send, %strFolder%{Enter}
CheckResult("Explorer_F4Esc", "Is your Explorer folder changed to """ . strFolder . """?")

/*
if (A_Language <> "0407")
{
	; Alt-D Method
	strFolder := A_MyDocuments
	Info("Method 3)`nIn the next step, the folder in Explorer should change to your """ . strFolder . """ folder.")
	Send, !d
	Sleep, 500 ; long delay for safety
	Send, %strFolder%{Enter}
	CheckResult("Explorer_AltD", "Is your Explorer folder changed to """ . strFolder . """?")
}
else
	Info("Method 2)`nThis method is skipped for German language.")
*/

; ControlSend Method
strFolder := A_ScriptDir
Info("Method 3`nIn this step, the folder in Explorer should change to your """ . strFolder . """ folder.")
ControlSetText, Edit1, %strFolder%, A
ControlSend, Edit1, {Enter}, A
CheckResult("Explorer_CtrlSend", "Is your Explorer folder changed to """ . strFolder . """?")

; ControlSend Method
strFolder := "C:\Windows"
Info("Method 4)`nIn this last step, the folder in Explorer should change to your """ . strFolder . """ folder.")
Explorer_Navigate(strFolder)
CheckResult("Explorer_Shell", "Is your Explorer folder changed to """ . strFolder . """?")

Send, !{F4}

Info("One last test? We will test if changing folder in Dialog box works well on your system.`n`nThe script will run Notepad and open the ""Open"" dialog box.")
run, Notepad, , , strPID
Sleep, 500 ; long delay for safety
Send, ^o
strFolder := "C:\"
Info("In the next step, the file list in the dialog box should change to your """ . strFolder . """ folder.")
ControlFocus, Edit1, A
ControlGetText, strOldText, Edit1, A
ControlSetText, Edit1, %strFolder%, A
ControlSend, Edit1, {Enter}, A
ControlSetText, Edit1, %strOldText%, A
CheckResult("Notepad_ControlSend", "Is your file list now showing your """ . strFolder . """ folder?")
Info("Thank you. The script will now close Notepad.")
Send, !{F4}
Sleep, 500 ; long delay for safety
WinActivate, ahk_pid %strPID%
Sleep, 500 ; long delay for safety
Send, !{F4}

FileDelete, %strLogFile%
FileAppend, %strLog%, %strLogFile%
Info("Log saved to """ . strLogFile . """. This file will now be opened in Notepad and DELETED from your A_Temp folder.`n`nPlease post the content of this file to the forum thread. Thank you for your help!")
run, Notepad %strLogFile%, , , strPID
WinActivate, ahk_pid %strPID%
sleep, 1000
FileDelete, %strLogFile%
return

; ------------------

ControlEdit1Exist()
{
	ControlGet, strOut, Enabled,, Edit1, A
	return not ErrorLevel
}

Info(str)
{
	MsgBox, 4161, Diag, %str%`n`nPress OK to continue.
	IfMsgBox, Cancel
		ExitApp
}
	
CheckResult(strTestName, strQuestion)
{
	MsgBox, 4131, Diag, %strQuestion%
	IfMsgBox, Cancel
		ExitApp
	IfMsgBox, Yes
		strResult := 1
	IfMsgBox, No
		strResult := 0
	Log(strTestName, strResult)
}

Log(strTestName, strResult)
{
	strLog := strLog . strTestName . " | " . strResult . " | " . strEnv . "`r`n"
}

Explorer_Navigate(FullPath, hwnd="") {  ; by Learning one
    hwnd := (hwnd="") ? WinExist("A") : hwnd ; if omitted, use active window
    WinGet, ProcessName, ProcessName, % "ahk_id " hwnd
    if (ProcessName != "explorer.exe")  ; not Windows explorer
        return
    For pExp in ComObjCreate("Shell.Application").Windows
    {
        if (pExp.hwnd = hwnd) { ; matching window found
            pExp.Navigate("file:///" FullPath)
            return
        }
    }
}
