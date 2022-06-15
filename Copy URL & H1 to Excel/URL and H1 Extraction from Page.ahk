#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ;Allows windowing functions and commands to operate on windows whose titles contain WinTitle anywhere instead of just at the beginning.

global Wait := 350 ;Define a real number and save it in a variable called "Wait" - this number will be used later to tell AHK how long (in milliseconds) to pause between steps.
global BrowserWindow := "" ;Define what AHK should search for in the Title of the Browser window.
global ExcelFile := "" ;Define what AHK should search for in the Title of the Excel file.

Esc:: ;A loop breaker that stops the running script and reloads it in order to break loops so they don't run wild.
Reload, %A_AHKPath% "%A_ScriptDir%\URL and H1 Extraction from Page.ahk"
ExitApp
return

;Before triggering the script, ensure that all the pages that need to be scraped are open in various tabs, the starting tab is active, and the cell from where you want to start filling is selected.
^Numpad1::
	MouseGetPos, xpos, ypos, window ;Get the coordinates of the cursor and which window it's in so that AHK can put the cursor back there once the script is done. This is aboslutely not essential but it annoys me when my cursor is not where I left it, so I programmed this in.
Loop, 2 ;Enter how many times you want the script to loop
{
	Sleep, %Wait% ;Adding a small delay to compensate for any kind of PC lag is good practice. Always use enough pauses to prevent your macro doing something nonsensical by mistake
		;This loop activates the browser, copies the URL to Excel and goes one cell to the right, then goes back to the brower, copies the H1 to Excel as text, and goes to the next row, and finally, goes back to the browser window and goes to the next tab.
	WinActivate, %BrowserWindow%
	AddressBar()
	Copy()
	WinActivate, %ExcelFile%
	Sleep, %Wait%
	Paste()
	Send, {Tab}
	Sleep, %Wait%
	WinActivate, %BrowserWindow%
	Sleep, %Wait%
	Click, 80, 420, 3 ;Replace these coordinates with the correct ones for your setup. Use WindoSpy.ahk to find the coordinates easily.
	Copy()
	WinActivate, %ExcelFile%
	Sleep, %Wait%
	PasteSpecialText()
	Sleep, %Wait%
	Send, {Down}
	Sleep, %Wait%
	Send, {Home} 
	Sleep, %Wait%
	WinActivate, %BrowserWindow%
	SwitchTab()
}
MouseMove %xpos%, %ypos%, %window% ;Move the cursor to where it originally was before the loop started.
return

;Essential functions that have been used in the script above are defined below. 

Copy() { ;Copy & Pause
	Send, {CtrlDown}c{CtrlUp}
	Sleep, %Wait%
}

Paste() { ;Paste & Pause
	Send, {CtrlDown}v{CtrlUp}
	Sleep, %Wait%
}

PasteSpecialText() { ;Paste & Pause
	Sleep, %Wait%
	Send, {CtrlDown}{AltDown}v{AltUp}{CtrlUp}
	Sleep, %Wait%
	Sleep, %Wait%
	Send, t
	Sleep, %Wait%
	Send, {Enter}
}

AddressBar() { ;Activate Address bar & pause
Send, {CtrlDown}l{CtrlUp}
Sleep, %Wait%
}

SwitchTab() { ;Go to next browser tab & pause
Send, {CtrlDown}{Tab}{CtrlUp}
Sleep, %Wait%
}
