
/*
- press f3 to copy current pic (either on explorer/desktop or windows' img viewer)
	to %picsfolder%, following pattern in list.txt.
- +enter to load copied img in irfanview, so you can crop it (select with left mouse button,
	crop with ^y, and save with ^s - can disable dialogs for this in irfanview's options)

- missing files start with a :, so can look for them easily

- needs LVS.ahk, a list.txt in script's folder, and optionally irfanview

- list.txt has lines in the form
	%SheetNum% %Cell%; %Name%
	like
		1 A1; pic 1
		2 B14; pic 2
- list.txt references the worksheet index, not its name, to avoid problems with special chars, encoding, valid filenames, etc.

- to be used in conjunction with some excel vba code, that inserts pics in a spreadsheet based on their filenames

- originally for Ahk Unicode 1.1.15.0, untested in other versions
*/


/*		
TODO
- auto set pic orientation based on exif tag, maybe using jhead on dest file. When inserting on excel, it ignores exif orientation.

*/


#noenv

	irfanview := A_ScriptDir "\irfanview.exe"
	picsfolder := A_ScriptDir "\renamed-pics"
	listfile := A_ScriptDir "\list.txt"
	
	if not folderexist(picsfolder)
		if not foldercreate(picsfolder)
			msgboxandexit("could not create folder " picsfolder "`nAborting.")
	
	fileread, listtxt_original, % listfile
	if errorlevel
		msgboxandexit("could not load list file " listfile "`nAborting.")
	
	validclasses =
	( Ltrim
		Photo_Lightweight_Viewer
		ExploreWClass
		CabinetWClass
		Progman
		WorkerW
	)
	loop, parse, validclasses, `n, `r
		groupadd, validwindows, ahk_class %a_loopfield%
	
	LVS_Init("callback", "cell|name", 1, -1, True, False)

return

#ifwinactive, ahk_group validwindows
	f3::
		sourcefile := getSelectedText()
		if (errorlevel or sourcefile = "" or not fileexist(sourcefile))
		{
			msgbox, error on getting valid filepath via Clipboard. Aborting.
			return
		}

		listtxt_markedmissing := ""
		loop, parse, listtxt_original, `n, `r
		{
			if isempty(A_LoopField)
				continue
			
			picid := nthfield(A_LoopField, 1, ";")
			picname := nthfield(A_LoopField, 2, ";")
			
			if not FileExist(picsfolder "\" picid ".*")  ; will match any extension
				listtxt_markedmissing .= picid ";" ":" picname
			else
				listtxt_markedmissing .= picid ";" picname

			listtxt_markedmissing .= "`n"
		}
		
		LVS_SetList(listtxt_markedmissing, ";", False, True)
		LVS_UpdateColOptions()
		LVS_Show()
	return
#ifwinactive


callback(picid, escaped=False) {
	global picsfolder, sourcefile, irfanview
	
	LVS_Hide()
	
	if (escaped || picid == "")
		return

	opendestfileinirfanview := GetKeyState("Shift") ? True : False ; get this info early, file copying can take some time
	
	SplitPath, sourcefile,,, fileextension
	
	destfile := picsfolder "\" picid "." fileextension
	FileCopy, % sourcefile, % destfile, 1  ; TODO: if there's a file with different extension, will not overwrite
	
	if opendestfileinirfanview
		run, "%irfanview%" "%destfile%"
}


folderexist(folder) {
	return instr(fileexist(folder), "D") ? True : False
}

foldercreate(folder) {
	filecreatedir, % folder
	return errorlevel
}

msgboxandexit(txt) {
	msgbox % txt
	exitapp
}

isempty(txt) {
	return RegExMatch(txt, "^\s*$") ? True : False
}

nthfield(txt, n, fieldsdelimiter) {
	stringsplit, fields, txt, % fieldsdelimiter, %A_Space%%A_Tab%

	if (fields0 < n)
		return ""
	
	return fields%n%
}

; returns the selected text while preserving the clipboard.
; sets errorlevel to 1 if cannot retrieve text via ^c
; based on ManaUser's, http://www.autohotkey.com/forum/topic27797.html
; here used to grab filepath from explorer / windows' img viewer
getSelectedText(Keep=0)		; keep: keep retrieved text in clipboard?
{
	savedClip := ClipboardAll
	Clipboard := ""
	Send ^c
	ClipWait 1.5
	if ErrorLevel
	{
		Clipboard := savedClip
		ErrorLevel := 1
		return ""
	}
	
	selectedText := Clipboard
	
	if not Keep
		Clipboard := savedClip
	
	ErrorLevel := 0
	return selectedText
}


#include %A_ScriptDir%\LVS.ahk