; getactiveurl.ahk
#SingleInstance, force
sciahk = C:\Program Files\AutoHotkey\SciTE\SciTE.exe
; AutoHotkey Version: AutoHotkey 1.1
; Language:           English
; Platform:           Win7 SP1 / Win8.1 / Win10
; Author:             Antonio Bueno <user atnbueno of Google's popular e-mail service>
; https://www.autohotkey.com/boards/viewtopic.php?t=3702
; Short description:  Gets the URL of the current (active) browser tab for most modern browsers
; Last Mod:           2016-05-19
;~ #Include %A_scriptdir%\bbb-1-function.ahk
;~ #Include %A_scriptdir%\bbb-1-function.ahk

;~ DO NOT USE FUNCTION Link

;Clipboard =




Menu, Tray, Icon, % A_WinDir "\system32\netshell.dll", 86 ; Shows a world icon in the system tray

ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass,MozillaWindowClass"

return

::getlink::  ;This gets the current browser link
			nTime := A_TickCount
			sURL := GetActiveBrowserURL()
			MsgBox %sURL%
			Clipboard = %sURL%
			MsgBox %Clipboard%
return
::sendlink::
sendinput %sURL%
Return

::wd.::
MsgBox, %a_workingdir%
return




zeropage:
			winget, ActiveID, ID, A
			WinActivate, ahk_exe firefox.exe
			send, 0
			WinActivate, ahk_id %ActiveId%
return

^!k::  ;sends webpage title
			;~ mtclient()

			mtclient(){
			global
			;~ email config for shareable document  ^!l
			SetKeyDelay 111755,111755,111755
			;~ winget, ActiveID, ID, A
			;~ ; MsgBox %activeID% %ID% %A%
			;~ WinActivate, ahk_exe firefox.exe

			;~ send, 0
			;~ winactivate, ahk_id %ActiveID%
			;~ browserid=getactiveurl
			;~ MsgBox %broswerid%

			sURL =
			sURL :=
			title =
			title :=
			scname =
			scbody =
			scname :=
			scbody :=
			nTime := A_TickCount
			sURL := GetActiveBrowserURL()
			if (sURL = "")
				{
					run b:\bat\ahkstart.bat
				}
			Else
			{
			;MsgBox %sURL%
			;gosub zeropage
			

			WinGetTitle, title, A
			WinGetClass, sClass, A
			MsgBox, %title%  %sURL%
			sleep 1000
	    	gosub reglink
			Gui +LastFound +OwnDialogs +AlwaysOnTop
			inputbox, scname, Enter Name of Script
			Gui +LastFound +OwnDialogs +AlwaysOnTop
			InputBox, scbody, Enter Additional Text  ;add additional text I have not found out how to paste multiple paragraphs.
			Gui +LastFound +OwnDialogs +AlwaysOnTop
			;MsgBox,  %scname% %title%` %sURL% %scbody%
			FileAppend,`r`:`:%scname%`:`:`nstline =`n`(`n%title%`n%sURL%`n`r%scbody%`n`)`nclipw`(`)`nreturn`n`:`:%scname%r.`:`:`nrun %sURL%`nreturn`n`n, %A_ScriptDir%\bbb-1-sharable-docs.ahk

			setKeyDelay -1, -1, -1
			sleep 1500
			;reloads updated files
			run, b:\bat\ahkstart.bat

/*

batchfile named ahkstart.bat
echo off
TASKKILL /F /IM "AutoHotkey.exe"
TASKKILL /F /IM "AutoHotkeyU32.exe"

echo off

timeout 3 >  null
IF EXIST b:\%NUL% start /b b:\bbb-1-startup-work.ahk
IF EXIST b:\%NUL% start /b b:\bbb-getactiveur2-works.ahk
IF EXIST b:\%NUL% start /b b:\WindowSnipping.ahk

REM IF EXIST b:\%NUL% start /b b:\HotkeysForPastingImages.ahk
REM IF EXIST b:\%NUL% start /min b:\ahkstudio.ahk
REM start /min TrueImage.exe
copy b:\bbb-1-sharable-docs.ahk  b:\bbb-1-sharable-docs-backup.ahk
CMD /c
echo on
EXIT
*/
			}

}

::mesmes::
MsgBox %sURL%
return

::workdir::
MsgBox, %A_Workingdir%
Return


::scriptdir.::
MsgBox %A_ScriptDir%
return

::surl.::
SendInput, %sURL%
return

^!e::  ;sends webpage title
			;~ mtclient()

			mtrun(){
				global
			SetKeyDelay 11755,11755,11755
			FileDelete, %A_scriptdir%\runll.txt
			sleep 1000
			nTime := A_TickCount
			sURL := GetActiveBrowserURL()

			fileappend,`:`:`:`:`n, %A_scriptdir%\runll.txt
			sleep 1000
			fileappend, run`, %sURL%`n, %A_scriptdir%\runll.txt
			;~ fileappend, run, %sURL%`n, %A_scriptdir%\runll.txt
			sleep 1000
			FileAppend    return`n,%A_scriptdir%\runll.txt
			Sleep 1000
			RunWait, %sciahk% %A_scriptdir%\runll.txt
			FileRead, newpost, %A_WorkingDir%\runll.txt
			;MsgBox, %newpost%
			StringReplace, %newpost%, %newpost%, `n, , All
			;MsgBox, %newpost%
			Run, %sciahk% %A_WorkingDir%\runll.txt
			sleep 100

			setKeyDelay -1, -1, -1
			;~ SendInput, return
			}
return




^#k::  ;sends webpage title and link only
    mtfriend(){
			;~ email config for shareable document
			global
			nTime := A_TickCount
			sURL := GetActiveBrowserURL()
			MsgBox %sURL%
			WinGetTitle, title, A
			WinGetClass, sClass, A
			gosub reglink
			run, mailto:%getemailadd%?subject= addfile %title%&Body=%title% `%0A%sURL%
			}
Return

; a

::wgclass::
WinGetClass, class, A
MsgBox, The active window's class is "%class%".
Return




reglink:  ; Used to clean out or adjust unnecessary or incompatible link text
        ClipWait, 2
		SetKeyDelay 100,100
		Clipboard = %title%
		Clipboard := RegExReplace(clipboard,"stline `= `(([0-9)])`)","stline =")
		Clipboard := RegExReplace(clipboard,"&","and")
		Clipboard := RegExReplace(clipboard,"%","`%")
		Clipboard := RegExReplace(clipboard,"YouTube","")
		Clipboard := RegExReplace(clipboard,"???","")
		Clipboard := RegExReplace(clipboard,"??? Mozilla Firefox","")
		Clipboard := RegExReplace(clipboard,"Google","")
		Clipboard := RegExReplace(clipboard,"-.+YouTube.+???","")
		Clipboard := RegExReplace(clipboard,"Mozilla Firefox","")
		Clipboard := RegExReplace(clipboard,"Chrome","")
		Clipboard := RegExReplace(clipboard,"(0)","")
		Clipboard := RegExReplace(clipboard,"(1)","")
		Clipboard := RegExReplace(clipboard,"(2)","")
		Clipboard := RegExReplace(clipboard,"(3)","")
		Clipboard := RegExReplace(clipboard,"(4)","")
		Clipboard := RegExReplace(clipboard,"(5)","")
		Clipboard := RegExReplace(clipboard,"(6)","")
		Clipboard := RegExReplace(clipboard,"(7)","")
		Clipboard := RegExReplace(clipboard,"(8)","")
		Clipboard := RegExReplace(clipboard,"(9)","")
		Clipboard := RegExReplace(clipboard,"`(`)","")
		Clipboard := RegExReplace(clipboard,"{`(}{)`}","")
		title = %Clipboard%
		;~ Clipboard := RegExReplace(clipboard,"(.+?)","")
		;~ MsgBox % clipboard
		return



GetActiveBrowserURL() {
	global ModernBrowsers, LegacyBrowsers
	WinGetClass, sClass, A
	If sClass In % ModernBrowsers
		Return GetBrowserURL_ACC(sClass)
	Else If sClass In % LegacyBrowsers
		Return GetBrowserURL_DDE(sClass) ; empty string if DDE not supported (or not a browser)
	Else
		Return ""
}

; "GetBrowserURL_DDE" adapted from DDE code by Sean, (AHK_L version by maraskan_user)
; Found at http://autohotkey.com/board/topic/17633-/?p=434518

GetBrowserURL_DDE(sClass) {
	WinGet, sServer, ProcessName, % "ahk_class " sClass
	StringTrimRight, sServer, sServer, 4
	iCodePage := A_IsUnicode ? 0x04B0 : 0x03EC ; 0x04B0 = CP_WINUNICODE, 0x03EC = CP_WINANSI
	DllCall("DdeInitialize", "UPtrP", idInst, "Uint", 0, "Uint", 0, "Uint", 0)
	hServer := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", sServer, "int", iCodePage)
	hTopic := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "WWW_GetWindowInfo", "int", iCodePage)
	hItem := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "0xFFFFFFFF", "int", iCodePage)
	hConv := DllCall("DdeConnect", "UPtr", idInst, "UPtr", hServer, "UPtr", hTopic, "Uint", 0)
	hData := DllCall("DdeClientTransaction", "Uint", 0, "Uint", 0, "UPtr", hConv, "UPtr", hItem, "UInt", 1, "Uint", 0x20B0, "Uint", 10000, "UPtrP", nResult) ; 0x20B0 = XTYP_REQUEST, 10000 = 10s timeout
	sData := DllCall("DdeAccessData", "Uint", hData, "Uint", 0, "Str")
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hServer)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hTopic)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hItem)
	DllCall("DdeUnaccessData", "UPtr", hData)
	DllCall("DdeFreeDataHandle", "UPtr", hData)
	DllCall("DdeDisconnect", "UPtr", hConv)
	DllCall("DdeUninitialize", "UPtr", idInst)
	csvWindowInfo := StrGet(&sData, "CP0")
	StringSplit, sWindowInfo, csvWindowInfo, `" ;"; comment to avoid a syntax highlighting issue in autohotkey.com/boards
	Return sWindowInfo2
}

GetBrowserURL_ACC(sClass) {
	global nWindow, accAddressBar
	If (nWindow != WinExist("ahk_class " sClass)) ; reuses accAddressBar if it's the same window
	{
		nWindow := WinExist("ahk_class " sClass)
		accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindow))
	}
	Try sURL := accAddressBar.accValue(0)
	If (sURL == "") {
		WinGet, nWindows, List, % "ahk_class " sClass ; In case of a nested browser window as in the old CoolNovo (TO DO: check if still needed)
		If (nWindows > 1) {
			accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindows2))
			Try sURL := accAddressBar.accValue(0)
		}
	}
	If ((sURL != "") and (SubStr(sURL, 1, 4) != "http")) ; Modern browsers omit "http://"
		sURL := "http://" sURL
	If (sURL == "")
		nWindow := -1 ; Don't remember the window if there is no URL
	Return sURL
}

; "GetAddressBar" based in code by uname
; Found at http://autohotkey.com/board/topic/103178-/?p=637687

GetAddressBar(accObj) {
	Try If ((accObj.accRole(0) == 42) and IsURL(accObj.accValue(0)))
		Return accObj
	Try If ((accObj.accRole(0) == 42) and IsURL("http://" accObj.accValue(0))) ; Modern browsers omit "http://"
		Return accObj
	For nChild, accChild in Acc_Children(accObj)
		If IsObject(accAddressBar := GetAddressBar(accChild))
			Return accAddressBar
}

IsURL(sURL) {
	Return RegExMatch(sURL, "^(?<Protocol>https?|ftp)://(?<Domain>(?:[\w-]+\.)+\w\w+)(?::(?<Port>\d+))?/?(?<Path>(?:[^:/?# ]*/?)+)(?:\?(?<Query>[^#]+)?)?(?:\#(?<Hash>.+)?)?$")
}

; The code below is part of the Acc.ahk Standard Library by Sean (updated by jethrow)
; Found at http://autohotkey.com/board/topic/77303-/?p=491516

Acc_Init()
{
	static h
	If Not h
		h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}
Acc_ObjectFromWindow(hWnd, idObject = 0)
{
	Acc_Init()
	If DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
	Return ComObjEnwrap(9,pacc,1)
}
Acc_Query(Acc) {
	Try Return ComObj(9, ComObjQuery(Acc,"{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
}
Acc_Children(Acc) {
	If ComObjType(Acc,"Name") != "IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	Else {
		Acc_Init(), cChildren:=Acc.accChildCount, Children:=[]
		If DllCall("oleacc\AccessibleChildren", "Ptr",ComObjValue(Acc), "Int",0, "Int",cChildren, "Ptr",VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*",cChildren)=0 {
			Loop %cChildren%
				i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i), Children.Insert(NumGet(varChildren,i-8)=9?Acc_Query(child):child), NumGet(varChildren,i-8)=9?ObjRelease(child):
			Return Children.MaxIndex()?Children:
		} Else
			ErrorLevel := "AccessibleChildren DllCall Failed"
	}
}


;~ run, mailto:%getemailadd%?subject= %title%&Body= `:`:%scriptname%`:`: `%0Astline `=`%0A `(`%0A%title% `%0A%sURL%`%0A`)`%0Aclipw`(`)`%0Areturn`%0A`;`~%title% %sURL%