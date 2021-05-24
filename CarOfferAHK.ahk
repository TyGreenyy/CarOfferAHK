; =====================================================================================
; AHK Version ...: AHK 1.1.32.00 (Unicode 64-bit) - December 22, 2020
; Platform ......: Windows 10
; Language ......: English (en-US)
; Author ........: TyGreeny
; =====================================================================================

reloadAsAdmin_Task()         ;Runs reloadAsAdmin task
Suspend, On                     ;Suspend Script for Update

directoryMove(){
    if !(A_ScriptDir = A_MyDocuments){
        FileCopy, %A_ScriptDir%\CarOfferAHK.ah k, %A_MyDocuments%\CarOfferAHK.ahk
        Sleep, 3000
        Run, %A_MyDocuments%\CarOfferAHK.ahk  /restart
        Sleep, 200
        ExitApp
    } else {
        FileDelete, C:\Users\%A_UserName%\Downloads\CarOfferAHK.ahk
        }
    return
}

updateScript() {                     ;Create Directory Structure - Update script from Github

  if !FileExist("%A_MyDocuments%\CarOfferAHK\")
      FileCreateDir, %A_MyDocuments%\CarOfferAHK\              ;Create Directory Structure 
      FileCreateDir, %A_MyDocuments%\CarOfferAHK\resources\             ;Create Directory Structure

 UrlDownloadToFile, https://raw.githubusercontent.com/TyGreenyy/CarOfferAHK/main/resources/DealerCodes.ini, %A_MyDocuments%\CarOfferAHK\resources\DealerCodes.ini
  UrlDownloadToFile, https://raw.githubusercontent.com/TyGreenyy/CarOfferAHK/main/resources/CarOfferAHK_Rounded.ico, %A_MyDocuments%\CarOfferAHK\resources\CarOfferAHK_Rounded.ico
  UrlDownloadToFile, https://github.com/TyGreenyy/CarOfferAHK/raw/main/resources/imageres.dll, %A_MyDocuments%\CarOfferAHK\resources\imageres.dll

  whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")               ;Create HTTP request to check version
  whr.Open("GET", "https://raw.githubusercontent.com/TyGreenyy/CarOfferAHK/main/version.md", true)
  whr.Send()
  ; Using 'true' above and the call below allows the script to remain responsive.
  whr.WaitForResponse()
  global version := whr.ResponseText

  RegExMatch(trim(version), "[0-9]" , version)                  ;Checks version against Github version
  if (version = 8){                                             ;Downloads new .ahk if version does not match
      } else {
         UrlDownloadToFile, https://raw.githubusercontent.com/TyGreenyy/CarOfferAHK/main/CarOfferAHK.ahk, %A_MyDocuments%\CarOfferAHK.ahk
        }
}

;{ ============================== Copy From ===============================================
;} ====================================================================================

;{ ============================== Notes ===============================================
; All directives(that are settings) and many commands appear here irvrespective of its need
; The directives/commands that are default are commented by ";;" ,
; those not needed by ";;;" and those not desired by ";~ "
;} ====================================================================================

; directoryMove()
; updateScript()
caseMenu.__new()                ;creates the caseMenu
trayMenu()                      ;creates th etraymenu

#SingleInstance Force
#include %A_ScriptDir%

SetWorkingDir %A_ScriptDir%     ;Set script path as working directory.

global SCR_Name                 ;For tray menu options | SCR_Name, no ext
SplitPath, A_ScriptFullPath, , , , SCR_Name,

global SCR_UserDir              ;~ SCR_UserDir = C:\Users\%USER%
SplitPath, A_MyDocuments, , SCR_UserDir, , ,s

global SCR_Path = A_ScriptDir   ;Sets path for tray menu options

#NoEnv                           ;Donot use default environmental variables.
#Persistent                     ;Keeps a script permanently running
#SingleInstance Force           ;Replaces old instance with new one
#UseHook                        ;For machine level hotkeys. Removes need to use $ in hotkeys
#Hotstring * ? B C K0 Z         ;Hotstring global settings
#InstallKeybdHook               ;Installs Keyboard hook
#InstallMouseHook               ;Installs Mouse hook
#KeyHistory 250                 ;maximum number of keyboard and mouse events displayed
#MaxHotkeysPerInterval 200      ;Rate a warning dialog will be displayed
#MaxMem 256                     ;Allows variables with heavy memory usage.
#MaxThreads 255                 ;Allows # of pseudo-threads to run simultaneously.
#MaxThreadsPerHotkey 1          ;Hotkey can not be launched when it is already running
Thread, interrupt, 50           ;Minimum time for interupt
SendMode Input                  ;How the script sends simulated keys.
Process, Priority, , R          ;Runs Script at High process priority for best performance
SetBatchLines, -1               ;Never sleep for best performance
SetKeyDelay, 0, 0               ;Smallest ossible delay
SetMouseDelay, 0                ;Smallest possible delay
SetDefaultMouseSpeed, 0         ;Move the mouse instantly
SetWinDelay, 0                  ;Smallest possible delay
SetControlDelay, 0              ;Smallest possible delay
SetTitleMatchMode, 2            ;A window's title can contain WinTitle anywhere inside it
BlockInput, Mouse               ;Keyboard/mouse is blocked during Click or MouseClick
CoordMode, ToolTip, Screen      ;Tooltip co-ords are specified in global co-ords
CoordMode, Pixel, Screen        ;Co-ords for image/pixelSearch are specified in global co-ords
CoordMode, Mouse, Screen        ;Mouse co-ords are specified in global co-ords
CoordMode, Caret, Screen        ;A_CaretX/Y are specified in global co-ords
CoordMode, Menu, Screen         ;Co-ords for "Menu, Show" are specified in global co-ords
SetNumLockState, On

trayMenu(){
	ifexist %A_MyDocuments%\CarOfferAHK\resources\CarOfferAHK_Rounded.ico
		Menu, Tray, Icon, %A_MyDocuments%\CarOfferAHK\resources\imageres.dll, 2
}

Suspend, Off

;{==================================ToggleKeys=========================================
;} ====================================================================================

; CapsLockOffTimer(t:=60000){
;  if ((A_TimeIdleKeyboard>t) AND GetKeyState("CapsLock","T")){
;     SetCapsLockState,Off
;     ; SetTimer, CapsLockOffTimer, Off
;     return True
;  }
;  return False
;}

return
DateAction:
	DatePaste := StrReplace(A_ThisMenuItem, "Paste: ", "")
	SendInput %DatePaste%{Raw}%A_EndChar%
Return

class caseMenu {
	__new(){

		for _, j in [["truePeopleSearch","Search &True People","14"],["phoneFormatPara","Format P&hone: (###) ###-####","12"]]

		{
			act:=ObjBindMethod(this,"textFormat",j[1])
			Menu, caseMenu, Add, % j[2], % act
			Menu, caseMenu, Icon, % j[2], %A_MyDocuments%\CarOfferAHK\resources\imageres.dll , % j[3]
		}

		FormatTime, OutputVar, %A_Now%, MMM dd, yyyy 'at' h:mm tt CST
		for _, j in [[OutputVar]]

		{
			Menu, caseMenu, add, Paste: %OutputVar%, DateAction
			Menu, caseMenu, Icon, Paste: %OutputVar%, %A_MyDocuments%\CarOfferAHK\resources\imageres.dll , 5
		}

		;Seperator One
		Menu, caseMenu, Add

		for _, j in [["VINAnalysis","Search &VIN","2"],["jiraFunc","Search &JIRA","6"],["carGurus","Search Car&Gurus","1"],["searchHubspot","Search &Hubspot","5"],["handsellCopy","Copy Handsell Unit","16"],["vehicleinfopaste","&Year Make Model - VIN","17"],["yearmakemodelformat2","Year Make Model - V&IN - $","17"]]

		{
			act:=ObjBindMethod(this,"textFormat",j[1])
			Menu, caseMenu, Add, % j[2], % act
			Menu, caseMenu, Icon, % j[2], %A_MyDocuments%\CarOfferAHK\resources\imageres.dll , % j[3]
		}

		;Seperator Two
		Menu, caseMenu, Add

		for _, j in [["searchVXDealer","Search V&XDealer","25"],["dealerstats","Search DealerStats","3"],["dealerCDS","Search Dealer &CDS","17"],["groupWholeSale","Search Wholesale Units","17"],["auctionCaps","Search Auction Caps","22"],["dealerExclu","Search Dealer Exclusions","23"],["matrixOverview","Search Matrix Overview","19"],["dealerAccepts","Search Dealer &Accepts","2"],["dealerOG","Search Dealer &OfferGuards","10"],["dealerPuts","Search Dealer P&uts","9"]]

		{
			act:=ObjBindMethod(this,"textFormat",j[1])
			Menu, caseMenu, Add, % j[2], % act
			Menu, caseMenu, Icon, % j[2], %A_MyDocuments%\CarOfferAHK\resources\imageres.dll , % j[3]
		}

		;Seperator Four
		Menu, caseMenu, Add

		for _, j in [["phoneFormat","###-###-####"],["phoneFormatPara","(###) ###-####"],["googlefileDownload","Google Drive DL Link"],["fileRename","File &Rename"],["cleanRename","Clean Rename"],["urlencode","URL Encode"]]

		{
		act:=ObjBindMethod(this,"textFormat",j[1])
		Menu, Submenu1, Add, % j[2], % act
		Menu, caseMenu, Add, &Format, :Submenu1
		}

		for _, j in [["U","&UPPER CASE"],["L","&lower case"],["T","&Title Case"],["S","&Sentence case."]]

		{
			act:=ObjBindMethod(this,"caseChange",j[1])
			Menu, Submenu2, Add, % j[2], % act
			Menu, caseMenu, Add, &Change Case, :Submenu2

		}

		for _, j in [["TGL_ExpWinGetSel","Copy Full FilePath(s)"],["TGL_ExpWinGetDir","Get File Path"],["TGL_ExpWinGetSel_Private","Copy Filename + Ext"]]

		{
			act:=ObjBindMethod(this,"textFormat",j[1])
			Menu, Submenu6, Add, % j[2], % act
			Menu, caseMenu, Add, &Windows Explorer, :Submenu6

		}

		for _, j in []

		{
			act:=ObjBindMethod(this,"textFormat",j[1])
			Menu, caseMenu, Add, % j[2], % act
		}

		Menu, Submenu1, Add

		for _, j in [["addQuotes","Add Quotes"],["addPercent","Variables"],["removeformat","Remove Format"]]

	{
		act:=ObjBindMethod(this,"textFormat",j[1])
		Menu, Submenu1, Add, % j[2], % act
	}

			/*
			for _, i in ["&Capslock","&Numlock","Sc&rollLock","I&nsert"] {
			act:=ObjBindMethod(this,"toggle",strReplace(i,"&"))
			Menu, caseMenu, Add, % i, % act
			}
			*/
	return
	}

	show() {
/*        for _, i in ["&Numlock","Sc&rollLock","I&nsert"]
		;~ sleep, 500*
		for _, i in ["&Numlock","Sc&rollLock","I&nsert"]
			Menu, caseMenu, % GetKeyState(strReplace(i,"&"),"T")?"Check":"Uncheck", % i
*/
		Menu, caseMenu, Show, % A_CaretX, % (A_Carety+25)
		return
		}

		caseChange(type){ ; type: U=UPPER, L=Lower, T=Title, S=Sentence, I=Invert
			text:=caseChange(getSelectedText(), type)
				oldClip:=ClipboardAll
				clipboard:=text
				Send ^v
				sleep 100
				Clipboard:=oldClip
			return
			}

		toggle(key){
			if (key=Insert)
				Send, {Insert}
			else if (key=Capslock)
				SetCapsLockState, % GetKeyState("CapsLock","T")?"Off":"On"
			else if (key=Numlock)
				SetNumLockState, % GetKeyState("NumLock","T")?"Off":"On"
			else if (key=Scrolllock)
				SetScrollLockState, % GetKeyState("ScrollLock","T")?"Off":"On"
			return
			}

		textFormat(type){
				%type%()
			return
			}
		}

DateAction(type) {
		SendInput %A_ThisMenuItem%{Raw}%A_EndChar%
	Return
}

Togglekeys_check(){
	return {c:getkeyState("Capslock","T"), n:getkeyState("Numlock","T"), s:getkeyState("ScrollLock","T"), i:getkeyState("Insert","T")}
}

strRemove(parent,strlist) {
	for _,str in strlist
		parent:=strReplace(parent,str)
	return parent
}

caseChange(text,type){ ; type: U=UPPER, L=Lower, T=Title, S=Sentence, I=Invert
	static X:= ["iOS","iPhone","iPad","I","AHK","AutoHotkey","Dr","Mr","Ms","Mrs","BK","DEA","AKJ","via","v","vs","GPG","PGP","to","of","and","on","with","USB" ]
	;~ list of words that should not be modified for S,T
	if (type="S") { ;Sentence case.
		text := RegExReplace(RegExReplace(text, "(.*)", "$L{1}"), "(?<=[^a-zA-Z0-9_-]\s|\n).|^.", "$U{0}")
	} else if (type="I") ;iNVERSE
		text:=RegExReplace(text, "([A-Z])|([a-z])", "$L1$U2")
	else text:=RegExReplace(text, "(.*)", "$" type "{1}")

	if (type="S" OR type="T")
		for _, word in X ;Parse the exceptions
			text:= RegExReplace(text,"i)\b" word "\b", word)
	return text
}

TGL_ExpWinGetSel(){
	WinGet, hwndtmp, ID, A
	WinActivate, %hwndtmp%,
	Clipboard := JEE_ExpWinGetSel(hwndtmp)
}

TGL_ExpWinGetSel_Private(){
	WinGet, hwndtmp, ID, A
	WinActivate, %hwndtmp%,
	linestoParse := JEE_ExpWinGetSel(hwndtmp)
	Loop, parse, linestoParse, `n, `r  ; Specifying `n prior to `r allows both Windows and Unix files to be parsed.
	{
		SplitPath, A_LoopField, newFilename
		if (A_Index = 1)
			Clipboard := newfilename
		else
			Clipboard := Clipboard . "`n" newfilename
		; MsgBox, 4, , Line number %A_Index% is %A_LoopField%.`n`nContinue?
		; IfMsgBox, No, break
	}
}

TGL_ExpWinGetDir(){
	Clipboard := JEE_ExpWinGetDir()
}

cutFunc(){
	Send, ^x
}
copyFunc(){
	fileExplorerName := getSelectedText()
	Clipboard := fileExplorerName
}
pasteFunc(){
	Clip(ClipBoard)
}
deleteFunc(){
	Send, {Del}
}

addQuotes(){
	oldClip := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
		oldClip := oldClip2
	oldClip = "%oldClip%"
	Clipboard := oldClip
	Clip(oldClip)
}

addPercent(){
	oldClip := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
			oldClip := oldClip2
	oldClip = %oldClip%
	oldClip = `%%oldClip%`%
	oldClip := oldClip " "
	Clipboard := oldClip
	Clip(oldClip)
}

prefaceTime(){
	oldClip := Clip()
	FormatTime, TimeString, R, MMM dd',' yyyy 'at' h:mm tt 'CST'
	oldClip = %oldClip%
	oldClip := TimeString ": " oldClip
	Clip(oldClip)
}

removeformat(){
	oldClip := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
		oldClip := oldClip2
	oldClip := Clip()
	oldClip := RegExReplace(ClipBoard , "\R\R\K\R+")
	Clipboard := oldClip
	Clip(oldClip)
}

searchTermClean(){
	searchTerm := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
		searchTerm := oldClip2
	searchTerm := LTrim(searchTerm)
	searchTerm := StrReplace(StrReplace(searchTerm, "`n", ""), "`r", "")
	searchTerm := RegExReplace(RegExReplace(searchTerm, "[\']", ""), "\W", " ")
	searchTerm := StrSplit(searchTerm, A_Space)
	searchTerm := searchTerm[1] " " searchTerm[2] " " searchTerm[3] " " searchTerm[4]
	searchTerm := Trim(searchTerm)
	return searchTerm
}

searchHubspot(){
	searchTerm := searchTermClean()
	hubspotID := getDealerID(Trim(searchTerm))
	hubspotID := hubspotID[3]
	if !(hubspotID = "") {
		openLink := "https://app.hubspot.com/contacts/5712725/company/" . hubspotID
		toast("Opening Dealer Record in Hubspot", openLink, ,2000)
	} else {
		openLink := "https://app.hubspot.com/reports-dashboard/5712725/view/4177402?globalSearchQuery=" . searchTerm
			toast("Opening Hubspot and Searching:", searchTerm, ,2000)
	}
	Shellrun(openLink)
}

groupWholeSale(){
	searchTerm := searchTermClean()
	GroupID := getDealerID(Trim(searchTerm))
	dealershipID := GroupID[1]
	GroupID := GroupID[5]
	if !(GroupID = "") {
		openLink := "http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.DealerInventoryOffers&p_groupid="  . GroupID . "&p_reportType=&p_wholesale=W&LinkHref=True"
			toast("Opening Dealer Group Wholesale", openLink, ,5000)
			Shellrun(openLink)
	} else if !(dealershipID = "") {
		openLink := "http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.DealerInventoryOffers&p_dealershipid="  . dealershipID . "&p_reportType=&p_wholesale=W&LinkHref=True"
			toast("Opening Dealer Wholesale", openLink, ,2000)
			Shellrun(openLink)
	} else {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)

	}

}

searchVXDealer(){
	searchTerm := searchTermClean()
	searchTerm := StrReplace(searchTerm, " ", "%20")
	if (searchTerm = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
	}
	openlink := "https://admin.pearlsolutions.com/Home/Portal#/dealerships?search=" . searchTerm
	toast("Opening Dealer Search in VX Admin", openLink, ,2000)
	ShellRun(openlink)
}

getDealerID(searchTerm){
   FileRead, FileText, %A_MyDocuments%\CarOfferAHK\resources\DealerCodes.ini
		Data := {}
		SearchLine := "i)" . searchTerm
		Loop, Parse, FileText, `n , `r
		{
		if (RegExMatch(A_Loopfield, SearchLine)){
			linetext := StrSplit(A_Loopfield, "&&", "&&")
				return linetext
			}
		}
 }

getGroupID(searchTerm){
   FileRead, FileText, %A_MyDocuments%\CarOfferAHK\resources\DealerCodes.ini
	Data := {}
		SearchLine := "i)" . searchTerm
		Loop, Parse, FileText, `n , `r
		{
		if (RegExMatch(A_Loopfield, SearchLine)){
			linetext := StrSplit(A_Loopfield, A_tab)
				return linetext[9]
			}
		}
 }

dealerstats(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.DealerStatsSummary&p_dealershipIDs=%dealershipID%
	toast("Opening DealerStatsSummary", openLink, ,2000)
	;~ openlink := URI_URLEncode(openLink)
	ShellRun(openLink)
	return
}

auctionCaps(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.Auction.MMRCapsNew&p_dealershipID=%dealershipID%
	toast("Opening Dealer Auction Caps", openLink, ,2000)
	;~ openlink := URI_URLEncode(openLink)
	ShellRun(openLink)
	return
}

;~ http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.DealerExclusions&rdAjaxCommand=RefreshElement&pDealershipID=

dealerExclu(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.DealerExclusions&pDealershipID=%dealershipID%
	toast("Opening Dealer Exclusions", openLink, ,2000)
	;~ openlink := URI_URLEncode(openLink)
	ShellRun(openLink)
	return
}

dealerCDS(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.carOfferInvPerf&islDealer=%dealershipID%&LinkHref=True=
	openlink := URI_URLEncode(openLink)
	toast("Opening Dealer CDS Report", openLink, ,2000)
	ShellRun(openLink)
	return
}

;=http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.carOfferInvPerf&islDealer=9614&LinkHref=True

dealerAccepts(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.transactionAcceptDetails&p_dealershipID=%dealershipID%&p_offerAccept=1&p_ReportLevel=DETAIL&rdAgReset=True&LinkHref=True&rdRequestForwarding=Form
	openlink := URI_URLEncode(openLink)
	toast("Opening Dealer Accepted Offers", openLink, ,2000)
	ShellRun(openLink)
	return
}

dealerOG(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.transactionAcceptDetails&p_dealershipID=%dealershipID%&p_offerAccept=&p_og=1&p_ReportLevel=DETAIL&rdAgReset=True&LinkHref=True&rdRequestForwarding=Form
	openlink := URI_URLEncode(openLink)
	toast("Opening Dealer OfferGuards", openLink, ,2000)
	ShellRun(openLink)
	return
}

dealerPuts(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.transactionAcceptDetails&p_dealershipID=%dealershipID%&p_offerAccept=&p_putAccept=1&p_ReportLevel=DETAIL&rdAgReset=True&LinkHref=True&rdRequestForwarding=Form
	openlink := URI_URLEncode(openLink)
	toast("Opening Dealer Put Bids", openLink, ,2000)
	ShellRun(openLink)
	return
}

VINAnalysis(){
	searchTerm := Trim(searchTermClean())
	if !RegExMatch(searchTerm, "[A-Za-z0-9_]") {
		toast("No Match Found", "`nHighlight a Full Vin Number", ,5000)
		return
	}
	if !(StrLen(searchTerm) = 17){
		; Check VIN
		toast("No Match Found", "`nHighlight a Full Vin Number", ,5000)
		return
	}
	openLink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.MatrixAnalysis&p_action=R&p_requestID=&VIN=%searchTerm%&rdDisableRememberExpansion_menuTreeSide=True&p_vin=%searchTerm%&rdShowElementHistory=
	toast("Searching VIN Analysis", openLink, ,2000)
	shellrun(openLink)
	Clipboard := searchTerm
	return
}

matrixOverview(){
	searchTerm := searchTermClean()
	Sleep, 200
	Clipboard := searchTerm
	dealershipID := getDealerID(searchTerm)
	dealershipID := dealershipID[1]
	if (dealershipID  = "") {
		toast("No Match Found", "`nHighlight a Dealer Name", ,5000)
		return
	}
	openlink = http://ops.pearlsolutions.com/rdPage.aspx?rdReport=Caroffer.MatrixOverview&p_dealershipID=%dealershipID%&rdAgReset=True&LinkHref=True&rdRequestForwarding=Form
	openlink := URI_URLEncode(openLink)
	toast("Opening Dealer Matrix Overview", openLink, ,2000)
	ShellRun(openLink)
	return
}

jiraFunc(){
	searchTerm := searchTermClean()
	openLink = https://caroffer.atlassian.net/secure/QuickSearch.jspa?searchString=%searchTerm%
	shellrun(openLink)
	toast("Searching JIRA", openLink, ,2000)
	clipboardLink = %searchTerm%
	Clipboard := clipboardLink
}

carGurus(){
	searchTerm := Trim(searchTermClean())
	if !RegExMatch(searchTerm, "[A-Za-z0-9_]") {
		toast("No Match Found", "`nHighlight a Full Vin Number", ,5000)
		return
	}
	if !(StrLen(searchTerm) = 17){
		toast("No Match Found", "`nHighlight a Full Vin Number", ,5000)
		return
	}
	carGurus := "https://www.cargurus.com/Cars/instantMarketValueFromVIN.action?startUrl=%2F&++++carDescription.vin%0D%0A=" . searchTerm
	ShellRun(carGurus)
	toast("Searching CarGurus Instant Market Value", carGurus, ,2000)
	Clipboard := carGurus
}

removeTabs(){
	oldClip := searchTermClean()
	oldClip := StrReplace(oldClip, "`n",  " ")
	oldClip := StrReplace(oldClip, "`r",  " ")
	oldClip := StrReplace(oldClip, A_tab,  " ")
	oldClip := StrReplace(oldClip, A_Space,  " ")
	Clipboard := oldClip
	Clip(oldClip)
	return
}

yearmakemodelformat(){
	oldClip := Clipboard
		oldClip2 := Clip()
		if (oldClip2 != "")
		  oldClip := oldClip2
	tempvar := RegExMatch(oldClip, "([A-HJ-NPR-Z\d]{8}[\dX][A-HJ-NPR-Z\d]{8})", VinNumbervar)
	tempvar := RegExMatch(oldClip, "((19|20)\d\d{1})([\w\s].*)(\n|\s)", YearMakeModelVin2)
	YearMakeModelVin2 := StrReplace(YearMakeModelVin2, "`n",  " ")
	YearMakeModelVin2 := StrReplace(YearMakeModelVin2, "`r",  "")
	YearMakeModelVin2 := StrReplace(YearMakeModelVin2, A_tab,  " ")
	YearMakeModel := caseChange(Yeavar,"T")
	YearMakeModelVin2 := caseChange(YearMakeModelVin2,"T")
	YearMakeModelVin := YearMakeModelVin2 " - " VinNumbervar
	Clipboard := YearMakeModelVin
	return
}

yearmakemodelformat2(){
	oldClip := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
	  oldClip := oldClip2
	  StringSplit, searchTermArray, oldClip, `n,
	  Loop, %searchTermArray0%
	   {
		searchTermTemp := searchTermArray%A_Index%
		if (RegExMatch(searchTermTemp, "((19|20)\d\d{1}\s)([\w\s].*)(\n|\s)", YearMakeModelVin2)){
			searchTerm2%A_Index% := searchTermTemp
			searchTerm2 := StrReplace(StrReplace(StrReplace(caseChange(searchTerm2%A_Index%,"T"), A_Tab,  " "), "`r",  ""), "`n",  " ")
		}
		if (RegExMatch(searchTermTemp, "([A-HJ-NPR-Z\d]{8}[\dX][A-HJ-NPR-Z\d]{8})", VinNumbervar)){
			searchTerm%A_Index% := searchTermTemp
			searchTerm := StrReplace(searchTerm%A_Index%, A_Tab,  " ")
			searchTerm := StrReplace(StrReplace(searchTerm, "`r",  ""), "`n",  " ")
			; ArrayIndex += 1
			searchTermFinal%A_Index%  := searchTerm2 " - " searchTerm
			ArrayIndexLess := (%A_Index% - 1)
			searchTermFinal := searchTermFinal%ArrayIndexLess% "`n" searchTermFinal%A_Index%
		}
		if (RegExMatch(searchTermTemp, "[\r\n]*(Sell Today)[\r\n]", offerVar)){
			nextIndex := (A_Index + 1)
			offerVar := searchTermArray%nextIndex%
			searchTermFinal := searchTermFinal " - Sell Today: " offerVar
		}
	   }
	if (searchTermFinal){
	} else {
		toast("Tip:", "`nHighlight Right to Left`nfrom Sell Today to the Year Make Model ", ,5000)
	}
	Clipboard := searchTermFinal
	toast("Success! Copied to Clipboard:", searchTermFinal, ,3000)
	return
}

vehicleinfopaste(){
	searchTerm := Clip()
	Loop, parse, searchTerm, `n, `r
	{
	  if A_Index = 1
	  {
		Clipboard_One := A_LoopField
	  }
		Clipboard_Last := A_LoopField
	}
	Clipboard_One := caseChange(Clipboard_One, "T")
	searchTerm := Clipboard_One " - " Clipboard_Last
	Clipboard := searchTerm
		return
}

handsellCopy(){
	Send, ^c
	Sleep, 50
	test := Clipboard
	Vehicle := StrSplit(test, "`n")
	Vehicle := Vehicle[1]
	Vehicle := StrReplace(Vehicle, "`n",  "")
	Vehicle := StrReplace(Vehicle, "`r",  "")
	RegExMatch(test, "VIN:(.*)", VIN2)
	VIN := StrSplit(VIN2, ":")
	VIN := VIN[2]
	RegExMatch(test, "(\$.*)", Price)
		RegExMatch(Price, "\d+\,\d+", Price)
		Price := StrReplace(StrReplace(Price, ",", ""), "$", "")
		Price := (Price + 1500)
		Price := "$" SubStr(Price, -(StrLen(Price)), (StrLen(Price)-3)) "," SubStr(Price, -2, 3)
	RegExMatch(test, "Miles:(.*)", Miles)
		if (VIN2 = "") OR (Price = "") OR (Miles = "") {
			toast("Tip:", "`nHighlight Right to Left`nfrom Recall Count to the Year Make Model ", ,5000)
			Return
		}
	html := WinClip.GetHTML()
	regexformula := "href\s*=\s*['""]?([^""']+)"
	regex := RegExMatch(html, regexformula, hreflink)
	hreflink := StrSplit(hreflink, "=""")
	texttopaste := Vehicle " - " VIN " - " Miles " - " Price "`n" hreflink[2]
	toast("Success! Copied to Clipboard:", texttopaste, ,3000)
	Clipboard := texttopaste
	return
}

googlefileDownload(){
	searchTerm := Clip()
	foundPos := InStr(searchTerm, "file/d/")
	searchTerm := SubStr(searchTerm, foundPos)
	searchTerm := StrSplit(searchTerm, "/")
	searchTerm := searchTerm[3]
	googlefileDownload := "https://drive.google.com/uc?export=download&id=" . searchTerm
	Clipboard := googlefileDownload
		return
}

truePeopleSearch(){
	phoneFormat := phoneBaseFormat()
		if (phoneFormat = "")
			toast("Tip:", "`nHighlight a 10 Digit Phone Number", ,5000)
			Exit
	phonenumber := "("phoneFormat[4]")"phoneFormat[5]"-"phoneFormat[6]
	openLink := "https://www.truepeoplesearch.com/results?phoneno=" phonenumber
	openlink := URI_URLEncode(openLink)
	ShellRun(openLink)
		toast("Opening People Search", openLink, ,2000)
		return
}

phoneBaseFormat(){
	oldClip := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
		oldClip := oldClip2
	oldClip := RegExReplace(oldClip, "[^0-9]", "")
	if (!RegExMatch(oldClip, "^\d{9,11}$")) {
		toast("Tip:", "`nHighlight a 10 Digit Phone Number", ,5000)
		Exit
		}
	phoneSection1 := SubStr(oldClip, 1, 3)
	phoneSection2 := SubStr(oldClip, 4, 3)
	phoneSection3 := SubStr(oldClip, 7, 4)
	phoneNumber1 := RegExReplace(oldClip, "^.*(\d{3})[^\d]*(\d{3})[^\d]*(\d{4})$", "$1-$2-$3", 1)
	phoneNumber2 := RegExReplace(oldClip, "^.*(\d{3})[^\d]*(\d{3})[^\d]*(\d{4})$", "($1) $2-$3", 1)
	phoneNumber3 := RegExReplace(oldClip, "^.*(\d{3})[^\d]*(\d{3})[^\d]*(\d{4})$", "($1)$2-$3", 1)
	Clipboard := oldClip
	return [phoneNumber1, phoneNumber2, phoneNumber3, phoneSection1, phoneSection2, phoneSection3]
}

phoneFormat(){
	phoneFormat := phoneBaseFormat()
	Clip(phoneFormat[1])
		return
}

phoneFormatPara(){
	phoneFormat := phoneBaseFormat()
	Clip(phoneFormat[2], true)
		return
}

fileRename(){
	Send, {F2}
	oldClip := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
		oldClip := oldClip2
	oldClip := StrReplace(StrReplace(StrReplace(StrReplace(oldClip, "_", " "), "-", " "), "   ", " "), "  ", " ")
		oldClip := StrReplace(StrReplace(caseChange(oldClip, "T"), "#", " "), "-", " ")
		oldClip = %oldClip%
		oldClip := StrReplace(oldClip, A_Space, "_")
	Clipboard := oldClip
	Clip(oldClip, true)
	return
}

cleanRename(){
	oldClip := Clipboard
	oldClip2 := Clip()
	if (oldClip2 != "")
		oldClip := oldClip2
	oldClip := StrReplace(StrReplace(StrReplace(StrReplace(oldClip, "_", " "), "-", " "), "   ", " "), "  ", " ")
		oldClip := StrReplace(StrReplace(caseChange(oldClip, "T"), "#", " "), "-", " ")
		oldClip = %oldClip%
	Clipboard := oldClip
	Clip(oldClip, true)
		return
}

;=========================================JEEE========================================
;functions for Desktop and Explorer folder windows:
;JEE_ExpWinGetSel ;get selected files (list)
;JEE_ExpWinGetDir ;get Explorer window folder
;JEE_ExpWinSetDir ;set Explorer window folder (navigate) (or open file/url in new window)
;=====================================================================================
;=====================================================================================

JEE_ExpWinGetSel(hWnd:=0, vSep:="`n")
{
	local oItem, oWin, oWindows, vCount, vOutput, vWinClass
	DetectHiddenWindows, On
	(!hWnd) && hWnd := WinExist("A")
	WinGetClass, vWinClass, % "ahk_id " hWnd
	if (vWinClass = "CabinetWClass") || (vWinClass = "ExploreWClass")
	{
		for oWin in ComObjCreate("Shell.Application").Windows
			if (oWin.HWND = hWnd)
				break
	}
	else if (vWinClass = "Progman") || (vWinClass = "WorkerW")
	{
		oWindows := ComObjCreate("Shell.Application").Windows
		VarSetCapacity(hWnd, 4, 0)
		;SWC_DESKTOP := 0x8 ;VT_BYREF := 0x4000 ;VT_I4 := 0x3 ;SWFO_NEEDDISPATCH := 0x1
		oWin := oWindows.FindWindowSW(0, "", 8, ComObject(0x4003, &hWnd), 1)
	}
	else
		return

	vCount := oWin.Document.SelectedItems.Count
	vOutput := ""
	VarSetCapacity(vOutput, (260+StrLen(vSep))*vCount*2)
	for oItem in oWin.Document.SelectedItems
		if !(SubStr(oItem.path, 1, 3) = "::{")
			vOutput .= oItem.path vSep
	oWindows := oWin := oItem := ""
	return SubStr(vOutput, 1, -StrLen(vSep))
}

;=====================================================================================

JEE_ExpWinGetDir(hWnd:=0)
{
	local oWin, oWindows, vDir, vWinClass
	DetectHiddenWindows, On
	(!hWnd) && hWnd := WinExist("A")
	WinGetClass, vWinClass, % "ahk_id " hWnd
	if (vWinClass = "CabinetWClass") || (vWinClass = "ExploreWClass")
	{
		for oWin in ComObjCreate("Shell.Application").Windows
			if (oWin.HWND = hWnd)
				break
	}
	else if (vWinClass = "Progman") || (vWinClass = "WorkerW")
	{
		oWindows := ComObjCreate("Shell.Application").Windows
		VarSetCapacity(hWnd, 4, 0)
		;SWC_DESKTOP := 0x8 ;VT_BYREF := 0x4000 ;VT_I4 := 0x3 ;SWFO_NEEDDISPATCH := 0x1
		oWin := oWindows.FindWindowSW(0, "", 8, ComObject(0x4003, &hWnd), 1)
	}
	else
		return
	vDir := oWin.Document.Folder.Self.Path
	oWindows := oWin := ""
	return vDir
}
;======================================================================================
;{==================================WinClipAPI=========================================
;} ====================================================================================

class WinClip_base
{
	__Call( aTarget, aParams* ) {
		if ObjHasKey( WinClip_base, aTarget )
			return WinClip_base[ aTarget ].Call( this, aParams* ) ;updated for AHKv2 compatibility
			;return WinClip_base[ aTarget ].( this, aParams* )
		throw Exception( "Unknown function '" aTarget "' requested from object '" this.__Class "'", -1 )
	}

	Err( msg ) {
		throw Exception( this.__Class " : " msg ( A_LastError != 0 ? "`n" this.ErrorFormat( A_LastError ) : "" ), -2 )
	}

	ErrorFormat( error_id ) {
		VarSetCapacity(msg,1000,0)
		if !len := DllCall("FormatMessageW"
					,"UInt",FORMAT_MESSAGE_FROM_SYSTEM := 0x00001000 | FORMAT_MESSAGE_IGNORE_INSERTS := 0x00000200      ;dwflags
					,"Ptr",0        ;lpSource
					,"UInt",error_id    ;dwMessageId
					,"UInt",0           ;dwLanguageId
					,"Ptr",&msg         ;lpBuffer
					,"UInt",500)            ;nSize
			return
		return  strget(&msg,len)
	}
}

class WinClipAPI_base extends WinClip_base
{
	__Get( name ) {
		if !ObjHasKey( this, initialized )
			this.Init()
		else
			throw Exception( "Unknown field '" name "' requested from object '" this.__Class "'", -1 )
	}
}

class WinClipAPI extends WinClip_base
{
	memcopy( dest, src, size ) {
		return DllCall( "msvcrt\memcpy", "ptr", dest, "ptr", src, "uint", size )
	}
	GlobalSize( hObj ) {
		return DllCall( "GlobalSize", "Ptr", hObj )
	}
	GlobalLock( hMem ) {
		return DllCall( "GlobalLock", "Ptr", hMem )
	}
	GlobalUnlock( hMem ) {
		return DllCall( "GlobalUnlock", "Ptr", hMem )
	}
	GlobalAlloc( flags, size ) {
		return DllCall( "GlobalAlloc", "Uint", flags, "Uint", size )
	}
	OpenClipboard() {
		return DllCall( "OpenClipboard", "Ptr", 0 )
	}
	CloseClipboard() {
		return DllCall( "CloseClipboard" )
	}
	SetClipboardData( format, hMem ) {
		return DllCall( "SetClipboardData", "Uint", format, "Ptr", hMem )
	}
	GetClipboardData( format ) {
		return DllCall( "GetClipboardData", "Uint", format )
	}
	EmptyClipboard() {
		return DllCall( "EmptyClipboard" )
	}
	EnumClipboardFormats( format ) {
		return DllCall( "EnumClipboardFormats", "UInt", format )
	}
	CountClipboardFormats() {
		return DllCall( "CountClipboardFormats" )
	}
	GetClipboardFormatName( iFormat ) {
		size := VarSetCapacity( bufName, 255*( A_IsUnicode ? 2 : 1 ), 0 )
		DllCall( "GetClipboardFormatName", "Uint", iFormat, "str", bufName, "Uint", size )
		return bufName
	}
	GetEnhMetaFileBits( hemf, ByRef buf ) {
		if !( bufSize := DllCall( "GetEnhMetaFileBits", "Ptr", hemf, "Uint", 0, "Ptr", 0 ) )
			return 0
		VarSetCapacity( buf, bufSize, 0 )
		if !( bytesCopied := DllCall( "GetEnhMetaFileBits", "Ptr", hemf, "Uint", bufSize, "Ptr", &buf ) )
			return 0
		return bytesCopied
	}
	SetEnhMetaFileBits( pBuf, bufSize ) {
		return DllCall( "SetEnhMetaFileBits", "Uint", bufSize, "Ptr", pBuf )
	}
	DeleteEnhMetaFile( hemf ) {
		return DllCall( "DeleteEnhMetaFile", "Ptr", hemf )
	}
	ErrorFormat(error_id) {
		VarSetCapacity(msg,1000,0)
		if !len := DllCall("FormatMessageW"
					,"UInt",FORMAT_MESSAGE_FROM_SYSTEM := 0x00001000 | FORMAT_MESSAGE_IGNORE_INSERTS := 0x00000200      ;dwflags
					,"Ptr",0        ;lpSource
					,"UInt",error_id    ;dwMessageId
					,"UInt",0           ;dwLanguageId
					,"Ptr",&msg         ;lpBuffer
					,"UInt",500)            ;nSize
			return
		return  strget(&msg,len)
	}
	IsInteger( var ) {
		if (var+0 == var) && (Floor(var) == var) ;test for integer while remaining v1 and v2 compatible
			return True
		else
			return False
	}
	LoadDllFunction( file, function ) {
			if !hModule := DllCall( "GetModuleHandleW", "Wstr", file, "UPtr" )
					hModule := DllCall( "LoadLibraryW", "Wstr", file, "UPtr" )

			ret := DllCall("GetProcAddress", "Ptr", hModule, "AStr", function, "UPtr")
			return ret
	}
	SendMessage( hWnd, Msg, wParam, lParam ) {
		 static SendMessageW

		 If not SendMessageW
				SendMessageW := this.LoadDllFunction( "user32.dll", "SendMessageW" )

		 ret := DllCall( SendMessageW, "UPtr", hWnd, "UInt", Msg, "UPtr", wParam, "UPtr", lParam )
		 return ret
	}
	GetWindowThreadProcessId( hwnd ) {
		return DllCall( "GetWindowThreadProcessId", "Ptr", hwnd, "Ptr", 0 )
	}
	WinGetFocus( hwnd ) {
		GUITHREADINFO_cbsize := 24 + A_PtrSize*6
		VarSetCapacity( GuiThreadInfo, GUITHREADINFO_cbsize, 0 )    ;GuiThreadInfoSize = 48
		NumPut(GUITHREADINFO_cbsize, GuiThreadInfo, 0, "UInt")
		threadWnd := this.GetWindowThreadProcessId( hwnd )
		if not DllCall( "GetGUIThreadInfo", "uint", threadWnd, "UPtr", &GuiThreadInfo )
				return 0
		return NumGet( GuiThreadInfo, 8+A_PtrSize,"UPtr")  ; Retrieve the hwndFocus field from the struct.
	}
	GetPixelInfo( ByRef DIB ) {
		;~ typedef struct tagBITMAPINFOHEADER {
		;~ DWORD biSize;              0
		;~ LONG  biWidth;             4
		;~ LONG  biHeight;            8
		;~ WORD  biPlanes;            12
		;~ WORD  biBitCount;          14
		;~ DWORD biCompression;       16
		;~ DWORD biSizeImage;         20
		;~ LONG  biXPelsPerMeter;     24
		;~ LONG  biYPelsPerMeter;     28
		;~ DWORD biClrUsed;           32
		;~ DWORD biClrImportant;      36

		bmi := &DIB  ;BITMAPINFOHEADER  pointer from DIB
		biSize := numget( bmi+0, 0, "UInt" )
		;~ return bmi + biSize
		biSizeImage := numget( bmi+0, 20, "UInt" )
		biBitCount := numget( bmi+0, 14, "UShort" )
		if ( biSizeImage == 0 )
		{
			biWidth := numget( bmi+0, 4, "UInt" )
			biHeight := numget( bmi+0, 8, "UInt" )
			biSizeImage := (((( biWidth * biBitCount + 31 ) & ~31 ) >> 3 ) * biHeight )
			numput( biSizeImage, bmi+0, 20, "UInt" )
		}
		p := numget( bmi+0, 32, "UInt" )  ;biClrUsed
		if ( p == 0 && biBitCount <= 8 )
			p := 1 << biBitCount
		p := p * 4 + biSize + bmi
		return p
	}
	Gdip_Startup() {
		if !DllCall( "GetModuleHandleW", "Wstr", "gdiplus", "UPtr" )
					DllCall( "LoadLibraryW", "Wstr", "gdiplus", "UPtr" )

		VarSetCapacity(GdiplusStartupInput , 3*A_PtrSize, 0), NumPut(1,GdiplusStartupInput ,0,"UInt") ; GdiplusVersion = 1
		DllCall("gdiplus\GdiplusStartup", "Ptr*", pToken, "Ptr", &GdiplusStartupInput, "Ptr", 0)
		return pToken
	}
	Gdip_Shutdown(pToken) {
		DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)
		if hModule := DllCall( "GetModuleHandleW", "Wstr", "gdiplus", "UPtr" )
			DllCall("FreeLibrary", "Ptr", hModule)
		return 0
	}
	StrSplit(str,delim,omit := "") {
		if (strlen(delim) > 1)
		{
			;StringReplace,str,str,% delim,ƒ,1      ;■¶╬
			str := StrReplace(str, delim, "ƒ")
			delim := "ƒ"
		}
		ra := Array()
		loop, parse,str,% delim,% omit
			if (A_LoopField != "")
				ra.Insert(A_LoopField)
		return ra
	}
	RemoveDubls( objArray ) {
		while True
		{
			nodubls := 1
			tempArr := Object()
			for i,val in objArray
			{
				if tempArr.haskey( val )
				{
					nodubls := 0
					objArray.Remove( i )
					break
				}
				tempArr[ val ] := 1
			}
			if nodubls
				break
		}
		return objArray
	}
	RegisterClipboardFormat( fmtName ) {
		return DllCall( "RegisterClipboardFormat", "ptr", &fmtName )
	}
	GetOpenClipboardWindow() {
		return DllCall( "GetOpenClipboardWindow" )
	}
	IsClipboardFormatAvailable( iFmt ) {
		return DllCall( "IsClipboardFormatAvailable", "UInt", iFmt )
	}
	GetImageEncodersSize( ByRef numEncoders, ByRef size ) {
		return DllCall( "gdiplus\GdipGetImageEncodersSize", "Uint*", numEncoders, "UInt*", size )
	}
	GetImageEncoders( numEncoders, size, pImageCodecInfo ) {
		return DllCall( "gdiplus\GdipGetImageEncoders", "Uint", numEncoders, "UInt", size, "Ptr", pImageCodecInfo )
	}
	GetEncoderClsid( format, ByRef CLSID ) {
		;format should be the following
		;~ bmp
		;~ jpeg
		;~ gif
		;~ tiff
		;~ png
		if !format
			return 0
		format := "image/" format
		this.GetImageEncodersSize( num, size )
		if ( size = 0 )
			return 0
		VarSetCapacity( ImageCodecInfo, size, 0 )
		this.GetImageEncoders( num, size, &ImageCodecInfo )
		loop,% num
		{
			pici := &ImageCodecInfo + ( 48+7*A_PtrSize )*(A_Index-1)
			pMime := NumGet( pici+0, 32+4*A_PtrSize, "UPtr" )
			MimeType := StrGet( pMime, "UTF-16")
			if ( MimeType = format )
			{
				VarSetCapacity( CLSID, 16, 0 )
				this.memcopy( &CLSID, pici, 16 )
				return 1
			}
		}
		return 0
	}
}

;{=====================================================================================
;} ====================================================================================

;{================================WinCLip==============================================
;} ====================================================================================
class WinClip extends WinClip_base
{
	__New()
	{
		this.isinstance := 1
		this.allData := ""
	}

	_toclipboard( ByRef data, size )
	{
		if !WinClipAPI.OpenClipboard()
			return 0
		offset := 0
		lastPartOffset := 0
		WinClipAPI.EmptyClipboard()
		while ( offset < size )
		{
			if !( fmt := NumGet( data, offset, "UInt" ) )
				break
			offset += 4
			if !( dataSize := NumGet( data, offset, "UInt" ) )
				break
			offset += 4
			if ( ( offset + dataSize ) > size )
				break
			if !( pData := WinClipAPI.GlobalLock( WinClipAPI.GlobalAlloc( 0x0042, dataSize ) ) )
			{
				offset += dataSize
				continue
			}
			WinClipAPI.memcopy( pData, &data + offset, dataSize )
			if ( fmt == this.ClipboardFormats.CF_ENHMETAFILE )
				pClipData := WinClipAPI.SetEnhMetaFileBits( pData, dataSize )
			else
				pClipData := pData
			if !pClipData
				continue
			WinClipAPI.SetClipboardData( fmt, pClipData )
			if ( fmt == this.ClipboardFormats.CF_ENHMETAFILE )
				WinClipAPI.DeleteEnhMetaFile( pClipData )
			WinClipAPI.GlobalUnlock( pData )
			offset += dataSize
			lastPartOffset := offset
		}
		WinClipAPI.CloseClipboard()
		return lastPartOffset
	}

	_fromclipboard( ByRef clipData )
	{
		if !WinClipAPI.OpenClipboard()
			return 0
		nextformat := 0
		objFormats := object()
		clipSize := 0
		formatsNum := 0
		while ( nextformat := WinClipAPI.EnumClipboardFormats( nextformat ) )
		{
			if this.skipFormats.hasKey( nextformat )
				continue
			if ( dataHandle := WinClipAPI.GetClipboardData( nextformat ) )
			{
				pObjPtr := 0, nObjSize := 0
				if ( nextformat == this.ClipboardFormats.CF_ENHMETAFILE )
				{
					if ( bufSize := WinClipAPI.GetEnhMetaFileBits( dataHandle, hemfBuf ) )
						pObjPtr := &hemfBuf, nObjSize := bufSize
				}
				else if ( nSize := WinClipAPI.GlobalSize( WinClipAPI.GlobalLock( dataHandle ) ) )
					pObjPtr := dataHandle, nObjSize := nSize
				else
					continue
				if !( pObjPtr && nObjSize )
					continue
				objFormats[ nextformat ] := { handle : pObjPtr, size : nObjSize }
				clipSize += nObjSize
				formatsNum++
			}
		}
		structSize := formatsNum*( 4 + 4 ) + clipSize  ;allocating 4 bytes for format ID and 4 for data size
		if !structSize
		{
			WinClipAPI.CloseClipboard()
			return 0
		}
		VarSetCapacity( clipData, structSize, 0 )
		; array in form of:
		; format   UInt
		; dataSize UInt
		; data     Byte[]
		offset := 0
		for fmt, params in objFormats
		{
			NumPut( fmt, &clipData, offset, "UInt" )
			offset += 4
			NumPut( params.size, &clipData, offset, "UInt" )
			offset += 4
			WinClipAPI.memcopy( &clipData + offset, params.handle, params.size )
			offset += params.size
			WinClipAPI.GlobalUnlock( params.handle )
		}
		WinClipAPI.CloseClipboard()
		return structSize
	}

	_IsInstance( funcName )
	{
		if !this.isinstance
		{
			throw Exception( "Error in '" funcName "':`nInstantiate the object first to use this method!", -1 )
			return 0
		}
		return 1
	}

	_loadFile( filePath, ByRef Data )
	{
		f := FileOpen( filePath, "r","CP0" )
		if !IsObject( f )
			return 0
		f.Pos := 0
		dataSize := f.RawRead( Data, f.Length )
		f.close()
		return dataSize
	}

	_saveFile( filepath, byRef data, size )
	{
		f := FileOpen( filepath, "w","CP0" )
		bytes := f.RawWrite( &data, size )
		f.close()
		return bytes
	}

	_setClipData( ByRef data, size )
	{
		if !size
			return 0
		if !ObjSetCapacity( this, "allData", size )
			return 0
		if !( pData := ObjGetAddress( this, "allData" ) )
			return 0
		WinClipAPI.memcopy( pData, &data, size )
		return size
	}

	_getClipData( ByRef data )
	{
		if !( clipSize := ObjGetCapacity( this, "allData" ) )
			return 0
		if !( pData := ObjGetAddress( this, "allData" ) )
			return 0
		VarSetCapacity( data, clipSize, 0 )
		WinClipAPI.memcopy( &data, pData, clipSize )
		return clipSize
	}

	__Delete()
	{
		ObjSetCapacity( this, "allData", 0 )
		return
	}

	_parseClipboardData( ByRef data, size )
	{
		offset := 0
		formats := object()
		while ( offset < size )
		{
			if !( fmt := NumGet( data, offset, "UInt" ) )
				break
			offset += 4
			if !( dataSize := NumGet( data, offset, "UInt" ) )
				break
			offset += 4
			if ( ( offset + dataSize ) > size )
				break
			params := { name : this._getFormatName( fmt ), size : dataSize }
			ObjSetCapacity( params, "buffer", dataSize )
			pBuf := ObjGetAddress( params, "buffer" )
			WinClipAPI.memcopy( pBuf, &data + offset, dataSize )
			formats[ fmt ] := params
			offset += dataSize
		}
		return formats
	}

	_compileClipData( ByRef out_data, objClip )
	{
		if !IsObject( objClip )
			return 0
		;calculating required data size
		clipSize := 0
		for fmt, params in objClip
			clipSize += 8 + params.size
		VarSetCapacity( out_data, clipSize, 0 )
		offset := 0
		for fmt, params in objClip
		{
			NumPut( fmt, out_data, offset, "UInt" )
			offset += 4
			NumPut( params.size, out_data, offset, "UInt" )
			offset += 4
			WinClipAPI.memcopy( &out_data + offset, ObjGetAddress( params, "buffer" ), params.size )
			offset += params.size
		}
		return clipSize
	}

	GetFormats()
	{
		if !( clipSize := this._fromclipboard( clipData ) )
			return 0
		return this._parseClipboardData( clipData, clipSize )
	}

	iGetFormats()
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return 0
		return this._parseClipboardData( clipData, clipSize )
	}

	Snap( ByRef data )
	{
		return this._fromclipboard( data )
	}

	iSnap()
	{
		this._IsInstance( A_ThisFunc )
		if !( dataSize := this._fromclipboard( clipData ) )
			return 0
		return this._setClipData( clipData, dataSize )
	}

	Restore( ByRef clipData )
	{
		clipSize := VarSetCapacity( clipData )
		return this._toclipboard( clipData, clipSize )
	}

	iRestore()
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return 0
		return this._toclipboard( clipData, clipSize )
	}

	Save( filePath )
	{
		if !( size := this._fromclipboard( data ) )
			return 0
		return this._saveFile( filePath, data, size )
	}

	iSave( filePath )
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
					return 0
		return this._saveFile( filePath, clipData, clipSize )
	}

	Load( filePath )
	{
		if !( dataSize := this._loadFile( filePath, dataBuf ) )
			return 0
		return this._toclipboard( dataBuf, dataSize )
	}

	iLoad( filePath )
	{
		this._IsInstance( A_ThisFunc )
		if !( dataSize := this._loadFile( filePath, dataBuf ) )
			return 0
		return this._setClipData( dataBuf, dataSize )
	}

	Clear()
	{
		if !WinClipAPI.OpenClipboard()
			return 0
		WinClipAPI.EmptyClipboard()
		WinClipAPI.CloseClipboard()
		return 1
	}

	iClear()
	{
		this._IsInstance( A_ThisFunc )
		ObjSetCapacity( this, "allData", 0 )
	}

	Copy( timeout := 1, method := 1 )
	{
		this.Snap( data )
		this.Clear()    ;clearing the clipboard
		if( method = 1 )
			SendInput, ^{Ins}
		else
			SendInput, ^{vk43sc02E} ;ctrl+c
		ClipWait,% timeout, 1
		if ( ret := this._isClipEmpty() )
			this.Restore( data )
		return !ret
	}

	iCopy( timeout := 1, method := 1 )
	{
		this._IsInstance( A_ThisFunc )
		this.Snap( data )
		this.Clear()    ;clearing the clipboard
		if( method = 1 )
			SendInput, ^{Ins}
		else
			SendInput, ^{vk43sc02E} ;ctrl+c
		ClipWait,% timeout, 1
		bytesCopied := 0
		if !this._isClipEmpty()
		{
			this.iClear()   ;clearing the variable containing the clipboard data
			bytesCopied := this.iSnap()
		}
		this.Restore( data )
		return bytesCopied
	}

	Paste( plainText := "", method := 1 )
	{
		ret := 0
		if ( plainText != "" )
		{
			this.Snap( data )
			this.Clear()
			ret := this.SetText( plainText )
		}
		if( method = 1 )
			SendInput, +{Ins}
		else
			SendInput, ^{vk56sc02F} ;ctrl+v
		this._waitClipReady( 3000 )
		if ( plainText != "" )
		{
			this.Restore( data )
		}
		else
			ret := !this._isClipEmpty()
		return ret
	}

	iPaste( method := 1 )
	{
		this._IsInstance( A_ThisFunc )
		this.Snap( data )
		if !( bytesRestored := this.iRestore() )
			return 0
		if( method = 1 )
			SendInput, +{Ins}
		else
			SendInput, ^{vk56sc02F} ;ctrl+v
		this._waitClipReady( 3000 )
		this.Restore( data )
		return bytesRestored
	}

	IsEmpty()
	{
		return this._isClipEmpty()
	}

	iIsEmpty()
	{
		return !this.iGetSize()
	}

	_isClipEmpty()
	{
		return !WinClipAPI.CountClipboardFormats()
	}

	_waitClipReady( timeout := 10000 )
	{
		start_time := A_TickCount
		sleep 100
		while ( WinClipAPI.GetOpenClipboardWindow() && ( A_TickCount - start_time < timeout ) )
			sleep 100
	}

	iSetText( textData )
	{
		if ( textData = "" )
			return 0
		this._IsInstance( A_ThisFunc )
		clipSize := this._getClipData( clipData )
		if !( clipSize := this._appendText( clipData, clipSize, textData, 1 ) )
			return 0
		return this._setClipData( clipData, clipSize )
	}

	SetText( textData )
	{
		if ( textData = "" )
			return 0
		clipSize :=  this._fromclipboard( clipData )
		if !( clipSize := this._appendText( clipData, clipSize, textData, 1 ) )
			return 0
		return this._toclipboard( clipData, clipSize )
	}

	GetRTF()
	{
		if !( clipSize := this._fromclipboard( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, "Rich Text Format" ) )
			return ""
		return strget( &out_data, out_size, "CP0" )
	}

	iGetRTF()
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, "Rich Text Format" ) )
			return ""
		return strget( &out_data, out_size, "CP0" )
	}

	SetRTF( textData )
	{
		if ( textData = "" )
			return 0
		clipSize :=  this._fromclipboard( clipData )
		if !( clipSize := this._setRTF( clipData, clipSize, textData ) )
					return 0
		return this._toclipboard( clipData, clipSize )
	}

	iSetRTF( textData )
	{
		if ( textData = "" )
			return 0
		this._IsInstance( A_ThisFunc )
		clipSize :=  this._getClipData( clipData )
		if !( clipSize := this._setRTF( clipData, clipSize, textData ) )
					return 0
		return this._setClipData( clipData, clipSize )
	}

	_setRTF( ByRef clipData, clipSize, textData )
	{
		objFormats := this._parseClipboardData( clipData, clipSize )
		uFmt := WinClipAPI.RegisterClipboardFormat( "Rich Text Format" )
		objFormats[ uFmt ] := object()
		sLen := StrLen( textData )
		ObjSetCapacity( objFormats[ uFmt ], "buffer", sLen )
		StrPut( textData, ObjGetAddress( objFormats[ uFmt ], "buffer" ), sLen, "CP0" )
		objFormats[ uFmt ].size := sLen
		return this._compileClipData( clipData, objFormats )
	}

	iAppendText( textData )
	{
		if ( textData = "" )
			return 0
		this._IsInstance( A_ThisFunc )
		clipSize := this._getClipData( clipData )
		if !( clipSize := this._appendText( clipData, clipSize, textData ) )
			return 0
		return this._setClipData( clipData, clipSize )
	}

	AppendText( textData )
	{
		if ( textData = "" )
			return 0
		clipSize :=  this._fromclipboard( clipData )
		if !( clipSize := this._appendText( clipData, clipSize, textData ) )
			return 0
		return this._toclipboard( clipData, clipSize )
	}

	SetHTML( html, source := "" )
	{
		if ( html = "" )
			return 0
		clipSize :=  this._fromclipboard( clipData )
		if !( clipSize := this._setHTML( clipData, clipSize, html, source ) )
			return 0
		return this._toclipboard( clipData, clipSize )
	}

	iSetHTML( html, source := "" )
	{
		if ( html = "" )
			return 0
		this._IsInstance( A_ThisFunc )
		clipSize := this._getClipData( clipData )
		if !( clipSize := this._setHTML( clipData, clipSize, html, source ) )
			return 0
		return this._setClipData( clipData, clipSize )
	}

	_calcHTMLLen( num )
	{
		while ( StrLen( num ) < 10 )
			num := "0" . num
		return num
	}

	_setHTML( ByRef clipData, clipSize, htmlData, source )
	{
		objFormats := this._parseClipboardData( clipData, clipSize )
		uFmt := WinClipAPI.RegisterClipboardFormat( "HTML Format" )
		objFormats[ uFmt ] := object()
		encoding := "UTF-8"
		htmlLen := StrPut( htmlData, encoding ) - 1   ;substract null
		srcLen := 2 + 10 + StrPut( source, encoding ) - 1      ;substract null
		StartHTML := this._calcHTMLLen( 105 + srcLen )
		EndHTML := this._calcHTMLLen( StartHTML + htmlLen + 76 )
		StartFragment := this._calcHTMLLen( StartHTML + 38 )
		EndFragment := this._calcHTMLLen( StartFragment + htmlLen )

		;replace literal continuation section to make cross-compatible with v1 and v2
		html := "Version:0.9`r`n"
		html .= "StartHTML:" . StartHTML . "`r`n"
		html .= "EndHTML:" . EndHTML . "`r`n"
		html .= "StartFragment:" . StartFragment . "`r`n"
		html .= "EndFragment:" . EndFragment . "`r`n"
		html .= "SourceURL:" . source . "`r`n"
		html .= "<html>`r`n"
		html .= "<body>`r`n"
		html .= "<!--StartFragment-->`r`n"
		html .= htmlData . "`r`n"
		html .= "<!--EndFragment-->`r`n"
		html .= "</body>`r`n"
		html .= "</html>`r`n"

		sLen := StrPut( html, encoding )
		ObjSetCapacity( objFormats[ uFmt ], "buffer", sLen )
		StrPut( html, ObjGetAddress( objFormats[ uFmt ], "buffer" ), sLen, encoding )
		objFormats[ uFmt ].size := sLen
		return this._compileClipData( clipData, objFormats )
	}

	_appendText( ByRef clipData, clipSize, textData, IsSet := 0 )
	{
		objFormats := this._parseClipboardData( clipData, clipSize )
		uFmt := this.ClipboardFormats.CF_UNICODETEXT
		str := ""
		if ( objFormats.haskey( uFmt ) && !IsSet )
			str := strget( ObjGetAddress( objFormats[ uFmt ],  "buffer" ), "UTF-16" )
		else
			objFormats[ uFmt ] := object()
		str .= textData
		sLen := ( StrLen( str ) + 1 ) * 2
		ObjSetCapacity( objFormats[ uFmt ], "buffer", sLen )
		StrPut( str, ObjGetAddress( objFormats[ uFmt ], "buffer" ), sLen, "UTF-16" )
		objFormats[ uFmt ].size := sLen
		return this._compileClipData( clipData, objFormats )
	}

	_getFiles( pDROPFILES )
	{
		fWide := numget( pDROPFILES + 0, 16, "uchar" ) ;getting fWide value from DROPFILES struct
		pFiles := numget( pDROPFILES + 0, 0, "UInt" ) + pDROPFILES  ;getting address of files list
		list := ""
		while numget( pFiles + 0, 0, fWide ? "UShort" : "UChar" )
		{
			lastPath := strget( pFiles+0, fWide ? "UTF-16" : "CP0" )
			list .= ( list ? "`n" : "" ) lastPath
			pFiles += ( StrLen( lastPath ) + 1 ) * ( fWide ? 2 : 1 )
		}
		return list
	}

	_setFiles( ByRef clipData, clipSize, files, append := 0, isCut := 0 )
	{
		objFormats := this._parseClipboardData( clipData, clipSize )
		uFmt := this.ClipboardFormats.CF_HDROP
		if ( append && objFormats.haskey( uFmt ) )
			prevList := this._getFiles( ObjGetAddress( objFormats[ uFmt ], "buffer" ) ) "`n"
		objFiles := WinClipAPI.StrSplit( prevList . files, "`n", A_Space A_Tab )
		objFiles := WinClipAPI.RemoveDubls( objFiles )
		if !objFiles.MaxIndex()
			return 0
		objFormats[ uFmt ] := object()
		DROP_size := 20 + 2
		for i,str in objFiles
			DROP_size += ( StrLen( str ) + 1 ) * 2
		VarSetCapacity( DROPFILES, DROP_size, 0 )
		NumPut( 20, DROPFILES, 0, "UInt" )  ;offset
		NumPut( 1, DROPFILES, 16, "uchar" ) ;NumPut( 20, DROPFILES, 0, "UInt" )
		offset := &DROPFILES + 20
		for i,str in objFiles
		{
			StrPut( str, offset, "UTF-16" )
			offset += ( StrLen( str ) + 1 ) * 2
		}
		ObjSetCapacity( objFormats[ uFmt ], "buffer", DROP_size )
		WinClipAPI.memcopy( ObjGetAddress( objFormats[ uFmt ], "buffer" ), &DROPFILES, DROP_size )
		objFormats[ uFmt ].size := DROP_size
		prefFmt := WinClipAPI.RegisterClipboardFormat( "Preferred DropEffect" )
		objFormats[ prefFmt ] := { size : 4 }
		ObjSetCapacity( objFormats[ prefFmt ], "buffer", 4 )
		NumPut( isCut ? 2 : 5, ObjGetAddress( objFormats[ prefFmt ], "buffer" ), 0, "UInt" )
		return this._compileClipData( clipData, objFormats )
	}

	SetFiles( files, isCut := 0 )
	{
		if ( files = "" )
			return 0
		clipSize := this._fromclipboard( clipData )
		if !( clipSize := this._setFiles( clipData, clipSize, files, 0, isCut ) )
			return 0
		return this._toclipboard( clipData, clipSize )
	}

	iSetFiles( files, isCut := 0 )
	{
		this._IsInstance( A_ThisFunc )
		if ( files = "" )
			return 0
		clipSize := this._getClipData( clipData )
		if !( clipSize := this._setFiles( clipData, clipSize, files, 0, isCut ) )
			return 0
		return this._setClipData( clipData, clipSize )
	}

	AppendFiles( files, isCut := 0 )
	{
		if ( files = "" )
			return 0
		clipSize := this._fromclipboard( clipData )
		if !( clipSize := this._setFiles( clipData, clipSize, files, 1, isCut ) )
			return 0
		return this._toclipboard( clipData, clipSize )
	}

	iAppendFiles( files, isCut := 0 )
	{
		this._IsInstance( A_ThisFunc )
		if ( files = "" )
			return 0
		clipSize := this._getClipData( clipData )
		if !( clipSize := this._setFiles( clipData, clipSize, files, 1, isCut ) )
			return 0
		return this._setClipData( clipData, clipSize )
	}

	GetFiles()
	{
		if !( clipSize := this._fromclipboard( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_HDROP ) )
			return ""
		return this._getFiles( &out_data )
	}

	iGetFiles()
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_HDROP ) )
			return ""
		return this._getFiles( &out_data )
	}

	_getFormatData( ByRef out_data, ByRef data, size, needleFormat )
	{
		needleFormat := (WinClipAPI.IsInteger( needleFormat ) ? needleFormat : WinClipAPI.RegisterClipboardFormat( needleFormat ))
		if !needleFormat
			return 0
		offset := 0
		while ( offset < size )
		{
			if !( fmt := NumGet( data, offset, "UInt" ) )
				break
			offset += 4
			if !( dataSize := NumGet( data, offset, "UInt" ) )
				break
			offset += 4
			if ( fmt == needleFormat )
			{
				VarSetCapacity( out_data, dataSize, 0 )
				WinClipAPI.memcopy( &out_data, &data + offset, dataSize )
				return dataSize
			}
			offset += dataSize
		}
		return 0
	}

	_DIBtoHBITMAP( ByRef dibData )
	{
		;http://ebersys.blogspot.com/2009/06/how-to-convert-dib-to-bitmap.html
		pPix := WinClipAPI.GetPixelInfo( dibData )
		gdip_token := WinClipAPI.Gdip_Startup()
		DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", &dibData, "Ptr", pPix, "Ptr*", pBitmap )
		DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", hBitmap, "int", 0xffffffff )
		DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		WinClipAPI.Gdip_Shutdown( gdip_token )
		return hBitmap
	}

	GetBitmap()
	{
		if !( clipSize := this._fromclipboard( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
			return ""
		return this._DIBtoHBITMAP( out_data )
	}

	iGetBitmap()
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
			return ""
		return this._DIBtoHBITMAP( out_data )
	}

	_BITMAPtoDIB( bitmap, ByRef DIB )
	{
		if !bitmap
			return 0
		if !WinClipAPI.IsInteger( bitmap )
		{
			gdip_token := WinClipAPI.Gdip_Startup()
			DllCall("gdiplus\GdipCreateBitmapFromFileICM", "wstr", bitmap, "Ptr*", pBitmap )
			DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", hBitmap, "int", 0xffffffff )
			DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
			WinClipAPI.Gdip_Shutdown( gdip_token )
			bmMade := 1
		}
		else
			hBitmap := bitmap, bmMade := 0
		if !hBitmap
				return 0
		;http://www.codeguru.com/Cpp/G-M/bitmap/article.php/c1765
		if !( hdc := DllCall( "GetDC", "Ptr", 0 ) )
			goto, _BITMAPtoDIB_cleanup
		hPal := DllCall( "GetStockObject", "UInt", 15 ) ;DEFAULT_PALLETE
		hPal := DllCall( "SelectPalette", "ptr", hdc, "ptr", hPal, "Uint", 0 )
		DllCall( "RealizePalette", "ptr", hdc )
		size := DllCall( "GetObject", "Ptr", hBitmap, "Uint", 0, "ptr", 0 )
		VarSetCapacity( bm, size, 0 )
		DllCall( "GetObject", "Ptr", hBitmap, "Uint", size, "ptr", &bm )
		biBitCount := NumGet( bm, 16, "UShort" )*NumGet( bm, 18, "UShort" )
		nColors := (1 << biBitCount)
	if ( nColors > 256 )
		nColors := 0
	bmiLen  := 40 + nColors * 4
		VarSetCapacity( bmi, bmiLen, 0 )
		;BITMAPINFOHEADER initialization
		NumPut( 40, bmi, 0, "Uint" )
		NumPut( NumGet( bm, 4, "Uint" ), bmi, 4, "Uint" )   ;width
		NumPut( biHeight := NumGet( bm, 8, "Uint" ), bmi, 8, "Uint" ) ;height
		NumPut( 1, bmi, 12, "UShort" )
		NumPut( biBitCount, bmi, 14, "UShort" )
		NumPut( 0, bmi, 16, "UInt" ) ;compression must be BI_RGB

		; Get BITMAPINFO.
		if !DllCall("GetDIBits"
							,"ptr",hdc
							,"ptr",hBitmap
							,"uint",0
							,"uint",biHeight
							,"ptr",0      ;lpvBits
							,"ptr",&bmi  ;lpbi
							,"uint",0)    ;DIB_RGB_COLORS
			goto, _BITMAPtoDIB_cleanup
		biSizeImage := NumGet( &bmi, 20, "UInt" )
		if ( biSizeImage = 0 )
		{
			biBitCount := numget( &bmi, 14, "UShort" )
			biWidth := numget( &bmi, 4, "UInt" )
			biHeight := numget( &bmi, 8, "UInt" )
			biSizeImage := (((( biWidth * biBitCount + 31 ) & ~31 ) >> 3 ) * biHeight )
			;~ dwCompression := numget( bmi, 16, "UInt" )
			;~ if ( dwCompression != 0 ) ;BI_RGB
				;~ biSizeImage := ( biSizeImage * 3 ) / 2
			numput( biSizeImage, &bmi, 20, "UInt" )
		}
		DIBLen := bmiLen + biSizeImage
		VarSetCapacity( DIB, DIBLen, 0 )
		WinClipAPI.memcopy( &DIB, &bmi, bmiLen )
		if !DllCall("GetDIBits"
							,"ptr",hdc
							,"ptr",hBitmap
							,"uint",0
							,"uint",biHeight
							,"ptr",&DIB + bmiLen     ;lpvBits
							,"ptr",&DIB  ;lpbi
							,"uint",0)    ;DIB_RGB_COLORS
			goto, _BITMAPtoDIB_cleanup
_BITMAPtoDIB_cleanup:
		if bmMade
			DllCall( "DeleteObject", "ptr", hBitmap )
		DllCall( "SelectPalette", "ptr", hdc, "ptr", hPal, "Uint", 0 )
		DllCall( "RealizePalette", "ptr", hdc )
		DllCall("ReleaseDC","ptr",hdc)
		if ( A_ThisLabel = "_BITMAPtoDIB_cleanup" )
			return 0
		return DIBLen
	}

	_setBitmap( ByRef DIB, DIBSize, ByRef clipData, clipSize )
	{
		objFormats := this._parseClipboardData( clipData, clipSize )
		uFmt := this.ClipboardFormats.CF_DIB
		objFormats[ uFmt ] := { size : DIBSize }
		ObjSetCapacity( objFormats[ uFmt ], "buffer", DIBSize )
		WinClipAPI.memcopy( ObjGetAddress( objFormats[ uFmt ], "buffer" ), &DIB, DIBSize )
		return this._compileClipData( clipData, objFormats )
	}

	SetBitmap( bitmap )
	{
		if ( DIBSize := this._BITMAPtoDIB( bitmap, DIB ) )
		{
			clipSize := this._fromclipboard( clipData )
			if ( clipSize := this._setBitmap( DIB, DIBSize, clipData, clipSize ) )
				return this._toclipboard( clipData, clipSize )
		}
		return 0
	}

	iSetBitmap( bitmap )
	{
		this._IsInstance( A_ThisFunc )
		if ( DIBSize := this._BITMAPtoDIB( bitmap, DIB ) )
		{
			clipSize := this._getClipData( clipData )
			if ( clipSize := this._setBitmap( DIB, DIBSize, clipData, clipSize ) )
				return this._setClipData( clipData, clipSize )
		}
		return 0
	}

	GetText()
	{
		if !( clipSize := this._fromclipboard( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_UNICODETEXT ) )
			return ""
		return strget( &out_data, out_size / 2, "UTF-16" )
	}

	iGetText()
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, this.ClipboardFormats.CF_UNICODETEXT ) )
			return ""
		return strget( &out_data, out_size / 2, "UTF-16" )
	}

	GetHtml()
	{
		if !( clipSize := this._fromclipboard( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, "HTML Format" ) )
			return ""
		return strget( &out_data, out_size, "CP0" )
	}

	iGetHtml()
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return ""
		if !( out_size := this._getFormatData( out_data, clipData, clipSize, "HTML Format" ) )
			return ""
		return strget( &out_data, out_size, "CP0" )
	}

	_getFormatName( iformat )
	{
		if this.formatByValue.HasKey( iformat )
			return this.formatByValue[ iformat ]
		else
			return WinClipAPI.GetClipboardFormatName( iformat )
	}

	iGetData( ByRef Data )
	{
		this._IsInstance( A_ThisFunc )
		return this._getClipData( Data )
	}

	iSetData( ByRef data )
	{
		this._IsInstance( A_ThisFunc )
		return this._setClipData( data, VarSetCapacity( data ) )
	}

	iGetSize()
	{
		this._IsInstance( A_ThisFunc )
		return ObjGetCapacity( this, "alldata" )
	}

	HasFormat( fmt )
	{
		if !fmt
			return 0
		return WinClipAPI.IsClipboardFormatAvailable( WinClipAPI.IsInteger( fmt ) ? fmt : WinClipAPI.RegisterClipboardFormat( fmt )  )
	}

	iHasFormat( fmt )
	{
		this._IsInstance( A_ThisFunc )
		if !( clipSize := this._getClipData( clipData ) )
			return 0
		return this._hasFormat( clipData, clipSize, fmt )
	}

	_hasFormat( ByRef data, size, needleFormat )
	{
		needleFormat := WinClipAPI.IsInteger( needleFormat ) ? needleFormat : WinClipAPI.RegisterClipboardFormat( needleFormat )
		if !needleFormat
			return 0
		offset := 0
		while ( offset < size )
		{
			if !( fmt := NumGet( data, offset, "UInt" ) )
				break
			if ( fmt == needleFormat )
				return 1
			offset += 4
			if !( dataSize := NumGet( data, offset, "UInt" ) )
				break
			offset += 4 + dataSize
		}
		return 0
	}

	iSaveBitmap( filePath, format )
	{
		this._IsInstance( A_ThisFunc )
		if ( filePath = "" || format = "" )
			return 0
		if !( clipSize := this._getClipData( clipData ) )
			return 0
		if !( DIBsize := this._getFormatData( DIB, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
			return 0
		gdip_token := WinClipAPI.Gdip_Startup()
		if !WinClipAPI.GetEncoderClsid( format, CLSID )
			return 0
		DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", &DIB, "Ptr", WinClipAPI.GetPixelInfo( DIB ), "Ptr*", pBitmap )
		DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "wstr", filePath, "Ptr", &CLSID, "Ptr", 0 )
		DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		WinClipAPI.Gdip_Shutdown( gdip_token )
		return 1
	}

	SaveBitmap( filePath, format )
	{
		if ( filePath = "" || format = "" )
			return 0
		if !( clipSize := this._fromclipboard( clipData ) )
			return 0
		if !( DIBsize := this._getFormatData( DIB, clipData, clipSize, this.ClipboardFormats.CF_DIB ) )
			return 0
		gdip_token := WinClipAPI.Gdip_Startup()
		if !WinClipAPI.GetEncoderClsid( format, CLSID )
			return 0
		DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", &DIB, "Ptr", WinClipAPI.GetPixelInfo( DIB ), "Ptr*", pBitmap )
		DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "wstr", filePath, "Ptr", &CLSID, "Ptr", 0 )
		DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		WinClipAPI.Gdip_Shutdown( gdip_token )
		return 1
	}

	static ClipboardFormats := { CF_BITMAP : 2 ;A handle to a bitmap (HBITMAP).
								,CF_DIB : 8  ;A memory object containing a BITMAPINFO structure followed by the bitmap bits.
								,CF_DIBV5 : 17 ;A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits.
								,CF_DIF : 5 ;Software Arts' Data Interchange Format.
								,CF_DSPBITMAP : 0x0082 ;Bitmap display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data.
								,CF_DSPENHMETAFILE : 0x008E ;Enhanced metafile display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data.
								,CF_DSPMETAFILEPICT : 0x0083 ;Metafile-picture display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data.
								,CF_DSPTEXT : 0x0081 ;Text display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data.
								,CF_ENHMETAFILE : 14 ;A handle to an enhanced metafile (HENHMETAFILE).
								,CF_GDIOBJFIRST : 0x0300 ;Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST.Handles associated with clipboard formats in this range are not automatically deleted using the GlobalFree function when the clipboard is emptied. Also, when using values in this range, the hMem parameter is not a handle to a GDI object, but is a handle allocated by the GlobalAlloc function with the GMEM_MOVEABLE flag.
								,CF_GDIOBJLAST : 0x03FF ;See CF_GDIOBJFIRST.
								,CF_HDROP : 15 ;A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by passing the handle to the DragQueryFile function.
								,CF_LOCALE : 16 ;The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text. An application that pastes text from the clipboard can retrieve this format to determine which character set was used to generate the text. Note that the clipboard does not support plain text in multiple character sets. To achieve this, use a formatted text data type such as RTF instead. The system uses the code page associated with CF_LOCALE to implicitly convert from CF_TEXT to CF_UNICODETEXT. Therefore, the correct code page table is used for the conversion.
								,CF_METAFILEPICT : 3 ;Handle to a metafile picture format as defined by the METAFILEPICT structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle.
								,CF_OEMTEXT : 7 ;Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
								,CF_OWNERDISPLAY : 0x0080 ;Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The hMem parameter must be NULL.
								,CF_PALETTE : 9 ;Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.If the clipboard contains data in the CF_PALETTE (logical color palette) format, the application should use the SelectPalette and RealizePalette functions to realize (compare) any other data in the clipboard against that logical palette.When displaying clipboard data, the clipboard always uses as its current palette any object on the clipboard that is in the CF_PALETTE format.
								,CF_PENDATA : 10 ;Data for the pen extensions to the Microsoft Windows for Pen Computing.
								,CF_PRIVATEFIRST : 0x0200 ;Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message.
								,CF_PRIVATELAST : 0x02FF ;See CF_PRIVATEFIRST.
								,CF_RIFF : 11 ;Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
								,CF_SYLK : 4 ;Microsoft Symbolic Link (SYLK) format.
								,CF_TEXT : 1 ;Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
								,CF_TIFF : 6 ;Tagged-image file format.
								,CF_UNICODETEXT : 13 ;Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
								,CF_WAVE : 12 } ;Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.

	static                       WM_COPY := 0x301
								,WM_CLEAR := 0x0303
								,WM_CUT := 0x0300
								,WM_PASTE := 0x0302

	static skipFormats := {      2      : 0 ;"CF_BITMAP"
								,17     : 0 ;"CF_DIBV5"
								,0x0082 : 0 ;"CF_DSPBITMAP"
								,0x008E : 0 ;"CF_DSPENHMETAFILE"
								,0x0083 : 0 ;"CF_DSPMETAFILEPICT"
								,0x0081 : 0 ;"CF_DSPTEXT"
								,0x0080 : 0 ;"CF_OWNERDISPLAY"
								,3      : 0 ;"CF_METAFILEPICT"
								,7      : 0 ;"CF_OEMTEXT"
								,1      : 0 } ;"CF_TEXT"

	static formatByValue := {    2 : "CF_BITMAP"
								,8 : "CF_DIB"
								,17 : "CF_DIBV5"
								,5 : "CF_DIF"
								,0x0082 : "CF_DSPBITMAP"
								,0x008E : "CF_DSPENHMETAFILE"
								,0x0083 : "CF_DSPMETAFILEPICT"
								,0x0081 : "CF_DSPTEXT"
								,14 : "CF_ENHMETAFILE"
								,0x0300 : "CF_GDIOBJFIRST"
								,0x03FF : "CF_GDIOBJLAST"
								,15 : "CF_HDROP"
								,16 : "CF_LOCALE"
								,3 : "CF_METAFILEPICT"
								,7 : "CF_OEMTEXT"
								,0x0080 : "CF_OWNERDISPLAY"
								,9 : "CF_PALETTE"
								,10 : "CF_PENDATA"
								,0x0200 : "CF_PRIVATEFIRST"
								,0x02FF : "CF_PRIVATELAST"
								,11 : "CF_RIFF"
								,4 : "CF_SYLK"
								,1 : "CF_TEXT"
								,6 : "CF_TIFF"
								,13 : "CF_UNICODETEXT"
								,12 : "CF_WAVE" }
}
;{=====================================================================================
;} ====================================================================================

;======================URL Encode===================

URI_Encode(URI, RE="[0-9A-Za-z]"){
	VarSetCapacity(Var, StrPut(URI, "UTF-8"), 0), StrPut(URI, &Var, "UTF-8")
	While Code := NumGet(Var, A_Index - 1, "UChar")
		Res .= (Chr:=Chr(Code)) ~= RE ? Chr : Format("%{:02X}", Code)
	Return, Res
}

URI_Decode(URI){
	Pos := 1
	While Pos := RegExMatch(URI, "i)(%[\da-f]{2})+", Code, Pos)
	{
		VarSetCapacity(Var, StrLen(Code) // 3, 0), Code := SubStr(Code,2)
		Loop, Parse, Code, `%
			NumPut("0x" A_LoopField, Var, A_Index-1, "UChar")
		Decoded := StrGet(&Var, "UTF-8")
		URI := SubStr(URI, 1, Pos-1) . Decoded . SubStr(URI, Pos+StrLen(Code)+1)
		Pos += StrLen(Decoded)+1
	}
	return, URI
}

;----------------------------------

URI_URLEncode(URL){ ; keep ":/;?@,&=+$#."
	return URI_Encode(URL, "[0-9a-zA-Z:/;?@,&=+$#.]")
}

URI_URLDecode(URL){
	return URI_Decode(URL)
}

;Msgbox % URI_Encode("hi hello")
;Msgbox % URI_Decode("hi%20hello")

;~ ^j::
;~ ;Msgbox % URI_Encode("hi hello")
;~ Msgbox % URI_Encode("On British TV's ""Top of the Pops" "this Booker T. & the MGs hit might be titled ""Spring Onions")
;~ return

;=========================================Reload==================================

reloadAsAdmin(force:=True){
	if A_IsAdmin
		return 0
	Run % "*RunAs " ( A_IsCompiled ? "" : """"  A_AhkPath """" )  " """ A_ScriptFullpath """", %A_ScriptDir%, UseErrorLevel
	return _reloadAsAdmin_Error(e,force)
}

;http://ahkscript.org/boards/viewtopic.php?t=4334
reloadAsAdmin_Task(force:=True) { ;  By SKAN,  http://goo.gl/yG6A1F,  CD:19/Aug/2014 | MD:22/Aug/2014
; Asks for UAC only first time

  TASK_CREATE := 0x2,  TASK_LOGON_INTERACTIVE_TOKEN := 3
  e:=""
  Try TaskSchd := ComObjCreate( "Schedule.Service" ),    TaskSchd.Connect()
  Catch e
	  return _reloadAsAdmin_Error(e,force)

  CmdLine       := ( A_IsCompiled ? "" : """"  A_AhkPath """" )  A_Space  ( """" A_ScriptFullpath """"  )
  TaskName      := A_ScriptName " @" SubStr( "000000000"  DllCall( "NTDLL\RtlComputeCrc32"
				   , "Int",0, "WStr",CmdLine, "UInt",StrLen( CmdLine ) * 2, "UInt" ), -9 )

  Try {
	Try TaskRoot := TaskSchd.GetFolder("\AHK-ReloadAsAdmin")
	catch
		TaskRoot := TaskSchd.GetFolder("\"), TaskName:="[AHK-ReloadAsAdmin]" TaskName
	RunAsTask := TaskRoot.GetTask( TaskName )
  }
  TaskExists    := !A_LastError

  if !A_IsAdmin {
	if TaskExists {
		RunAsTask.Run("")
		ExitApp
	} else reloadAsAdmin(force)
  }

  else if !TaskExists {
	XML := "
	(LTrim Join
	  <?xml version=""1.0"" ?><Task xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task""><Regi
	  strationInfo /><Triggers><LogonTrigger><Enabled>true</Enabled><Delay>PT30S</Delay></LogonTrigger>
	  </Triggers><Principals><Principal id=""Author""><LogonType>InteractiveToken</LogonT
	  ype><RunLevel>HighestAvailable</RunLevel></Principal></Principals><Settings><MultipleInstancesPolic
	  y>Parallel</MultipleInstancesPolicy><DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries><
	  StopIfGoingOnBatteries>false</StopIfGoingOnBatteries><AllowHardTerminate>false</AllowHardTerminate>
	  <StartWhenAvailable>false</StartWhenAvailable><RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAva
	  ilable><IdleSettings><StopOnIdleEnd>true</StopOnIdleEnd><RestartOnIdle>false</RestartOnIdle></IdleS
	  ettings><AllowStartOnDemand>true</AllowStartOnDemand><Enabled>true</Enabled><Hidden>false</Hidden><
	  RunOnlyIfIdle>false</RunOnlyIfIdle><DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteApp
	  Session><UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine><WakeToRun>false</WakeToRun><
	  ExecutionTimeLimit>PT0S</ExecutionTimeLimit></Settings><Actions Context=""Author""><Exec>
	  <Command>"   (  A_IsCompiled ? A_ScriptFullpath : A_AhkPath )       "</Command>
	  <Arguments>" ( !A_IsCompiled ? """" A_ScriptFullpath  """" : "" )   "</Arguments>
	  <WorkingDirectory>" A_ScriptDir "</WorkingDirectory></Exec></Actions></Task>
	)"

	TaskRoot.RegisterTask( TaskName, XML, TASK_CREATE, "", "", TASK_LOGON_INTERACTIVE_TOKEN )

  }

  return TaskName
}

_reloadAsAdmin_Error(e,force){
	if force {
		MsgBox, 4112, FATAL ERROR!!, Couldn't restart script!!`nError Code: %e%
		ExitApp
	}
	return 1
}

use_TrayIcon(Script, Action) { ; use tray icon actions of a running AHK script
	static a := { Open: 65300, Help:    65301, Spy:   65302, Reload: 65303
				, Edit: 65304, Suspend: 65305, Pause: 65306, Exit:   65307 }
	DetectHiddenWindows, On
	PostMessage, 0x111, % a[Action],,, %Script% - AutoHotkey
}

;================================================================================

;=========================================Clip==================================
Clip(Text="", Reselect="")
{
	Static BackUpClip, Stored, LastClip
	If (A_ThisLabel = A_ThisFunc) {
		If (Clipboard == LastClip)
			Clipboard := BackUpClip
		BackUpClip := LastClip := Stored := ""
	} Else {
		If !Stored {
			Stored := True
			BackUpClip := ClipboardAll ; ClipboardAll must be on its own line
		} Else
			SetTimer, %A_ThisFunc%, Off
		LongCopy := A_TickCount, Clipboard := "", LongCopy -= A_TickCount ; LongCopy gauges the amount of time it takes to empty the clipboard which can predict how long the subsequent clipwait will need
		If (Text = "") {
			SendInput, ^c
			ClipWait, LongCopy ? 0.6 : 0.2, True
		} Else {
			Clipboard := LastClip := Text
			ClipWait, 10
			SendInput, ^v
		}
		SetTimer, %A_ThisFunc%, -700
		Sleep 20 ; Short sleep in case Clip() is followed by more keystrokes such as {Enter}
		If (Text = "")
			Return LastClip := Clipboard
		Else If ReSelect and ((ReSelect = True) or (StrLen(Text) < 3000))
			SendInput, % "{Shift Down}{Left " StrLen(StrReplace(Text, "`r")) "}{Shift Up}"
	}
	return
	Clip:
	return Clip()
}

getSelectedText()
{
/*
Returns selected text without disrupting the clipboard. However, if the clipboard contains a large amount of data, some of it may be lost
*/
	clipOld:=ClipboardAll
	Clipboard:=""
	Send, ^c
	ClipWait, 0.1, 1
	clipNew:=Clipboard
	Clipboard:=clipOld

	;Special for explorer
	WinGet, w, ID, A
	WinGetClass, c, ahk_id %w%
	if c in Progman,WorkerW,Explorer,CabinetWClass
		SplitPath, clipNew,,,, clipNew2

	return clipNew2?clipNew2:clipNew
}

;======================================delayedtimer.ahk==================================

class delayedTimer {
	set(f0,t0,runatStart:=False){
		if !isObject(this.obj)
			this.obj:=[]
		return this.obj.push({f:f0, t:t0, r:runatStart})
	}
	start(runNow:=False){
		for _,item in this.obj {
			if !item.t
				continue
			f:=item.f
			setTimer, % f, % item.t
		}
		if runNow
			this.firstRun()
	}
	firstRun() {
		this.running:=True
		f:=ObjBindMethod(this,"_firstRun")
		setTimer, % f, -100
	}
	_firstRun(){
		for _,item in this.obj
			if item.r
				item.f.call()
		this.running:=False
		return this.reset()
	}
	reset(){
		return this.obj:=[]
	}
}

;======================================ShellRun.ahk ===============================

;====================================== EXAMPLE ==================================================
;ShellRun("Taskmgr.exe")  ;Task manager
;ShellRun("Notepad", A_ScriptFullPath)  ;Open a file with notepad
;ShellRun("Notepad",,,"RunAs")  ;Open untitled notepad as administrator
;================================================================================================
/*
  ShellRun by Lexikos
	requires: AutoHotkey_L
	license: http://creativecommons.org/publicdomain/zero/1.0/

  Credit for explaining this method goes to BrandonLive:
  http://brandonlive.com/2008/04/27/getting-the-shell-to-run-an-application-for-you-part-2-how/

  Shell.ShellExecute(File [, Arguments, Directory, Operation, Show])
  http://msdn.microsoft.com/en-us/library/windows/desktop/gg537745

  Parameters for ShellRun
  1 application to launch
  2 command line parameters
  3 working directory for the new process
  4 "Verb" (For example, pass "RunAs" to run as administrator)
  5 Suggestion to the application about how to show its window - see the msdn link for possible values
  */

ShellRun(prms*)
{
	shellWindows := ComObjCreate("{9BA05972-F6A8-11CF-A442-00A0C90A8F39}")

	desktop := shellWindows.Item(ComObj(19, 8)) ; VT_UI4, SCW_DESKTOP

	; Retrieve top-level browser object.
	if ptlb := ComObjQuery(desktop
		, "{4C96BE40-915C-11CF-99D3-00AA004AE837}"  ; SID_STopLevelBrowser
		, "{000214E2-0000-0000-C000-000000000046}") ; IID_IShellBrowser
	{
		; IShellBrowser.QueryActiveShellView -> IShellView
		if DllCall(NumGet(NumGet(ptlb+0)+15*A_PtrSize), "ptr", ptlb, "ptr*", psv:=0) = 0
		{
			; Define IID_IDispatch.
			VarSetCapacity(IID_IDispatch, 16)
			NumPut(0x46000000000000C0, NumPut(0x20400, IID_IDispatch, "int64"), "int64")

			; IShellView.GetItemObject -> IDispatch (object which implements IShellFolderViewDual)
			DllCall(NumGet(NumGet(psv+0)+15*A_PtrSize), "ptr", psv
				, "uint", 0, "ptr", &IID_IDispatch, "ptr*", pdisp:=0)

			; Get Shell object.
			shell := ComObj(9,pdisp,1).Application

			; IShellDispatch2.ShellExecute
			shell.ShellExecute(prms*)

			ObjRelease(psv)
		}
		ObjRelease(ptlb)
	}
}

; ShellRunEdge(prms)
; msgbox % prms
; {
;     IApplicationActivationManager := ComObjCreate("{45BA127D-10A8-46EA-8AB7-56EA9078943C}", "{2e941141-7f97-4756-ba1d-9decde894a3d}")
;         DllCall(NumGet(NumGet(IApplicationActivationManager+0)+3*A_PtrSize), "Ptr", IApplicationActivationManager, "Str", "Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge", "Str", prms, "UInt", 0, "IntP", processId)
;       ObjRelease(IApplicationActivationManager)
; }

;======================================ShellRun.ahk ===============================

;====================================== EXAMPLE ===================================

; toast("Menu Tip:", "https://google.com", ,2000)

toast(t, msg, s:=14, l:=2000) {
		c:=0xf5f8fa, o:=Bold, f:=Segoe UI
		global GUI_handle:= "Toast_GUI" 1224
		Gui, %GUI_handle%: New, -Caption +ToolWindow +AlwaysOnTop +hwndHWND
				this.hwnd:=hwnd
		Gui, %GUI_handle%: Color, 0x222222
		Gui, %GUI_handle%: +LastFoundExist
		WinSet, Trans, 200
		pX := A_CaretX, pY:= A_CaretY
		if (pX = "" OR py = ""){
			pX := (A_ScreenWidth-680)
			pY := (A_ScreenHeight-160)
		}
		Gui, %GUI_handle%: Margin, 50, 20 
		Gui, %GUI_handle%: Font, Bold s%s% c%c% %o%, %f% 
		Gui, %GUI_handle%: Add, Text, +Center w600, %t% `n %msg%
		; Gui, %GUI_handle%: Font, norm s%s% c%c% %o%, %f%
		; Gui, %GUI_handle%: Add, Text, xp y+%m%, %msg%
		Gui, %GUI_handle%: Show, autosize x%pX% y%pY%, NoActivate
		setTimer, closeGUI, % "-"l
}

closeGUI:
		Gui, %GUI_handle%: Destroy
return

;}================================END OF FUNCTIONS=======================================

;{=================================Start Hotkeys=========================================

; Triggers for Menu
NumpadEnter::Send {NumpadEnter}

NumpadEnter & NumpadAdd::
	casemenu.Show()
return

*CapsLock::
keywait, Capslock, T0.2
if (ErrorLevel){
	^CapsLock::
	caseMenu.show()
	return
	}
Send, {CapsLock Up}
return

~>+CapsLock::
ScrollLock::
~Insert::
return

; Restart Script
^Esc::
	if (escPresses > 0){
		 escPresses += 1
		 return
	}
	escPresses := 1
	KeyWait, Esc, U
	time1 := A_TimeSinceThisHotkey
	SetTimer, escKey, -300 ; Wait for more presses within a 300 millisecond window.
return

escKey:
	if (escPresses = 1){
		if (time1>1000)
		use_TrayIcon(A_ScriptFullPath, "Open")
	} else if (escPresses = 2) {
	use_TrayIcon(A_ScriptFullPath, "Reload")
	}
	escPresses := 0
return

;;=======================================================================================

;;=======================================================================================
;;================================== Copy To ============================================
;;=======================================================================================
return
ExitApp
