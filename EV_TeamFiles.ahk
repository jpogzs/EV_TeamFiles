/*
------------------------------------------
TeamFiles
Team Files Monitoring Script

F12 - Toggle Transparent
F11 - Toggle Always On Top
F2 to copy filename
Doubleclick files to open folder
Or right click for menu
Auto update is 5mins interval
Click Update for manual update

For code explanation contact James Pogi
Hope it helps! :)
------------------------------------------
*/

;pinalitan ko ng mga flags ng MsgBox para mas visible.

#NoEnv
#Warn
#MaxThreadsPerHotkey, 2
#SingleInstance Force

title = EV_TeamFiles
version = 1.0.0.3

;nilayan ko nito para mdaling iswitch ng enveronment. need lang din gumamit ng shortcut, immodify lang din yung target para gumana.

/*


DatabaseName = % A_Args[1]

if (A_Args[1] != "production")
    {
         if (A_Args[1] != "mainline")
            {
                MsgBox, 262160, %title%, Could not connect to the server at this time. `n`n(For further info email james.pogio@eagleview.com)
                ExitApp
            }
    }

;~ DatabaseName = mainline

*/
;~ 
DatabaseName = production

PCUserName = %A_UserName%

T := 0
A := 0

global Query := []

;binago ko yung string format para flexi.
AppPath := Format("\\mnl-fs01\PhillyTools\{1}\apps\{2}", DatabaseName, title)
libpath := A_ScriptDir
dbpath := Format("\\mnl-fs01\PhillyTools\{1}\database\reference", DatabaseName)
SQL := new SQLiteObj(libpath "\sqlite3.dll", dbpath Format("\{1}.s3db",DatabaseName))
if (SQL=0)
    return

VersionQuery := []
VersionQuery := SQL.Query("SELECT AppName,Revision,Status,DateRelease,ID,AppVersion FROM tblAppVersion WHERE AppName = """ title """ ")
AppName := VersionQuery[1]
Revision := VersionQuery[2]
AppStatus := VersionQuery[3]
DateRelease := VersionQuery[4]
AppID := VersionQuery[5]
AppVersion := VersionQuery[6] ;eto yung ginamit ko instead yung Revision kasi di ko ma palitan file version number pag nag compiled ng app. pwede naman kaso gagamit pa ng third-party tool hehe.

If (AppStatus!=1)
{
	MsgBox, 262192, %title%, Could not connect to the server at this time. `n Please try again later.`n`n(For further info email james.pogio@eagleview.com)
	ExitApp
}

If (version!=AppVersion)
{
	 ;inalis ko muna yun function para sa OK yung copy ng updated app kasi sa server pag .exe na sya di na ma-uupadate kasi app running na yung app. sama nlang natin sa deployment ng installer ng philltools.
	MsgBox, 262193, %title%, Please Update your Phillytools. Your version is outdated.`n(For further info email james.pogio@eagleview.com)
	IfMsgBox, Cancel 
		ExitApp
}

UserQuery := []
Cols := "tblUser.User_ID,tblUser.UserName,tblUser.TechNumber,tblUser.TeamNumber,tblUserRoleAccess.Uber,tblUserRoleAccess.Lead,tblUserRoleAccess.Admin,tblUserRoleAccess.SupOIC,tblUserRoleAccess.Coord,tblUserRoleAccess.SubTL" 
Table1 := "tblUser"
Table2 := "tblUserRoleAccess"
InnerJoin := "tblUser.User_ID = tblUserRoleAccess.RoleAccess_ID"
User := PCUserName
Condition := "Username = " """" User """"
UserQuery := SQL.Query("SELECT " Cols " FROM " Table1 " INNER JOIN " Table2 "  ON " InnerJoin "  WHERE " Condition " COLLATE NOCASE ")
UserID := UserQuery[1]
UserName := UserQuery[2]
TechNumber := UserQuery[3]
TeamNumber := UserQuery[4]
RoleAccessUber := UserQuery[5]
RoleAccessLead := UserQuery[6]
RoleAccessAdmin := UserQuery[7]
RoleAccessSupOIC := UserQuery[8]
RoleAccessCoord := UserQuery[9]
RoleAccessSubTL := UserQuery[10]

;Nag add ako role ko ng isa pang validition 
if (RoleAccessUber != 1 && RoleAccessLead != 1 && RoleAccessAdmin != 1 && RoleAccessSupOIC != 1 && RoleAccessCoord != 1 && RoleAccessSubTL != 1)
{
  MsgBox, 262192, %title%, You are not allowed to use this tool.
  ExitApp
}

RoleQuery := []
RoleQuery := SQL.Query("SELECT tblPrefRole.RoleName FROM tblUserRole INNER JOIN tblPrefRole ON tblUserRole.Role = tblPrefRole.Role_ID WHERE tblUserRole.Role_ID = " UserID " ")
Role := RoleQuery[1]

Menu, Tray, NoStandard
Menu, Tray, Tip, F12 - Toggle Transparent`nF11 - Toggle Always On Top`nF2 to copy filename`nDoubleclik files to open folder`nOr right click for menu`nAuto update is 5mins interval`nClick Update for manual update
Menu, Tray, Add, Exit, GuiClose


TSelect:

Gui, +AlwaysOnTop +ToolWindow
Gui -Resize -MinimizeBox
Gui, Color, 141414
Gui, Font, cGreen s6
Gui, Add, Text,, Version %version%
Gui, Font, cGreen s12, Courier
Gui, Add, Text, Section, User: %UserName%
Gui, Add, Text, y+3, Team: %TeamNumber%
Gui, Add, Text, y+3, Role: %Role%
Gui, Font, cGreen s12, Courier

TeamNumSelect := []
TeamLimit := 150
Loop, %TeamLimit%
	TeamNumSelect .= A_Index . "|"
Gui, Add, ComboBox, ys R20 w55 hwndhc vTeamNum Choose%TeamNumber%  +0x40, %TeamNumSelect%
SendMessage, 0xC5, 3, 0, Edit1, AHK_ID %hc%
if (RoleAccessUber!=1)
	GuiControl, Disable, TeamNum
Gui, Font, cGreen s10, Courier
Gui, Add, Button, gUpdate vButton w55, OK
Gui, Add, Text, xs y+1,
Gui, Show, x1500 y155 AutoSize, %title%
Return

Update:

Gui, Submit
If (TeamNum < 1 OR TeamNum > TeamLimit)
{
	Gui, Destroy
	Goto TSelect
}

WinGetPos, winX, winY,
Gui, Destroy

Gui, +AlwaysOnTop
Gui, Font, cGreen s10, Courier
Gui, Color, 141414
If (T = 1)
{
	Gui +LastFound
	Gui, -Caption +ToolWindow
}

WinSet, TransColor, 141414
Gui, Font, cGreen s20, Courier
Gui, Add, Text, vTeam cGreen, TEAM %TeamNum%
Gui, Font, cGreen s10, Courier
Gui, Add, Button, xm+220 yp w80 gRefresh vBupdate, &Update
TeamQuery := []
TeamQuery := SQL.Query("SELECT TeamName FROM tblPrefTeamName WHERE TeamNumber = " TeamNum " ")
TeamName := TeamQuery[1]
Gui, Add, Text,xs y+3 cGreen, Team %TeamName%
GuiControl, Hide, BUpdate
Gui, Add, Text, xs y+1, --------------------------------------
Gui, Font, cYellow s10, Courier
Gui, Add, Text, xs y+1 vTupdate, Loading...............................
Gui, Font, cGreen s10, Courier
Gui, Add, Text, xs y+1, --------------------------------------

Gui, Show, x%winX% y%winY% h160 NA, %title%

Gui Color, , 141414
Gui, Font, cGreen s10, Courier
Gui, Add, CheckBox, xm y+3 Section Checked1 vLiveCheck gLiveCheck, Live:
Gui, Add, Text, x+5 cGreen vVlivefiles, ...
Gui, Add, Text, xm+120 ys, New   : 
Gui, Add, Text, x+5 yp cGreen vVcountnw, ...
Gui, Add, Text, xm+225 ys, Tech : 
Gui, Add, Text, x+5 cGreen vVtechcount, ...
Gui, Add, CheckBox, xm y+3 Section Checked1 vP4Check gP4Check,     P4  :
Gui, Add, Text, x+5 cGreen vVp4, ...
Gui, Add, Text, xm+120 ys, InProg: 
Gui, Add, Text, x+5 yp cGreen vVcountip, ...
Gui, Add, Text, xm+225 ys, Files: 
Gui, Add, Text, x+5 cGreen vVfilecount, ...
Gui, Add, CheckBox, xm y+3 Section vTrainingCheck gTrainingCheck,  Test:
Gui, Add, Text, x+5 cGreen vVtrainingfiles, ...
Gui, Add, Text, xm+120 ys, Rework: 
Gui, Add, Text, x+5 yp cGreen vVcountrw, ...
Gui, Add, Text, xs y+1, --------------------------------------

Gui, Add, ListView,xs y+3 vLV gLV -ReadOnly -E0x200 w300 -Hdr,Filenumb|FT|ProdType|TechNo|Box|Path|Prio
LV_ModifyCol("Hdr")
LV_ModifyCol(1,"Hdr")
LV_ModifyCol(2,"Hdr")
LV_ModifyCol(3,"Hdr")
LV_ModifyCol(4,"Hdr")
LV_ModifyCol(5,"AutoHdr")
LV_ModifyCol(6,"0 Integer")
LV_ModifyCol(7,"0 Integer")
LV_ModifyCol("Left")

Menu, MyContextMenu, Add, Copy Filenumber, CopyFilenum
Menu, MyContextMenu, Add, Open Folder, OpenFolder
Menu, SubMenu, Add, Filenumber, FilenumSort
Menu, SubMenu, Add, Priority, PrioritySort
Menu, SubMenu, Add, Product, ProductSort
Menu, SubMenu, Add, Tech, TechSort
Menu, MyContextMenu, Add, Sort By, :Submenu

Refresh:
GuiControl,, Bupdate, Updating
countnw := 0
countip := 0
countrw := 0
totalfiles := 0
techcount := 0
techcountnw := 0
techcountip := 0
techcountrw := 0
hasfiles := 0
techfileproduct := 0
livefiles := 0
p4 := 0
trainingfiles := 0
fileprio := 0
ArrayList := []
listboxarray := []
ArrayTech := []
ArrayLink := []
ArrayBox := []
ArrayFilenum := []
ArrayFileType := []
ArrayProductType := []
ArrayFilePrio := []
ArrayLive := []
ArrayP4 := []
ArrayTraining := []

FileList := ""
SPath := "S:\1_InBox\Day Shift\*TCT"
Loop, Files, %SPath%%TeamNum%, D
{
	FileList .= A_LoopFileName "`n"
			Loop, Parse, Filelist, `n,
			if(A_LoopField > 0)
				tech = %A_LoopField%
		ParentFolder := A_LoopFileFullPath
		Loop, Files, %ParentFolder%\New\*.*
		{
			if (A_LoopFileExt = "evm")
			{
				++techcountnw
				++countnw
				Gosub FileData
				link = %ParentFolder%\New
				box := "New"
				Gosub FileDataDisplay
			}
		}
		Loop, Files, %ParentFolder%\Inprogress\*.*
		{
			if (A_LoopFileExt = "evm")
			{
				++techcountip
				++countip
				Gosub FileData
				link = %ParentFolder%\InProgress
				box := "InP"
				Gosub FileDataDisplay
			}
		}
		Loop, Files, %ParentFolder%\Rework\*.*
		{
			if (A_LoopFileExt = "evm")
			{
				++techcountrw
				++countrw
				Gosub FileData
				link = %ParentFolder%\Rework
				box := "Rwk"
				Gosub FileDataDisplay
			}
		}
		hasfiles := techcountnw+techcountip+techcountrw
		If(hasfiles>=1)
			++techcount
		techcountnw := 0
		techcountip := 0
		techcountrw := 0
		totalfiles := countnw+countip+countrw
}
totalfiles := countnw+countip+countrw

GuiControl, Text, Vlivefiles, %livefiles%
GuiControl,, Vp4, %p4%
GuiControl,, Vtrainingfiles, %trainingfiles%
GuiControl,, Vtechcount, %techcount%
GuiControl,, Vfilecount, %totalfiles%
GuiControl,, Vpriocount, %livefiles%
GuiControl,, Vcountnw, %countnw%
GuiControl,, Vcountip, %countip%
GuiControl,, Vcountrw, %countrw%
GuiControl, Font, Tupdate,
GuiControl,, Tupdate, Last update: %A_Hour%:%A_Min%:%A_Sec%
GuiControl,, Bupdate, Update
GuiControl, Show, BUpdate

Reload:
LV_Delete()
Gui, Submit, NoHide
GuiControlGet, LiveCheckX,, LiveCheck
GuiControlGet, P4CheckX,, P4Check
GuiControlGet, TrainingCheckX,, TrainingCheck

For index, element in ArrayTech
	{
		rowcount := 0
		listboxarray .= ArrayList[A_Index] "|"
		LV_Add("",ArrayFilenum[A_Index],ArrayFileType[A_Index],ArrayProductType[A_Index],ArrayTech[A_Index],ArrayBox[A_Index],ArrayLink[A_Index],ArrayFilePrio[A_Index])
		rowcount := LV_GetCount()
		LV_GetText(PrioText, rowcount, 7)
		If (LiveCheckX=0)
			{
			If (PrioText="Live")
				LV_Delete(rowcount)
			}
		If (P4CheckX=0)
			{
			If (PrioText="P4")
				LV_Delete(rowcount)
			}
		If (TrainingCheckX=0)
			{
			If (PrioText="Training")
				LV_Delete(rowcount)
			}
	}


LV_ModifyCol(2,"Sort")


rowcount := LV_GetCount()
excess := rowcount-5
if (rowcount > 35)
	GuiControl, Move, LV, % "h600" . "w300"
else
	GuiControl, Move, LV, % "h" . (85+(excess*17)) . "w300"
Gui, Show, AutoSize NA, %title%
SetTimer, Refresh, 300000
Return



LiveCheck:
Goto Reload

P4Check:
Goto Reload

TrainingCheck:
Goto Reload

FileData:
GetPos1 := InStr(A_LoopFileName, "_", false, 1, 1)+1		; 1st underscore to Right
GetPos2 := InStr(A_LoopFileName, "_", false, 1, 1)-1		; 1st underscore to Left
GetPos3 := InStr(A_LoopFileName, "_", false, 1, 2)+1		; 2nd underscore to Right
GetPos4 := InStr(A_LoopFileName, "_", false, 1, 3)-1		; 3rd underscore to Left
GetPos5 := GetPos4-GetPos3+1
GetPos6 := GetPos3-GetPos1-1
filename := SubStr(A_LoopFileName, 1, GetPos4)
filenum := SubStr(A_LoopFileName, GetPos1, GetPos6)
filetype := SubStr(A_LoopFileName, 1, GetPos2)
if (filetype=="P4")
{
	++p4
	fileprio := "P4"
}
livefileslist := "P1,P2,P3,1H,2H,3H"
if filetype in %livefileslist%
{
	++livefiles
	fileprio := "Live"
}
filetypelist := "P1,P2,P3,P4,1H,2H,3H"
if filetype not in %filetypelist%
{
	++trainingfiles
	fileprio := "Training"
}
producttype := SubStr(A_LoopFileName, GetPos3, GetPos5)
techname := SubStr(tech, 1, 6)
Return

FileDataDisplay:
ArrayTech.Push(techname)
ArrayLink.Push(link)
ArrayBox.Push(box)
ArrayFilenum.Push(filenum)
ArrayFileType.Push(filetype)
ArrayProductType.Push(producttype)
ArrayFilePrio.Push(fileprio)
if (livefiles>0)
{
	Gui, Font, cRed s10, Courier
	GuiControl, Font, LiveCheck
	GuiControl, Font, Vlivefiles
}
GuiControl, Text, Vlivefiles, %livefiles%
GuiControl,, Vp4, %p4%
GuiControl,, Vtrainingfiles, %trainingfiles%
GuiControl,, Vtechcount, %techcount%
GuiControl,, Vfilecount, %totalfiles%
GuiControl,, Vpriocount, %livefiles%
if (livefiles>0)
	GuiControl, Font, Vpriocount
Gui, Font, cGreen s10, Courier
GuiControl,, Vcountnw, %countnw%
GuiControl,, Vcountip, %countip%
GuiControl,, Vcountrw, %countrw%
Return

F12::
T := !T
If T
	{
	GuiControl,, Vtransparent, 1
	Gui, Color, 141414
	Gui +LastFound
	WinSet, TransColor, 141414
	Gui, -Caption +ToolWindow
	Gui, Show, AutoSize NA, %title%
	}
If not T
	{
		GuiControl,, Vtransparent, 0
		Gui, +Caption -ToolWindow
		Gui, Color, 141414
		Gui -LastFound
		WinSet, TransColor, 202020
		Gui, Show, AutoSize, %title%
	}
Return

F11::
A := !A
if A
	Gui, +AlwaysOnTop
if not A
	Gui, -AlwaysOnTop
Return

MenuHandler:
Return

GuiContextMenu:
if (A_GuiControl != "LV")
    return
Menu, MyContextMenu, Show, %A_GuiX%, %A_GuiY%
return

CopyFilenum:
clipboard := ""
FocusedRowNumber := LV_GetNext(0, "F")
if not FocusedRowNumber
    return
LV_GetText(copiedfilenum, FocusedRowNumber)
clipboard := copiedfilenum
Return

OpenFolder:
FocusedRowNumber := LV_GetNext(0, "F")
if not FocusedRowNumber
    return
LV_GetText(fpath, FocusedRowNumber,6)
Run, %fpath%
Return

LV:
FocusedRowNumber := LV_GetNext(0, "F")
if not FocusedRowNumber
    return
LV_GetText(fpath, FocusedRowNumber,6)
Run, %fpath%
Return

Sort:
Return

FilenumSort:
LV_ModifyCol(1,"Sort")
Return

PrioritySort:
LV_ModifyCol(2,"Sort")
Return

ProductSort:
LV_ModifyCol(3,"Sort")
Return

TechSort:
LV_ModifyCol(4,"Sort")
Return

GuiClose:
ExitApp



class SQLiteObj
{
    
    __new(Path_SQLDLL, Path_DB)
	{
        if FileExist(Path_SQLDLL)=""
		{ 
			MsgBox, 262192, EV_TeamFiles, sqlite3.dll not found!`n`n(For further info email james.pogio@eagleview.com)
			ExitApp
            ;~ return 0
        }
        This.Library := DllCall("LoadLibrary", "Str", Path_SQLDLL, "Ptr")
        This.Callback := RegisterCallback("QueryCallback", "F C", 4)
        This.Open(Path_DB)
    }

    Open(Path_DB)
	{
        DB := 0
		SQLITE_OPEN_READONLY := 0x00000001 ;ginawa ko pala itong read only mode since mag read lang nman talaga yun need nya.
        VarSetCapacity(Address, StrPut(Path_DB, "UTF-8"), 0)
        StrPut(Path_DB, &Address, "UTF-8")
        DllCall("sqlite3.dll\sqlite3_open_v2" ; https://www.sqlite.org/c3ref/open.html
                , "Ptr", &Address
                , "PtrP", DB
                , "Int", SQLITE_OPEN_READONLY ;change from "2" value
                , "Ptr", 0
                , "CDecl Int")
        This.Database := DB
    }

    Query(Q)
	{
        Query := []
        VarSetCapacity(Address, StrPut(Q, "UTF-8"), 0)
        StrPut(Q, &Address, "UTF-8")        
        DllCall("sqlite3.dll\sqlite3_exec" ; https://www.sqlite.org/c3ref/exec.html
                , "Ptr", This.Database
                , "Ptr", &Address
                , "Ptr", This.Callback)
        return Query
    }
}

QueryCallback(ParamFromCaller, Columns, Values, Names)
{
    Loop %Columns%
        Query.Push(StrGet(NumGet(Values+0, (A_Index-1)*8, "uInt"), "UTF-8"))
    return 0
}
