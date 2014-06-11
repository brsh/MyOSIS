'==============================================================================================
' MyOSIS.vbs
' Gather tons of system information
'
' Date  		Who		What
' 2014/06/11	BDS		Created
'
'==============================================================================================
'PRESHOW
'This requires the declaration of variables (helps prevent typos)
Option Explicit

'Since we want this to be silent (no user sees it ever!), we "disable" all errors
'For the purposes of development, this is disabled
'On Error Resume Next

'These are constants for our sql connection later....
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
Const ScriptVer = "0.1"
'The database server for TESTING
Const DBServer = "."
Const DBFile = "MyOSIS"
'The database for PRODUCTION
'Const DBServer = "um... don't have one"
'Const DBFile = "nope"

'Ah, the declaration of our variables.
Dim ComputerName, ComputerDomain, UserName, DomainName, TimeZone, MachineModel, MachineSerial
Dim MACAddress, IPAddress, SessionName, ClientName, LogonDC, OpSys, LastBoot, FullUserName
Dim wShell, wshNetwork, objWMIService, colLection, objItem, objConn, objRS
Dim EnviroTypes, WshSysEnv, x, LogInOut, ExitCounter, DBFile, DBServer

'Some environment variables are wacky, 
'this array helps us cycle through the types to find the one we're looking for.
EnviroTypes = Array("Volatile", "Process", "System", "User")

'The wShell object gives us access to various pieces of system info....
'It also lets us write to the event Log
Set wShell = WScript.CreateObject("WScript.Shell")
WriteEventLog "Script Started. Processing Log" & LogInOut & "."

'The WSH Network object gives us access to names: user, domain, and computer
Err.Clear
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
If Err.Number <> 0 Then 
	WriteEventLog "An error getting the WScript.Network object: " & Err.Description
	Err.clear
End If

'Create the Base WMI Object
'WMI is more powerful then the Shell Object
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'==============================================================================================
'ENTR'ACTE
'Get Command Line arguments for log in or out (the function is at the end of the script).
LogInOut = ParseCLI()










'==============================================================================================
'CURTAIN CALL
'Clean up
WriteEventLog "Script Ended."

Set wShell = Nothing
Set wshNetwork = Nothing
Set objConn = Nothing
Set objRS = Nothing

'That's all folks...
WScript.Quit


'==============================================================================================
'ENCORE
'These are all the "handy modules" we call throughout the script

Sub WriteEventLog(sText)
	wShell.LogEvent 0, "MyOSIS.vbs: " & sText
End Sub

Function ParseCLI()
	'This reads throught the command line switches and acts accordingly
	Dim oArgs, i, sArgs, sT
	On Error Resume Next
	Set oArgs = WScript.Arguments ' create object with collection
	'then cycle through 'em
	For i = 0 to oArgs.Count - 1
		'The temp variable just saves me some typing
		sT = UCase(oArgs(i))
		sT = Replace(sT, "/", "")
		sT = Replace(sT, "-", "")
		Select Case Left(sT, 1)
			Case "I"
				sArgs = "IN"
			Case "O"
				sArgs = "OUT"
			Case Else
				'The off chance the script is just run w/out a valid switch
				sArgs = "NONE"
		End Select
	Next
	sArgs = Trim(sArgs)
	If Len(sArgs) = 0 Then sArgs = "NONE"
	ParseCLI = sArgs
End Function

Function ReadReg(sWhich)
	'This is a silly reg-reader for the name information (user, domain, and computer)
	'It's rarely accurate, but it's a last resort type thing
	'The User and Domain names tend to be blank....
	On Error Resume Next
	Dim sRet, sKey, sVal
	sRet = ""
	'This is the location for the User and Domain names, we rest it for Machine names
	sKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
	
	Select Case UCase(sWhich)
		'Here we specify the different values to find (and the different regpath for machine name)
		Case "NAME"
			sVal = "DefaultUserName"
		Case "SYSTEM"
			sKey = "HKLM\SYSTEM\CurrentControlSet\Services\TCPIP\Parameters\"
			sVal = "Hostname"
		Case "DOMAIN"
			sVal = "DefaultDomainName"
	End Select
	
	sRet = wShell.RegRead(sKey & sVal)
	ReadReg = sRet
End Function

Function TrimNoData(s)
	'Just a little function to clean up the strings a little
	'This strips any extra spaces and sets empty strings to "No Data"
	Dim sRet
	sRet = Trim(s)
	If Len(sRet) = 0 Then sRet = "No Data"
	TrimNoData = sRet
End Function

Function StringIsEmpty(s)
	'If the string is empty, this returns TRUE, otherwise FALSE
	StringIsEmpty = CBool(Len(Trim(s)) = 0)
End Function

Function EnVarCycle(s)
	'In order to search all the possible Environment Variables
	'We use this routine until we find the value specified.
	'EnviroTypes is defined earlier in the script, and includes all the types....
	Dim sRet
	sRet = ""
	For x = 0 To UBound(EnviroTypes)
		Set WshSysEnv = wShell.Environment(EnviroTypes(x))
		sRet = WshSysEnv(s)
		If sRet <> "" Then
			'We got what we came for...
			Exit For
		End If
	Next
	EnVarCycle = sRet
End Function

Function WMIDateStringToDate(dtmDate)
On Error Resume Next
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

Sub DoHelp
	Dim sB
	sB = ""
	sB = sB & string(76, "-") & VbCrLf
	sB = sB & "SpaceMailer" & VbCrLf
	sB = sB & string(76, "-") & VbCrLf
	sB = sB & "Gathers drive space information for all local hard drives and emails" & VbCrLf
	sB = sB & "that information in a report to admins." & VbCrLf
	sB = sB & VbCrLf
	sB = sB & "Usage: spacemailer [percent] [email1] [email2] [email...]" & VbCrLf
	sB = sB & VbCrLf
	sB = sB & "Options: "  & VbCrLf
	sB = sB & VbCrLf
	sB = sB & "percent    if freespace is below this percent, issues an alert (default=20)" & VbCrLf
	sB = sB & "email      email addresses for the report (separate by spaces; default=Is)" & VbCrLf
	sB = sB & VbCrLf
	sB = sB & VbCrLf
	
	WScript.Echo sB
	WScript.quit
End Sub

Sub OpenDB()
	Const adOpenStatic = 3
	Const adLockOptimistic = 3
	Const adUseClient = 3

	'The next 2 lines create the Connection and RecordSet objects 
	Set oConn = CreateObject("ADODB.Connection")
	Set oRS = CreateObject("ADODB.Recordset")

	'Just in case something failed earlier, we'll clear the error to be fresh
	Err.Clear

	'The next line opens the Connection to the specified server and database, using TCP/IP (DBMSS0CN)
	oConn.Open "DRIVER={SQL Server}; server=" & DBServer & "; database=" & DBFile & "; Network=DBMSSOCN; User id=LoginOutWriter; password=3v3nl3sss3cur3!"

	'If the Connection was successful, Then
	If Err.Number = 0 Then
		'Create a local recordset with fields based on the actual DB Table (much quicker for large DBs)
		oRS.CursorLocation = adUseClient
		oRS.Open "SELECT TOP 0 * FROM SystemSpace", oConn, adOpenStatic, adLockOptimistic
		If Err.Number <> 0 Then
			WriteEventLog "An Error occured selecting columns from the Database: " & Err.Description
			Err.Clear
		End If
	Else
		bToEmail = True
		WriteEventLog "An Error occured connecting to the Database: " & Err.Description
		Err.Clear
	End If
End Sub

Sub CloseDB
	oRS.Close
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing
End Sub 

Sub WriteDB(sSystem, sDrive, sFormat, sSize, sFree)
	'Now the database stuff.
	'Create a new row in the local recordset and populate our new info
' 	oRS.AddNew
' 	oRS("systemname") = sSystem
' 	oRS("driveletter") = Left(sDrive, 1)
' 	oRS("driveformat") = sFormat
' 	oRS("drivesizeb") = sSize
' 	oRS("drivefreeb") = sFree
' 	oRS("coldate") = Now()
' 	oRS("sversion") = ScriptVer
' 
' 	'Send our new row to the server and save it
' 	oRS.Update
	If Err.Number <> 0 Then
		'Write the error to the event Log
		WriteEventLog "An Error occured updating the Database: " & Err.Description
		Err.Clear
	End If
End Sub

Function PadRight(sOrig)
	'right justify a column of text so "123" becomese "   123"
	Dim sRet
	sRet = String(12 - Len(sOrig), " ") & sOrig
	PadRight = sRet
End Function

Function PadPercent(sOrig)
	'right justify a column of text so "123" becomese "   123"
	Dim sRet
	sRet = String(8 - Len(sOrig), " ") & sOrig
	PadPercent = sRet
End Function

Function PadLeft(sOrig)
	'left justify a column - padding the end with spaces
	Dim sRet
	If iWidth < Len(sOrig) Then iWidth = Len(sOrig) + 2
	sRet = sOrig & String(iWidth - Len(sOrig), " ")
	PadLeft = sRet
End Function

Sub Echo(sText)
	'Set up the output - either to the screen or elsewhere
	If bToScreen Then WScript.Echo sText
	msgBody = msgBody & sText & VbCrLf
End Sub

Function PercNum(sNumber)
	'Here, we format a 0.xxxxxx number into a percent
	Dim sRet
	sRet = FormatNumber(sNumber * 100, 1) & "%"
	PercNum = PadPercent(sRet)
End Function

Function GBNum(sNumber)
	'Here, we format an number to a drive size (so 123 becomes 124b for bytes; 27232130 becomes 27mb)
	Dim sRet, iLen
 	If Left(sNumber / 1024, 1) = "9" Then
 		iLen = Len(sNumber) - 1
 	Else
		iLen = Len(sNumber)
 	End If
	Select Case iLen
		Case 0, 1, 2, 3
			'bytes
			sRet = FormatNumber(sNumber, 2) & "B"
		Case 4, 5, 6
			'kilobytes
			sRet = FormatNumber(sNumber / 1024, 2) & "kb"
		Case 7, 8, 9
			'megabytes
			sRet = FormatNumber(sNumber / 1048576, 2) & "mb"
		Case 10, 11, 12
			'gigabytes
			sRet = FormatNumber(sNumber / 1073741824, 2) & "gb"
		Case 13, 14, 15
			'terabytes
			sRet = FormatNumber(sNumber / 1099511627776, 2) & "tb"
		Case 16, 17, 18
			'petabytes
			sRet = FormatNumber(sNumber / 1125899906842624, 2) & "pb"
		Case 19, 20, 21
			'exabytes
			sRet = FormatNumber(sNumber / 1.152921504606847e+18, 2) & "eb"
		Case 22, 23, 24
			'zettabytes
			sRet = FormatNumber(sNumber / 1.180591620717411e+21, 2) & "zb"
		Case 25, 26, 27
			'yottabytes
			sRet = FormatNumber(sNumber / 1.208925819614629e+24, 2) & "yb"		
		Case Else
			'bytes
			sRet = FormatNumber(sNumber, 2) & "B"

	End Select
	sRet = PadRight(sRet)
	GBNum = sRet
End Function

Function Ping(strComputer)
	Set png = wShell.exec("ping -a -n 1 " & strComputer)
	Do until png.status = 1 : wscript.sleep 10 : Loop
	strPing = png.stdout.readall

	'NOTE:  The string being looked for in the Instr is case sensitive.
	'Do not change the case of any character which appears on the
	'same line as a Case InStr.  As this will result in a failure.
	Select Case True
	Case InStr(strPing, "Request timed out") > 1
		strReply = "Request timed out"
		strCname = lcase(getcName(strPing))
		Ping = False
	Case InStr(strPing, "could not find host") > 1
		strReply = "Host not reachable"
		strCname = lcase(getcName(strPing))
		Ping = False
	Case InStr(strPing, "Reply from") > 1
		strReply = "Ping Succesful"
		strCname = lcase(getcName(strPing))
		Ping = True
	End Select
	If strCName = "" Then
		strCname = "N/A"
		'Ping = False
	End If
End Function