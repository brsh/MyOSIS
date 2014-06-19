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
Const ScriptVer = "0.1"

'Since we want this to be silent (no user sees it ever!), we "disable" all errors
'For the purposes of development, this is disabled
'On Error Resume Next

'These are constants for our sql connection later....
'The database server for TESTING
Const DBServer = "."
Const DBFile = "MyOSIS"
'The database for PRODUCTION
'Const DBServer = "um... don't have one"
'Const DBFile = "nope"

'Ah, the declaration of our variables.
Dim oSystem, cNet(), oDNSEntries
Set oSystem = New System
Set oDNSEntries = New DNSEntry


'Dim sComputerName, sComputerDomain, sUserName, sDomainName, sTimeZone, sMachineModel, sMachineSerial
Dim sMACAddress, sIPAddress, sSessionName, sClientName, sLogonDC, sOpSys, sLastBoot, sFullUserName
Dim wShell, wNetwork, oWMIService, colLection, oItem, oConn, oRS
Dim aEnviroTypes, wSysEnv, x, iExitCounter

'Some environment variables are wacky, 
'this array helps us cycle through the types to find the one we're looking for.
aEnviroTypes = Array("Volatile", "Process", "System", "User")

'The wShell object gives us access to various pieces of system info....
'It also lets us write to the event Log
Set wShell = WScript.CreateObject("WScript.Shell")
WriteEventLog "Script Started. Processing Log."

'The WSH Network object gives us access to names: user, domain, and computer
Err.Clear
Set wNetwork = WScript.CreateObject( "WScript.Network" )
If Err.Number <> 0 Then 
	WriteEventLog "An error getting the WScript.Network object: " & Err.Description
	Err.clear
End If

'Create the Base WMI Object
'WMI is more powerful then the Shell Object
Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'==============================================================================================
'ENTR'ACTE
'Get Command Line arguments
ParseCLI

Echo oSystem.Name
SearchDNS oSystem.Name & ".", false

GetNetworkInfo

Dim i
For i = 1 To oDNSEntries.UIndex
	Echo oDNSEntries.Record(i) & " / " & oDNSEntries.RecordType(i) & " / " & oDNSEntries.Address(i)
Next


'==============================================================================================
'CURTAIN CALL
'Clean up
WriteEventLog "Script Ended."

Set oSystem = Nothing
Set wShell = Nothing
Set wNetwork = Nothing
Set oConn = Nothing
Set oRS = Nothing
Set oWMIService = Nothing

'That's all folks...
WScript.Quit


'==============================================================================================
'ENCORE
'These are all the "handy modules" we call throughout the script

Sub SearchDNS(sSearch, bExact)
	Dim aOut, i, j, sZone, aCnames, sRecord, sType, sAddress, aAll, sAll
	aOut = RunIt("dnscmd SFO-DC-01 /enumzones /primary /secondary /forward")

	For i = 0 To UBound(aOut)
		If InStr(aOut(i), "Primary") then
			sZone = TrimNoData(Split(Trim(aOut(i)), " ")(0))
			
			aCnames = RunIt("dnscmd SFO-DC-01 /enumrecords " & sZone & " @ /type a,cname | find /I " & Chr(34) & sSearch & Chr(34) )
			For j = 0 To Ubound(aCnames)
				sAll = Replace(aCnames(j), " ", vbTab)
				sAll = Replace(sAll, vbTab & vbTab, vbTab)
				aAll = Split(sAll & vbtab, vbtab)
 				sRecord = TrimNoData(aAll(0))
 				if len(sRecord) > 0 Then 
	 				sType = TrimNoData(aAll(ubound(aAll) - 2))
 					sAddress = TrimNoData(aAll(ubound(aAll) - 1))
 					If bExact Then 
 						If sAddress = sSearch then
 							oDNSEntries.Add sRecord & "." & sZone, sType, sAddress
 						End If
 					Else 
						oDNSEntries.Add sRecord & "." & sZone, sType, sAddress
 					End if
 				End If
			Next
		End If
	Next
End Sub

Sub GetNetworkInfo
Dim p_colItems, i, x, sResult, p_oItem, p_sIPAddress, p_sCaption, p_sIPSubnet, p_sDNSServerSearchOrder
Set p_colItems = oWMIService.ExecQuery ("Select * From Win32_NetworkAdapter") 'WHERE NetConnectionStatus=2
i = 0
For Each p_oItem in p_colItems
	If p_oItem.MACAddress <> vbNull Then 
   		ReDim Preserve cNet(i)
   		Set cNet(i) = New Network
		cNet(i).MACAddress = p_oItem.MACAddress & ""
		cNet(i).AdapterType = p_oItem.AdapterType
		cNet(i).Caption = p_oItem.Caption
		cNet(i).Description = p_oItem.Description
		cNet(i).Manufacturer = p_oItem.Manufacturer
		cNet(i).Name = p_oItem.Name
		cNet(i).NetConnectionID = p_oItem.NetConnectionID
		cNet(i).NetConnectionStatus = p_oItem.NetConnectionStatus
		cNet(i).NetEnabled = p_oItem.NetEnabled
		cNet(i).PhysicalAdapter = p_oItem.PhysicalAdapter
		cNet(i).ProductName = p_oItem.ProductName
		cNet(i).Speed = p_oItem.Speed
		i = i + 1
	End If
Next

Set p_oItem = Nothing
Set p_colItems = Nothing

' Define query to get information - IPEnabled restricts the information to active Adaptors
Set p_colItems = oWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = TRUE")

For x = 0 to UBound(cNet)
' Get each adaptor from the table

For Each p_oItem In p_colItems
	i = 0
	If p_oItem.MACAddress = cNet(x).MACAddress Then
		' Get each IP address for the adaptor
		For Each p_sIPAddress In p_oItem.IPAddress
			' check to see if it is not a 0 ip address
			If p_sIPAddress = "0.0.0.0" Then
				'Don't do anything
			Else
				'skip if an IPv6 address
				If InStr(p_sIPAddress, "::") = 0 Then
					' Set up the correct adaptor name by stringing the first 12 characters and also the MAC address
					p_sCaption = p_oItem.Caption
					p_sCaption = Right(p_sCaption, Len(p_sCaption) - 11)
					' Format DHCP info if required
					'If p_oItem.DHCPEnabled and blnShowDHCP Then
						'If blnShowDHCPExpire Then
							'strDHCP = " (Expires: " & fnDisplayDate(objItem.DHCPLeaseExpires) & ")"
						'End If
					'Else
					'	strDHCP = ""
					'End If
					p_sIPSubnet = p_oItem.IPSubnet(i)
					i = i + 1
					'p_IP, p_Mask, p_IsDHCP, p_DHCPServer, p_DHCPExpire, p_DHCPSet
					cNet(x).AddIP p_sIPAddress, p_sIPSubnet, p_oItem.DHCPEnabled, p_oItem.DhcpServer, p_oItem.DHCPLeaseExpires, p_oItem.DHCPLeaseObtained
					
					If Left(cNet(x).IPFormated, 3) = "   " Then
						cNet(x).IPFormated = cNet(x).IPFormated & vbTab & vbTab & vbTab & vbTab & p_sIPAddress + p_sIPSubnet & vbCrLf
					Else
						cNet(x).IPFormated = "       IP Address:" & vbTab & p_sIPAddress + p_sIPSubnet & vbCrLf
					End If
					If Not IsNull(p_oItem.DHCPServer) Then
						cNet(x).IPFormated = cNet(x).IPFormated & vbTab & "       DHCP Server:" & vbTab & p_oItem.DHCPServer  & vbCrLf
					End If
					If Not IsNull(p_oItem.DefaultIPGateway) Then
						cNet(x).IPFormated = cNet(x).IPFormated & vbTab & "       Gateway:" & vbTab & vbTab & Join(p_oItem.DefaultIPGateway, ", ") & vbCrLf
					End If
	
					If Not IsNull(p_oItem.DNSServerSearchOrder) Then 
						p_sDNSServerSearchOrder = Join(p_oItem.DNSServerSearchOrder, ", ")
						cNet(x).IPFormated = cNet(x).IPFormated & vbTab & "       DNS Servers:" & vbTab & p_sDNSServerSearchOrder & vbCrLf
					End If
					p_sDNSServerSearchOrder = ""
				End If
			End If
		Next
	End If
Next
Next
 	Dim hold, masks, dhcpYes, dhcpServer, dhcpSet, dhcpExpire
' 
 For x = 0 To Ubound(cNet)
' Echo Ubound(cNet(x).IPAddresses)
 	hold = cNet(x).IPAddresses
' 	masks = cNet(x).MaskOctet
' 	dhcpYes = cNet(x).IsDHCPed
' 	dhcpServer = cNet(x).DHCPServerAddress
' 	dhcpSet = cNet(x).DHCPTimeSet
' 	dhcpExpire = cNet(x).DHCPTimeExpires
 	For i = 1 To Ubound(hold) 
 		SearchDNS hold(i), True
' 		Echo hold(i)
' 		Echo masks(i)
' 		Echo "DHCP: " & dhcpYes(i)
' 		If dhcpYes(i) Then
' 			Echo "DHCP Server: " & dhcpServer(i)
' 			Echo "DHCP Set: " & dhcpSet(i)
' 			Echo "DHCP Exp: " & dhcpExpire(i)
' 		End If
 	Next
 Next


sResult = ""

For x = 0 to UBound(cNet)
	If cNet(x).NetConnectionStatus <> "Unknown" Then
		If Not Instr(UCase(cNet(x).Name), "BLUETOOTH") > 0 then
			sResult = sResult & vbTab & cNet(x).Name & vbCrLf
			sResult = sResult & vbTab & cNet(x).MacAddress & " | " & cNet(x).NetEnabled & " | " & cNet(x).NetConnectionStatus & " | " & cNet(x).Speed & vbCrLf
			sResult = sResult & vbTab & cNet(x).IPFormated
			sresult = sResult & vbCrLf
		End If 
	End If
Next

Echo sResult

End Sub

Function RunIt(ByVal sCommandLine)
	Dim pRun, sOut, maOut, iTimeOut, sThrowAway, sTrash
	sThrowAway = ""
	iTimeOut = 0
'Echo sCommandLine
	Set pRun = wShell.exec("%comspec% /c " & sCommandLine)
	Do Until (pRun.status <> 0) Or (iTimeOut = 3000) 
		Do While Not pRun.StdOut.AtEndOfStream
			sThrowAway = sThrowAway & vbCrLf & pRun.StdOut.ReadLine
			WScript.Sleep 1
		Loop
		iTimeOut = iTimeout + 1 
	Loop
	If pRun.Status = 0 Then pRun.Terminate 
	sOut = sThrowAway
	maOut = Split(sOut, vbCrLf)
	
	RunIt = maOut

End Function

Sub WriteEventLog(sText)
	wShell.LogEvent 0, "MyOSIS.vbs: " & sText
End Sub

Sub ParseCLI()
	'This reads throught the command line switches and acts accordingly
	Dim oArgs, i, sA
	Set oArgs = WScript.Arguments
	For i = 0 to oArgs.Count - 1
		sA = oArgs(i)
		'Let's see if this is a number
		If IsNumeric(sA) Then 
			'aha, if it's between 0 and 100 then it's our new Alert Threshold
			If (sA < 100) And (sA > 0) Then AlertThresholdPercent = sA
		End If
		
		Select Case Left(UCase(sA), 2)
			Case "/?", "-?", "/H", "-H", "-HELP"
				DoHelp
			Case "/s", "/S", "-s", "-S"
				'Send the output to the screen, rather than an email
				bToScreen = True
				bToEmail = False
			Case "/d", "/D", "-d", "-D"
				bDebug = True
				'The database server for TESTING
		End Select
	Next
End Sub

Function TrimNoData(s)
	'Just a little function to clean up the strings a little
	'This strips any extra spaces and sets empty strings to "No Data"
	Dim sRet, iChar
		s = Trim(s)
        ' remove all non-printable characters
		s = Replace(s, vbCrLf, "")

		s = Replace(s, vbTab, "")

        For iChar = 0 To 31
            While InStr(s, Chr(iChar)) > 0
                s = Replace(s, Chr(iChar), "")
            Wend
        Next

        For iChar = 127 To 255
            While InStr(s, Chr(iChar)) > 0
                s = Replace(s, Chr(iChar), "")
            Wend
        Next
	If Len(s) = 0 Then s = ""
	TrimNoData = s
End Function

Function StringIsEmpty(s)
	'If the string is empty, this returns TRUE, otherwise FALSE
	StringIsEmpty = CBool(Len(Trim(s)) = 0)
End Function

Function fnSubstring(p_strData,p_intStart,p_intLength )
   Dim intLen
   intLen = Len(p_strdata)

   If p_intStart < 1 Or p_intStart > intLen Then
      fnSubstring = ""
   Else
      If p_intLength > intLen - p_intStart + 1 Then
         p_intLength = intLen - p_intStart + 1
      End If
      fnSubstring = Right(Left(p_strData, p_intStart + p_intLength - 1), p_intLength)
   End If 
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
'On Error Resume Next
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

Sub DoHelp
	Dim sB
	sB = ""
	sB = sB & string(74, "-") & VbCrLf
	sB = sB & "MyOSIS" & VbCrLf
	sB = sB & string(74, "-") & VbCrLf
	sB = sB & "Gathers boatloads of system information and outputs the results" & vbCrLf 
	sB = sB & "to text, html, or SQL db." & vbCrLf
	sB = sB & VbCrLf
	sB = sB & "Usage: " & vbCrLf 
	sB = sB & "  myosis.vbs [ -options ] [ host ]" & VbCrLf
	sB = sB & vbCrLf
	sB = sB & "Options: "  & VbCrLf
	sB = sB & vbCrLf
	sB = sB & " host              system to query (local system is default)" & vbCrLf
	sB = sB & vbCrLf
	sB = sB & " Output type (select ONE; -screen is default)" & vbCrLf
	sB = sB & "   -d | -db        save to database" & vbCrLf
	sB = sB & "   -t | -txt       save to text" & VbCrLf
	sB = sB & "   -w | -web       save to html" & VbCrLf
	sB = sB & "   -s | -screen    output to screen (default)" & VbCrLf
	sB = sB & "   -q | -query     save SQL update statement" & VbCrLf
	sB = sB & vbCrLf
	sB = sB & "Examples:" & vbCrLf
	sB = sB & vbCrLf
	sB = sB & " myosis.vbs -t:file.txt calcium" & vbCrLf
	sB = sB & " myosis.vbs -w:""Note the quotes for spaces.html"""  & vbCrLf
	sB = sB & VbCrLf
	sB = sB & VbCrLf
	
	WScript.Echo sB
	WScript.quit
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
	WScript.Echo sText
End Sub

Function PercNum(sNumber)
	'Here, we format a 0.xxxxxx number into a percent
	Dim sRet
	sRet = FormatNumber(sNumber * 100, 1) & "%"
	PercNum = PadPercent(sRet)
End Function

Function GBNum(sNumber)
	'Here, we format numbers to drive sizes (so 123 becomes 123b for bytes; 27232130 becomes 27mb)
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
			sRet = FormatNumber(sNumber / 1024, 2) & "KB"
		Case 7, 8, 9
			'megabytes
			sRet = FormatNumber(sNumber / 1048576, 2) & "MB"
		Case 10, 11, 12
			'gigabytes
			sRet = FormatNumber(sNumber / 1073741824, 2) & "GB"
		Case 13, 14, 15
			'terabytes
			sRet = FormatNumber(sNumber / 1099511627776, 2) & "TB"
		Case 16, 17, 18
			'petabytes
			sRet = FormatNumber(sNumber / 1125899906842624, 2) & "PB"
		Case 19, 20, 21
			'exabytes
			sRet = FormatNumber(sNumber / 1.152921504606847e+18, 2) & "EB"
		Case 22, 23, 24
			'zettabytes
			sRet = FormatNumber(sNumber / 1.180591620717411e+21, 2) & "ZB"
		Case 25, 26, 27
			'yottabytes (really? I'm gonna see this in my life??)
			sRet = FormatNumber(sNumber / 1.208925819614629e+24, 2) & "YB"		
		Case Else
			'bytes
			sRet = FormatNumber(sNumber, 2) & "B"
	End Select
	sRet = PadRight(sRet)
	GBNum = sRet
End Function

'*****************************************************************
'Classes
'You're like an out of work school teacher - no class.
Class System
	'This Class creates a computer object so we don't need multiple arrays
	Public ID					'Just a handle for the object
	Public Domain				'The domain
	Public CommentLocal			'
	Public CommentAD			'
	Public TimeLastBoot			'
	Public Administrators		'
	Public LocalUserProfiles	'
	Public EnvironmentVariables	'
	
	Public Property Get DN
		' Use the NameTranslate object to convert the NT name of the computer To
		' the Distinguished name required for the LDAP provider. Computer names
		' must end with "$". Returns comma delimited string to calling code.
		Dim oTrans, oDomain, sResponse
		' Constants for the NameTranslate object.
		Const ADS_NAME_INITTYPE_GC = 3
		Const ADS_NAME_TYPE_NT4 = 3
		Const ADS_NAME_TYPE_1779 = 1
		Set oTrans = CreateObject("NameTranslate")
		Set oDomain = getObject("LDAP://rootDse")
		oTrans.Init ADS_NAME_INITTYPE_GC, ""
		oTrans.Set ADS_NAME_TYPE_NT4, wNetwork.UserDomain & "\" & Me.Name & "$"
		sResponse = oTrans.Get(ADS_NAME_TYPE_1779)
		'Set DN To upper Case
		DN = UCase(sResponse)
	End Property
	
	Public Property Get Name
		Name = wNetwork.ComputerName
	End Property
	

End Class

Class OpSys
	Public Name
	Public Edition
	Public ServicePack
	Public Architecture
	Public SystemType
	Public Features
	Public Roles
	Public TimeInstall
	

End Class

Class Memory
	Public Total
	Public Free
	Public PageFileSize
	Public PageFileType
End Class

Class Hardware
	Public Make
	Public Model
	Public SerialNumber
	Public BIOSVersion
End Class

Class DiskDrive
	'Array
	Public Mount
	Public Label
	Public SpaceTotal
	Public SpaceFree
	Public Count
End Class

Class Network
	'This class creates a printer object so we don't need multiple arrays
Public AdapterType
Public Caption
Public Description
Public MACAddress
Public Manufacturer
Public Name
Public NetConnectionID
Private internalConnectionStatus

Public Property Get NetConnectionStatus
	NetConnectionStatus = internalConnectionStatus
End Property
Public Property Let NetConnectionStatus(ByVal statIn)
	Select Case statIn
		Case 0 
			internalConnectionStatus = "Disconnected"
		Case 1
			internalConnectionStatus = "Connecting"
		Case 2 
			internalConnectionStatus = "Connected"
		Case 3 
			internalConnectionStatus = "Disconnecting"
		Case 4 
			internalConnectionStatus = "Hardware not present"
		Case 5 
			internalConnectionStatus = "Hardware disabled"
		Case 6 
			internalConnectionStatus = "Hardware malfunction"
		Case 7 
			internalConnectionStatus = "Media disconnected"
		Case 8 
			internalConnectionStatus = "Authenticating"
		Case 9 
			internalConnectionStatus = "Authentication succeeded"
		Case 10 
			internalConnectionStatus = "Authentication failed"
		Case 11
			internalConnectionStatus = "Invalid address"
		Case 12
			internalConnectionStatus = "Credentials required"
		Case Else
			internalConnectionStatus = "Unknown"
	End Select
End Property
Private Internal_NetEnabled
Public Property Get NetEnabled
	NetEnabled = Internal_NetEnabled
End Property
Public Property Let NetEnabled (ByVal statIn)
	
	If statIn Then
		Internal_NetEnabled = "Enabled"
	ElseIf Not statIn Then
		Internal_NetEnabled = "Disabled"
	Else
		Internal_NetEnabled = "Unknown Status"
	End If
End Property

Public NetworkAddresses
Private internal_PhysicalAdapter
Public Property Get PhysicalAdapter
	PhysicalAdapter = internal_PhysicalAdapter
End Property
Public Property Let PhysicalAdapter(ByVal statIn)
	If statIn Then
		internal_PhysicalAdapter = "Physical Adapter"
	ElseIf Not statIn Then
		internal_PhysicalAdapter = "Virtual Adapter"
	Else
		internal_PhysicalAdapter = "Unknown Adapter"
	End If
End Property
Public ProductName
Private internal_Speed
Public Property Get Speed
	Speed = internal_Speed
End Property
Public Property Let Speed(ByVal statIn)
	If IsNumeric(statIn) Then
		Select Case Len(statIn)
			Case 0, 1, 2, 3
				internal_Speed = statIn & "bps / " & statIn/8 & "Bps"
			Case 4, 5, 6
				internal_Speed = statIn/1000 & "kbps / " & (statIn/8)/1000 & "KBps"
			Case 7, 8, 9
				internal_Speed = statIn/1000000 & "mbps / " & (statIn/8)/1000000 & "MBps"
			Case 10, 11, 12
				internal_Speed = statIn/1000000000 & "gbps / " & (statIn/8)/1000000000 & "GBps"
			Case 13, 14, 15
				internal_Speed = statIn/1000000000000 & "gbps / " & (statIn/8)/1000000000000 & "GBps"
			Case 16, 17, 18
				internal_Speed = Round(statIn/1000000000000000, 2) & "gbps / " & Round((statIn/8)/1000000000000000, 2) & "GBps"
			Case 19, 20, 21
				internal_Speed = Round(statIn/1000000000000000000, 2) & "gbps / " & Round((statIn/8)/1000000000000000000, 2) & "GBps"
			Case Else
				internal_Speed = statIn & "bps / " & statIn/8 & "Bps"
		End Select
	Else
		internal_Speed = "Unknown bps"
	End If
End Property
Public Gateway
Public DNS
Public IPFormated

	Private m_IP()
	Private m_Mask()
	Private m_IsDHCP()
	Private m_DHCPTimeExpire()
	Private m_DHCPTimeSet()
	Private m_DHCPServer()

	Private Function FormatDHCPTime(p_strDate)
		Dim strYear, strMonth, strDay, strHour, strMinute, strSecond
		strYear =   fnSubstring(p_strDate,1,4)
		strMonth =  fnSubstring(p_strDate,5,2)   
		strDay =    fnSubstring(p_strDate,7,2)   
		strHour =   fnSubstring(p_strDate,9,2)   
		strMinute = fnSubstring(p_strDate,11,2)   
		strSecond = fnSubstring(p_strDate,13,2)   
		FormatDHCPTime = cdate(strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond)
	End Function

	Private Sub Class_Initialize
		ReDim m_IP(0)
		ReDim m_Mask(0)
		ReDim m_IsDHCP(0)
		ReDim m_DHCPTimeExpire(0)
		ReDim m_DHCPTimeSet(0)
		ReDim m_DHCPServer(0)
	End Sub

	Public Sub AddIP(p_IP, p_Mask, p_IsDHCP, p_DHCPServer, p_DHCPExpire, p_DHCPSet)
		Dim sp_Count
		sp_Count = Me.UIndex + 1
		ReDim Preserve m_IP(sp_Count)
		ReDim Preserve m_Mask(sp_Count)
		ReDim Preserve m_DHCPServer(sp_Count)
		ReDim Preserve m_IsDHCP(sp_Count)
		ReDim Preserve m_DHCPTimeExpire(sp_Count)
		ReDim Preserve m_DHCPTimeSet(sp_Count)

		m_IP(sp_Count) = p_IP
		m_Mask(sp_Count) = p_Mask
		If p_IsDHCP Then
			m_DHCPTimeSet(sp_Count) = FormatDHCPTime(p_DHCPSet)
			m_DHCPTimeExpire(sp_Count) = FormatDHCPTime(p_DHCPExpire)
			m_IsDHCP(sp_Count) = True
			m_DHCPServer(sp_Count) = p_DHCPServer
		Else
			m_DHCPTimeSet(sp_Count) = ""
			m_DHCPTimeExpire(sp_Count) = ""
			m_IsDHCP(sp_Count) = False
			m_DHCPServer(sp_Count) = ""
		End If
	End Sub
	
	Public Property Get UIndex
		UIndex = UBound(m_IP)
	End Property

	Public Property Get LIndex
		UIndex = LBound(m_IP)
	End Property

	Public Property Get IPAddresses
		IPAddresses = m_IP
	End Property
	
	Public Property Get DHCPTimeSet
		DHCPTimeSet = m_DHCPTimeSet
	End Property
	Public Property Get IsDHCPed
		IsDHCPed = m_IsDHCP
	End Property
	Public Property Get DHCPServerAddress
		DHCPServerAddress = m_DHCPServer
	End Property
	Public Property Get DHCPTimeExpires
		DHCPTimeExpires = m_DHCPTimeExpire
	End Property
	
	Public Property Get MaskDecimal
		'Yeah, this doesn't work yet :)
		Dim p_RetVal
		Select Case m_Mask
			Case "255.255.255.255"
				p_RetVal = "/32"
			Case "255.255.255.254"
				p_RetVal = "/31"
			Case "255.255.255.252"
				p_RetVal = "/30"
			Case "255.255.255.248"
				p_RetVal = "/29"
			Case "255.255.255.240"
				p_RetVal = "/28"
			Case "255.255.255.224"
				p_RetVal = "/27"
			Case "255.255.255.192"
				p_RetVal = "/26"
			Case "255.255.255.128"
				p_RetVal = "/25"
			Case "255.255.255.0"
				p_RetVal = "/24"
			Case "255.255.254.0"
				p_RetVal = "/23"
			Case "255.255.252.0"
				p_RetVal = "/22"
			Case "255.255.248.0"
				p_RetVal = "/21"
			Case "255.255.240.0"
				p_RetVal = "/20"
			Case "255.255.224.0"
				p_RetVal = "/19"
			Case "255.255.192.0"
				p_RetVal = "/18"
			Case "255.255.128.0"
				p_RetVal = "/17"
			Case "255.255.0.0"
				p_RetVal = "/16"
			Case "255.254.0.0"
				p_RetVal = "/15"
			Case "255.252.0.0"
				p_RetVal = "/14"
			Case "255.248.0.0"
				p_RetVal = "/13"
			Case "255.240.0.0"
				p_RetVal = "/12"
			Case "255.224.0.0"
				p_RetVal = "/11"
			Case "255.192.0.0"
				p_RetVal = "/10"
			Case "255.128.0.0"
				p_RetVal = "/9"
			Case "255.0.0.0"
				p_RetVal = "/8"
			Case "254.0.0.0"
				p_RetVal = "/7"
			Case "252.0.0.0"
				p_RetVal = "/6"
			Case "248.0.0.0"
				p_RetVal = "/5"
			Case "240.0.0.0"
				p_RetVal = "/4"
			Case "224.0.0.0"
				p_RetVal = "/3"
			Case "192.0.0.0"
				p_RetVal = "/2"
			Case Else
				p_RetVal = "/?"
		End Select
		MaskDecimal = p_RetVal
	End Property 
	
	Public Property Get MaskOctet
		MaskOctet = m_Mask
	End Property 

'Private Function SubnetToSlash()

End Class

Class DNSEntry
	Private m_Record()
	Private m_Type()
	Private m_Address()
	
	Private Sub Class_Initialize
		ReDim m_Record(0)
		ReDim m_Type(0)
		ReDim m_Address(0)
	End Sub

	Public Sub Add(p_Record, p_Type, p_Address)
		Dim sp_Count
		sp_Count = Me.UIndex + 1
		ReDim Preserve m_Record(sp_Count)
		ReDim Preserve m_Type(sp_Count)
		ReDim Preserve m_Address(sp_Count)

		m_Record(sp_Count) = p_Record
		m_Address(sp_Count) = p_Address
		Select Case UCase(p_Type)
			Case "A"
				m_Type(sp_Count) = "Alias"
			Case "CNAME"
				m_Type(sp_Count) = "CName"
			Case Else
				m_Type(sp_Count) = "Unknown"
		End Select
	End Sub
	
	Public Property Get Address(p_I)
		Address = m_Address(p_I)
	End Property 	
	Public Property Get Record(p_I)
		Record = m_Record(p_I)
	End Property 	
	Public Property Get RecordType(p_I)
		RecordType = m_Type(p_I)
	End Property 
	
	Public Property Get UIndex
		UIndex = UBound(m_Address)
	End Property

	Public Property Get LIndex
		UIndex = LBound(m_Address)
	End Property
	
End Class

'End Classes
'*****************************************************************
