'VBScript

'pstMgr.vbs by CommVault Professional Services
	'Edited on 11/15/2010 to add /RL option. /RL can be used with /InputFile: to read a computer list file for /Report output. /Input file replaces /ADPath which is used with /R option.
	'Edited on 11/17/2010 to add filtering. Filter.txt is created in c:\pstmgrlogs automatically with some default entries. It can be modfied to add more entries
		'cont...fixed /C to name copied file consistent with other options and to avoid overwrite of existing PST's with same name
		'cont...fixed uppercase/lowercase mismatch with LegacyExchangeDN comparison with file owner name.
'This script has twenty (20) switches: (/R /Report: /C /CL /P /InputFile: /ADPath: /Server: /Share: /Source: /Destination: /L /S /D /DL /E /MTime: /G /? /H
'This script derives security from the user logged in running the script. 
'Run the script while logged in as a Domain Admin
'Make sure that there are write priviledges on destination share especially with /M option!!!!.
'Logon Script, DeleteLocalPSTs.vbs is deployed to startup folder for all users. Each user with PST's on workstation will have to logon for deletion of PST's
'With /? or /H switch will give usage instructions
'Must be run via cmd as cscript pstmgr.vbs <OPTIONS>


'------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' START SCRIPT
'
'------------------------------------------------------------------------------------------------------------------------------------------------------------

'Check if cscript

checkifcscript

Set WshShell = WScript.CreateObject("WScript.Shell")
Sub checkifcscript
  strengine = LCase(Mid(WScript.FullName, InstrRev(WScript.FullName,"\")+1))
  If Not strengine="cscript.exe" Then
	wscript.echo "Use CMD cscript pstmgr.vbs <OPTIONS>"
    WScript.Quit
  End If
End Sub

WScript.Echo "CommVault Professional Services"
WScript.Echo "Latest Edit: 11/17/2010"
Wscript.echo
WScript.Echo "/? or /H for Usage Instructions"
Wscript.echo
strMyComputer = "."

Set objWMI2 = GetObject("winmgmts:\\" & strMyComputer & "\root\cimv2")					
Set objNetwork = Wscript.CreateObject("WScript.Network")									
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")					
Set objFSO = CreateObject("Scripting.FileSystemObject")								
Set objShell = Wscript.CreateObject("wscript.shell")




strDirectory = "c:\pstmgrlogs"
strFile = "\pstmgr.log"
clist = "\computerlist"
CurrentDate = Now

'Check if log directory already exists. If not create it.
If objFSO.FolderExists(strDirectory) Then
	Set objFolder = objFSO.GetFolder(strDirectory)
	Else
	Set objFolder = objFSO.CreateFolder(strDirectory)
End If

'Check if log file already exists. If not create it.
If objFSO.FileExists(strDirectory & strFile) Then
	Set objFolder = objFSO.GetFolder(strDirectory)
	Else
	Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
End If 

set objFile = nothing
set objFolder = Nothing

' ForAppending = 8 ForReading = 1, ForWriting = 2
Const ForAppending = 8

'Open log file for appending
Set objReportFile = objFSO.OpenTextFile (strDirectory & strFile, ForAppending, True)

objReportFile.WriteBlankLines(2)

objReportFile.WriteLine("----------------------------------------------------------------------------------------------------------------------")
objReportFile.WriteLine("							SCRIPT RUN: " & FormatDateTime(CurrentDate, vbGeneralDate))
objReportFile.WriteLine("----------------------------------------------------------------------------------------------------------------------")

objReportFile.WriteBlankLines(2)


'Filters
'strDirectory = "c:\pstmgrlogs"
strFile2 = "c:\pstmgrlogs\filter.txt"
'Set objTextFile = objFSO.OpenTextFile(strFile,1)
gettingFileFilters = false
gettingFolderFilters = false
strFileFilters = ""
strFolderFilters = ""



'Check if log file already exists. If not create it.
If objFSO.FileExists(strFile2) Then
	wscript.echo "There is a filter file"
	Set objTextFile = objFSO.OpenTextFile(strFile2,1)
	Do Until objTextFile.AtEndOfStream
            strNextLine = objTextFile.ReadLine
            if InStr(strNextLine,"[file]") Then
                        ' We are getting file filters
                        gettingFileFilters = true
                        gettingFolderFilters = false
            End If
            if InStr(strNextLine,"[folder]") Then
                        ' we are getting folder filters
                        gettingFolderFilters = true
                        gettingFileFilters = false
            End If 
            If gettingFileFilters Then
                        header = StrComp(strNextLine,"[file]",1)
                        if len(strNextLine)>1 and header <> 0 Then
                                    strFileFilters = strFileFilters & strNextLine & ","
                        End If
            End If    
            If gettingFolderFilters Then
                        header = StrComp(strNextLine,"[folder]",1)
                        If len(strNextLine)>1 and header <> 0 Then
                                    strFolderFilters = strFolderFilters & strNextLine & ","
                        End If
            End If    
	Loop

fileFilters = split(strFileFilters,",")
folderFilters = split(strFolderFilters,",")

objTextFile.Close
Else
	wscript.echo "No Filter File. Need to create one"
	Set objFile = objFSO.CreateTextFile(strFile2)
	objFile.close
	Set objFile2 = objFSO.OpenTextFile (strFile2, ForAppending, True)
	objFile2.writeline "[file]"
	objFile2.writeline "exdapsttemplate.pst"
	objFile2.writeline "unicodepsttemplate.pst"
	objFile2.writeline "[folder]"
	objFile2.writeline "c:\$recycle.bin"
	objFile2.writeline "c:\RECYCLER"
	objFile2.close
	set objFile2 = nothing
	
	Set objFile3 = objFSO.OpenTextFile(strFile2,1)
	Do Until objFile3.AtEndOfStream
            strNextLine = objFile3.ReadLine
            if InStr(strNextLine,"[file]") Then
                        ' We are getting file filters
                        gettingFileFilters = true
                        gettingFolderFilters = false
            End If
            if InStr(strNextLine,"[folder]") Then
                        ' we are getting folder filters
                        gettingFolderFilters = true
                        gettingFileFilters = false
            End If 
            If gettingFileFilters Then
                        header = StrComp(strNextLine,"[file]",1)
                        if len(strNextLine)>1 and header <> 0 Then
                                    strFileFilters = strFileFilters & strNextLine & ","
                        End If
            End If    
            If gettingFolderFilters Then
                        header = StrComp(strNextLine,"[folder]",1)
                        If len(strNextLine)>1 and header <> 0 Then
                                    strFolderFilters = strFolderFilters & strNextLine & ","
                        End If
            End If    
	Loop

fileFilters = split(strFileFilters,",")
folderFilters = split(strFolderFilters,",")
End If 


'set objFolder = Nothing


'----------------------------------------------------------------------------------------------------------
'
' Parse input parameters
'
'----------------------------------------------------------------------------------------------------------

Set colNamedArguments = WScript.Arguments.Named
If colNamedArguments.Exists("H") OR colNamedArguments.Exists("?") Then
  Call Usage()
  WScript.Quit
End If

adpath = colNamedArguments.Item("ADPath")
'ex. LDAP://ou=pstgroup,DC=company,DC=com
strServer = colNamedArguments.Item("Server")
strShare = colNamedArguments.Item("Share")
booCopy = colNamedArguments.Exists("C")
objStartFolder = colNamedArguments.Item("Source")
strCopy1 = colNamedArguments.Item("Destination")
ReportFile = colNamedArguments.Item("Report")
ExMerge = colNamedArguments.Exists("E")
Report = colNamedArguments.Exists("R")
ReportList = colNamedArguments.Exists("RL")
ListCreate = colNamedArguments.Exists("L")
ReadList = colNamedArguments.Exists("CL")
InputFile = colNamedArguments.Item("InputFile")
Reboot = colNamedArguments.Exists("S")
Deploy = colNamedArguments.Exists("D")
DeleteDeploy = colNamedArguments.Exists("DL")
booPrepare = colNamedArguments.Exists("P")
strMod = colNamedArguments.Item("MTime")
GetLegacy = colNamedArguments.Exists("G")

'------------------------------------------------------------------------------
'
'	Main Program Logic
'
'------------------------------------------------------------------------------

If DeleteDeploy Then 'Delete DeleteLocalPSTs.vbs from All Users startup directory specified in computerlist
 DeleteLoginScript
 End If

If Deploy Then 'Copy DeleteLocalPSTs.vbs to All Users startup directory to be executed upon logon. Will close default profiles and delete psts
 DeployLoginScript
 End If

If Report Then 'Create Report file with PST's found on workstations specified in computerlist. Directory must exist!
 LocalReport
 End If
 
If ReportList Then 'Create Report file with PST's found on workstations specified in computerlist. Directory must exist!
 LocalReportList
 End If

If ExMerge Then 'Create ExMerge ready PST's
 SetupExmerge
 End If
 
If booPrepare Then 'Copy PST's to central share and maintain file ownership for use with PSTDiscovery tool
 PreparePSTs
End If
 
If booCopy Then 'Copy PST's to central share and create pst migration folder structure for subclient. Uses LDAP path
 LDAPCopy
 End If

If ListCreate Then 'Create computerlist file based on computers found in LDAP path. Can be used as input to other options.
 CreateComputerList
 End If

If ReadList Then 'Copy PST's to central share and create pst migration folder structure for subclient. Uses computerlist
 ComputerListCopy
 End If

If Reboot Then
 RebootRemoteComputer
 End If

If GetLegacy Then
 GetLegacyExchangeDN
 End If

'------------------------------------------------------------------------------
'
'	SubRoutines
'
'------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'GET LEGACYEXCHANGEDN INFO

Sub GetLegacyExchangeDN

	Const ADS_PROPERTY_NOT_FOUND = &h8000500D

	On Error Resume Next
	
	'Create Report File. Directory must exist!
	Set objReportFile4 = objFSO.CreateTextFile(ReportFile,True)
	
	objReportFile.WriteLine("/G Report Option Ran")
	objReportFile.WriteBlankLines(1)
	objReportFile.WriteLine("LegacyExchangeDN and SAMAccountName for all users in /ADPath" & adpath)
	objReportFile.WriteBlankLines(1)

	Set objUsers = GetObject(adpath)

	For Each objUser in objUsers
		LoginName = objUser.Get("sAMAccountName") 
		Name = objUser.Get("legacyExchangeDN")

		If err.number = ADS_PROPERTY_NOT_FOUND Then

			objReportFile.WriteLine("Error:  " & vbTab & Hex(ADS_PROPERTY_NOT_FOUND) & " NO LegacyExchangeDN VALUE!" & vbTab & "sAMAccountName: " & vbTab & LoginName)
			objReportFile4.WriteLine("Error:  " & vbTab & Hex(ADS_PROPERTY_NOT_FOUND) & " NO LegacyExchangeDN VALUE!" & vbTab & "sAMAccountName: " & vbTab & LoginName)
			wscript.echo "Error: " & Hex(ADS_PROPERTY_NOT_FOUND) & " NO LegacyExchangeDN VALUE!" & " sAMAccountName: " & LoginName
			Err.Clear
		Else
			WScript.Echo "legacyExchangeDN: " & Name & "  sAMAccountName: " & LoginName
			objReportFile.WriteLine("legacyExchangeDN: " & vbTab & Name & vbTab & "sAMAccountName: " & vbTab & LoginName) 
			objReportFile4.WriteLine("legacyExchangeDN: " & vbTab & Name & vbTab & "sAMAccountName: " & vbTab & LoginName)
			
		End If
	Next
	objReportFile.Close
	objReportFile4.Close
	
End Sub 'End Sub GetLegacyExchangeDN


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'MAKE LOCAL REPORT FILE WITH INFORMATION REGARDING ALL PST'S FOUND ON ALL WORKSTATIONS

Sub LocalReport

	On Error Resume Next
	Err.Clear
	
	Const ADS_PROPERTY_NOT_FOUND = &h8000500D

	'Determine if proper input variables
	If adpath = "" Then
		wscript.echo "Need to use /Report and /ADPath:LDAP:// for /R option to work!"
		objReportFile.WriteLine("Need to use /Report and /ADPath:LDAP:// for /R option to work!")
		wscript.quit
	End If
	
	If ReportFile = "" Then
		wscript.echo "Need to use /Report and /ADPath:LDAP:// for /R option to work!"
		objReportFile.WriteLine("Need to use /Report and /ADPath:LDAP:// for /R option to work!")
		wscript.quit
	End If
	
	'Create Report File. Directory must exist!
	Set objReportFile3 = objFSO.CreateTextFile(ReportFile,True)
	objReportFile.WriteLine("/R Report Option Ran")
	objReportFile.WriteBlankLines(1)
	
	'Query AD for all computers in LDAP path provide as argument
	Const ADS_SCOPE_SUBTREE = 2
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	objConnection.Provider = ("ADsDSOObject")
	objConnection.Open "Active Directory Provider"
	objCommand.ActiveConnection = objConnection
	objCommand.Properties("Page Size") = 10000 
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE  
	objCommand.CommandText = "SELECT Name, operatingSystemVersion FROM '" & adpath & "' WHERE objectClass='computer'"

	'Loop through all computers found in LDAP path provided and add pst's found on each to report file 
	Set objRecordSet = objCommand.Execute  
	objRecordSet.MoveFirst 
	wscript.echo "Looking for PST's in LDAP path and writing to " & ReportFile
	wscript.echo
	
	Do Until objRecordSet.EOF

		strComputer = objRecordSet.Fields("Name").Value
		If strComputer = "" Then
		wscript.echo "No Computer Value"
		wscript.quit
		End If 
		Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
		strPingResults = LCase(WshExec.StdOut.ReadAll)
		
		If InStr(strPingResults, "reply from") Then
			WScript.Echo strComputer & " responded to ping."
			WScript.Echo
			Set objWshScriptExec = objShell.Exec("ping.exe " & strComputer)
			Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")
			Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")
			
			For Each objDisk in colDisks
				X = objDisk.DeviceID
				Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")
				For Each objFile in colFiles
				
				If InArray(objFile.Drive & objFile.Path,folderFilters) Then
                        WScript.Echo "PST was filtered due to a folder filter"
						objReportFile.WriteLine("PST was filtered due to a folder filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a folder filter" & vbTab & strComputer & vbTab & objFile.Name)
				ElseIf InArray(objFile.FileName & "." & objFile.Extension,fileFilters) Then
                        WScript.Echo "PST was filtered due to a file filter"
						objReportFile.WriteLine("PST was filtered due to a file filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a file filter" & vbTab & strComputer & vbTab & objFile.Name)
				Else 
				
					wscript.echo "PST NOT FILTERED"
					dateTime.Value = objFile.LastModified
					modDate = DateDiff("d", dateTime.GetVarDate(True), Date)
					wscript.echo "Checking File Modification Date"
					If int(modDate) => int(strMod) Then
						wscript.echo "File Modification Match - Write to " & ReportFile
						'wscript.echo objFile.Name
						wscript.echo
						  'This following is necessary for paths with apostrophes
						  A = split(objFile.Name,"\")
						  B = join(A,"\\")
						  A = split(B,"'")
						  FName = join(A,"\'")
						  Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_LogicalFileSecuritySetting WHERE Path='" & FName &"'")
					  
						For Each objItem in colItems
							RetVal = objItem.GetSecurityDescriptor(wmiSecurityDescriptor)
							Set G = wmiSecurityDescriptor.Owner
							NTName = G.Domain & "\" & G.Name
							NewName = TranslateNT4toDn(NTName)
							Set objUser =  GetObject("LDAP://" & NewName)
							LegacyName = objUser.Get("legacyExchangeDN")
							'WScript.Echo LegacyName
							If err.number <> 0 Then
								Wscript.echo "No LegacyExchangeDN Value!"
								wscript.echo objFile.Name
								wscript.echo
								wscript.echo
								objReportFile.WriteLine("No LegacyExchangeDN. Check your security access." & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
								objReportFile3.WriteLine("No LegacyExchangeDN" & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))
							Else		
								'WScript.Echo LegacyName
								TT = split(LegacyName,"=")
								'wscript.echo TT
								strExtension = TT(UBound(TT) - 0)
								wscript.echo "Checking to see if LegacyExchangeDN matches file owner"
								If LCase(strExtension) = LCase(G.Name) Then
									WScript.echo "LegacyExchangeDN Match!"
									WScript.Echo LegacyName
									wscript.echo
									objReportFile.WriteLine("LegacyExchangeDN Match!")
									objReportFile.WriteLine(strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
									objReportFile3.WriteLine("LegacyExchangeDN Match!" & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
								Else
									
									WScript.echo "LegacyExchangeDN Mismatch!"
									wscript.echo "Error: " & Hex(ADS_PROPERTY_NOT_FOUND)
									wscript.echo
									objReportFile.WriteLine("LegacyExchangeDN Mismatch!")
									objReportFile.WriteLine(strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
									objReportFile3.WriteLine("LegacyExchangeDN Mismatch!" & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
								End If
							End if
						Next
					Else
					wscript.echo "File Modification No Match - Don't Report"
					wscript.echo objFile.Name
					
					End If	
				End If
				Next
			Next
		Else	
		WScript.Echo strComputer & " did not respond to ping."
		WScript.Echo	 
		End If	
		objRecordSet.MoveNext
		'wscript.echo
	Loop
	objReportFile.Close
	objReportFile3.Close
	
End Sub 'End Sub LocalReport


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'MAKE LOCAL REPORT FILE WITH INFORMATION REGARDING ALL PST'S FOUND ON ALL WORKSTATIONS

Sub LocalReportList

	On Error Resume Next
	Err.Clear
	
	Const ADS_PROPERTY_NOT_FOUND = &h8000500D
	intHighNumber = 100000
	intLowNumber = 1
	Const ForReading = 1
	
	'Determine if proper input variables
	If InputFile = "" Then
		wscript.echo "Need to use /InputFile: for /CL option to work!"
		objReportFile.WriteLine("Need to use /InputFile: for /CL option to work!")
		wscript.quit
	End If
	
	If objFso.FileExists(InputFile) = False Then
		wscript.echo "Need Valid Input!"
		wscript.quit
    End If
	
	If ReportFile = "" Then
		wscript.echo "Need to use /Report and /ADPath:LDAP:// for /R option to work!"
		objReportFile.WriteLine("Need to use /Report and /ADPath:LDAP:// for /R option to work!")
		wscript.quit
	End If
	
	'Create Report File. Directory must exist!
	Set objReportFile3 = objFSO.CreateTextFile(ReportFile,True)
	objReportFile.WriteLine("/RL Report Option Ran")
	objReportFile.WriteBlankLines(1)

Set objReportFile2 = objFSO.OpenTextFile(InputFile, ForReading, True)
	'Loop through one computer per line until reach end of file.
	Do While objReportFile2.AtEndOfStream <> True
		strComputer = objReportFile2.ReadLine

	wscript.echo "Looking for PST's in Computer List and writing to " & ReportFile
	wscript.echo
	
	

		If strComputer = "" Then
		wscript.echo "No Computer Value"
		wscript.quit
		End If 
		Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
		strPingResults = LCase(WshExec.StdOut.ReadAll)
		
		If InStr(strPingResults, "reply from") Then
			WScript.Echo strComputer & " responded to ping."
			WScript.Echo
			Set objWshScriptExec = objShell.Exec("ping.exe " & strComputer)
			Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")
			Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")
			
			For Each objDisk in colDisks
				X = objDisk.DeviceID
				Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")
				For Each objFile in colFiles
				
				If InArray(objFile.Drive & objFile.Path,folderFilters) Then
                        WScript.Echo "PST was filtered due to a folder filter"
						objReportFile.WriteLine("PST was filtered due to a folder filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a folder filter" & vbTab & strComputer & vbTab & objFile.Name)
				ElseIf InArray(objFile.FileName & "." & objFile.Extension,fileFilters) Then
                        WScript.Echo "PST was filtered due to a file filter"
						objReportFile.WriteLine("PST was filtered due to a file filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a file filter" & vbTab & strComputer & vbTab & objFile.Name)
				Else 
				
					'wscript.echo "PST NOT FILTERED"
					
					dateTime.Value = objFile.LastModified
					modDate = DateDiff("d", dateTime.GetVarDate(True), Date)
					wscript.echo "Checking File Modification Date"
					If int(modDate) => int(strMod) Then
						wscript.echo "File Modification Match - Write to " & ReportFile
						'wscript.echo objFile.Name
						wscript.echo
						  'This following is necessary for paths with apostrophes
						  A = split(objFile.Name,"\")
						  B = join(A,"\\")
						  A = split(B,"'")
						  FName = join(A,"\'")
						  Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_LogicalFileSecuritySetting WHERE Path='" & FName &"'")
					  
						For Each objItem in colItems
							RetVal = objItem.GetSecurityDescriptor(wmiSecurityDescriptor)
							Set G = wmiSecurityDescriptor.Owner
							NTName = G.Domain & "\" & G.Name
							NewName = TranslateNT4toDn(NTName)
							Set objUser =  GetObject("LDAP://" & NewName)
							LegacyName = objUser.Get("legacyExchangeDN")
							'WScript.Echo LegacyName
							If err.number <> 0 Then
								Wscript.echo "No LegacyExchangeDN Value!."
								wscript.echo objFile.Name
								wscript.echo
								wscript.echo
								objReportFile.WriteLine("No LegacyExchangeDN. Check your security access." & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
								objReportFile3.WriteLine("No LegacyExchangeDN" & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))
							Else		
								'WScript.Echo LegacyName
								TT = split(LegacyName,"=")
								'wscript.echo TT
								strExtension = TT(UBound(TT) - 0)
								wscript.echo "Checking to see if LegacyExchangeDN matches file owner"
								If LCase(strExtension) = LCase(G.Name) Then
									WScript.echo "LegacyExchangeDN Match!"
									WScript.Echo LegacyName
									wscript.echo
									objReportFile.WriteLine("LegacyExchangeDN Match!")
									objReportFile.WriteLine(strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
									objReportFile3.WriteLine("LegacyExchangeDN Match!" & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
								Else
									
									WScript.echo "LegacyExchangeDN Mismatch!"
									wscript.echo "Error: " & Hex(ADS_PROPERTY_NOT_FOUND)
									wscript.echo
									objReportFile.WriteLine("LegacyExchangeDN Mismatch!")
									objReportFile.WriteLine(strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
									objReportFile3.WriteLine("LegacyExchangeDN Mismatch!" & vbTab & strComputer & vbTab & G.Name & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True))	
								End If
							End if
						Next
					Else
					wscript.echo "File Modification No Match - Don't Report"
					wscript.echo objFile.Name
					
					End If	
				End If
					
				Next
			Next
		Else	
		WScript.Echo strComputer & " did not respond to ping."
		WScript.Echo	 
		End If	

	Loop
	
	objReportFile.Close
	objReportFile3.Close
	'objReportFile2.Close
	
End Sub 'End Sub LocalReportList


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



'READ COMPUTER LIST FROM INPUT FILE AND COPY PSTS FOUND ON COMPUTERS ON EACH LINE


Sub ComputerListCopy

On Error Resume Next
Err.Clear
	Const ADS_PROPERTY_NOT_FOUND = &h8000500D
	intHighNumber = 100000
	intLowNumber = 1
	Const ForReading = 1
	objReportFile.WriteLine(" /CL Option Ran. PST's Copied from input file")
	objReportFile.WriteBlankLines(1)
	
	'Determine if proper input variables
	If InputFile = "" Then
		wscript.echo "Need to use /InputFile: for /CL option to work!"
		objReportFile.WriteLine("Need to use /InputFile: for /CL option to work!")
		wscript.quit
	End If
	
	If objFso.FileExists(InputFile) = False Then
		wscript.echo "Need Valid Input!"
		wscript.quit
    End If
	
	If strServer = "" Then
		wscript.echo "Need /Server: Value!"
		objReportFile.WriteLine("Need /Server: Value!")
		wscript.quit
	End If
	
	If strShare = "" Then
		wscript.echo "Need /Share: Value!"
		objReportFile.WriteLine("Need /Share: Value!")
		wscript.quit
	End If
	
	'Determine an available drive letter to use to map share to copy pst's to
	Set objDictionary = CreateObject("Scripting.Dictionary")
	Set colDisks = objWMI2.ExecQuery("Select * from Win32_LogicalDisk")

	For Each objDisk in colDisks
		objDictionary.Add objDisk.DeviceID, objDisk.DeviceID
	Next
	
	strDrive = " "
	
	For i = 67 to 90																	
		If objDictionary.Exists(Chr(i) & ":") Then
		Else
			strDrive = Chr(i) & ":"	
			Exit For
		End If
	Next
	
	If strDrive = " " Then
		Wscript.Echo("There are no available drive letters on this computer.")
		Wscript.Quit
	End If
	
	'Map Network drive on computer script is run from based on first available drive letter. 
	Set objNetwork = Wscript.CreateObject("WScript.Network")
	objNetwork.MapNetworkDrive strDrive, "\\" & strServer & "\" & strShare
	If err.number <> 0 Then
		wscript.echo "Couldn't map drive! PST's won't get copied"
		wscript.echo "Make sure /server and /share options specified!"
		objReportFile.WriteLine("Couldn't map drive. PST's won't get copied")
		objReportFile.WriteLine("Make sure /server and /share options specified!")
		objReportFile.WriteBlankLines(1)
		wscript.quit
	End If
	
	strDir = strDrive & "\" & "\" & "CommVault PST Staging Directory" & "\"
	If objFSO.FolderExists(strDir) Then
		Set objFolder = objFSO.GetFolder(strDir)
		Else
		Set objFolder = objFSO.CreateFolder(strDir)
	End If
	
	'Open computer list file for reading


	Set objReportFile2 = objFSO.OpenTextFile(InputFile, ForReading, True)
	
	wscript.echo "Looking for PST's in computerlist and copying to " & "\\" & strServer & "\" & strShare
	wscript.echo
	
	'Loop through one computer per line until reach end of file.
	Do While objReportFile2.AtEndOfStream <> True
		strComputer = objReportFile2.ReadLine
		Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
		strPingResults = LCase(WshExec.StdOut.ReadAll)
		
		If InStr(strPingResults, "reply from") Then
			WScript.Echo strComputer & " responded to ping."
			WScript.Echo
			
			'Kill outlook if running to allow for copy of pst's open in outlook. Sleep for 10 seconds to give Outlook process some time to stop
			Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")
			Set colItems = objWMI.ExecQuery ("Select * From Win32_Process Where Name = 'outlook.exe'")
			
			For Each Process in colItems
				wscript.echo "If Outlook.exe is running kill it"
				wscript.echo
				errResult = Process.Terminate															
				If errResult <> 0 Then
				wscript.echo "Error: " & errResult & " Couldn't Kill Outlook.exe"
				objReportFile.WriteLine("Unable to Terminate Outlook. Error: " & errResult)
				End If
			Next
			
			WScript.Sleep(10000)
			
			'Make sure to scan only local disks for PST's
			Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")
			
			For Each objDisk in colDisks
				X = objDisk.DeviceID
				Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")
				
				For Each objFile in colFiles
				
					If InArray(objFile.Drive & objFile.Path,folderFilters) Then
                        WScript.Echo "PST was filtered due to a folder filter"
						objReportFile.WriteLine("PST was filtered due to a folder filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a folder filter" & vbTab & strComputer & vbTab & objFile.Name)
					ElseIf InArray(objFile.FileName & "." & objFile.Extension,fileFilters) Then
                        WScript.Echo "PST was filtered due to a file filter"
						objReportFile.WriteLine("PST was filtered due to a file filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a file filter" & vbTab & strComputer & vbTab & objFile.Name)
					Else 
				
					'wscript.echo "PST NOT FILTERED"
					dateTime.Value = objFile.LastModified
					modDate = DateDiff("d", dateTime.GetVarDate(True), Date)
					
					If int(modDate) => int(strMod) Then
						wscript.echo "File Modification Match - Copy"
						wscript.echo objFile.Name
						'This following is necessary for paths with apostrophes
						A = split(objFile.Name,"\")
						B = join(A,"\\")
						A = split(B,"'")
						FName = join(A,"\'")
						Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_LogicalFileSecuritySetting WHERE Path='" & FName &"'")
							  
						For Each objItem in colItems
							'Create random number so that each PST is copied to it's own folder to avoid overwrite if same name. A folder instead of file is created per the behavior of XCOPY.
							Randomize
							intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
							tt = left(intNumber,3)
							RetVal = objItem.GetSecurityDescriptor(wmiSecurityDescriptor)
							Set G = wmiSecurityDescriptor.Owner
							NTName = G.Domain & "\" & G.Name
							NewName = TranslateNT4toDn(NTName)
							Set objUser =  GetObject("LDAP://" & NewName)
							LegacyName = objUser.Get("legacyExchangeDN")
							kk = split(LegacyName,"=")
							strExtension = kk(UBound(kk) - 0)
							If LCase(strExtension) = LCase(G.Name) Then
								WScript.echo "LegacyExchangeDN Match!"
								WScript.Echo LegacyName
								wscript.echo
								strDir2 = strDir & G.Name & "\"
									
								If objFSO.FolderExists(strDir2) Then
									Set objFolder = objFSO.GetFolder(strDir2)
								Else
									Set objFolder = objFSO.CreateFolder(strDir2)
								End If
									
								strCopy = strDrive & G.Name & "_" & tt & "\"
								MyString = Mid(objFile.Name,3)
								X2 = left(X, 1)
								FinalString = "\\" & strComputer & "\" & X2 & "$" & MyString
								wscript.echo "Copying....."
								WScript.Echo FinalString & " Copied To..."
								strCopy = strDir2 & "\" & strComputer & "_" & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "." & objFile.Extension
								WScript.Echo strCopy
								wscript.echo
								objFSO.Copyfile Finalstring, strCopy
								objReportFile.Writeline("PST Successfully Copied")
								objReportFile.Writeline(objFile.Name & " Copied to...")
								objReportFile.Writeline(strCopy)
								objReportFile.WriteBlankLines(1)
							
								If errResult <> 0 Then
									objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
								End If
							Else
								WScript.echo "LegacyExchangeDN Mismatch for file owner: " & G.Name 
								wscript.echo "Error: " & Hex(ADS_PROPERTY_NOT_FOUND)
								objReportFile.WriteLine ("LegacyExchangeDN Mismatch! for file owner: " & G.Name)
								objReportFile.WriteLine ("Error: " & Hex(ADS_PROPERTY_NOT_FOUND))
								MyString = Mid(objFile.Name,3)
								X2 = left(X, 1)
								FinalString = "\\" & strComputer & "\" & X2 & "$" & MyString
								strDir4 = strDrive & "\" & "\" & "CommVault PST Staging Directory" & "\" & "LegacyExchangeDN MISMATCH" & "\"
								
								If objFSO.FolderExists(strDir4) Then
									Set objFolder = objFSO.GetFolder(strDir4)
								Else
									Set objFolder = objFSO.CreateFolder(strDir4)
								End If
								
								strDir3 = strDir4 & G.Name & "\"
										
								If objFSO.FolderExists(strDir3) Then
									Set objFolder = objFSO.GetFolder(strDir3)
								Else
									Set objFolder = objFSO.CreateFolder(strDir3)
								End If
								
								strCopy5 = strDir3 & strComputer & "_" & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "." & objFile.Extension
								wscript.echo "Copying....."
								wscript.echo
								wscript.Echo FinalString & " Copied To..."
								WScript.Echo strCopy5
								wscript.echo
									
								objFSO.Copyfile FinalString, strCopy5
								objReportFile.Writeline("PST Successfully Copied")
								objReportFile.Writeline(objFile.Name & " Copied to...")
								objReportFile.Writeline(strCopy5)
								objReportFile.WriteBlankLines(1)
					
								If errResult <> 0 Then
									objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
									Wscript.Echo "Result1: " & errResult
								End If
							End If
						Next
					Else
					wscript.echo "File Modification No Match - Don't Copy"
					wscript.echo objFile.Name
					wscript.echo
					objReportFile.Writeline("File Modification No Match - Don't Copy")
					objReportFile.Writeline(objFile.Name)
					objReportFile.WriteLine
					End If
				End If
				Next
			Next
		Else
		WScript.Echo strComputer & " did not respond to ping."
		WScript.Echo
		End If		
	Loop
	'close log file and remove network drive after copying.
	objReportFile.Close
	objNetwork.RemoveNetworkDrive strDrive
	
End Sub 'End Sub ComputerListCopy

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'READ COMPUTER LIST FROM INPUT FILE AND COPY PSTS FOUND ON COMPUTERS ON EACH LINE FOR PREPARATION FOR PSTDISCOVERY TOOL


Sub PreparePSTs
	
	On Error Resume Next
	Err.Clear

	intHighNumber = 100000
	intLowNumber = 1
	Const ForReading = 1
	objReportFile.WriteLine(" /P Option Ran. PST's Copied from input file to PSTDiscovery Tool SharedPSTLocation")
	objReportFile.WriteBlankLines(1)
	
	'Determine if proper input variables
	If InputFile = "" Then
		wscript.echo "Need to use /InputFile: /Server: and /Share: for /P option to work!"
		objReportFile.WriteLine("Need to use /InputFile: /Server: and /Share: for /P option to work!")
		wscript.quit
	End If
	
	If objFso.FileExists(InputFile) = False Then
		wscript.echo "Need Valid Input!"
		wscript.quit
    End If

	If strServer = "" Then
		wscript.echo "Need /Server: Value!"
		objReportFile.WriteLine("Need /Server: Value!")
		wscript.quit
	End If
	
	If strShare = "" Then
		wscript.echo "Need /Share: Value!"
		objReportFile.WriteLine("Need /Share: Value!")
		wscript.quit
	End If
	
	'Determine an available drive letter to use to map share to copy pst's to
	Set objDictionary = CreateObject("Scripting.Dictionary")
	Set colDisks = objWMI2.ExecQuery("Select * from Win32_LogicalDisk")

	For Each objDisk in colDisks
		objDictionary.Add objDisk.DeviceID, objDisk.DeviceID
	Next
	
	strDrive = " "
	
	For i = 67 to 90																	
		If objDictionary.Exists(Chr(i) & ":") Then
		Else
			strDrive = Chr(i) & ":"	
			Exit For
		End If
	Next
	
	If strDrive = " " Then
		Wscript.Echo("There are no available drive letters on this computer.")
		Wscript.Quit
	End If
	
	'Map Network drive on computer script is run from based on first available drive letter. 
	Set objNetwork = Wscript.CreateObject("WScript.Network")
	objNetwork.MapNetworkDrive strDrive, "\\" & strServer & "\" & strShare
	If err.number <> 0 Then
		wscript.echo "Couldn't map drive! PST's won't get copied"
		wscript.echo "Make sure /server and /share options specified!"
		objReportFile.WriteLine("Couldn't map drive. PST's won't get copied")
		objReportFile.WriteLine("Make sure /server and /share options specified!")
		objReportFile.WriteBlankLines(1)
		wscript.quit
	End If
	
	strDir = strDrive & "\" & "\" & "CommVault PSTDiscovery Tool SharedPSTLocation" & "\"

	If objFSO.FolderExists(strDir) Then
		Set objFolder = objFSO.GetFolder(strDir)
	Else
		Set objFolder = objFSO.CreateFolder(strDir)
	End If
	
	'Open computer list file for reading
	Set objReportFile2 = objFSO.OpenTextFile(InputFile, ForReading, True)
	
	wscript.echo "Looking for PST's in computerlist and copying to " & "\\" & strServer & "\" & strShare
	wscript.echo
	
	'Loop through one computer per line until reach end of file.
	Do While objReportFile2.AtEndOfStream <> True
		strComputer = objReportFile2.ReadLine
		Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
		strPingResults = LCase(WshExec.StdOut.ReadAll)
		
		If InStr(strPingResults, "reply from") Then
			WScript.Echo strComputer & " responded to ping."
			WScript.Echo
			objReportFile.WriteLine(strComputer & " responded to ping. Scanning for PST's...")
			objReportFile.WriteBlankLines(1)
			'Kill outlook if running to allow for copy of pst's open in outlook. Sleep for 10 seconds to give Outlook process some time to stop
			Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")
			Set colItems = objWMI.ExecQuery ("Select * From Win32_Process Where Name = 'outlook.exe'")
			
			For Each Process in colItems
				errResult = Process.Terminate															
				If errResult <> 0 Then
					objReportFile.WriteLine("Unable to Terminate Outlook. Error: " & errResult)
				End If
			Next
			
			WScript.Sleep(10000)
			
			'Make sure to scan only local disks for PST's
			Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")
			
			For Each objDisk in colDisks
				X = objDisk.DeviceID
				Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")
				
				For Each objFile in colFiles
				
					If InArray(objFile.Drive & objFile.Path,folderFilters) Then
                        WScript.Echo "PST was filtered due to a folder filter"
						objReportFile.WriteLine("PST was filtered due to a folder filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a folder filter" & vbTab & strComputer & vbTab & objFile.Name)
					ElseIf InArray(objFile.FileName & "." & objFile.Extension,fileFilters) Then
                        WScript.Echo "PST was filtered due to a file filter"
						objReportFile.WriteLine("PST was filtered due to a file filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a file filter" & vbTab & strComputer & vbTab & objFile.Name)
					Else 
				
						wscript.echo "PST NOT FILTERED"
						dateTime.Value = objFile.LastModified
						modDate = DateDiff("d", dateTime.GetVarDate(True), Date)

						If int(modDate) => int(strMod) Then
							wscript.echo "File Modification Match. Copy PST!"
							
							objReportFile.WriteLine("File Modification Match. Copy PST!")
							'This following is necessary for paths with apostrophes
							A = split(objFile.Name,"\")
							B = join(A,"\\")
							A = split(B,"'")
							FName = join(A,"\'")
							Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_LogicalFileSecuritySetting WHERE Path='" & FName &"'")
							
							For Each objItem in colItems
								'Create random number so that each PST is copied to it's own folder to avoid overwrite if same name. A folder instead of file is created per the behavior of XCOPY.
								Randomize
								intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
								tt = left(intNumber,3)
								RetVal = objItem.GetSecurityDescriptor(wmiSecurityDescriptor)
								Set G = wmiSecurityDescriptor.Owner
								NTName = G.Domain & "\" & G.Name
								NewName = TranslateNT4toDn(NTName)	
								Set objUser =  GetObject("LDAP://" & NewName)
								LegacyName = objUser.Get("legacyExchangeDN")
								kk = split(LegacyName,"=")
								strExtension = kk(UBound(kk) - 0)
								
								If LCase(strExtension) = LCase(G.Name) Then
									strDir2 = strDir & strComputer & "_" & G.Name & "_"
									strCopy = strDir2 & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "\"
									MyString = Mid(objFile.Name,3)
									X2 = left(X, 1)
									FinalString = "\\" & strComputer & "\" & X2 & "$" & MyString
									WScript.Echo FinalString & " Copied To..."
									WScript.Echo strCopy
									wscript.echo
									x = objFile.Name
									'XCOPY is used to maintain file ownership for each PST to help determine the proper mailbox association
									errResult = objShell.Run("XCOPY """ & FinalString & """ """ & strCopy & """ /O /i", 2, True)
									objReportFile.Writeline("PST Successfully Copied")
									objReportFile.Writeline(objFile.Name & " Copied to...")
									objReportFile.Writeline(strCopy)
									objReportFile.WriteBlankLines(1)
									
									If errResult <> 0 Then
										objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
									End If
								Else
									WScript.echo "LegacyExchangeDN Mismatch for file owner: " & G.Name 
									wscript.echo "Error: " & Hex(ADS_PROPERTY_NOT_FOUND)
									objReportFile.WriteLine ("LegacyExchangeDN Mismatch! for file owner: " & G.Name)
									objReportFile.WriteLine ("Error: " & Hex(ADS_PROPERTY_NOT_FOUND))
									MyString = Mid(objFile.Name,3)
									X2 = left(X, 1)
									FinalString = "\\" & strComputer & "\" & X2 & "$" & MyString
									strDir7 = strDir & "LegacyExchangeDN MISMATCH" & "\"
									
									If objFSO.FolderExists(strDir7) Then
										Set objFolder = objFSO.GetFolder(strDir7)
									Else
										Set objFolder = objFSO.CreateFolder(strDir7)
									End If
										
									strDir6 = strDir7 & strComputer & "_" & G.Name & "_"
									strCopy6 = strDir6 & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "\"	
									errResult = objShell.Run("XCOPY """ & FinalString & """ """ & strCopy6 & """ /O /i", 2, True)
									wscript.Echo FinalString & " Copied To..."
									WScript.Echo strCopy6
									wscript.echo
									objReportFile.Writeline("PST Successfully Copied")
									objReportFile.Writeline(objFile.Name & " Copied to...")
									objReportFile.Writeline(strCopy6)
									objReportFile.WriteBlankLines(1)
									
									If errResult <> 0 Then
										objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
									End If
								End If	
							Next
						Else
						wscript.echo "File Modification No Match - Don't Copy"
						wscript.echo objFile.Name
						wscript.echo
						objReportFile.Writeline("File Modification No Match - Don't Copy")
						objReportFile.Writeline(objFile.Name)
						objReportFile.WriteLine
						End If
					End If
				Next
			Next	
		Else
		WScript.Echo strComputer & " did not respond to ping."
		WScript.Echo
		objReportFile.Writeline(strComputer & " did not respond to ping. Not scanning for PST's")
		objReportFile.WriteBlankLines(1)
		End If
	Loop
	'close log file and remove network drive after copying.
	objReportFile.Close
	objNetwork.RemoveNetworkDrive strDrive

End Sub 'End Sub PreparePSTs


'-----------------------------------------------------------------------------------------------
'CREATE COMPUTER LIST FOR INPUT FILE

Sub CreateComputerList

		'Determine if proper input variables
	If adpath = "" Then
		wscript.echo "Need to use /ADPath:LDAP:// for /L option to work! Will create C:\pstmgrlogs\computerlist"
		objReportFile.WriteLine("Need to use /ADPath:LDAP:// for /L option to work! Will create C:\pstmgrlogs\computerlist")
		wscript.quit
	End If
	
	wscript.echo "Create File with list of computers found in LDAP Path provided: " & adpath
	wscript.echo
	objReportFile.WriteLine("/L Create Computer List Option Ran against: " & adpath )
	objReportFile.WriteBlankLines(1)
	
	'Create report file. Each time will overwrite preexisting report file. Directory must exist!
	Set objReportFile2 = objFSO.CreateTextFile(strDirectory & clist,True)
	Const ADS_SCOPE_SUBTREE = 2
	
	'Connect to Active Directory
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	objConnection.Provider = ("ADsDSOObject")
	objConnection.Open "Active Directory Provider"
	objCommand.ActiveConnection = objConnection
	
	'Limit to 10000 results
	objCommand.Properties("Page Size") = 10000 
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE  
	objCommand.CommandText = "SELECT Name, operatingSystemVersion FROM '" & adpath & "' WHERE objectClass='computer'"
	Set objRecordSet = objCommand.Execute  
	objRecordSet.MoveFirst
	
	'Loop through all computers found in LDAP path provided by users and write one computer per line to file
	Do Until objRecordSet.EOF
		objReportFile.WriteLine(objRecordSet.Fields("Name").Value)
		objReportFile2.WriteLine(objRecordSet.Fields("Name").Value)
		wscript.echo objRecordSet.Fields("Name").Value
		objRecordSet.MoveNext
	Loop
	
End Sub 'End Sub CreateComputerList

'-----------------------------------------------------------------------------------------------
'READ LDAP PATH AND COPY PSTS FOUND ON EACH COMPUTER IN LDAP CONTAINER SUCH AS DOMAIN OR OU

Sub LDAPCopy

	On Error Resume Next
	Err.Clear

	intHighNumber = 100000
	intLowNumber = 1

	objReportFile.WriteLine("/C Option Ran. PST's Copied based on computers in LDAP path specified")
	objReportFile.WriteBlankLines(1)
	
	'Determine if proper input variables
	If adpath = "" Then
		wscript.echo "Need to use /ADPath:LDAP:// /Server: and /Share: for /C option to work!"
		objReportFile.WriteLine("Need to use /ADPath:LDAP:// /Server: and /Share: for /C option to work!")
		wscript.quit
	End If
	
	If strServer = "" Then
		wscript.echo "Need /Server: Value!"
		objReportFile.WriteLine("Need /Server: Value!")
		wscript.quit
	End If
	
	If strShare = "" Then
		wscript.echo "Need /Share: Value!"
		objReportFile.WriteLine("Need /Share: Value!")
		wscript.quit
	End If
	
	'Determine an available drive letter to use to map share to copy pst's to
	Set objDictionary = CreateObject("Scripting.Dictionary")
	Set colDisks = objWMI2.ExecQuery("Select * from Win32_LogicalDisk")				

	For Each objDisk in colDisks
		objDictionary.Add objDisk.DeviceID, objDisk.DeviceID
	Next

	strDrive = " "
	
	For i = 67 to 90																		
		If objDictionary.Exists(Chr(i) & ":") Then
		Else
			strDrive = Chr(i) & ":"															
			Exit For
		End If
	Next

	If strDrive = " " Then
		Wscript.Echo("There are no available drive letters on this computer.")				
		Wscript.Quit
	End If
	
	'Map Network drive on computer script is run from based on first available drive letter.
	Set objNetwork = Wscript.CreateObject("WScript.Network")								
	objNetwork.MapNetworkDrive strDrive, "\\" & strServer & "\" & strShare
	
	If err.number <> 0 Then
		wscript.echo "Couldn't map drive! PST's won't get copied"
		wscript.echo "Make sure /server and /share options specified!"
		objReportFile.WriteLine("Couldn't map drive. PST's won't get copied")
		objReportFile.WriteLine("Make sure /server and /share options specified!")
		objReportFile.WriteBlankLines(1)
		wscript.quit
	End If
	
	strDir = strDrive & "\" & "\" & "CommVault PST Staging Directory" & "\"
	
	If objFSO.FolderExists(strDir) Then
		Set objFolder = objFSO.GetFolder(strDir)
		Else
		Set objFolder = objFSO.CreateFolder(strDir)
	End If
	
	Const ADS_SCOPE_SUBTREE = 2
	
	'Connect to Active Directory
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	objConnection.Provider = ("ADsDSOObject")
	objConnection.Open "Active Directory Provider"
	objCommand.ActiveConnection = objConnection
	'Limit to 10000 results
	objCommand.Properties("Page Size") = 10000 
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE  
	objCommand.CommandText = "SELECT Name, operatingSystemVersion FROM '" & adpath & "' WHERE objectClass='computer'"
	
	'Loop through all computers found in LDAP path provided
	Set objRecordSet = objCommand.Execute  
	objRecordSet.MoveFirst 
	wscript.echo "Looking for PST's in LDAP path " & adpath & " and copying to " & "\\" & strServer & "\" & strShare
	Do Until objRecordSet.EOF
		wscript.echo
		strComputer = objRecordSet.Fields("Name").Value
		If strComputer = "" Then
		wscript.echo "No Computer Value"
		wscript.quit
		End If 
		Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
		strPingResults = LCase(WshExec.StdOut.ReadAll)
		
		If InStr(strPingResults, "reply from") Then
			WScript.Echo strComputer & " responded to ping." 
			WScript.Echo
			objReportFile.WriteLine(strComputer & " responded to ping. Scanning for PST's...")
			objReportFile.WriteBlankLines(1)
			Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")
			
			'Kill Outlook process to copy PST's open and sleep for 10 seconds to give Outlook process time to exit
			Set colItems = objWMI.ExecQuery ("Select * From Win32_Process Where Name = 'outlook.exe'")
			  For Each Process in colItems
				errResult = Process.Terminate															' Kill Outlook
					If errResult <> 0 Then
					  objReportFile.WriteLine("Unable to Terminate Outlook. Error: " & errResult)			 
					End If
			  Next

			WScript.Sleep(10000)
			t = 0
			'Make sure to scan only local disks for PST's
			Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")

			For Each objDisk in colDisks
				X = objDisk.DeviceID
				Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")

				For Each objFile in colFiles
					If InArray(objFile.Drive & objFile.Path,folderFilters) Then
                        WScript.Echo "PST was filtered due to a folder filter"
						objReportFile.WriteLine("PST was filtered due to a folder filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a folder filter" & vbTab & strComputer & vbTab & objFile.Name)
					ElseIf InArray(objFile.FileName & "." & objFile.Extension,fileFilters) Then
                        WScript.Echo "PST was filtered due to a file filter"
						objReportFile.WriteLine("PST was filtered due to a file filter " & strComputer)
						wscript.echo objFile.Name
						objReportFile.WriteLine objFile.Name
						objReportFile3.WriteLine("PST was filtered due to a file filter" & vbTab & strComputer & vbTab & objFile.Name)
					Else 
				
					wscript.echo "PST NOT FILTERED"				
					dateTime.Value = objFile.LastModified
					modDate = DateDiff("d", dateTime.GetVarDate(True), Date)

					If int(modDate) => int(strMod) Then
						wscript.echo "File Modification Match. Copy PST!"
						
						objReportFile.WriteLine("File Modification Match. Copy PST!")
						'The following is necessary for paths with apostrophes
						A = split(objFile.Name,"\")
						B = join(A,"\\")
						A = split(B,"'")
						FName = join(A,"\'")
						Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_LogicalFileSecuritySetting WHERE Path='" & FName &"'")
						
						For Each objItem in colItems
							'Create random number so that each PST is copied to it's own folder to avoid overwrite if same name. A folder instead of file is created per the behavior of XCOPY.
							Randomize
							intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
							tt = left(intNumber,3)
							RetVal = objItem.GetSecurityDescriptor(wmiSecurityDescriptor)
							Set G = wmiSecurityDescriptor.Owner
							NTName = G.Domain & "\" & G.Name
							NewName = TranslateNT4toDn(NTName)
							Set objUser =  GetObject("LDAP://" & NewName)
							LegacyName = objUser.Get("legacyExchangeDN")
							kk = split(LegacyName,"=")
							strExtension = kk(UBound(kk) - 0)
							If LCase(strExtension) = LCase(G.Name) Then
								strDir2 = strDir & G.Name
								
								If objFSO.FolderExists(strDir2) Then
									Set objFolder = objFSO.GetFolder(strDir2)
									Else
									Set objFolder = objFSO.CreateFolder(strDir2)
								End If
								
								MyString = Mid(objFile.Name,3)
								X2 = left(X, 1)
								FinalString = "\\" & strComputer & "\" & X2 & "$" & MyString
								'strCopy = strDir2 & "\" & strCom 'Original AD Copy entry
								strCopy = strDir2 & "\" & strComputer & "_" & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "." & objFile.Extension
								wscript.Echo FinalString & " Copied To..."
								WScript.Echo strCopy5
								wscript.echo
								objFSO.Copyfile Finalstring, strCopy, TRUE
								objReportFile.Writeline("PST Successfully Copied")
								
								objReportFile.Writeline(objFile.Name & " Copied to...")
								objReportFile.Writeline(strCopy)
								objReportFile.WriteBlankLines(1)
								
								If errResult <> 0 Then
									objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
									Wscript.Echo "Result1: " & errResult
								End If
							Else
								WScript.echo "LegacyExchangeDN Mismatch for file owner: " & G.Name 
								wscript.echo "Error: " & Hex(ADS_PROPERTY_NOT_FOUND)
								wscript.echo
								objReportFile.WriteLine ("LegacyExchangeDN Mismatch! for file owner: " & G.Name)
								objReportFile.WriteLine ("Error: " & Hex(ADS_PROPERTY_NOT_FOUND))
								MyString = Mid(objFile.Name,3)
								X2 = left(X, 1)
								FinalString = "\\" & strComputer & "\" & X2 & "$" & MyString
								strDir4 = strDrive & "\" & "\" & "CommVault PST Staging Directory" & "\" & "LegacyExchangeDN MISMATCH" & "\"
								If objFSO.FolderExists(strDir4) Then
								Set objFolder = objFSO.GetFolder(strDir4)
								Else
								Set objFolder = objFSO.CreateFolder(strDir4)
								End If
								strDir3 = strDir4 & G.Name & "\"
										
								If objFSO.FolderExists(strDir3) Then
									Set objFolder = objFSO.GetFolder(strDir3)
								Else
									Set objFolder = objFSO.CreateFolder(strDir3)
								End If
									strCopy5 = strDir3 & strComputer & "_" & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "." & objFile.Extension
									wscript.Echo FinalString & " Copied To..."
									WScript.Echo strCopy5
									wscript.echo
									objFSO.Copyfile FinalString, strCopy5
									objReportFile.Writeline("PST Successfully Copied")
									objReportFile.Writeline(objFile.Name & " Copied to...")
									objReportFile.Writeline(strCopy5)
									objReportFile.WriteBlankLines(1)
								
								If errResult <> 0 Then
									objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
									Wscript.Echo "Result1: " & errResult
									objReportFile.WriteBlankLines(1)
								End If	
							End If
						t = t + 1
						Next
					Else
					wscript.echo "File Modification No Match - Don't Copy"
					wscript.echo objFile.Name
					wscript.echo
					objReportFile.Writeline("File Modification No Match - Don't Copy")
					objReportFile.Writeline(objFile.Name)
					objReportFile.WriteLine
					End If
					End If
				Next
			Next
		Else
			WScript.Echo strComputer & " did not respond to ping."
			WScript.Echo
			objReportFile.Writeline(strComputer & " did not respond to ping. Not scanning for PST's")
			objReportFile.WriteBlankLines(1)
		End If
		objRecordSet.MoveNext
		wscript.echo
	Loop
	'close log file and remove network drive after copying.
	objReportFile.Close
	objNetwork.RemoveNetworkDrive strDrive

End Sub 'End Sub LDAPCopy

'-----------------------------------------------------------------------------------------------
'COPY FILES ALREADY MIGRATED TO EXMERGE STAGING DIRECTORY

Sub SetupExmerge

	objReportFile.WriteLine("/E Copy to Exmerge Staging Directory Option Ran")
	
		If objStartFolder = "" Then
			wscript.echo "Need /Source and /Destination: Value!"
			objReportFile.WriteLine("Need /Source and /Destination: Value!")
			wscript.quit
		End If
		
		Dim strCopy1, errResult, count, col
		count = 0
		Set objFolder = objFSO.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		ShowSubfolders objFolder

End Sub

Sub ShowSubFolders(Folder)

on error resume next

	If strcopy1 = "" Then
		wscript.echo "Need /Source and /Destination: Value!"
		objReportFile.WriteLine("Need /Source and /Destination: Value!")
		wscript.quit
	End If
	
		For Each Subfolder in Folder.SubFolders
			Set objFolder = objFSO.GetFolder(Subfolder.Path)
			NewFileName = Subfolder.Name
			Set colFiles = objFolder.Files
			
			For Each objFile in colFiles
				strCopy2 = strcopy1 & NewFileName & ".pst"
				errResult = objFSO.MoveFile(objFile, strCopy2)

				If errResult.number <> 0 Then
					wscript.echo errResult.number & " PST Already Exists. Please remove first!"	
					err.clear
				End If
				count = count + 1
				'The loop only runs through once to overwrite existing file since each pst in source will be copied to destination with same name, i.e. user/mailbox alias, per Exmerge requirements
				Exit For
			Next
			ShowSubFolders Subfolder
		Next

End Sub 'End Sub SetupExmerge

'-----------------------------------------------------------------------------------------------
'Reboot remote workstations to force logon/logoff for deploy logon script to delete PST's

Sub RebootRemoteComputer

	If InputFile = "" Then
		wscript.echo "Need /InputFile: Value!"
		objReportFile.WriteLine("Need /InputFile: Value!")
		wscript.quit
	End If
	
	If objFso.FileExists(InputFile) = False Then
		wscript.echo "Need Valid Input!"
		wscript.quit
    End If

	objReportFile.WriteLine("/S Reboot Remote Computers Option Ran")
	Const ForReading = 1
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objReportFile2 = objFSO.OpenTextFile(InputFile, ForReading, True)

		Do While objReportFile2.AtEndOfStream <> True
	    	strComputer = objReportFile2.ReadLine
			Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
			strPingResults = LCase(WshExec.StdOut.ReadAll)
		
			If InStr(strPingResults, "reply from") Then
				WScript.Echo strComputer & " responded to ping. It will be rebooted!" 
				WScript.Echo
				objReportFile.WriteLine(strComputer & " responded to ping. It will be rebooted!")
				strShutdown = "shutdown -r -t 0 -f -m \\" & strComputer 
				set objShell = CreateObject("WScript.Shell") 
				objShell.Run strShutdown 
			Else
				WScript.Echo strComputer & " did not respond to ping."
				WScript.Echo
				objReportFile.WriteLine(strComputer & "did not respond to ping.")
			End If
		Loop

End Sub 'End Sub RebootRemoteComputer

'-----------------------------------------------------------------------------------------------
'Copy CloseOutlook.vbs script to remote workstations

Sub DeployLoginScript

	If InputFile = "" Then
		wscript.echo "Need /InputFile: Value!"
		objReportFile.WriteLine("Need /InputFile: Value!")
		wscript.quit
	End If
	
	If objFso.FileExists(InputFile) = False Then
		wscript.echo "Need Valid Input!"
		wscript.quit
    End If

	Const ForReading = 1
	objReportFile.WriteLine("/D Deploy PST-Delete Logon Script Option Ran")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	CurrentPath = objFSO.GetAbsolutePathName(".")
	
	Set objReportFile2 = objFSO.OpenTextFile(InputFile, ForReading, True)

	Do While objReportFile2.AtEndOfStream <> True
		strComputer = objReportFile2.ReadLine
		Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
		strPingResults = LCase(WshExec.StdOut.ReadAll)
		
		If InStr(strPingResults, "reply from") Then
			WScript.Echo strComputer & " responded to ping." 
			WScript.Echo
			objReportFile.WriteLine(strComputer & " responded to ping. Scanning for PST's...")
			objReportFile.WriteBlankLines(1)
			ScriptFolder = "\Documents and Settings\All Users\Start Menu\Programs\Startup\DeleteLocalPSTs.vbs"
			ScriptFolder2 = "\Documents and Settings\All Users\Start Menu\Programs\Startup"
			X2 = "C$"
			FinalString = "\\" & strComputer & "\" & X2 & ScriptFolder
			FinalString2 = "\\" & strComputer & "\" & X2 & ScriptFolder2
			'wscript.echo FinalString
			If objFSO.FileExists(FinalString) Then
			
				wscript.echo "DeleteLocalPSTs.vbs already exists in " & strcomputer & " All Users Startup Directory"
				wscript.echo
				objReportFile.WriteLine("DeleteLocalPSTs.vbs already exists in " & strcomputer & " All Users Startup Directory")

			Else
			
				set objShell2 = CreateObject("Shell.Application")
				Set objFolder = objShell2.NameSpace(FinalString2) 
				ScriptPath = CurrentPath & "\DeleteLocalPSTs.vbs"
				objFolder.CopyHere ScriptPath
				objReportFile.WriteLine("DeleteLocalPSTs.vbs Copied To " & FinalString)
				wscript.echo "DeleteLocalPSTs.vbs Copied To " & FinalString
				wscript.echo

			End If
		Else
			WScript.Echo strComputer & " did not respond to ping."
			WScript.Echo
			objReportFile.WriteLine(strComputer & "did not respond to ping.")
		End If
	Loop
	objReportFile.Close
	
End Sub 'End Sub DeployLoginScript

'-----------------------------------------------------------------------------------------------
'Delete CloseOutlook.vbs script on remote workstations

Sub DeleteLoginScript

	If InputFile = "" Then
		wscript.echo "Need /InputFile: Value!"
		objReportFile.WriteLine("Need /InputFile: Value!")
		wscript.quit
	End If

	If objFso.FileExists(InputFile) = False Then
		wscript.echo "Need Valid Input!"
		wscript.quit
    End If
	
	Const ForReading = 1
	objReportFile.WriteLine("/DL Delete PST-Delete Logon Script Option Ran")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	X2 = "C$"
	Set objReportFile2 = objFSO.OpenTextFile(InputFile, ForReading, True)
	ScriptFolder = "\Documents and Settings\All Users\Start Menu\Programs\Startup\DeleteLocalPSTs.vbs"
	Do While objReportFile2.AtEndOfStream <> True
		strComputer = objReportFile2.ReadLine
		Set WshExec = objShell.Exec("ping -n 3 -w 2000 " & strComputer) 'send 3 echo requests, waiting 2secs each
		strPingResults = LCase(WshExec.StdOut.ReadAll)
		
		If InStr(strPingResults, "reply from") Then
			FinalString = "\\" & strComputer & "\" & X2 & ScriptFolder
			WScript.Echo strComputer & " responded to ping." 
			WScript.Echo
			objReportFile.WriteLine(strComputer & " responded to ping. Scanning for PST's...")
			objReportFile.WriteBlankLines(1)
			If objFSO.FileExists(FinalString) Then
				'wscript.echo FinalString
				objFSO.DeleteFile FinalString
				wscript.echo "Deploy PST Deletion Script " & FinalString & " Deleted"
				wscript.echo
				objReportFile.WriteLine("Deploy PST Deletion Script " & FinalString & " Deleted")
			Else
				wscript.echo "All Deploy PST Deletion Scripts Already Deleted"
				wscript.echo
				objReportFile.WriteLine("All Deploy PST Deletion Scripts Already Deleted")

			End If
		
		Else
			WScript.Echo strComputer & " did not respond to ping."
			WScript.Echo
			objReportFile.WriteLine(strComputer & "did not respond to ping.")
		End If
	Loop
	objReportFile.Close

End Sub 'End Sub DeleteLoginScript


'-----------------------------------------------------------------------------------------------
'Translate NT4 names to dn

function TranslateNT4toDn(nt4name)
dim nto
dim result
const ADS_NAME_INITTYPE_SERVER = 2
const ADS_NAME_INITTYPE_GC = 3
const ADS_NAME_TYPE_1779 = 1
const ADS_NAME_TYPE_NT4 = 3

TranslateNT4toDn=""
set nto = CreateObject("NameTranslate")
'on error resume next
nto.Init ADS_NAME_INITTYPE_GC , ""
nto.set ADS_NAME_TYPE_NT4, nt4name

' If fails return a blank value
'
TranslateNT4toDn = nto.Get(ADS_NAME_TYPE_1779)
'WScript.Echo TranslateNT4toDn
on error goto 0
end function 'End function TranslateNT4toDn

Function InArray(strValue,arrName)
            Dim i
            For i = 0 to UBound(arrName)
                        If InStr(UCase(strValue),UCase(arrName(i))) and len(arrName(i))>1 Then
                                    InArray = TRUE
                                    Exit Function
                        End If
            Next
            InArray = FALSE 
End Function

'-----------------------------------------------------------------------------------------------
'USAGE INSTRUCTIONS

Sub Usage()

  WScript.Echo "/E - Copy PST files after migration to ExMerge staging directory"
  WScript.Echo "/Source - Source directory for /E option"
  WScript.Echo "/Destination - Destination directory for /E option"
  WScript.Echo "/S - Reboot computers specified in /InputFile to force users to logon. To be used after /D and before /DL options."
  WScript.Echo "/D - Deploy Delete PST login script on computers specified in /InputFile. The closeoutlook.vbs script should be in the same directory as this script"
  WScript.Echo "/DL - Delete Delete PST login script on computers specified in /InputFile"
  Wscript.Echo "/R - Generates user specified local report. Will be in tab-separated format for input into spreadsheet"
  Wscript.Echo "/RL - Generates user specified local report. Will be in tab-separated format for input into spreadsheet. Uses /InputFile"
  Wscript.Echo "/Report - Specifies report file to be created per /R option. Directory Must Exist!"
  WScript.Echo "/G - Get ExchangeLegacyDN info for /ADPath: and create /Report:"
  WScript.Echo "/Server - a server where share is located for PST's to be copied to"
  WScript.Echo "/Share - a share name for PST's to be copied to"
  WScript.Echo "/ADPath - LDAP Path for use with /R, /C /G options"
  WScript.Echo "/InputFile - File containing list of computers, one per line, as input"
  WScript.Echo "/MTime - How many days since file modification. For use with copy options."
  WScript.Echo "/C - Copy all pst files from computers found in LDAP path"
  WScript.Echo "/CL - Copy all pst files specified in computer list input file"  
  WScript.Echo "/P - Copy all pst files specified in computer list input file and maintain file ownership. For use with PSTDiscovery tool"
  WScript.Echo "/L - Create computer list file for input file in log directory, c:/pstmgrlogs/computerlist"
  WScript.Echo "Usage: cscript pstmgr.vbs [/Server:<Server Name> /Share:<Share Name>] [/Source:<Source Directory> /Destination:<Destination Directory] [/InputFile:<Computer List>] [/ADPath:<LDAP Path>] [/Report:<Report File>] [/R] [/C] [/P] [/CL] [/E] [/L] [/MTime] [/D] [/DL] [/S]"
  
  
End Sub



'------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' END SCRIPT
'
'------------------------------------------------------------------------------------------------------------------------------------------------------------