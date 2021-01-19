'VBScript

'Hound.vbs by CommVault Professional Services
'edited on 11/17/2010 to fix upper case/lower case mismatch issue with LegacyExchangeDN comparison

'Should be added as a Logon Script to a Group Policy containing users.
'Will copy, delete, report on PST files found on computer that user logs onto.

'Run with /? or /H to get usage information

'The move option /M close profile functionality only works against the default MAPI profile. All PST's will be deleted but other associated profiles will need to have their PST's closed manually.
'The /C /M options will create folder structure necessary for Exchange Mailbox Archiver. Will verify file owner legacyexchangedn and if not match then segment.
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' START SCRIPT
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
On Error Resume Next
err.clear

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
WScript.Echo date
Wscript.echo
WScript.Echo "/? or /H for Usage Instructions"
Wscript.echo

Const ADS_PROPERTY_NOT_FOUND = &h8000500D

Set colNamedArguments = WScript.Arguments.Named
If colNamedArguments.Exists("H") OR colNamedArguments.Exists("?") Then
  Call Usage()
  WScript.Quit
End If

strServer = colNamedArguments.Item("Server")
strShare = colNamedArguments.Item("Share")
booReport = colNamedArguments.Exists("R")
booCopy = colNamedArguments.Exists("C")
booMove = colNamedArguments.Exists("M")
booPrepare = colNamedArguments.Exists("P")
strMod = colNamedArguments.Item("MTime")

strMyComputer = "."

'-----------------------------------------------------------------------------------------------------------------
'
' Create a report of PST's found in C:\Windows\pst_<host name>.tsv
'
'-----------------------------------------------------------------------------------------------------------------

Set objWMI = GetObject("winmgmts:\\" & strMyComputer & "\root\cimv2")						' Using WMI Class
Set objNetwork = Wscript.CreateObject("WScript.Network")									' Using WSH Network Object
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")									' Needed to convert UTC to Normal Date
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set tFolder = objFSO.GetSpecialFolder(TemporaryFolder)										' Creating a name for a temp file in temp dir
Set objShell = Wscript.CreateObject("wscript.shell")
tName = objFSO.GetTempName
rFileName = tFolder.Path & "\pst_" & objNetwork.ComputerName & ".tsv"


Set objReportFile = objFSO.CreateTextFile(rFileName,True)									' Creating the report file in temp dir
wscript.echo "Created report file: " & rFileName
wscript.echo
objReportFile.Write(vbTab & vbTab & vbTab & strHeader1)
objReportFile.WriteBlankLines(2)
objReportFile.WriteLine("Domain" & vbTab & "LegacyExchangeDN" & vbTab & "File Owner" & vbTab & "Computer" & vbTab & "PST File Name " & vbTab & "Size of PST (bytes)" & vbTab & "Date Modified"  & vbTab & "Days Since Modified")


Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")

For Each objDisk in colDisks

	X = objDisk.DeviceID
	Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")

	For Each objFile in colFiles	
		dateTime.Value = objFile.LastModified
		modDate = DateDiff("d", dateTime.GetVarDate(True), Date)

		If int(modDate) => int(strMod) Then
			wscript.echo "File Modification Match"
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
					If err.number = ADS_PROPERTY_NOT_FOUND Then
					LegacyName = "NO VALUE"
					'wscript.echo LegacyName
					End If

					kk = split(LegacyName,"=")
					strExtension = kk(UBound(kk) - 0)
					If LCase(strExtension) = LCase(G.Name) Then
						WScript.echo "LegacyExchangeDN Match!"
						WScript.Echo LegacyName & " " & G.Name
						wscript.echo objFile.Name
						wscript.echo
						objReportFile.WriteLine(G.Domain & vbTab & LegacyName & vbTab & G.Name & vbTab & objNetwork.ComputerName & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True) & vbTab & int(modDate))
					Else
						WScript.echo "LegacyExchangeDN MisMatch!"
						WScript.Echo LegacyName & " " & G.Name
						wscript.echo objFile.Name
						wscript.echo
						objReportFile.WriteLine(G.Domain & vbTab & LegacyName & vbTab & G.Name & vbTab & objNetwork.ComputerName & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True) & vbTab & int(modDate))
					End If
			Next
		Else
		wscript.echo "No Modification Match"
		wscript.echo objFile.Name
		wscript.echo
		objReportFile.WriteLine("NA" & vbTab & 	"NA" & vbTab & "NA" & vbTab & objNetwork.ComputerName & vbTab & objFile.Name & vbTab & objFile.FileSize & vbTab & dateTime.GetVarDate(True) & vbTab & int(modDate))

		End If
	Next
Next

objReportFile.WriteBlankLines(2)
objReportFile.WriteBlankLines(2)

If Not (booReport OR booCopy OR booMove or booPrepare) Then
  WScript.Echo "Report was saved in the file " & rFileName
  objReportFile.Close
  WScript.Quit
End If

'-----------------------------------------------------------------------------
'
'	Determining available letter to mount a share
'
'-----------------------------------------------------------------------------

Set objDictionary = CreateObject("Scripting.Dictionary")
Set colDisks = objWMI.ExecQuery("Select * from Win32_LogicalDisk")				' Win32_LogicalDisk

For Each objDisk in colDisks
    objDictionary.Add objDisk.DeviceID, objDisk.DeviceID
Next

strDrive = " "
For i = 67 to 90																		' Looping through letters starting with A to Z
    If objDictionary.Exists(Chr(i) & ":") Then
    Else
        strDrive = Chr(i) & ":"															' If a letter is available - exit the loop
        Exit For
    End If
Next

If strDrive = " " Then
    Wscript.Echo("There are no available drive letters on this computer.")				' If no letters available - exit the script
    Wscript.Quit
End If


'-------------------------------------------------------------------------------
'
'	Mapping a Network Share
'
'-------------------------------------------------------------------------------

Set objNetwork = Wscript.CreateObject("WScript.Network")								' Using WSH Network Object
objNetwork.MapNetworkDrive strDrive, "\\" & strServer & "\" & strShare			' No username and password if local admin
wscript.echo
WScript.Echo "Mapped Drive:" & strDrive

'--------------------------------------------------------------------------------
'
'	/M Remove PST Stores from within Outlook to avoid error and Kill Outlook.exe
'	
'--------------------------------------------------------------------------------

If booMove Then
	wscript.echo "Closing PST's in Outlook for default profile only!"
	Const olFolderCalendar = 9
	Const olFolderInbox = 6
	set olApp = CreateObject("Outlook.Application")
	set olNameSpace =olApp.GetNameSpace("MAPI")
	rootStoreID = olNameSpace.GetDefaultFolder(olFolderInbox).parent.storeId


	WScript.echo "Closing any opened .pst file to avoid conflict"
	Dim i, temp
	For i = olNameSpace.Folders.count To 1 Step -1
		temp = olNameSpace.Folders(i).storeID
		If Left(temp,75) <> Left(rootStoreID,75) Then
			' === At least the first 75 digits of the rootStoreID
			'     are the same for items that aren’t Personal Folders.
			'     Since they're not equal, this must be a 
			'     Personal Folder. Close it.
			olNameSpace.RemoveStore olNameSpace.Folders(i)
		End If
	Next

	Set colItems = objWMI.ExecQuery ("Select * From Win32_Process Where Name = 'outlook.exe'")
	wscript.echo "Killing Outllook.exe if running"
	For Each Process in colItems
		errResult = Process.Terminate															
		If errResult <> 0 Then
		  objReportFile.WriteLine("Unable to Terminate Outlook. Error: " & errResult)
		  Wscript.Echo "Termination Result: " & errResult
		  WScript.Quit
		End If
	Next
	
	WScript.Sleep(10000)

End If

'------------------------------------------------------------------------------
'
'	Copy or Move pst Files. /M Will delete all PSTs found!!
'
'------------------------------------------------------------------------------

If booCopy or booMove Then
	
	On Error Resume Next
	err.clear
	
	wscript.echo "Copying Files..."
	wscript.echo
	Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")
	strDir = strDrive & "\" & "\" & "CommVault PST Staging Directory" & "\"

	If objFSO.FolderExists(strDir) Then
		Set objFolder = objFSO.GetFolder(strDir)
		Else
		Set objFolder = objFSO.CreateFolder(strDir)
		wscript.echo "Created directory: " & strDir
	End If


	For Each objDisk in colDisks

		X = objDisk.DeviceID
		Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")

		For Each objFile in colFiles																
			dateTime.Value = objFile.LastModified

			modDate = DateDiff("d", dateTime.GetVarDate(True), Date)

			If int(modDate) => int(strMod) Then
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
					If err.number = ADS_PROPERTY_NOT_FOUND Then
						LegacyName = "NO VALUE"
						'wscript.echo LegacyName
					End If
					kk = split(LegacyName,"=")
					strExtension = kk(UBound(kk) - 0)
					
					If LCase(strExtension) = LCase(G.Name) Then

						strDir2 = strDir & G.Name
						
						If objFSO.FolderExists(strDir2) Then
							Set objFolder = objFSO.GetFolder(strDir2)
							Else
							Set objFolder = objFSO.CreateFolder(strDir2)
							wscript.echo "Created directory: " & strDir2
						End If
						
						x = objFile.Name
						strCopy = strDir2 & "\" & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "." & objFile.Extension
						errResult = objFile.Copy(strCopy)
						wscript.echo obFile.Name & " Copied To..."
						wscript.echo strCopy

						If errResult <> 0 Then
							objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
							Wscript.Echo "Result1: " & errResult
						End If
						
						If booMove Then
							wscript.echo "Deleting PSTs..."
							errResult = objFile.Delete															' Delete copied files
							If errResult <> 0 Then
								objReportFile.WriteLine("Unable to delete file " & objFile.Name & ". Error: " & errResult)
								Wscript.Echo "Result2: " & errResult
							End If
						End If
					Else

						strDir3 = strDir & "LegacyExchangeDN MISMATCH\"
						wscript.echo strDir3
						If objFSO.FolderExists(strDir3) Then
							Set objFolder = objFSO.GetFolder(strDir3)
							Else
							Set objFolder = objFSO.CreateFolder(strDir3)
							wscript.echo "Created directory: " & strDir3
						End If
						
						strDir4 = strDir3 & G.Name
						wscript.echo strDir3
						If objFSO.FolderExists(strDir4) Then
							Set objFolder = objFSO.GetFolder(strDir4)
							Else
							Set objFolder = objFSO.CreateFolder(strDir4)
							wscript.echo "Created directory: " & strDir4
						End If
						
						x = objFile.Name
						
						strCopy = strDir4 & "\" & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "." & objFile.Extension
						wscript.echo strCopy
						errResult = objFile.Copy(strCopy)
						wscript.echo obFile.Name & " Copied To..."
						wscript.echo strCopy

						If errResult <> 0 Then
							objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
							Wscript.Echo "Result1: " & errResult
						End If
						
						If booMove Then
							errResult = objFile.Delete															' Delete copied files
							If errResult <> 0 Then
								objReportFile.WriteLine("Unable to delete file " & objFile.Name & ". Error: " & errResult)
								Wscript.Echo "Result2: " & errResult
							End If
						End If
						
					End If
						
				Next			
			End If
		Next
	Next

End If

'------------------------------------------------------------------------------
'
'	Copy/Delete pst Files as preparation for PSTDiscovery Tool
'
'------------------------------------------------------------------------------

If booPrepare Then

	On Error Resume Next
	err.clear

	strDir = strDrive & "\" & "\" & "CommVault PSTDiscovery Tool SharedPSTLocation" & "\"

	If objFSO.FolderExists(strDir) Then
		Set objFolder = objFSO.GetFolder(strDir)
		Else
		Set objFolder = objFSO.CreateFolder(strDir)
	End If

	Set colDisks = objWMI.ExecQuery("Select * FROM Win32_LogicalDisk where DriveType = 3")

	For Each objDisk in colDisks

		X = objDisk.DeviceID
		Set colFiles = objWMI.ExecQuery ("Select * FROM CIM_Datafile Where Extension = 'pst'AND Drive='" & X & "' ")

		For Each objFile in colFiles																' Output to report file with PST file name and attributes
			dateTime.Value = objFile.LastModified
			modDate = DateDiff("d", dateTime.GetVarDate(True), Date)

			If int(modDate) => int(strMod) Then

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
					
					If err.number = ADS_PROPERTY_NOT_FOUND Then
						LegacyName = "NO VALUE"
						wscript.echo LegacyName
					End If
					
					kk = split(LegacyName,"=")
					strExtension = kk(UBound(kk) - 0)
					If LCase(strExtension) = LCase(G.Name) Then
						WScript.echo "LegacyExchangeDN Match!"
						WScript.Echo LegacyName
						wscript.echo
					
						strDir2 = strDir & objNetwork.ComputerName & "_" & G.Name & "_"   
						strCopy = strDir2 & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "\"
						'strCopy = strDrive	   
						x = objFile.Name
						errResult = objShell.Run("XCOPY """ & x & """ """ & strCopy & """ /o /i", 2, True)
							
						If errResult <> 0 Then
							objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
							Wscript.Echo "Result1: " & errResult
						End If
						
						If booMove Then
							errResult = objFile.Delete															' Delete copied files
							If errResult <> 0 Then
								objReportFile.WriteLine("Unable to delete file " & objFile.Name & ". Error: " & errResult)
								Wscript.Echo "Result2: " & errResult
							End If
						End If
					Else
						WScript.echo "LegacyExchangeDN MisMatch!"
						WScript.Echo LegacyName
						wscript.echo
						  

						strDir3 = strDir & "LegacyExchangeDN MISMATCH\"
						wscript.echo strDir3
						If objFSO.FolderExists(strDir3) Then
							Set objFolder = objFSO.GetFolder(strDir3)
							Else
							Set objFolder = objFSO.CreateFolder(strDir3)
						End If
						strDir2 = strDir3 & objNetwork.ComputerName & "_" & G.Name & "_" 
						x = objFile.Name
						
						strCopy = strDir2 & objFile.FileName & "_" & DateDiff("s", #1-1-1970 00:00:00#, Now) & "\"
						'strCopy = strDrive	   
						x = objFile.Name
						errResult = objShell.Run("XCOPY """ & x & """ """ & strCopy & """ /o /i", 2, True)

						If errResult <> 0 Then
							objReportFile.WriteLine("Unable to copy file " & objFile.Name & ". Error: " & errResult)
							Wscript.Echo "Result1: " & errResult
						End If
						
						If booMove Then
							errResult = objFile.Delete															' Delete copied files
							If errResult <> 0 Then
								objReportFile.WriteLine("Unable to delete file " & objFile.Name & ". Error: " & errResult)
								Wscript.Echo "Result2: " & errResult
							End If
						End If
					End If
				Next
			End If
		 Next
	Next

End If

'-----------------------------------------------------------------------------------------------
'Translate NT4 names to dn
'
'-----------------------------------------------------------------------------------------------

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


'---------------------------------------------------------------------------------------------------
'
' Copy final report
'
'---------------------------------------------------------------------------------------------------

objReportFile.Close
strDir = strDrive & "\" & "\" & "CommVault PST Reports" & "\"

If objFSO.FolderExists(strDir) Then
	Set objFolder = objFSO.GetFolder(strDir)
	Else
	Set objFolder = objFSO.CreateFolder(strDir)
End If

Set objReportFile = objFSO.GetFile(rFileName)
objReportFile.Copy(strDir & "\pst_" & objNetwork.ComputerName & ".tsv")
wscript.echo "Report File Copied"
wscipt.echo
objNetwork.RemoveNetworkDrive strDrive
wscript.echo "Network Drive Removed"
wscript.echo
WScript.Quit

Sub Usage()
  WScript.Echo "Usage: cscript hound [/Server:<Server Name> /Share:<Share Name>] [/C] [/D] [/P] [/S] [/R] [/MTime]"
  WScript.Echo "/Server - Server to copy PST's to"
  WScript.Echo "/Share - a CIFS share to copy reports and/or PST's to"
  WScript.Echo "/C - copy all pst files from local machine to central repository and create pst archiving folder structure"
  WScript.Echo "/P - copy all pst files from local machine to central repository and maintain file ownership for use with PSTDiscovery tool"
  WScript.Echo "/M - move pst files from local machine to central repository and create pst archiving folder structure and close any PSTs from default MAPI profile. Warning: Will delete all PSTs!!"
  WScript.Echo "/R - copy report file to /Server /Share"
  WScript.Echo "/MTime: - Only copy and/or report on PST's not modified in number of days specified."
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' END SCRIPT
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
