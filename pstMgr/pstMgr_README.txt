VBScript

CommVault Professional Services - Mark Richardson
2009

*****CHANGES SECTION*****

11/15/10 - /RL option added to allow for reporting against an input file.
11/17/10 - File and folder filters capability added. Filter.txt created automatically in c:\pstmgrlogs and always used. Can be modified to add additional filters

*****END CHANGES SECTION*****

pstMgr.vbs
DeleteLocalPSTs.vbs

Script Usage: (/R /Report: /C /CL /P /InputFile: /ADPath: /Server: /Share: /Source: /Destination: /L /S /D /DL /E /MTime: /G /? /H

/? or /H switch will give usage instructions

This script derives security from the user logged in running the script. 
Run the script while logged in as a Domain Admin
Make sure that there are write priviledges on destination share especially with /M option!.
Logon Script, DeleteLocalPSTs.vbs is deployed to startup folder for all users. Each user with PSTs on workstation will have to logon for deletion of PSTs
Must be run via cmd as cscript pstmgr.vbs <OPTIONS>
DeleteLocalPSTs.vbs must exist in the same directory as pstMgr.vbs
Directories must preexist for /R, /G options.
Log will be created as C:/pstmgrlogs/pstmgr.log. Directory will be created automatically
All copy target directories will be created automatically
/L Option will create computerlist file as C:\pstmgrlogs\computerlist
/G option is run against users in LDAP path
/InputFile should be in format of one computer per line

Filters:

A filter's file will be created in c:\pstmgrlogs automatically. If it is deleted it will just be recreated.
It is in the following format with the following default entries and can be edited to add additional filters or remove the default entries.

[file]
exdapsttemplate.pst
unicodepsttemplate.pst
[folder]
c:\$recycle.bin
c:\RECYCLER

[file] filters will filter out any files with the same name regardless of path
[folder] will filter out anything within the specified folder.

Target Directories are as such:
		Options: /C /CL
				IF file owner = legacyExchangeDN THEN Match 
			/server:<TARGET SERVER>/share:<TARGET SHARE>\CommVault PST Staging Directory\<FILE OWNER NAME>
				IF file owner NOT = legacyExchangeDN THEN Mismatch 
			/server:<TARGET SERVER>/share:<TARGET SHARE>\CommVault PST Staging Directory\LegacyExchangeDN MISMATCH\<FILE OWNER NAME>	
		Options: /P
				IF file owner = legacyExchangeDN THEN Match 
			/server:<TARGET SERVER>/share:<TARGET SHARE>\CommVault PSTDiscovery Tool SharedPSTLocation\<WORKSTATION_FILE OWNER_FILE NAME_UNIQUE ID>\filename.pst
				IF file owner NOT = legacyExchangeDN THEN Mismatch 
			/server:<TARGET SERVER>/share:<TARGET SHARE>\CommVault PSTDiscovery Tool SharedPSTLocation\LegacyExchangeDN MISMATCH\<WORKSTATION_FILE OWNER_FILE NAME_UNIQUE ID>\filename.pst
			
Usage: cscript pstmgr.vbs [/Server:<Server Name> /Share:<Share Name>] [/Source:<Source Directory> /Destination:<Destination Directory] 
[/InputFile:<Computer List>] [/ADPath:<LDAP Path>] [/Report:<Report File>] [/R] [/C] [/P] [/CL] [/E] [/L] [/MTime:<Time in Days>] [/G] [/D] [/DL] [/S]
	/Server:  a server where share is located for PSTs to be copied to
	/Share:  a share for PSTs to be copied to
	/R  Generates user specified local report. Will be in tabseparated format for input into spreadsheet. For use with /Report and /ADPath option
	/RL  Generates user specified local report. Will be in tabseparated format for input into spreadsheet. For use with /Report and /InputFile option
	/G  Get ExchangeLegacyDN info for /ADPath and create LegacyExchangeDN/user match report file
	/C  Copy all pst files from computers found in LDAP path to /Server /Share
	/Report:  Specifies report file to be created per /R option
	/P  Copy all pst files specified in computer list input file to /Server /Share and maintain file ownership. For use with PSTDiscovery tool
	/CL  Copy all pst files specified in computer list input file to /Server /Share
	/MTime:  How many days since file modification. For use with copy and report options. For use with /CL /C /P /R options
	/L  Create computer list file for input file in log directory, c:/pstmgrlogs/computerlist
	/E  Copy PST files after migration to ExMerge staging directory. NOTE: Not Necessary post 7.0!
	/Source:  Source directory for /E option
	/Destination:  Destination directory for /E option
	/ADPath:  LDAP Path for use with /R, /C /G /L options
	/InputFile:  File containing list of computers, one per line, as input. For use with /CL /P /S /D /DL options 
	/S  Reboot computers specified in /InputFile to force users to logon. To be used after /D and before /DL options.
	/D  Deploy Delete PST login script on computers specified in /InputFile. The DeleteLocalPSTs.vbs script should be in the same directory as this script
	/DL  Delete Delete PST login script on computers specified in /InputFile

REPORT OPTIONS -

Report on PSTs found on workstations.
 		
		cscript pstmgr.vbs /R /Report:c:<Report File> /ADPath:<LDAP Path> 
		ex. cscript pstmgr.vbs /R /Report:C:\reports\pstreport.tsv /ADPath:LDAP://ou=workstations,DC=company,DC=corp
		ex. cscript pstmgr.vbs /R /MTime:90 /Report:C:\reports\pstreport.tsv /ADPath:LDAP://ou=workstations,DC=company,DC=corp

		cscript pstmgr.vbs /RL /Report:c:<Report File> /InputFile:<Input file path>
		ex. cscript pstmgr.vbs /RL /Report:C:\reports\pstreport.tsv /InputFile:C:\pstmgrlogs\computerlist
		ex. cscript pstmgr.vbs /RL /MTime:90 /Report:C:\reports\pstreport.tsv /InputFile:C:\pstmgrlogs\computerlist
		
		cscript pstmgr.vbs /G /Report:c:<Report File> /ADPath:<LDAP Path> 
		ex. cscript pstmgr.vbs /G /Report:C:\reports\userlist.tsv /ADPath:LDAP://ou=pstusergroup,DC=company,DC=corp

COPY OPTIONS -

Copy PSTs from workstations to central share using LDAP Path.
 		
		cscript pstmgr.vbs /C /ADPath:<LDAP Path> /Server:<Destination Server> /Share:<Destination Folder>
		ex. >cscript pstmgr.vbs /C /ADPath:LDAP://ou=workstations,DC=company,DC=corp /Server:Server1 /Share:pstshare
		ex. >cscript pstmgr.vbs /C /MTime:180 /ADPath:LDAP://ou=workstations,DC=company,DC=corp /Server:Server1 /Share:pstshare

Copy PSTs from workstations to central share using Input File.
   A file with one computer per line can be used as input file for selection of computers the script will run against.
   The computer list file created with the /L option can be used or any file in the right format of one computer per line.
 		
		cscript pstmgr.vbs /CL /InputFile:<Input file path> /Server:<Destination Server> /Share:<Destination Folder>
		ex. >cscript pstmgr.vbs /CL /InputFile:C:\pstmgrlogs\computerlist /Server:Server1 /Share:pstshare
		ex. >cscript pstmgr.vbs /CL /MTime:180 /InputFile:C:\pstmgrlogs\computerlist /Server:Server1 /Share:pstshare

Copy PSTs from workstations to central share using Input File. Prepares for PSTDiscovery tool
   A file with one computer per line can be used as input file for selection of computers the script will run against.
   The computer list file created with the /L option can be used or any file in the right format.

		script pstmgr.vbs /P /InputFile:<Input file path> /Server:<Destination Server> /Share:<Destination Folder>
		ex. >cscript pstmgr.vbs /P /InputFile:C:\pstmgrlogs\computerlist /Server:Server1 /Share:pstshare
		ex. >cscript pstmgr.vbs /P /MTime: 180 /InputFile:C:\pstmgrlogs\computerlist /Server:Server1 /Share:pstshare

DELETE PST OPTIONS -

Outlook Profile only closed properly for default profile of user logging on!!
PST's cannot be recovered once deleted!!!

Reboot remote workstations specified in /InputFile to force logon. To be used with /DL and /D options
 		
		cscript pstmgr.vbs /S /InputFile:<Input file path>
		ex. cscript pstmgr.vbs /S /InputFile:C:\pstmgrlogs\computerlist

Deploy DeleteLocalPSTs.vbs logon script to remote workstations to be specified in /InputFile. DeleteLocalPSTs.vbs must be in same directory as this script
 		
		cscript pstmgr.vbs /D /InputFile:<Input file path>
		ex. cscript pstmgr.vbs /D /InputFile:C:\pstmgrlogs\computerlist

Delete DeleteLocalPSTs.vbs logon script.
 		
		cscript pstmgr.vbs /DL /InputFile:<Input file path>
		ex. cscript pstmgr.vbs /DL /InputFile:C:\pstmgrlogs\computerlist
		
VARIOUS OPTIONS -

Make computer list file to be used as template for input file for /InputFile option. 
	It will find all computers in LDAP path and then create file with one computer per line. The file can then be modified to remove or add computers. The file is created
	as c:\pstmgrlogs\computerlist
 		
		cscript pstmgr.vbs /L /ADPath:<LDAP Path> 
		ex. cscript pstmgr.vbs /L /ADPath:LDAP://ou=pstgroup,DC=company,DC=corp

Move PSTs after migration to staging folder for Exmerge which is necessary to get stubs from PST to mailbox with same folder structure
 		NOTE: This option is not necessary post 7.0!!
		cscript pstmgr.vbs /E /Source:<Source Folder> /Destination:<Destination Folder>
		ex. cscript pstmgr.vbs /E /Source:C:\Input Folder /Destination:C:\Output Folder\
		 	It is necessary that their be a trailing \ to the Destination path or else the script will not work correctly!!!!!!!!!!!!


