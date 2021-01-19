VBScript

CommVault Professional Services - Mark Richardson
2009
'edited on 11/17/2010 to fix upper case/lower case mismatch issue with LegacyExchangeDN comparison


Please rename the script files from .vbscript to .vbs
The file names were changed as a convenience to prevent the files from being blocked by most email scanners when emailing the package.

hound.vbs

Should be added as a Logon Script to a Group Policy.
Will copy, delete, report on PST files found on computer that user logs onto.
Users must be local administrators to their workstations.
Aggregation.xls can be used to collate reports into one spreadsheet. It is necessary to open macro and edit scan directory.

Run with /? or /H to get usage information

The move option /M close-profile functionality only works against the default MAPI profile. All PSTs will be deleted but other associated profiles will need to have their PSTs closed manually.
The /C /M options will create folder structure necessary for Exchange Mailbox Archiver. 
The /C /M /P /R options will match file owner to legacyexchangedn and if not match then segment.

Usage: cscript hound [/Server:<Server Name> /Share:<Share Name>] [/C] [/M] [/P] [/R] [/MTime:<Number of Days>]
   /Server: - Server to copy PSTs to
   /Share: - a CIFS share to copy reports and/or PSTs to
   /C - copy all pst files from local machine to central repository and create pst archiving folder structure
   /P - copy all pst files from local machine to central repository and maintain file ownership for use with PSTDiscovery tool
   /M - move pst files from local machine to central repository and create pst archiving folder structure and close any PSTs from default MAPI profile. Warning: Will delete all PSTs!!
   /R - copy report file to /Server /Share
   /MTime: - Optional. If not specified then all PST's will be included. Only copy and/or report on PSTs not modified in number of days specified.

Group Policy Object Editor > Scripts(Logon/Logoff) > Logon > Edit Script

Each parameter must be in its own set of quotes as follows:

	cscript.exe
	hound.vbs "/server:labex1"  "/share:PSTShare"  "/R" "MTime:180"

