;Written by KramWell.com - 19/MAR/2018
;Custom built application that significantly decreased the time taken to create a new user from 45 to 3 minutes.

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

ININUC := "NUC.INI"
INIDEP := "DEP.INI"

;Check if ini file exists
if !FileExist(ININUC){
	MsgBox % ININUC " file does not exist. Program will exit."
	Goto GuiClose
}

;Check if site file exists
if !FileExist(INIDEP){
	MsgBox % INIDEP " file does not exist. Program will exit."
	Goto GuiClose
}

IniRead, DOMAIN, %A_ScriptDir%\%ININUC%, ADSETTINGS, DOMAIN
IniRead, COMPANY, %A_ScriptDir%\%ININUC%, ADSETTINGS, COMPANY	
IniRead, PASSWORD, %A_ScriptDir%\%ININUC%, ADSETTINGS, PASSWORD

;Read site values in combobox
combo_item_sites := "--Select--|"
FileRead , data_from_nuc , %A_ScriptDir%\%ININUC%  ;  Reads into a variable
Loop , parse , data_from_nuc , `n  ;  Use a parse loop to add in the necessary delimiter used by comboboxes.
;we need to see if it contains [] if so only show
If InStr(A_LoopField, "[") AND (A_LoopField <> "[xxxxx]`r") AND (A_LoopField <> "[ADSETTINGS]`r"){
	;msgbox % "." A_LoopField "."
	combo_item_sites := combo_item_sites . "|" . A_LoopField 
	StringReplace , combo_item_sites , combo_item_sites , `r , , All  ;  remove the returns so we have a single line.	
	StringReplace , combo_item_sites , combo_item_sites , [ , , All  ;  remove the returns so we have a single line.		
	StringReplace , combo_item_sites , combo_item_sites , ] , , All  ;  remove the returns so we have a single line.			
}

;Read site values in combobox
combo_item_dep := "--Select--|"
FileRead , data_from_sites , %A_ScriptDir%\%INIDEP%  ;  Reads notepad.txt into a variable
Loop , parse , data_from_sites , `n  ;  Use a parse loop to add in the necessary delimiter used by comboboxes.
;we need to see if it containes [] if so only show
If InStr(A_LoopField, "["){
	combo_item_dep := combo_item_dep . "|" . A_LoopField 
	StringReplace , combo_item_dep , combo_item_dep , `r , , All  ;  remove the returns so we have a single line.	
	StringReplace , combo_item_dep , combo_item_dep , [ , , All  ;  remove the returns so we have a single line.		
	StringReplace , combo_item_dep , combo_item_dep , ] , , All  ;  remove the returns so we have a single line.			
}

;Create GUI
Gui,Add,Text,x10 y10 w100 h13,FirstName
Gui,Add,Edit,x10 y30 w100 h21 vFirstName gUPNandSAN,

Gui,Add,Text,x130 y10 w100 h13,LastName
Gui,Add,Edit,x130 y30 w100 h21 vLastName gUPNandSAN,

Gui,Add,Text,x250 y10 w100 h13,SamAccount
Gui,Add,Edit,x250 y30 w100 h21 vSamAccountName,

Gui,Add,Text,x10 y65 w50 h13,UPN
Gui,Add,Edit,x10 y85 w130 h21 vUPN,
Gui,Add,Text,x145 y90 w100 h21,%DOMAIN%

;Gui,Add,Text,x250 y70 w100 h13,Company
;Gui,Add,Text,x250 y90 w100 h21,%COMPANY%

Gui,Add,Text,x200 y65 w100 h13,Job Title
Gui,Add,Edit,x200 y85 w150 h21 vJobTitle,

Gui,Add,Text,x10 y120 w100 h13,Department
Gui,Add,DropDownList,x10 y140 w200 vDepartment, %combo_item_dep%

Gui,Add,Text,x220 y120 w100 h13,Location
Gui,Add,DropDownList,x220 y140 w130 vLocation gLocation, %combo_item_sites%

Gui,Add,Text,x10 y180 w100 h13,Extension
Gui,Add,Edit,x10 y200 w50 h21 vLyncExt,

Gui,Add,Button,x75 y200 w133 h23 gInquire vInquire,Lookup Free Numbers

Gui,Add,Button,x220 y200 w133 h23 gAllNumbers vAllNumbers,Show All Numbers

Gui,Add,Button,x290 y240 w63 h23 gCheckManager,PS Load
Gui,Add,Button,x220 y240 w63 h23 gReloadScript,Reload
Gui,Add,Button,x10 y240 w200 h23 gCreateUser,Create User

GuiControl,Disable, Inquire
GuiControl,Disable, LyncExt

Gui,Show,w360 h275,%COMPANY% - New User Creator. v0.7.2
return

;End Gui

AllNumbers:
	RunWait *RunAs PowerShell.exe "set-executionpolicy remotesigned"
	RunWait PowerShell.exe %A_WorkingDir%\skypeNumbers.ps1
Return

Location:
	GuiControlGet, Location
	
	IniRead, LYNCPREFIX, %A_ScriptDir%\%ININUC%, %Location%, LYNCPREFIX

	GuiControl,Disable, Inquire
	GuiControl,Disable, LyncExt
	
	If (LYNCPREFIX){	
		if (LYNCPREFIX != "ERROR"){
			GuiControl,Enable, Inquire
			GuiControl,Enable, LyncExt			
		}

	}
Return

Inquire:
	;bring up ini read of address details to add
	GuiControlGet, Location
		
	If (Location = "--Select--"){	
		MsgBox % "Please select Location"
		Return ;return if error, continue if not
	}	
	RunWait *RunAs PowerShell.exe "set-executionpolicy remotesigned"
	RunWait PowerShell.exe %A_WorkingDir%\skypeNumbers.ps1 -TypeLimit "SPARE" -LocationLimit %Location%
Return

ReloadScript:	
	Reload	
Return

CheckManager:	
	GuiControlGet, SamAccountName	
	GuiControlGet, LyncExt
	GuiControlGet, Location
	GuiControlGet, UPN

	LyncNumber := ""

	IniRead, LYNCPREFIX, %A_ScriptDir%\%ININUC%, %Location%, LYNCPREFIX

	if (LyncExt){
		If (LYNCPREFIX){	
			if (LYNCPREFIX != "ERROR"){
				;if LYNCPREFIX has value assign lync ext to it as long as that also has value.
				LyncNumber := LYNCPREFIX . LyncExt
			}

		}
	}	
	;MsgBox % LyncNumber

	RunWait *RunAs PowerShell.exe "set-executionpolicy remotesigned"

	If (SamAccountName = "") OR (UPN = ""){	
		RunWait PowerShell.exe %A_WorkingDir%\NewUserCreator.ps1
		Return
	}
	If (LyncExt = "") OR (LyncNumber = ""){
		RunWait PowerShell.exe %A_WorkingDir%\NewUserCreator.ps1 -EmailAddress %UPN%%DOMAIN% -samAccountName %SamAccountName%
		Return
	}
	RunWait PowerShell.exe %A_WorkingDir%\NewUserCreator.ps1 -EmailAddress %UPN%%DOMAIN% -samAccountName %SamAccountName% -LyncNumber %LyncNumber% -LyncExtension %LyncExt%
Return

;[string]$EmailAddress,
;[string]$samAccountName,
;[string]$LyncNumber,
;[string]$LyncExtension

CreateUser:
	
	;$Credentials = Get-Credential ;prompt enter creds to use AD-
	
	;Collect Variables here
	GuiControlGet, FirstName
	GuiControlGet, LastName
	GuiControlGet, LyncExt
	GuiControlGet, SamAccountName

	GuiControlGet, Department

	GuiControlGet, JobTitle

	GuiControlGet, UPN

	;bring up ini read of address details to add
	GuiControlGet, Location

	If (FirstName = "") OR (LastName = "") OR (SamAccountName = "") OR (UPN = "") OR (JobTitle = ""){	
		MsgBox % "You have blank fields, please correct."
		Return ;return if error, continue if not
	}	

	If (Department = "--Select--"){	
		MsgBox % "Please select Department"
		Return ;return if error, continue if not
	}		

	If (Location = "--Select--"){	
		MsgBox % "Please select Location"
		Return ;return if error, continue if not
	}	

	MsgBox, 4, ,% "Are you sure you would like to submit these details?"
	IfMsgBox, No
	Return  ; User pressed the "No" button.			
	Else	
		
	IniRead, DN, %A_ScriptDir%\%ININUC%, %Location%, DN
	IniRead, STREET, %A_ScriptDir%\%ININUC%, %Location%, STREET
	StringReplace, STREET, STREET, :, `r`n, All
	IniRead, CITY, %A_ScriptDir%\%ININUC%, %Location%, CITY
	IniRead, STATE, %A_ScriptDir%\%ININUC%, %Location%, STATE
	IniRead, ZIP, %A_ScriptDir%\%ININUC%, %Location%, ZIP
	IniRead, OFFICE, %A_ScriptDir%\%ININUC%, %Location%, OFFICE
	IniRead, USERDRIVE, %A_ScriptDir%\%ININUC%, %Location%, USERDRIVE

	;add to default groups
	IniRead, DEF_GROUP, %A_ScriptDir%\%ININUC%, %Location%, DEF_GROUP
	IniRead, MAIL_GROUP, %A_ScriptDir%\%ININUC%, %Location%, MAIL_GROUP
	
	IniRead, GROUP_DEPT, %A_ScriptDir%\%INIDEP%, %Department%, GROUP_DEPT

	;create country codes
	IniRead, COUNTRY, %A_ScriptDir%\%ININUC%, %Location%, COUNTRY
	IniRead, 2CODE, %A_ScriptDir%\%ININUC%, %Location%, 2CODE
	IniRead, COUNTRYCODE, %A_ScriptDir%\%ININUC%, %Location%, COUNTRYCODE

	IniRead, LYNCPREFIX, %A_ScriptDir%\%ININUC%, %Location%, LYNCPREFIX

	if (LyncExt){
		If (LYNCPREFIX){	
			if (LYNCPREFIX != "ERROR"){
				LyncNumber := LYNCPREFIX . LyncExt
			}

		}
	}	
	;MsgBox % LyncNumber
		
	TempFile := A_ScriptDir . "\tempOut.kwt"
	FileDelete, %TempFile%
	
	;this is for finding if the username is in use
	psScriptFindUser =
	(
		try {
			Import-Module ActiveDirectory -ErrorAction stop
			Get-ADUser -Identity '%SamAccountName%' -ErrorAction stop -Properties SamAccountName | select -expand SamAccountName
			Write-Output "USREXISTS"
			}
		catch {
			Write-Output $Error[0].ToString()
		}
	)
	
	;this executes the powershell and saves it into a file for viewing
	RunWait, PowerShell.exe -Command &{%psScriptFindUser%} > %TempFile%,,Hide
	FileRead,Output,%TempFile%
	FileDelete,%TempFile%
	
	;check to see if valid
	If InStr(Output, "USREXISTS"){
		MsgBox % "The username '" SamAccountName "' is already in use."
		Return ;return if error, continue if not
	}else{
		If !InStr(Output, "Cannot find an object with identity: '" SamAccountName){	
			MsgBox % Output
			Return ;return if error, continue if not
		}
	}
		
	;at this point we 
	
	psScriptCreateUser =
	(
		try {
			Import-Module ActiveDirectory -ErrorAction stop
			New-ADUser -Name '%LastName%, %FirstName%' -GivenName '%FirstName%' -Surname '%LastName%' -displayName '%LastName%, %FirstName%' -SamAccountName '%SamAccountName%' -Title '%JobTitle%' -Department '%Department%' -Description '%Department%' -Company '%COMPANY%' -StreetAddress """%STREET%""" -City '%CITY%' -State '%STATE%' -PostalCode '%ZIP%' -Office '%OFFICE%' -UserPrincipalName '%UPN%%DOMAIN%' -AccountPassword (ConvertTo-SecureString '%PASSWORD%' -AsPlainText -Force) -Path '%DN%' -Enabled $True -ChangePasswordAtLogon $true -ErrorAction stop
			Set-ADUser '%SamAccountName%' -Replace @{c='%2CODE%'} -ErrorAction stop
			Set-ADUser '%SamAccountName%' -Replace @{co='%COUNTRY%'} -ErrorAction stop
			Set-ADUser '%SamAccountName%' -Replace @{countrycode='%COUNTRYCODE%'} -ErrorAction stop
			Set-ADUser '%SamAccountName%' –Replace @{proxyAddresses = 'SMTP:%UPN%%DOMAIN%'}
			Add-ADGroupMember -Identity '%DEF_GROUP%' -Members '%SamAccountName%' -ErrorAction stop
			Add-ADGroupMember -Identity '%MAIL_GROUP%' -Members '%SamAccountName%' -ErrorAction stop
			Add-ADGroupMember -Identity '%GROUP_DEPT%' -Members '%SamAccountName%' -ErrorAction stop
			Write-Output "CREATEDUSR"%SamAccountName%
		}catch{
			Remove-ADUser -Identity '%SamAccountName%' -Confirm:$False
			Write-Output $Error[0].ToString()
		}	
	)
	
	;this executes the powershell and saves it into a file for viewing
	RunWait, PowerShell.exe -Command &{%psScriptCreateUser%} > %TempFile%,,Hide
	FileRead,Output,%TempFile%
	FileDelete,%TempFile%	
	
	;check to see if valid
	If !InStr(Output, "CREATEDUSR"SamAccountName){
		If (Output = ""){
			msgbox % "ERROR15: Something went wrong!`r`nTIP: Check permissions?"	
			Return ;return if error, continue if not
		}else{
			MsgBox % Output
			Return ;return if error, continue if not
		}
	}		
		
	;all good and created successfully	
	Msgbox % "User " SamAccountName " created successfully`r`nNow creating user share..."	

	Folder := USERDRIVE "\" SamAccountName
	;$UserName = "wellsma" 

	MsgBox % Folder

	TempFile := A_ScriptDir . "\tempOut.kwt"
	FileDelete, %TempFile%

	psScriptCreateFolder =
	(
		try {
			New-Item -ItemType directory -Path '%Folder%' -ErrorAction stop | Out-null
			try {
				$rule=new-object System.Security.AccessControl.FileSystemAccessRule ('%SamAccountName%', 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
				if(Test-Path '%Folder%') {
					try {
						$acl = Get-ACL -Path '%Folder%' -ErrorAction stop
						$acl.SetAccessRule($rule)
						Set-ACL -Path '%Folder%' -ACLObject $acl -ErrorAction stop
						Write-Output "FOLDERCREATED123"
					}catch{
						Remove-Item '%Folder%' -recurse
						Write-Output "ERROR16: " $Error[0].ToString()
					}            
				}else{            
					Write-Output "ERROR17: " $Error[0].ToString()
				}
			}catch{            
				Remove-Item '%Folder%' -recurse
				Write-Output "ERROR18: " $Error[0].ToString()           
			}		
		}catch{            
			Write-Output "ERROR19: " $Error[0].ToString()           
		}
	)
		
	;this executes the powershell and saves it into a file for viewing
	RunWait, PowerShell.exe -Command &{%psScriptCreateFolder%} > %TempFile%,,Hide
	FileRead,Output,%TempFile%
	FileDelete,%TempFile%	

	If !InStr(Output, "FOLDERCREATED123"){
		MsgBox %Output%		
		Return
		;add here option to open folder to create manually?
	}

	MsgBox % "YES! Folder has been created on drive:`r`n" Folder "`r`n`r`nOpening script..."

	RunWait *RunAs PowerShell.exe "set-executionpolicy remotesigned"

	If (LyncExt = "") OR (LyncNumber = ""){
		RunWait PowerShell.exe %A_WorkingDir%\NewUserCreator.ps1 -EmailAddress %UPN%%DOMAIN% -samAccountName %SamAccountName%
		Return
	}
	RunWait PowerShell.exe %A_WorkingDir%\NewUserCreator.ps1 -EmailAddress %UPN%%DOMAIN% -samAccountName %SamAccountName% -LyncNumber %LyncNumber% -LyncExtension %LyncExt%
	
Return

	
UPNandSAN:
	;add to UPN here
	GuiControlGet, FirstName ;get firstnameVar
	GuiControlGet, LastName  ;get lastnameVar		
	StringLower, FirstName, FirstName ;convert to lower
	StringLower, LastName, LastName   ;convert to lower
	
	GuiControl,,UPN,%FirstName%.%LastName% ;display in textboxUPN
	FirstName2Char := SubStr(FirstName, 1, 2)
	GuiControl,,SamAccountName,%LastName%%FirstName2Char% ;display in textboxSamaccountName
Return

	
GuiClose:
ExitApp 