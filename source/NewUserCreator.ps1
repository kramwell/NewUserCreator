#Second part of NewUserSetup.exe
Param (
[string]$EmailAddress,
[string]$samAccountName,
[string]$LyncNumber,
[string]$LyncExtension
)

$tick1 = ""
$tick2 = ""
$tick3 = ""
$tick4 = ""
$tick5 = ""
$tick6 = ""

#Main-function
function headerLogo {
	Write-Host "********************************************************************"
	Write-Host "**                User Creator / KramWell.com                     **"
	Write-Host "********************************************************************"
	Write-Host ""
}	 
	 
function Show-Menu
{
	cls
	headerLogo
	 
	Write-Host "Email Address:`t`t$EmailAddress"
	Write-Host "Logon Name:`t`t$samAccountName`n"	 

	Write-Host "Lync Number:`t`t$LyncNumber`n"	

	Write-Host "1: Press 1 - enable remote mailbox:`t`t[$tick1]"     
	Write-Host "2: Press 2 - sync with azure online:`t`t[$tick2]"
	Write-Host "3: Press 3 - modify Ad exchange:`t`t[$tick3]"
	Write-Host "4: Press 4 - Set Office license:`t`t[$tick4]"
	Write-Host "5: Press 5 - Enable UMM for user:`t`t[$tick5]"
	Write-Host "6: Press 6 - Enable VMM for user:`t`t[$tick6]"

	Write-Host "`nQ: Press Q - to quit program."
}

cls
headerLogo

If ($EmailAddress -eq ""){
$EmailAddress = (Read-Host "`nPlease enter email address e,g john.doe@example.com...")
}
If ($samAccountName -eq ""){
$samAccountName = (Read-Host "Please enter logon name e,g johndoe")
}
If ($LyncNumber -eq ""){
$LyncNumber = "NONE"
}

Write-Host "`nEmail Address:`t`t$EmailAddress"
Write-Host "Logon Name:`t`t$samAccountName"
Write-Host "Lync Number:`t`t$LyncNumber"

$Answer = Read-Host "`nIs this the information you want to use (y/N)"
If ($Answer.ToUpper() -ne "Y"){ 
	Write-Host "`n`nOK.  Please rerun the script from the program."
	Read-Host -Prompt "`nPress Enter to exit"
	Break
}

do
{
	Show-Menu
	$input = Read-Host "`nPlease make a selection"
	switch ($input)
	{
	 
		#####################################################################	 #enable remote mailbox
		'1' {
   
			cls
			headerLogo
			
			try {

				Write-Host "Connecting to on-prem EXCHANGE and enabling remote mailbox in the cloud...`n"
				$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionURI http://HYBRIDSERVER/powershell/ -Authentication kerberos -ErrorAction stop
				import-PSSession $session -ErrorAction stop
				Enable-RemoteMailbox "$EmailAddress" -remoteroutingaddress "$johndoe@example.mail.onmicrosoft.com" -ErrorAction stop
				#Remove-PSSession $session -ErrorAction stop	
				Write-Host "`n-Remote mailbox enabled" # -foregroundcolor "magenta"
				$tick1 = "DONE"	
			}catch{
				Write-Error $Error[0].ToString()
				Write-Warning "FATAL ERROR CAN NOT CONTINUE!"
			}
			Remove-PSSession $session
					
				Read-Host -Prompt "`nPress enter to continue.."			
		#####################################################################	#Connect and sync azure
		} '2' {
			cls
			headerLogo
			try {
				
				Write-Host "Connecting to azure and syncing database...`n"
				$session = New-PSSession -ComputerName "AZURESERVER" -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {Import-Module "C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1"} -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} -ErrorAction stop
				#Remove-PSSession $session -ErrorAction stop
				Write-Host "`n-Sync Complete" # -foregroundcolor "magenta"
				$tick2 = "DONE"
			}catch{
				Write-Error $Error[0].ToString()
				Write-Warning "FATAL ERROR CAN NOT CONTINUE!"
			}
			Remove-PSSession $session
			
				Read-Host -Prompt "`nPress enter to continue.."			

		#####################################################################	#modify exchange details on AD
        } '3' {
			cls
			headerLogo
			try {
				
				Write-Host "Connecting to AD to set exchange details...`n"
				$session = New-PSSession -ComputerName "DOMAINCONTROLLER" -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {Import-Module ActiveDirectory} -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {param($samAccountName) Set-ADUser "$samAccountName" -Replace @{msExchRecipientDisplayType = "-2147483642"} } -ArgumentList $samAccountName -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {param($samAccountName) Set-ADUser "$samAccountName" -Replace @{msExchRecipientTypeDetails = "2147483648"} } -ArgumentList $samAccountName -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {param($samAccountName) Set-ADUser "$samAccountName" -Replace @{msExchRemoteRecipientType = "4"} } -ArgumentList $samAccountName -ErrorAction stop
				#Remove-PSSession $session -ErrorAction stop
				Write-Host "`n-Modified exchange details" # -foregroundcolor "magenta"
				$tick3 = "DONE"
			}catch{
				Write-Error $Error[0].ToString()
				Write-Warning "FATAL ERROR CAN NOT CONTINUE!"
			}	
			Remove-PSSession $session
			Read-Host -Prompt "`nPress enter to continue.."	

		#####################################################################	#add license to Office 365
        } '4' {
			cls
			headerLogo
			try {

				Write-Host "Connecting to exchange to add license...`n"
				Write-Host "Please input Office365 admin username and password`n"

				$Credential = Get-Credential -ErrorAction stop

				$session = New-PSSession -ComputerName "EXCHANGE" -ErrorAction stop
				
				Invoke-Command -Session $session -ScriptBlock {param($Credential) Connect-MsolService -Credential $Credential} -ArgumentList $Credential -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {param($EmailAddress) Set-MsolUser -UserPrincipalName "$EmailAddress" -UsageLocation "UK" } -ArgumentList $EmailAddress -ErrorAction stop
				Invoke-Command -Session $session -ScriptBlock {param($EmailAddress) Set-MsolUserLicense -UserPrincipalName "$EmailAddress" -AddLicenses "EXAMPLE:SPE_E5" } -ArgumentList $EmailAddress -ErrorAction stop
				#Remove-PSSession $session -ErrorAction stop
				Write-Host "`n-Added enterprise license to Office365" # -foregroundcolor "magenta"
				$tick4 = "DONE"
			}catch{
				Write-Error $Error[0].ToString()
				Write-Warning "FATAL ERROR CAN NOT CONTINUE!"
			}	
			Remove-PSSession $session
			Read-Host -Prompt "`nPress enter to continue.."	

		#####################################################################	#enable VoiceMail in Office 365
        } '5' {
			cls
			headerLogo
			try {
				Write-Host "Connecting to office 365...`n"

				$UserCredential = Get-Credential -ErrorAction stop
				$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorAction stop
				Import-PSSession $Session -ErrorAction stop
				
				#if have skype number then set here ofterwise skip
				
				If ($LyncExtension -ne ""){

					Enable-UMMailbox -Identity "$EmailAddress" -UMMailboxPolicy "POLICY" -Extensions "$LyncExtension" -PIN 12345 -SIPResourceIdentifier "$EmailAddress" -PINExpired $true -ErrorAction stop
					Set-UMMailbox -Identity "$EmailAddress" -TUIAccessToCalendarEnabled $false -TUIAccessToEmailEnabled $false -ErrorAction stop
				}

				Add-RecipientPermission -Identity "$EmailAddress" -AccessRights SendAs -Trustee "user.access@example.com" -Confirm:$false -ErrorAction stop
				Remove-PSSession $Session
					
				Write-Host "`n-Added UMM to Office365" # -foregroundcolor "magenta"
				$tick5 = "DONE"
			}catch{
				Write-Error $Error[0].ToString()
				Write-Warning "FATAL ERROR CAN NOT CONTINUE!"
			}	
			Remove-PSSession $session
			Read-Host -Prompt "`nPress enter to continue.."	
			#Stop-Transcript

		#####################################################################	#enable VMM
        } '6' {
				cls
				headerLogo
			try {
		
				Write-Host "Connecting to skype to setup VMM...`n"

				$credential = get-credential
				$sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
				$session = New-PSSession -ConnectionUri 'https://SKYPESERVER/ocspowershell' -credential $credential -SessionOption $sessionOption -Verbose -ErrorAction stop
				Import-PSSession $session
					
					
				$aduser = Get-CsAdUser -Identity "$EmailAddress"

				##HERE WE NEED TO SAY IF NOT HAVING A LYNC NUMBER THEN SETUP FOR PC TO PC#

				If ($LyncExtension -ne ""){

					Enable-CsUser -Identity $aduser.identity -RegistrarPool "SKYPESERVER" -SipAddressType emailaddress -ErrorAction stop
					Write-Host "1 ok - waiting 25 seconds"
					Start-Sleep -s 25
					Set-CsUser -Identity "$EmailAddress" -LineUri "tel:$LyncNumber;ext=$LyncExtension" -EnterpriseVoiceEnabled $true -ErrorAction stop #-HostedVoiceMail $true
					Write-Host "2 ok"
					Set-CsClientPin -Identity "$EmailAddress" -Pin "12345" -ErrorAction stop
					Write-Host "3 ok"
					Grant-CsHostedVoicemailPolicy -Identity "$EmailAddress" -PolicyName CloudUM -ErrorAction stop
					Write-Host "4 ok"
					Set-CsUser -Identity "$EmailAddress" -HostedVoiceMail $true
					Write-Host "set voicemail"

					Remove-PSSession $session -ErrorAction stop

					Write-Host "`n-Added VMM to skype" # -foregroundcolor "magenta"	

					Write-Host "`nConnecting to AD to set telephone number...`n"
					$session = New-PSSession -ComputerName "DOMAINCONTROLLER" -ErrorAction stop
					Invoke-Command -Session $session -ScriptBlock {Import-Module ActiveDirectory} -ErrorAction stop
					Invoke-Command -Session $session -ScriptBlock {param($samAccountName, $LyncNumber) Set-ADUser "$samAccountName" -Replace @{telephoneNumber = "$LyncNumber"} } -ArgumentList $samAccountName, $LyncNumber -ErrorAction stop
					Remove-PSSession $session -ErrorAction stop
					Write-Host "`n-Modified telephone number" # -foregroundcolor "magenta"

				}else{

					#we need to just setup user for PC-to-PC access in skype
					Enable-CsUser -Identity $aduser.identity -RegistrarPool "SKYPESERVER" -SipAddressType emailaddress -ErrorAction stop
					Write-Host "1 ok - waiting 25 seconds"
					Start-Sleep -s 25
					Set-CsUser -Identity "$EmailAddress" -EnterpriseVoiceEnabled $false -ErrorAction stop #-HostedVoiceMail $true

					Remove-PSSession $session -ErrorAction stop

				}
				$tick6 = "DONE"
		
			}catch{
				Write-Error $Error[0].ToString()
				Write-Warning "FATAL ERROR CAN NOT CONTINUE!"
			}			
			Remove-PSSession $session
			Read-Host -Prompt "`nPress enter to continue.."	
			#Stop-Transcript
		} 'q' {
		return
		}
    }
}
until ($input -eq 'q')