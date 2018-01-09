Import-Module ActiveDirectory
Import-Module Microsoft.Online.SharePoint.PowerShell
Import-Module Microsoft.PowerShell.Management
#Import-Module  'C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.PowerShell.psd1'


#This PowerShell module allows you to connect to Exchange Online service.
#To connect, use: Connect-EXOPSSession -UserPrincipalName <your UPN>

#This PowerShell module allows you to connect Exchange Online Protection service too
#To connect, use: Connect-IPPSSession -UserPrincipalName <your UPN>
Set-Location ~

function Credentials(){
Add-Type -AssemblyName System.DirectoryServices.AccountManagement;
$Global:MyUPN = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.UserPrincipalName
$Global:UserCredential = Get-Credential -Credential $MyUPN
}
function Import-O365(){
Credentials
$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
$O365ComplianceSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
$EOPO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $O365Session
Import-PSSession $O365ComplianceSession
Import-PSSession $EOPO365Session
Connect-MsolService -Credential $UserCredential
#Connect-EXOPSSession -UserPrincipalName $MyUPN
#Connect-IPPSSession -UserPrincipalName $MyUPN
}

function prompt {
$Location = Get-Location
"$env:COMPUTERNAME : $env:USERNAME : $Location >"
}


