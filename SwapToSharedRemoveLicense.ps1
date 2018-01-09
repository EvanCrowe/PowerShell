function Credentials(){
Add-Type -AssemblyName System.DirectoryServices.AccountManagement;
$Global:MyUPN = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.UserPrincipalName
$Global:UserCredential = Get-Credential -Credential $MyUPN
}
function Import-O365(){
$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $O365Session
Connect-MsolService -Credential $UserCredential
Connect-AzureAD -Credential $UserCredential
}

Credentials
Import-O365


$Report = @()


#gets all mailboxes except for those filtered and sets them to shared mailboxes
$UserMailboxes = get-mailbox | ?{($_.RecipientTypeDetails -ne "SharedMailbox") -and
($_.primarySMTPAddress -ne ".@.com")} | select -ExpandProperty PrimarySMTPAddress
#Sets mailbox to shared
foreach($MailBox in $UserMailboxes){
Set-mailbox -Identity $MailBox -Type Shared
    $RepMailbox = New-Object psobject -Property @{
                            Mailbox = $Mailbox
                            MailboxStatus = "Shared"
                            }
                 $Report += $RepMailbox
}



#gets all users except for those filtered and removes license
$LicensedUPNs = get-msoluser | ?{($_.isLicensed -eq $True) -and
($_.UserPrincipalName -ne ".@.com")} 
#removes Licenses
foreach($UPN in $LicensedUPNs.UserPrincipalName){
$License = (get-msoluser -UserPrincipalName $UPN | select -ExpandProperty licenses).AccountSkuId
Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $License

    $RepLIC = New-Object psobject -Property @{
                            UPN = $UPN
                            LicenseStatus = "UnLicensed"
                            }
                 $Report += $RepLIC
}




$Report | Export-Csv C:\temp\ChangesMadeByScript.csv -NoTypeInformation

