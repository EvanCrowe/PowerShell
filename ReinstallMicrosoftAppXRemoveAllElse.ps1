#Array of Publisher values for Microsoft AppX packages
$MicrosoftPublisher = @("CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US",
                        "CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US")

#Uninstalls all non-Microsoft packages
Get-AppxPackage|Where-Object {$_.Publisher -cnotin $MicrosoftPublisher}|Remove-AppxPackage

#Reinstalls
Get-AppxPackage |Where-Object{$_.Publisher -in $MicrosoftPublisher} | ForEach-Object {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"}