<#
.Synopsis
   Gets PSD codes for Pennsylvania and Ohio, adds Vertex Codes if School.District matches.
   Exports to "Completed" directory
.DESCRIPTION
   Gets PSD codes for Pennsylvania and Ohio, adds Vertex Codes if School.District matches

   Uses the below listed websites to find PSD codes based off of Address
   https://thefinder.tax.ohio.gov
   https://munstats.pa.gov/public/findlocaltax.aspx

   Required fields
   FIRST.NAME	LAST.NAME	ADDR.ONE	ADDR.TWO	ZIP.CODE	CITY	STATE	LOCATION 	PSD.CODE	SCHOOL.DISTRICT	SD.CODE

   Imported document must be in .CSV format


.EXAMPLE
   get-psdcode -path Miller.csv
#>
function Get-PSDCode
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $SourceFile
    )

    Begin
    {
function Search-Ohio
{
    Param($Addr,
          $City,
          $ZIP
          )

    Begin{
            $ie = New-Object -comobject InternetExplorer.Application
            $ie.visible=$false
            $ie.navigate("https://thefinder.tax.ohio.gov/StreamlineSalesTaxWeb/AddressLookup/LookupByAddress.aspx?taxType=taxsummary")
            Start-Sleep -Milliseconds 1000
            While($ie.Busy){Start-Sleep -Milliseconds 1000}
            }
    Process
    {
    $addr
$ie.document.IHTMLDocument3_getElementById("txtAddress").value = $Addr
$ie.document.IHTMLDocument3_getElementById("txtCity").value = $City
$ie.document.IHTMLDocument3_getElementById("txtZip").value = $ZIP
Start-Sleep -Milliseconds 1000
While($ie.Busy){Start-Sleep -Milliseconds 500}
($ie.Document.IHTMLDocument3_getElementById("btnLookup")).click()
Start-Sleep -Milliseconds 1000
While($ie.Busy){Start-Sleep -Milliseconds 1000}
    }
    End
    {

    $OhioSD = $ie.Document.IHTMLDocument3_getElementById("lblSDNumber").textContent
    $SearchResultsOhio = New-Object psobject -Property @{
                            SchoolDistrict = $OhioSD
                            }
    $SearchResultsOhio

    Get-Process iexplore | Foreach-Object { $_.CloseMainWindow() } | Out-Null
$ie.quit() | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | Out-Null
Remove-variable ie | Out-Null
    }
}

function Search-Penn
{
    Param($Addr,
          $City,
          $ZIP
          )

    Begin{
$ie = New-Object -comobject InternetExplorer.Application
$ie.visible=$false
$ie.navigate("https://munstats.pa.gov/public/findlocaltax.aspx")
Start-Sleep -Milliseconds 1000
While($ie.Busy){Start-Sleep -Milliseconds 1000}
            }
    Process
    {
        $WorkAddress = $Addr
        $WorkCity = $City
        $WorkZIP = $ZIP

While($ie.Busy){Start-Sleep -Milliseconds 1000}

Start-Sleep -Milliseconds 1000

$ie.document.IHTMLDocument3_getElementById("ContentPlaceHolder1_txtHomeStreet").value = $Addr
$ie.document.IHTMLDocument3_getElementById("ContentPlaceHolder1_txtHomeCity").value = $City
$ie.document.IHTMLDocument3_getElementById("ContentPlaceHolder1_txtHomeZip").value = $ZIP

$ie.document.IHTMLDocument3_getElementById("ContentPlaceHolder1_txtWorkStreet").value = $WorkAddress
$ie.document.IHTMLDocument3_getElementById("ContentPlaceHolder1_txtWorkCity").value = $WorkCity
$ie.document.IHTMLDocument3_getElementById("ContentPlaceHolder1_txtWorkZip").value = $WorkZIP

While($ie.Busy){Start-Sleep -Milliseconds 1000}
($ie.Document.IHTMLDocument3_getElementById("ContentPlaceHolder1_btnViewInformation")).click()
While($ie.Busy){Start-Sleep -Milliseconds 1000}
Start-Sleep -Milliseconds 500
if(!($ie.Document.IHTMLDocument3_getElementById("ContentPlaceHolder1_lblHomeSchoolPSD").textContent)){
    try{
        ($ie.Document.IHTMLDocument3_getElementById("ContentPlaceHolder1_rblHome_0")).click()
        ($ie.Document.IHTMLDocument3_getElementById("ContentPlaceHolder1_rblWork_0")).click()
        ($ie.Document.IHTMLDocument3_getElementById("ContentPlaceHolder1_btnReport")).click()
    }
    catch{}
}

While($ie.Busy){Start-Sleep -Milliseconds 1000}
Start-Sleep -Milliseconds 3000


    }
    End
    {
    $PennPSDCode = $ie.Document.IHTMLDocument3_getElementById("ContentPlaceHolder1_lblHomeSchoolPSD").textContent
    $SchoolDistrictPenn = $ie.Document.IHTMLDocument3_getElementById("ContentPlaceHolder1_lblHomeSchoolDistrictName").textContent
    $SchoolDistrictPenn = $SchoolDistrictPenn -replace "Township","TWP"

    $SearchResultsPenn = New-Object psobject -Property @{
                            PDSCode = $PennPSDCode
                            SchoolDistrict = $SchoolDistrictPenn
                            }

$SearchResultsPenn

Get-Process iexplore | Foreach-Object { $_.CloseMainWindow() } | Out-Null
$ie.quit() | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | Out-Null
Remove-variable ie | Out-Null
    }
}
    }
    Process
    {
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement;
#$Global:MySAM = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.SamAccountName
$myinvocation.mycommand.name
$myinvocation.mycommand.path

$path = "$PSScriptRoot"

$RAW = Import-Csv -Path "$path\$SourceFile"

$Vertex = Import-Csv "$path\Resources\PA Vertex Guide.csv"

[array]$Processed = @()


foreach($UnProcessed in $RAW){


        $HomeAddress = $UnProcessed.'ADDR.ONE'
        $HomeCity = $UnProcessed.CITY
        $HomeState = $UnProcessed.STATE
        $HomeZIP = $UnProcessed.'ZIP.CODE'

        $WorkAddress = $HomeAddress
        $WorkCity = $HomeCity
        $WorkZIP = $HomeZIP

#Get School dist and SD code Penn
if(($UnProcessed.STATE -eq "Pennsylvania") -or ($UnProcessed.STATE -eq "PA")){
$lastname = $UnProcessed.'LAST.NAME'
Write-Host "Processing Pennsylvania Data for $lastname"
$SRP = Search-Penn -Addr $HomeAddress -City $HomeCity -ZIP $HomeZIP
$UnProcessed.'PSD.CODE' = $SRP.PDSCode
$UnProcessed.'SCHOOL.DISTRICT' = $SRP.SchoolDistrict

#Get Vertex Codes for Penn
foreach($SchDist in $Vertex){if($SchDist.'School District' -eq $UnProcessed.'SCHOOL.DISTRICT'){$UnProcessed.'SD.Code' = $SchDist.'SD Code'}}
}
#Get SD code Ohio
if(($UnProcessed.STATE -eq "Ohio") -or ($UnProcessed.STATE -eq "OH")){
$lastname = $UnProcessed.'LAST.NAME'
Write-Host "Processing Ohio Data for $lastname"
$SRO = Search-Ohio -Addr $HomeAddress -City $HomeCity -ZIP $HomeZIP
$UnProcessed.'SCHOOL.DISTRICT' = $SRO.SchoolDistrict
}
$Processed += $UnProcessed
}




    }
    End
    {
    #Gets date in a format safe for a file name
$Date = Get-Date -Format o | foreach {$_ -replace ":", "."}
#Saves to CSV
$Processed | Export-Csv "$Path\Completed\$SourceFile $Date.csv" -NoTypeInformation -Force

pause

$Error >> "$Path\ErrorLog\$Date.txt"
    }
}

Get-PSDCode -SourceFile 3.CSV