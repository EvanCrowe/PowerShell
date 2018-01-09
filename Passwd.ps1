
<#
.Synopsis
   Generates a pseudo-random ASCII string
.DESCRIPTION
   By default excludes extended and control characters
.EXAMPLE
   New-Password

         Entropy Password   Length
         ------- --------   ------
2.94770277922009 koC!;CmBD       9
.EXAMPLE
   New-Password -Length 16 -MinEntropy 2
.EXAMPLE
   New-Password -MinEntropy 7

        Entropy Password Length
        ------- -------- ------
2.8073549220576 SH!*GI6       7
#>
function New-Password
{
    [CmdletBinding()]
    #[Alias(npwd)]
    [OutputType([object])]
    Param(# Param1 help description
         [Parameter(Mandatory=$false,
                   Position=0)]
                   [int]
                   $Length,
        # Param2 help description
        [Parameter(Position=1)]
        [ValidateRange(0,5)]
                   [double]
                   $MinEntropy = 0,
        # Param3 help description
        [Parameter(Position=2)]
        [ValidateSet("ASCII","IncludeControl","IncludeExtended","IncludeCE")]
                   [String]
                   $CharSet = "ASCII"

    )

    Begin
    {

    function Get-Entropy
{
    Param ([Parameter(Mandatory = $True)]
           [ValidateNotNullOrEmpty()]
           [Byte[]]
           $Bytes
   )

   $FrequencyTable = @{}
   foreach($Byte in $Bytes){
           $FrequencyTable[$Byte]++
   }
   $Entropy = 0.0

   foreach($Byte in 0..255){
           $ByteProbability = ([Double]$FrequencyTable[[Byte]$Byte])/$Bytes.Length
       if($ByteProbability -gt 0){
          $Entropy += -$ByteProbability * [Math]::Log($ByteProbability, 2)
       }
   }
   $Entropy
}


function Get-CryptoRand{
    Param ([Parameter(Mandatory = $false)]
           [ValidateNotNullOrEmpty()]
           $Rolls = 20,
           [Parameter(Mandatory = $false)]
           [ValidateNotNullOrEmpty()]
           $Bound= 32..126
   )

Begin{

function Randomize-List{
    Param([array]$InputList)
    return $InputList | Get-Random -SetSeed ([Environment]::TickCount) -Count $InputList.Count;
    }

function rand{
    $bytes = new-object "System.Byte[]" $Rolls
    $rnd = new-object System.Security.Cryptography.RNGCryptoServiceProvider
    $rnd.GetBytes($bytes)
    $bytes = $bytes | sort -Unique | Where-Object {$_ -in $Bound}
    $bytes
    }
}

Process{
    $TMP = @()
    while($TMP.count -le $Rolls){
    $TMP += rand
    Write-Verbose $TMP.Length
    }
}

End{
    $Out = @()
    $out = Randomize-List -InputList $TMP
    $Out = $Out |Select-Object -First $Rolls
    $Out
}

}

switch ($CharSet){
"IncludeControl" {$Bound = 0..126}
"IncludeExtended" {$Bound = 32..254}
"IncludeCE" {$Bound = 0..254}
default {$Bound = 32..126}
}

#initilize array for storing chars for password
$PasswordString = @()

    }
    Process{

Write-Verbose "before 0 test length $Length"

if($Length -eq 0){$Length = Get-Random -SetSeed ([Environment]::TickCount) -Minimum 8 -Maximum 16}
#Sets loop count
$LoopCount = $Length * 10
Write-Verbose "after 0 test length $Length"


do{
#iterates
    $LoopCount = $LoopCount - 1
#Gets pseudo random char, puts in PasswordString array #FIX
    $PasswordASCIIDEC = (Get-CryptoRand -Bound $Bound -Rolls $Length)

    foreach($ASCIIDec in $PasswordASCIIDEC){
    Write-Verbose "ascii dec $ASCIIDec"
    $PasswordString += [char]$ASCIIDec
    }

#concatenates $PasswordString
    $Passwd = $PasswordString -join ''

#Finds Entropy of passwd
    $enc = [system.Text.Encoding]::UTF8
    [double]$PasswdEntropy = Get-Entropy $enc.GetBytes($passwd)
    Write-Verbose " passwd entropy $PasswdEntropy"
#If MinEntropy isn't set allows all possible values
    if($MinEntropy -eq 0){$MinEntropy = $PasswdEntropy}
Write-Verbose " min entropy inside do loop  $MinEntropy"
Write-Verbose "loop count $LoopCount"
}
until((($PasswdEntropy -ge $MinEntropy) -and ($Passwd.length -le $Length)) -or ($LoopCount -lt 0))

if($LoopCount -le 0){Write-Error "Unable to generate password of length $Length and entropy $MinEntropy, increase Length, or reduce complexity requirements"}


#convert to secure string
#TODO Add export to file
$SecureStringPassword = $Passwd  | ConvertTo-SecureString -AsPlainText -Force

         $Password = New-Object -TypeName psobject -Property @{
        'Password'=$passwd;
        'Length'=$passwd.length;
        'Entropy'=$PasswdEntropy;
        }


 #Tests object, outputs to screen and clipboard
if($LoopCount -gt 0){$Password}
# $Password.password|clip



 #Add-Type -AssemblyName System.speech
 #$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
 #$speak.Speak("$passwd")
 
#$myinvocation.mycommand.name
#$myinvocation.mycommand.path
#$myinvocation.MyCommand
    #TODO set AD password?
    #Say outloud?


    }
    End
    {
    $Length = $null
    $MinEntropy = $null
    $Bytes = $null
    $FrequencyTable = $null
    $Entropy = $null
    $ByteProbability = $null
    $Rolls = $null
    $Bound = $null
    $out = $null
    $bytes = $null
    $rnd = $null
    $MAXOut = $null
    $SelectRandomOut = $null
    $MaxLength = $null
    $PasswordString = $null
    $LoopCount = $null
    $charOut = $null
    $Passwd = $null
    $Password = $null
    }
}

New-Password -Verbose