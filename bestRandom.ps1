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

Get-CryptoRand