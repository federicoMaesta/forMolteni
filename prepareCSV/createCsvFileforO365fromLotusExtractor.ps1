#Stagin Version 0.1

Param (
    [Parameter(Position=0,Mandatory=$true)]
    [string]$file   

)

$arrForLic = @()

$filecontent = get-content $file
$filecontent[0] = $filecontent[0]  -replace ' ', ''
$filecontent | Set-content "tempfile.csv"

$users = Import-Csv "tempfile.csv"

foreach ($user in $users) {
   $arrForLic += $user.DestinationEmail.ToLower()

}


Remove-Item "tempfile.csv"

$arrForLic | Set-Content "usersO365.csv"
