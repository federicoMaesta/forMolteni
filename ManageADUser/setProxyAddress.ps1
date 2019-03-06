Param (
    [Parameter(Position=0,Mandatory=$true)]
    [string]$file   

)

$users = Import-Csv $file

foreach ($user in $users) {
#$user
    Get-ADUser -Filter {UserPrincipalName -Like $user}



<#
    Get-ADUser -Filter {UserPrincipalName -Like $UPNAD} | % {
        Set-ADUser $_.samAccountName -Add @{ProxyAddresses='SMTP:' + $user.UserPrincipalName}
    }
#>
}