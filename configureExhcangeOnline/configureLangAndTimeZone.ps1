#Stagin Version 0.1

Param (
    [Parameter(Position=0,Mandatory=$true)]
    [string]$file   

)

$users = Import-Csv $file

foreach ($user in $users) {
    Get-Mailbox -Identity $user | Set-MailboxRegionalConfiguration  -Language it-IT -TimeFormat HH:mm -TimeZone "W. Europe Standard Time" -Mailbo-LocalizeDefaultFolderName:$true
}
