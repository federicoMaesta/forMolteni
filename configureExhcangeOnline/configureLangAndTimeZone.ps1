Param (
    [Parameter(Position=0,Mandatory=$true)]
    [string]$file   

)

$LiveCred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber

Connect-MsolService -Credential $LiveCred

#Stagin Version 0.1

$users = get-content $file

foreach ($user in $users) {
    Get-Mailbox -Identity $user | Set-MailboxRegionalConfiguration  -Language it-IT -TimeFormat HH:mm -TimeZone "W. Europe Standard Time"
}