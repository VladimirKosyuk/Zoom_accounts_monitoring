#Created by https://github.com/VladimirKosyuk

# Compare exchange user mailboxes to zoom accounts, if not equal - send an email

# Build date: 14.12.2020

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#import modules
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Import-Module activedirectory
Import-Module PSZoom
#vars:email
$smtp_domain = ""
$msg_to = ""
$Smtp_srv = ""
$From = ""
$Subject = ""
$Body = ""
#vars:zoom auth keys 
$Global:ZoomApiKey    = ''  
$Global:ZoomApiSecret = ''
#vars:report folder
$Report = $PSScriptRoot+"\zoom_accounts.csv"
#delete previous report
Remove-item $Report -Force -ErrorAction SilentlyContinue

#get users list from exchange as array
$UPNs= (Get-ADUser -Filter * -Properties Mail,Enabled, DistinguishedName | ?{($_.Mail -ne $null) -and ($_.Enabled -eq $True) -and ($_.DistinguishedName -like "*OU=Users*") -and ($_.DistinguishedName -notlike "*OU=Disabled*") -and ($_.UserPrincipalName -like "*.*@$smtp_domain")}).UserPrincipalName
foreach ($UPN in $UPNs) {
$Exch_raw = get-user -Filter "UserPrincipalName -eq '$UPN'" |select WindowsEmailAddress, Department, Title
[array]$Exch_users += New-Object psobject -Property @{
    Email = $Exch_raw.WindowsEmailAddress
    Department = $Exch_raw.Department
    Title = $Exch_raw.Title
   }
}

#users list from zoom
$Zoom_users = (Get-ZoomUsers -AllPages -Status active)+(Get-ZoomUsers -AllPages -Status pending)

#if zoom users not equal to exchange users, than send an email without auth
if (($Diff_data = (Compare-Object -ReferenceObject $Zoom_users -DifferenceObject $Exch_users -Property email | ? {$_.SideIndicator -eq "=>"}).email) -ne $null){
foreach ($Diff in $Diff_data) {$Exch_users | ?{$_.Email -match $Diff} | Export-Csv -Append -Delimiter ';' -Path $Report -Encoding UTF8 -NoTypeInformation}
Send-MailMessage -To $msg_to -From $From -Subject $Subject -Body $Body -Attachments $Report -Port 25 -SmtpServer $Smtp_srv
} 
#rm vars
Remove-Variable -Name * -Force -ErrorAction SilentlyContinue