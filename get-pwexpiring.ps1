#Add Snap-in for Quest's "ActiveRoles" Commandlets (get-qaduser)
add-pssnapin quest.activeroles.admanagement

Import-Module activedirectory

#Import Exchange Information
[Reflection.Assembly]::LoadFile("c:\program files\microsoft\exchange\web services\1.1\Microsoft.Exchange.WebServices.dll")
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$s.Credentials = new-object net.networkcredential('xxxxxxx', 'xxxxxxx', 'domainname.com')
$s.AutoDiscoverUrl("password-notice@peoplematter.com")
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$softdel = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete


##get the Domain Policy for the maximum password age
$dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$root = $dom.GetDirectoryEntry()

## get the account policy
$search = [System.DirectoryServices.DirectorySearcher]$root
$search.Filter = "(objectclass=domainDNS)"
$result = $search.FindOne()

## maximum password age
$t = New-Object System.TimeSpan([System.Math]::ABS($result.properties["maxpwdage"][0]))

##want all accounts where password will expire in next 7 days
##password was set (max password age) - 7 days ago
$d = ($t.Days)* -1   ## max password age days ago
$d1 = $d + 7   ## 7 days on from max password age

##collect all e-mail addresses and notify techsupport of who all has been
#notified
$emails = @()

Get-QADUser -IncludeAllProperties | Where {($_.PwdLastSet -gt (Get-Date).AddDays($d)) -and ($_.PwdLastSet -lt (Get-Date).AddDays($d1)) } | % {

    #find days remaining
    $now = (get-date)
    $expires = $_.PwdLastSet.AddDays(90)
    $daysremaining = ($expires - $now).days
    write-host "Password last set: " $_.PwdLastSet "`n"
    write-host "Current date: " (Get-Date) "`n"
    write-host "Date password expires: " $_.PwdLastSet.AddDays(90) "`n"
    write-host "Days until password expires: " ($_.PwdLastSet.AddDays(90) - (Get-Date))

    
    $mail = new-object microsoft.exchange.webservices.data.emailmessage($s)
    $technotify = new-object microsoft.exchange.webservices.data.emailmessage($s)
    $techrecipient = "email@domainname.com"

    write-host "Account name: " $_.Name "`n"
    $name = $_.Name.replace(" ", ".")
    $email = invoke-expression "(get-aduser `"$name`" -properties mail | select mail).mail"

    $names += $name + " - " + $daysremaining + " days<BR>"
    
    $recipient = ""

    #Check to see if they have an e-mail address entry in AD
    if (!($email -match "[a-zA-Z]")) { 
        write-host "User `"$name`" doesn't have an e-mail address."
        $mail.Subject = "Password Notice: Unable to find an e-mail address for user: $name"
        $mail.Body = "Hi there TechSupport,<BR><BR>The `"password-notice`" script was unable to find an e-mail address in AD for user <b><u>$name</b></u>."
        $recipient = "email@domainname.com"
    } 
    else {
        $mail.Subject = "Notice: Your PeopleMatter account password will expire in $daysremaining days."
    }

    $office = invoke-expression "(get-aduser `"$name`" -properties office | select office).office"
    $pobox = invoke-expression "(get-aduser `"$name`" -properties pobox | select pobox).pobox"
    write-host "PO Box: " $pobox "`n"

    #Check to see if user has Argentina set as Office or whether the user is a
    #Mac Users
    if (($office -eq "Argentina" -OR $pobox -eq "Mac") -AND $recipient -eq "") {
        $mail.Body = "Hello from IT,<BR><BR>Your password will be expiring in $daysremaining days.  Please change it at your earliest convenience.  Please visit https://webmail.peoplematter.com and login, then, in the top-right corner, select options -> change password.<BR><BR>Sincerely,<BR> PeopleMatter IT"
    }

    #Chances are they're a Windows user, so give them instructions relevant to a
    #windows machine.
    elseif ($recipient -eq "") {
        $mail.Body = "Hello from IT,<BR><BR>Your password will be expiring in $daysremaining days.  Please change it at your earliest convenience.  You can hit ctrl + alt + del and choose `"change a password,`" or you can visit https://webmail.peoplematter.com and login, then, in the top-right corner, select options -> change password.<BR><BR>Sincerely,<BR> PeopleMatter IT"
    }

    write-host "Email address: " $email "`n"

    if ($recipient -eq "") {
        $recipient = $email
    }

    [void] $mail.ToRecipients.Add($recipient)

    $mail.Sendandsavecopy()

}
$technotify.Body = "Hello SysAdmins, The following users have been alerted of their passwords expiring:<BR>" + $names
$technotify.Subject = "Users have been notified of expiring passwords..."
[void] $mail.ToRecipients.Add($recipient)
[void] $technotify.ToRecipients.Add($techrecipient)
$technotify.Sendandsavecopy()

