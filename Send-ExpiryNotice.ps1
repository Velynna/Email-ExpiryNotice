#Requires -Version 5.1
#Requires -Module ActiveDirectory
function Send-ExpiryNotice {
    <#
        .SYNOPSIS
            Notifies AD users that their password is about to expire.
    
        .DESCRIPTION
            Lets users know their password will soon expire. Details the steps needed to change their
            password, and advises on what the password policy requires. Accounts for both standard Default
            Domain Policy based password policy and the fine grain password policy available in 2008 domains.
    
        .NOTES
            Updated by:     Antonya Johnston
            Date:           12-02-2020
            Author:         M. Ali (original AD query), Pat Richard, Lync MVP
            Version:        4.3.0
    
            Required:
                + Run as a scheduled task...
                + on a Windows Server 2012 or later...
                + Using local admin rights on server it's running on.
                + Exchange 2007 or later
                + Lync (n/a)
                + ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
            
            Acknowledgements:
                This script has been heavily edited from it's 3.0 predicessor. It has been modified to be 1 file.
                It has no images, but the content has been modified to be responsive.
    
                + Calculating time
                    http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/23fc5ffb-7cff-4c09-bf3e-2f94e2061f29/
                + Determine per user fine grained password settings
                    http://technet.microsoft.com/en-us/library/ee617255.aspx
                + Pat Richard
                    https://ucunleashed.com/318
        
        .LINK
            https://github.com/Velynna/New-AdPasswordReminder
        
        .INPUTS
            None. You cannot pipe objects to Send-ExpiryNotice.
        
        .EXAMPLE
            Send-ExpiryNotice
        
            Searches Active Directory for users who have passwords expiring soon, and emails them a reminder
            with instructions on how to change their password.
        
        .EXAMPLE
            Send-ExpiryNotice -Demo
        
            Searches Active Directory for users who have passwords expiring soon, and lists those users on
            the screen, along with days till expiration and policy setting
        
        .EXAMPLE
            Send-ExpiryNotice -PreviewUser [Username]
        
            Sends the HTML formatted email of the user specified via -PreviewUser. This is used to see what
            the HTML email will look like to the users.

        .EXAMPLE
            Send-ExpiryNotice -TestOU "OU=IT,OU=Users,DC=contoso,DC=com"
        
            Checks only users in a valid, supplied OrganizationalUnit.

        .EXAMPLE
            Send-ExpiryNotice -EmailOverride email@mail.com
        
            All email generated during the runtime will redirect to the entered email address. Emails for
            individual expiry notices will have the user's email address inserted into the Subject line.

        .EXAMPLE
            Send-ExpiryNotice -TestOU "OU=IT,OU=Users,DC=contoso,DC=com" -EmailOverride email@mail.com
        
            Checks only users in a valid, supplied OrganizationalUnit. Additionally, all email generated during
            the runtime will redirect to the entered email address. Used together this can prevent a flood of
            hundreds of emails going to a single test email address. Emails for individual expiry notices will
            have the user's email address inserted into the Subject line.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        # Runs the script in demo mode. No emails are sent to the User(s), and onscreen output includes those who
        # are expiring soon.
        [Parameter(ParameterSetName = 'Demo',
                    Mandatory=$false,
                    Position=0)] 
        [Switch]
        $Demo,
        
        # User name of User to send the preview email message to.
        [Parameter(ParameterSetName = 'Preview',
                    Mandatory=$false,
                    Position=1)] 
        [String]
        $PreviewUser = $null,

        # The OU you'd like to test with. Suggested user group: IT users. Use the DistinguishedName.
        [Parameter(ParameterSetName = 'Test',
                   Mandatory=$false,
                   Position=2)]
        [ValidateScript({
            try { [ADSI]::Exists("LDAP://$_") }
            catch { Throw "The path is invalid."} })]
        [String]
        $TestOU = $null,

        # All emails generated will redirect to the entered email address.
        [Parameter(ParameterSetName = 'Test',
                   Mandatory=$false,
                   Position=3)]
        [System.Net.Mail.MailAddress]
        $EmailOverride = $null
    )
    begin {
        Write-Debug "$(Get-Date -Format o) :: Entering Begin Block"
        Write-Debug ("$(Get-Date -Format o) :: `$Demo set to $Demo`n" +
              "DEBUG: $(Get-Date -Format o) :: `$PreviewUser set to $PreviewUser`n" +
              "DEBUG: $(Get-Date -Format o) :: `$TestOU set to $TestOU`n" +
              "DEBUG: $(Get-Date -Format o) :: `$EmailOverride set to $EmailOverride`n" +
              "DEBUG: $(Get-Date -Format o) :: Parameter Set in use is $($PSCmdlet.ParameterSetName)")
        Write-Verbose -Message "Collecting AD Domain information"
        $Domain = Get-ADDomain

        # Type-agnostic Constants
        Write-Verbose -Message "Initializing constants required for processing"
        $PDCEmulator = $Domain | Select-Object -ExpandProperty PDCEmulator
        $DomainDN = $Domain | Select-Object -ExpandProperty DistinguishedName
        [Int]$i = 0
        $Skipped = @()
        $ScriptName = $MyInvocation.MyCommand.Name

        $Props = @("PasswordExpired", "PasswordNeverExpires", "PasswordLastSet", "Name", "Mail")

        Write-Verbose -Message "Setting hard-coded notification variables"
        $DateFormat = 'd'

        #region Hard-coded Values
            # REPLACE THIS DATA
            $RootUsersDN = "OU=Users,$DomainDN"
            [Int]$DaysToWarn = 14 # How far out should emails begin being sent?
            $EmailFrom = "support@contoso.com"
            $TaskComplete = "taskcomplete@contoso.com" # Where should completion notice be sent to?
            $SmtpServer = "email.contoso.com"
            $CoName = "Contoso"
            $CoPhone = "(555) 555-5555"
            $CoSupport = "help@contoso.com"
            $StreetAddress = "1234 56th St Ste 789"
            $City = "City Name"
            $State = "ST"
            $Zip = "012345"
            $CoSite = "https://www.contoso.com"
            $PasswordResetSite = "https://subdomain.contoso.com"
            $MapsLink = "https://www.google.com/maps/place/..."

            # Calculations. Do nott replace!
            $CoTel = ($CoPhone -replace "\D+")
            $Address = "$StreetAddress<br />$City, $State $Zip"
            if ($TestOU) { $RootUsersDN = $TestOU }
            if ($EmailOverride) { $TaskComplete = $EmailOverride }
        #endRegion

        #region StyleSheet
            $FontFamily = "Calibri, Candara, Segoe, 'Segoe UI', Optima, Arial, sans-serif"
            $FontSize = "14px"
            $BodyFontColor = "#000000"
            $TitleColor = "#FFFFFF"
            $H1FontSize = "18px"
            $H1FontColor = "#FFFFFF"
            $H2FontColor = "#020086"
            $H3FontColor = "#217400"
            $BodyLinkColor = $H2FontColor
            $FooterLinkColor = $BodyFontColor
            $H1BgColor = "#000000"
            $CoNameSize = "16px"

            #Title bar color override. Uncomment to use.
            #$TitleBkgd = "#aa0000"

            #region Style Defaults
                $P1 = "
                    font-size: $FontSize; 
                    margin: 0 0 10px;
                "
                $P2 = "
                    font-size: $FontSize; 
                    margin: 0 0 10px 15px;
                "
                $P3 = "
                    font-size: $FontSize; 
                    margin: 0 0 10px 30px;
                "
                $H1 = "
                    color: $H1FontColor; 
                    font-size: $H1FontSize; 
                    margin: 9px 0;
                "
                $H2 = "
                    color: $H2FontColor; 
                    font-size: $FontSize; 
                    margin: 0 0 10px
                "
                $H3 = "
                    color: $H3FontColor; 
                    font-size: $FontSize; 
                    margin: 0 0 10px 15px;
                "

                $Link1 = "
                    color: $BodyLinkColor; 
                    text-decoration: none;
                "
                $Link2 = "
                    color: $FooterLinkColor; 
                    text-decoration: none;
                "

                $TitleBar = "
                    width: 100%; 
                    color: $TitleColor;
                    text-decoration: none; 
                    padding: 0;
                "
                $H1Bkgd = "
                    width: 100%; 
                    color: $H1FontColor; 
                    text-decoration: none; 
                    padding: 0;
                "
                $Table1 = "
                    border-spacing: 0; 
                    font-family: $FontFamily; 
                    color: $BodyFontColor; 
                    width: 100%; 
                    max-width: 1200px; 
                    margin: 0 auto;
                "
                $Table2 = "
                    border-spacing: 0; 
                    font-family: $FontFamily; 
                    color: $BodyFontColor; 
                    width: 100%; 
                    font-size: 12px; 
                    text-align: center;
                "
            #endRegion
        #endRegion
        
        Write-Debug "$(Get-Date -Format o) :: Exiting Begin Block"
    }
    process {
        Write-Debug "$(Get-Date -Format o) :: Entering Process Block"
        if ($Demo){
            Write-Verbose -Message "Demo mode"
            Write-Output -InputObject "`n"
            Write-Host ('{0,-25}{1,-8}{2,-12}{3,-25}' -f 'User', 'Expires', 'PolicyDays', 'Email'
                       ) -ForegroundColor Cyan
            Write-Host ('{0,-25}{1,-8}{2,-12}{3,-25}' -f '========================', '=======', '===========',
                        '===============') -ForegroundColor Cyan
        }
    
       Write-Verbose -Message "Setting event log configuration"
        [Object]$evt = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList ('Application')
        [String]$evt.Source = $ScriptName
        $InfoEvent = [Diagnostics.EventLogEntryType]::Information
        [String]$EventLogText = "Beginning processing"
        $evt.WriteEntry($EventLogText,$infoevent,70)
        $Timer =  [System.Diagnostics.Stopwatch]::StartNew()
        
        Write-Verbose -Message "Getting password policy configuration"
        $DomainPP = Get-ADDefaultDomainPasswordPolicy
        $DomainMaxPswdAge = $DomainPP | Select-Object -ExpandProperty MaxPasswordAge

        if ($PreviewUser){
            Write-Verbose -Message "Preview mode"
            $Users = Get-AdUser $PreviewUser -Properties $Props -Server $PDCEmulator
        } else {
            Write-Verbose -Message "Collecting list of all AD Users in $RootUsersDN"
            $GetParams = @{
                LDAPFilter = '(!(name=*$))'
                SearchScope = 'Subtree'
                SearchBase = $RootUsersDN
                Properties = $Props
                ResultSetSize = $null
                Server = $PDCEmulator
            }
            $Users = Get-AdUser @GetParams
        }
        if ($PreviewUser) { $Count = 1 }
        else { $Count = $Users.Count }
        Write-Debug "$(Get-Date -Format o) :: Total Users found in $RootUsersDN`: $Count"
        foreach ($User in $Users) {
            Write-Verbose -Message "Looping through users that meet the validation requirements."
            # While the OU in use is for enabled users, this is a final check to skip over disabled accounts.
            if ((!($User.PasswordExpired) -and !($User.PasswordNeverExpires) -and $User.Enabled) -or
                ($PreviewUser)) {
                Write-Debug "$(Get-Date -Format o) :: Entering #region Get-ADUserPasswordExpirationDate"
                #region Get-ADUserPasswordExpirationDate
                    Write-Verbose -Message "Checking PasswordLastSet date"
                    $PswdLastSet = $User.PasswordLastSet
                    if ($PswdLastSet) {
                        # Users being modified during runtime will error here. Use of SamAccountName
                        # vs. CN advised.
                        $AccountFGPP = Get-ADUserResultantPasswordPolicy $User.samAccountName
                        $AccountFGPP = $AccountFGPP | Select-Object -ExpandProperty MaxPasswordAge
                        if ($AccountFGPP) {
                            $MaxPswdAge = $AccountFGPP.TotalDays
                        } else {
                            $MaxPswdAge = $DomainMaxPswdAge.TotalDays
                        }
                        if (!($MaxPswdAge) -or ($MaxPswdAge.TotalMilliseconds -ne 0)) {
                            Write-Debug "$(Get-Date -Format o) :: `$MaxPswdAge set to $MaxPswdAge"
                            if ($PreviewUser){
                                $DaysTillExpire = 1
                            } else {
                                $Params = @{
                                    Start = (Get-Date)
                                    End = ($PswdLastSet + $MaxPswdAge).ToString()
                                }
                                $DaysTillExpire = [System.Math]::round(((New-TimeSpan @Params).TotalDays),0)
                            }
                            if ($DaysTillExpire -le $DaysToWarn){
                                Write-Debug "$(Get-Date -Format o) :: `$DaysTillExpire set to $DaysTillExpire"
                                Write-Debug "$(Get-Date -Format o) :: `$DaysToWarn set to $DaysToWarn"
                                Write-Verbose -Message "Preparing email for user."
                                $i++
                                if (!($Demo)) {
                                    if ($CustTitleBkgd) {
                                        $TitleBkgd = $CustTitleBkgd
                                    } else {
                                        if ($DaysTillExpire -le 2) {
                                            $TitleBkgd = '#E74C3C'
                                        } elseif ($DaysTillExpire -le 7) {
                                            # elseif only runs if the first statement is false, so no "between"
                                            # is required.
                                            $TitleBkgd = '#E67E22'
                                        } else {
                                            $TitleBkgd = $H1BgColor
                                        }
                                    }
                                    # Holds original values so it can clean before each loop.
                                    $TitleBarReset = $TitleBar
                                    $TitleBar += "background-color: $TitleBkgd`;"
                                    Write-Debug "$(Get-Date -Format o) :: `$TitleBkgd set to $TitleBkgd"
                                    $GivenName = $User.GivenName
                                    Write-Debug "$(Get-Date -Format o) :: `$GivenName set to $GivenName"
                                    $DateofExpiration = (Get-Date).AddDays($DaysTillExpire)
                                    # Second Get-Date used to reformat first call.
                                    $DateofExpiration = (Get-Date -Date ($DateofExpiration) -Format $DateFormat)
                                    Write-Debug ("$(Get-Date -Format o) :: `$DateofExpiration set to" +
                                                "$DateofExpiration")
                                    Write-Verbose -Message 'Assembling email message'
                                    [String]$EmailBody = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <!--[if !mso]><!-->
        <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <!--<![endif]-->
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title></title>
</head>
<!--[if (gte mso 9)|(IE)]>
<style type="text/css">
    table {border-collapse: collapse;}
</style>
<![endif]-->
<body style="margin: 0 !important; padding: 0;" bgcolor="#FFFFFF">
    <div style="width: 100%; table-layout: fixed; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;">
        <div style="max-width: 1200px; margin: 0 auto;">
            <!--[if (gte mso 9)|(IE)]>
            <table width="800" align="center" cellpadding="0" cellspacing="0" border="0">
            <tr>
            <td>
            <![endif]-->
            <table align="center" style="$Table1">
                <tr>
                    <td style="$TitleBar" align="center">
<p style="$H1">Your Password Expires in $DaysTillExpire Day(s)</p>
                    </td>
                </tr>
                <tr>
                    <td style="padding: 0;">
                        <table width="100%" style="$Table1">						
                            <tr>
                                <td style="width: 100%; padding: 10px;" align="left">
<p style="$P1">Hello $GivenName,<br /><br />
It's change time again! Your password for access to computers, and other single sign-on programs will expire in 
$DaysTillExpire day(s), on $DateofExpiration.</p>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="$H1Bkgd" align="center" bgcolor="$H1BgColor">
<p style="$H1">How to update your password</p>
                    </td>
                </tr>
                <tr>
                    <td style="padding: 0;">
                        <table width="100%" style="$Table1">						
                            <tr>
                                <td style="width: 100%; padding: 10px;" align="left">
<p style="$H2">On your company computer?</p>
<p style="$P2">Ensure you are first either in the building, <b>or</b> connect to the VPN, then you may update
your password on your computer by pressing Ctrl-Alt-Delete and selecting <i>Change a Password</i> from the
available options.</p>
<p style="$H2">On a phone, or tablet?</p>
<p style="$P2">Visit our password management site at <a href="$PasswordResetSite" target="_blank" style="$Link1">
$PasswordResetSite</a>, then click <i>Change My Password</i>. If you do not remember your password, click <i>
Forgot Password</i>, then follow the steps on the page. <br /><br />Once your password has been changed, follow
the instructions in the appropriate section(s) below.</p>
<p style="$H3">iOS</p>
<p style="$P3">Open the Mail app, and refresh the contents. You will be asked to enter your new password. If not
prompted right away, please try restarting your device.</p>
<p style="$H3">Android</p>
<p style="$P3">Restart your device, then navigate to the Email app. You should be asked to enter your new
password.</p>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="$H1Bkgd" align="center" bgcolor="$H1BgColor">
<p style="$H1">Contact Us</p>
                    </td>
                </tr>
                <tr>
                    <td style="font-size: 0; padding: 10px 0;" align="center">
                        <!--[if (gte mso 9)|(IE)]>
                        <table width="100%">
                        <tr>
                        <td width="400" valign="top">
                        <![endif]-->
                        <div style="width: 100%; max-width: 400px; display: inline-block; vertical-align: top;">
                            <table width="100%" style="$Table1">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table style="$Table2">
                                            <tr>
                                                <td style="padding: 12px 0 0;">
<p style="font-size: $CoNameSize; padding-top: 9px; margin: 0 0 10px;">
<a href="$CoSite" target="_blank" style="$Link2">$CoName</a></p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <!--[if (gte mso 9)|(IE)]>
                        </td><td width="400" valign="top">
                        <![endif]-->
                        <div style="width: 100%; max-width: 400px; display: inline-block; vertical-align: top;">
                            <table width="100%" style="$Table1">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table style="$Table2">
                                            <tr>
                                                <td style="padding: 12px 0 0;">
<p style="margin: 0;">Phone: <a href="tel:$CoTel" target="_blank" style="$Link2">$CoPhone</a></p>
<p style="margin: 0;">Email: <a href="mailto:$CoSupport" target="_blank" style="$Link2">$CoSupport</a></p>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <!--[if (gte mso 9)|(IE)]>
                        </td><td width="400" valign="top">
                        <![endif]-->
                        <div style="width: 100%; max-width: 400px; display: inline-block; vertical-align: top;">
                            <table width="100%" style="$Table1">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table style="$Table2">
                                            <tr>
                                                <td style="padding: 12px 0 0;">
<a href="$MapsLink" target="_blank" style="$Link2">$Address</a>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <!--[if (gte mso 9)|(IE)]>
                        </td>
                        </tr>
                        </table>
                        <![endif]-->
                    </td>
                </tr>
            </table>
            <!--[if (gte mso 9)|(IE)]>
            </td>
            </tr>
            </table>
            <![endif]-->
        </div>
    </div>
</body>
</html>
"@
                                    Write-Debug "$(Get-Date -Format o) :: `$EmailBody set to $EmailBody"
                                } # end if (!($Demo))
                                #region SendEmail to User
                                    if ($EmailOverride) { 
                                        $EmailTo = $EmailOverride
                                        $EmailSubject = ("To: $($User.mail) Your password expires in" +
                                                        "$DaysTillExpire day(s)")
                                    } else {
                                        $EmailTo = $User.mail
                                        $EmailSubject = "Your password expires in $DaysTillExpire day(s)"
                                    }
                                    if (!($Demo)) {
                                        Write-Debug "$(Get-Date -Format o) :: Entering #region SendEmail to User"
                                        if ($EmailTo) {
                                            Write-Verbose -Message "Sending message to $EmailTo"
                                            $MailParams = @{
                                                To = $EmailTo
                                                From = $EmailFrom
                                                Subject = $EmailSubject
                                                Body = $EmailBody
                                                Priority = "High"
                                                BodyAsHTML = $true
                                                SMTPServer = $SmtpServer
                                            }
                                            Send-MailMessage @MailParams
                                        } else {
                                            $Skipped += [PSCustomObject]@{
                                                Name = $User.Name
                                                Username = $User.samAccountName
                                                DaysRemaining = $DaysTillExpire
                                            }
                                        }
                                    } else {
                                        Write-Host ('{0,-25}{1,-8}{2,-12}{3,-25}' -f $User.Name, $DaysTillExpire,
                                                                $MaxPswdAge, $EmailTo)
                                    }
                                #endregion
                            } # end if ($DaysTillExpire -le $DaysToWarn)
                        } # end if (!($MaxPswdAge) -or ($MaxPswdAge.TotalMilliseconds -ne 0))
                    } # end if ($PswdLastSet)
                #endregion
            
                #region ClearVariables
                    # Without the clear, if there is an issue accessing an account's data, data from a previous
                    # user's will be used instead. This causes confusion.
                    ("User","PswdLastSet","AccountFGPP","MaxPswdAge","DaysTillExpire","GivenName","EmailBody",
                    "EmailSubject","DateofExpiration","EmailTo","TitleBkgd") |
                    ForEach-Object { 
                        Try { Clear-Variable $_ }
                        Catch {}
                    }
                    $TitleBar = $TitleBarReset
                #endRegion ClearVariables
            } # end Validation Check
        } # end ForEach User

        if ($Demo) { Write-Host "Accounts Processed: $i" }
        Write-Verbose -Message 'Writing summary event log entry'
        $Timer.Stop()
        $TimeElapsed = [System.Math]::Round($Timer.Elapsed.TotalMinutes,0)
        $EventLogText = "Send-ExpiryNotice.ps1 finished processing $i account(s) after $TimeElapsed minutes."
        $evt.WriteEntry($EventLogText,$infoevent,70)

        if (!($PreviewUser) -and !($Demo)) {
            Write-Verbose -Message 'Sending email confirmation'
            $CompletedSubject = "The Send-ExpiryEmail task has completed successfully"
            $CompletedBody = @"
<p>This message is to notify you that the scheduled task to send email notifications to users whose password are
expiring within $DaysToWarn days has completed successfully.</p>
<p>Total Processing Time (minutes): $TimeElapsed</p>
<p>Accounts Processed: $i</p>
<table>
    <thead>
        <tr>
            <th>Skipped User</th>
            <th>Username</th>
            <th>Days Remaining</th>
        </tr>
    </thead>
    <tbody>
"@
            foreach ($s in $Skipped) { 
                $CompletedBody += @"
        <tr>
            <td>$($s.Name)</td>
            <td>$($s.Username)</td>
            <td>$($s.DaysRemaining)</td>
        </tr>
"@}
            $CompletedBody += "</tbody></table>"

            $CompletedParams = @{
                To = $TaskComplete
                From = $EmailFrom
                Subject = $CompletedSubject
                Body = $CompletedBody
                BodyAsHTML = $true
                SMTPServer = $SmtpServer
            }
            Send-MailMessage @CompletedParams
        }
        Write-Debug "$(Get-Date -Format o) :: Exiting Process Block"
    }
    end {
        Write-Debug "$(Get-Date -Format o) :: Entering End Block"
        Write-Debug "$(Get-Date -Format o) :: Exiting End Block"
    }
}

# Invokation for scheduled task
Send-ExpiryNotice
