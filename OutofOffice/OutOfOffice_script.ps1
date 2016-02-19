
#*** VARIABLES / CONFIG SETTINGS **********************************************
# change this to be directory where script file is stored and will be run from
$myScriptPath  =  "C:\Scripts\Powershell\OutOfOffice"

# start and end times as text strings. used in calendar, auto replies, email etc. use date + time
$startTime     =  "08/03/2010 5:00 PM"    # format MM/dd/yyyy hh:mm AMPM
$endTime       =  "08/09/2010 8:00 AM"    # format MM/dd/yyyy hh:mm AMPM

# leave this be:
$startDt = [DateTime]::Parse($startTime)
$endDt = [DateTime]::Parse($endTime)

# calendar appointment subject and location
$apptCreate    =  $true
$apptSubject   =  "Out of Office"
$apptLocation  =  "Away"

# email address used for out of office auto reply and From for out of office email
$emailAddress  =  "my_email@mycompany.com"
$myName        =  "Geoff"

# internal and external messages for out of office automatic replies. Supports HTML
$autoReplySet  =  $true
$internalMsg   =  "I will be out of the office from <font color='blue'>" + $startDt.DayOfWeek.ToString() + " " + $startTime + "</font> to <font color='blue'>" + $endDt.DayOfWeek.ToString() + " " + $endTime + "</font>"
$externalMsg   =  $internalMsg

# this is who you want to email to notify ahead of time that you will be out of office
$emailSend     =  $true
# comma separate multiple addresses
$emailTo       =  "AppsTeam@mycompany.com, business_person@mycompany.com"
#$emailSubject =  [string]::Format("Out of office {0} - {1}", $startDt.ToShortDateString(), $endDt.ToShortDateString())
$emailSubject  =  "Out of office"
$emailBody     =  $internalMsg + ". Please see me before then if you need anything.<br/><br/>Thanks,<br/><br/>" + $myName + "<br/><br/>Sent by OutOfOffice.ps1"

#*** CONSTANTS ****************************************************************
if (!(test-path variable:\olFolderCalendar))
{ 
    New-Variable -Option constant -Name olFolderCalendar -Value 9
}    

if (!(test-path variable:\olAppointmentItem))     
{
    New-Variable -Option constant -Name olAppointmentItem  -Value 1
}    
    
if (!(test-path variable:\olOutOfOffice))         
{
    New-Variable -Option constant -Name olOutOfOffice  -Value 3
}    

if ($apptCreate)
{
    $outlook = new-object -com Outlook.Application

    #*** CREATE CALENDAR APPT *****************************************************
    $calendar = $outlook.Session.GetDefaultFolder($olFolderCalendar)
    $appt = $calendar.Items.Add($olAppointmentItem)
    $appt.Start = $startDt
    $appt.End = $endDt
    $appt.Subject = $apptSubject
    $appt.Location = $apptLocation
    $appt.BusyStatus = $olOutOfOffice
    $appt.Save()
}

#*** SET OUT OF OFFICE AUTOMATIC REPLIES **************************************

# Use of Set-MailboxAutoReplyConfiguration would need to be done on exchange server or via remote execution + permissions
# also tried Microsoft.Exchange.WebServices.dll and $service.SetUserOofSettings($Identity,$oof)... server also?

#$ns = $outlook.Session
# set out of office
#$ns.Stores[1].PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x661D000B", true)

if ($autoReplySet)
{
	# see http://telnetport25.wordpress.com/2008/03/16/quick-ish-tip-exchange-2007-setting-oof-for-users-via-powershell-2/
    $oofDll = [System.IO.Path]::Combine($myScriptPath, "EWSOofUtil.dll")
    [Reflection.Assembly]::LoadFile($oofDll)

    $oofutil = new-object EWSOofUtil.OofUtil 

    # Enabled,Disable,Scheduled
    # there are a lot of overloads. looks like for what we want we need to call main one with all the params
    # SetOof(string EmailAddress, string OofStatus, string InternalMessage, string ExternalMessage, DateTime DurationStartTime, DateTime DurationEndTime, 
    #        string InternalMessageLanguage, string ExternalMessageLanguage, string ExternalAudienceSetting, string UserName, string Password,string Domain,  string OofURL)
    $oofutil.setoof($emailAddress, "Scheduled", $internalMsg, $externalMsg, $startDt.ToUniversalTime(), $endDt.ToUniversalTime(), "", "", "", "", "", "", "")
}

#*** SEND OUT OF OFFICE EMAIL *************************************************

if ($emailSend)
{
    $smtpServer = "companyserver.domain.com"
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $mailMsg = new-object Net.Mail.MailMessage($emailAddress, $emailTo)
    $mailMsg.Subject = $emailSubject
    $mailMsg.Body = $emailBody
    $mailMsg.IsBodyHtml = $true
    $smtp.Send($mailMsg)
    echo "Sent email OOF notification to " $emailTo
    #$smtp.Send($emailAddress, $emailTo, $emailSubject, $emailBody)
}

#TODO: Update remedy. Use RemedyTFSMaster functionality? Or duplicate dlls and authentication? Or create web svc?
#TODO: Sharepoint calendar, google calendar
#TODO: Leave Request System web service use?
#TODO: Kronos and Nortel VOIP telephony research?