#Fabio's Out of Office PS Script
#searches next calendar entry with "Out of office" status and schedules a autoreply accourding to it
$executePath = $MyInvocation.MyCommand.Path  #Pfad an dem das Script zu zeit läuft
$executeDir = Split-Path $executePath

#find E-Mail Address for the current user
function Get-EMailAddress {
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    return $searcher.FindOne().Properties.mail
}



Function Get-OutlookOOFAppointments {
   param ( 
         [Int] $NumDays = 7,
         [DateTime] $Start = [DateTime]::Now,
         [DateTime] $End   = [DateTime]::Now.AddDays($NumDays),
         [int] $BusyStatus = 3 #0 = Free, 1 Tenative , 2= Busy, 3 = Out of office
   )
 
   Process {
      $outlook = New-Object -ComObject Outlook.Application
      $session = $outlook.Session
      $session.Logon()
 
      $apptItems = $session.GetDefaultFolder(9).Items
      $apptItems.Sort("[Start]")
      $apptItems.IncludeRecurrences = $true
      $apptItems = $apptItems
 
      $restriction = "[End] >= '{0}' AND [Start] <= '{1}' AND [BusyStatus] = '{2}'" -f $Start.ToString("g"), $End.ToString("g"), $BusyStatus
      
      $outlook = $session = $null;
      
      #find first OOF appointment
      return $apptItems.Find($restriction)   
      
   } 
}


Function Set-ExchangeAutoReply{
   param (
         [DateTime] $Start = [DateTime]::Now,
         [DateTime] $End   = [DateTime]::Now
   )

    "Enabling auto reply for :  " + $Start + " - "+  $End
    $EMailAdress = Get-EMailAddress
   
    # see http://telnetport25.wordpress.com/2008/03/16/quick-ish-tip-exchange-2007-setting-oof-for-users-via-powershell-2/
    $oofDll = [System.IO.Path]::Combine($executeDir, "EWSOofUtil.dll")
    [Reflection.Assembly]::LoadFile($oofDll)

    $oofutil = new-object EWSOofUtil.OofUtil 

    # Enabled,Disable,Scheduled
    # there are a lot of overloads. looks like for what we want we need to call main one with all the params
    # SetOof(string EmailAddress, string OofStatus, string InternalMessage, string ExternalMessage, DateTime DurationStartTime, DateTime DurationEndTime, 
    #        string InternalMessageLanguage, string ExternalMessageLanguage, string ExternalAudienceSetting, string UserName, string Password,string Domain,  string OofURL)
    #$oofutil.setoof($emailAddress, "Scheduled", $internalMsg, $externalMsg, $startDt.ToUniversalTime(), $endDt.ToUniversalTime(), "", "", "", "", "", "", "")

}


function main{
param ( 
         [Int] $NumDays = 7,
         [DateTime] $Start = [DateTime]::Now,
         [DateTime] $End   = [DateTime]::Now.AddDays($NumDays),
         [int] $BusyStatus = 3 #0 = Free, 1 Tenative , 2= Busy, 3 = Out of office
   )
  
  #look for an "out of office" appointment
  $apptItem = OutlookOOFAppointments -NumDays $NumDays -Start $Start -End $End -BusyStatus $BusyStatus
  
  If($apptItem -ne $null){
    Set-ExchangeAutoReply -Start $apptItem.Start -End $apptItem.End
  }else{
    write-Host  -ForegroundColor Yellow "No Out of Office appointment found for the next $NumDays days."
  }

}

main -NumDays 5

 



