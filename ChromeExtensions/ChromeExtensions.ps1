#|-----------------------------------------------------------------------|
#|  					  S U B F U N C T I O N S						 |
#V-----------------------------------------------------------------------V


# Check if script is running as Adminstrator and if not use RunAs 
# Use Check Switch to check if admin 
function Use-RunAs {    
    Param(
		[Switch]$Check
	) #end param
     
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
    if ($Check) { return $IsAdmin }     
    if ($MyInvocation.ScriptName -ne "") 
    {  
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $arg = "-file `"$($MyInvocation.ScriptName)`"" 
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'  
            } 
            catch 
            { 
                Write-Warning "Error - Konnte Script nicht mittels Run-As neugestartet werden"
                break               
            } 
            exit # Quit this session of powershell 
        }  
    }
    else  
    {  
        Write-Warning "Error - Bitte Script als .ps1 Datei speichern"  
        break 
    }  
}

#Prüft ob Policies vorhanden sind und löscht jene gegebenenfalls
function DeletePolicies{
	Param(
		[Parameter(Mandatory=$True,Position=1)]
		[string]$pPolicyPath
	) #end param
	
	#Prüfen ob Registry die Policy-Einträge vorhanden sind
	$PoliciesExists = Test-Path $pPolicyPath
	if($PoliciesExists){
	
		#Fragen ob Script mit Admin rechten aufgerufen wurde. Falls nein. Erneut starten (mit Admin privilegien)
		Use-RunAs
		
		#Delete Registry Entries that disable Extensions
		Remove-Item $pPolicyPath -Recurse
		
		#Screen output / Rückmeldung
		if($?){
			write-host "Policies erfolgreich gelöscht" -foregroundcolor "Green"
			write-host "Bitte Chrome neu starten um installierte Erweiterungen wieder zu aktivieren" -foregroundcolor "Green"
		}else{
			write-warning "ERROR: Policies konnten nicht gelöscht werden" -foregroundcolor "Red"
		}
	}else{
		write-host "INFO: Keine Policies gefunden. Keine Löschung notwendig" -foregroundcolor "White"
	}
	
	#Sleep Delay um Meldungen anzuzeigen
	Start-Sleep -s 4
}

#|-----------------------------------------------------------------------|
#|					 M A I N     P R O G R A M M E						 |
#V-----------------------------------------------------------------------V
$PolicyPath = "HKLM:\SOFTWARE\Policies\Google\Chrome"

DeletePolicies $PolicyPath







