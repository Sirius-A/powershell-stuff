#Fabio's local Group Membership Toolkit
#Released under the terms of the MIT X-license. 

#SETTINGS 
$computerName = "atvies99000srv.ww300.siemens.net" #Name des Computers/servers auf dem die lokale Gruppe gespeichert wird   (für Lokal "$env:computername" eintragen)
$Delimiter = "\" #benützt um CSV zu lesen
$HeaderCSV = ("Domain","Username")  #Name der Spalten im CSV File
$GroupName = "Clientele_Attachment_rw" #Name der Lokalen Gruppe, welche bearbeitet wird
$scriptpath = $MyInvocation.MyCommand.Path  #Pfad an dem das Script zu zeit läuft. nicht Ändern
$dir = Split-Path $scriptpath 				#Pfad an dem das Script zu zeit läuft. nicht Ändern
$LogFilePath = "$dir\GroupManagerLog.txt" #Pfad zur Logdatei. (wird automatisch erstellt wenn nich vorhanden) 
$userlistCSVPath = "$dir\userlist.csv"    #Pfad zur CSV User liste

# Check if script is running as Adminstrator and if not use RunAs 
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
                #Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'
				Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'
            } 
            catch 
            { 
                Write-Warning "Error - Konnte Script nicht mittels Run-As neustarten"
                break               
            } 
            exit # Quit this session of powershell 
        }  
    }
    else  
    {  
        Write-Warning "Error - Bitte Script als .ps1 Script speichern"  
        break 
    }  
}
Use-RunAs


# seting ausgeben
function printSettings { 
	Write-Host "" #newline
	Write-Host -ForegroundColor Cyan "Settings:"
	Write-Host -ForegroundColor White "Computername: " $computer.Name
	Write-Host -ForegroundColor White "Gruppenname: " $GroupName 
	Write-Host -ForegroundColor Gray "Userliste Pfad: " $userlistCSVPath	
	Write-Host -ForegroundColor Gray "Header für Userliste: "  $HeaderCSV
	Write-Host -ForegroundColor Gray "Trennzeichen für Userliste":  $Delimiter
}

cls #clear
Write-Host -ForegroundColor Cyan "Fabio's Group Membership Toolkit"

Write-Host -ForegroundColor White "Das Tool erlaubt das Modifizieren einer Gruppe anhand einer Userliste im CSV Format"

$LogArray = New-Object System.Collections.ArrayList
$title = "Bitte Aktion wählen"
$message = "Welche Aktion soll für die User ausgeführt werden?"
$add = New-Object System.Management.Automation.Host.ChoiceDescription "&Add", `
    "Die User werden zur Gruppe $GroupName hinzugefügt"
$remove = New-Object System.Management.Automation.Host.ChoiceDescription "&Remove", `
    "Die User werden aus der Gruppe $GroupName entfernt"
$escape = New-Object System.Management.Automation.Host.ChoiceDescription "&Escape", `
    "Abbruch; Keine Änderungen"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($add, $remove, $escape)


# MAIN Function
if([ADSI]::Exists("WinNT://$computerName,computer")) { 
	$computer = [ADSI]"WinNT://$computerName,computer"
	$Group = $computer.psbase.children.find($GroupName) 
	if(-NOT([string]::IsNullOrEmpty($Group.Guid))){ #Gruppe gefunden ?
		$CSVpathOK = Test-Path $userlistCSVPath
		if($CSVpathOK -eq $true){ #CSV file gefunden?
			printSettings
			$UserList = Import-Csv $userlistCSVPath -Delimiter $Delimiter -Header $HeaderCSV
			If($?){
				# List Users found in CSV and let User choose what to do
				
				Write-Host -ForegroundColor DarkGreen ""
				Write-Host -ForegroundColor Cyan "Im CSV-File gefundene User:"
				foreach($User in $UserList){
					Write-Host -ForegroundColor White "  " $User.Domain "\" $User.Username -Separator ""
				}
				$modus = $host.ui.PromptForChoice($title, $message, $options, 0)
				if ($modus -ne 2){
					Write-Host -ForegroundColor Cyan "User werden verarbeitet.."
					Write-Host -ForegroundColor DarkGreen ""
					
					#Add / Remove Users from Group
					
					foreach($User in $UserList){
						$date = Get-Date
						$currtime = $date.ToShortDateString() + " " + $date.ToShortTimeString()
						Try{
							switch ($modus){
								0 { #Add
									$Group.Add("WinNT://" + $User.Domain + "/" + $User.Username)
									Write-Host -ForegroundColor Green "Added " $User.Domain "/" $User.Username  -Separator ""
									[void]$LogArray.add( "[" + $currtime + "]: " + $User.Domain + "/" + $User.Username + " added to group " + $GroupName )
								}
								1 { #remove
									$Group.Remove("WinNT://" + $User.Domain + "/" + $User.Username)
									Write-Host -ForegroundColor Green "Removed " $User.Domain "/" $User.Username -Separator ""
									[void]$LogArray.add( "[" + $currtime + "]: " + $User.Domain + "/" + $User.Username + " removed from group " + $GroupName )
								}
							}
						}Catch [System.Management.Automation.MethodInvocationException]{
							Write-Host -ForegroundColor Red "Fehler für " $User.Domain "\" $User.Username ":   " $_.Exception.Message -Separator ""
							[void]$LogArray.add( "[" + $currtime + "]: " + $User.Domain + "/" + $User.Username + " FEHLER: " + $_.Exception.Message)
						}						
					}
					
					#Write into log file
					$LogArray | Out-File $LogFilePath -Append
					If($? -eq $true){
						Write-Host "" #new line
						Write-Host "Aktivitäten wurden im Logfile protokolliert. (Pfad:" $LogFilePath ")"
					}else{
						Write-Host -ForegroundColor Red "Fehler beim Schreiben ins Logfile! (Pfad:" $LogFilePath ")" 
					}
				}
			}else{
				write-Host -ForegroundColor Red "CSV File konnte nicht gelesen werden"
			}
		}else{
			write-host -ForegroundColor Red "CSV File nicht gefunden; Script beendet"
		}
	}else{
		write-host -ForegroundColor Red "Gruppe" $GroupName "nicht gefunden; Script abgebrochen"
	}
}else{
	write-Host -ForegroundColor Red "Computer nicht gefunden"
}

Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White ""

Write-Host -ForegroundColor White "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
