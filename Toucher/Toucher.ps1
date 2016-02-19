#Fabio's local Group Membership Toolkit
#Released under the terms of the MIT X-license. 

#SETTINGS 
$scriptpath = $MyInvocation.MyCommand.Path  #Pfad an dem das Script zu zeit läuft. nicht Ändern
$dir = Split-Path $scriptpath 				#Pfad an dem das Script zu zeit läuft. nicht Ändern
$pathlistPath = "$dir\pathlist.txt"         #Pfad zur Pfad Liste
$now = Get-Date                             #Jetztiges Datum + Zeit
$LogFilePath = "$dir\ToucherLog.txt" #Pfad zur Logdatei. (wird automatisch erstellt wenn nicht vorhanden) 

# seting ausgeben
function printSettings { 
	Write-Host "" #newline
	Write-Host -ForegroundColor Cyan "Settings:"
	Write-Host -ForegroundColor Gray "Eingelesenes File: " $pathlistPath
    Write-Host -ForegroundColor Gray "Zeitpunkt: " $now

    Write-Host -ForegroundColor Cyan "" #newline
    Write-Host -ForegroundColor Cyan "Hinweis:"
    Write-Host -ForegroundColor White "Da alle Files rekursiv gezählt werden müssen, kann das Auslesen der Pfade etwas lange dauern"

}

cls #clear
Write-Host -ForegroundColor Cyan "Fabio's File touch script"
Write-Host -ForegroundColor White "Das Script modifizert rekursiv das Updatedatum von allen Files in den angegebenen Unterverzeichnissen"

$LogArray = New-Object System.Collections.ArrayList
$title = "Bitte Aktion wählen"
$message = "Welche Aktion soll für die aufgelisteten Pfade durchgeführt werden?"
$update = New-Object System.Management.Automation.Host.ChoiceDescription "&Update", `
    "Alle Files in den aufgewählten Verzeichnissen werden rekursiv mit einem neuen datum beschrieben"
$escape = New-Object System.Management.Automation.Host.ChoiceDescription "&Abbruch", `
    "Abbruch; Keine Änderungen"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($update, $escape)

# MAIN Function
$CSVpathOK = Test-Path $pathlistPath
if($CSVpathOK -eq $true){ #CSV file gefunden?
    $PathList = Get-Content $pathlistPath
    If($?){
        # List Paths found in CSV and let User choose what to do
        printSettings

        Write-Host -ForegroundColor DarkGreen ""
        Write-Host -ForegroundColor Cyan "Im File gefundene Verzeichnisse:"
        foreach($Path in $PathList){
            $pathOK = Test-Path $Path
            if ($pathOK){
                $noFiles = (Get-ChildItem -Recurse "$Path").Count + 1 #count starts with 0
                Write-Host -ForegroundColor White "$Path   ($noFiles Files)"
            }else{
                Write-Host -ForegroundColor Red "$Path   nicht gefunden."
                Write-Host -ForegroundColor Red "Script wurde abgebrochen..."
                $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                exit

            }
        }
        $modus = $host.ui.PromptForChoice($title, $message, $options, 0)
        if ($modus -ne 1){
            Write-Host -ForegroundColor Cyan "Files werden verarbeitet.."
            Write-Host -ForegroundColor DarkGreen "" #Line Break

            foreach($Path in $PathList){
                $pathOK = Test-Path $Path
                if ($pathOK){
                    $Files = Get-ChildItem -Recurse "$Path"
                    foreach($file in $Files){
                        try{
                             $file.LastWriteTime=$now #Fehlermeldung unterbinden
                        }catch [Exception]{
                            Write-host -ForegroundColor red $file.FullName "konnte nicht geändert werden" # File konnte nicht geändert werden
                            $err = $true
                           $file.FullName + " konnte nicht geändert werden" | Out-File $LogFilePath -Append
                        } #end try/catch touch file 
                    }
                    Write-Host -ForegroundColor Green "$Path    aktualisiert"
                } #pathOK ?
            } #foreach($Path in $PathList)
        }else{

        } #modus Selektion
    }else{
    write-Host -ForegroundColor Red "CSV File konnte nicht gelesen werden"
    }
}else{
write-host -ForegroundColor Red "CSV File nicht gefunden; Script beendet"
}

Write-Host -ForegroundColor White ""
Write-Host -ForegroundColor White ""
if($err -eq $true) {
    write-Host -ForegroundColor Yellow "Fehler wurden im Logfile protokolliert"
}

Write-Host -ForegroundColor White ""

Write-Host -ForegroundColor White "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")