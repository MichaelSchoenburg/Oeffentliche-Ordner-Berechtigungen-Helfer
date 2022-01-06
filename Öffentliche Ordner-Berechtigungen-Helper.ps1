<#
.SYNOPSIS
    Öffentliche Ordner-Berechtigungen-Helfer
.DESCRIPTION
    Dieses Script fügt die "veröffentlichender Bearbeiter"-Berechtigung auf die ausgewählten öffentlichen Ordner für einen öffentlichen Benutzer hinzu.
    Dieses Script ist für die Einrichtung neuer Mitarbeiter gedacht, da hiermit keine Berechtigungen editiert werden können.
.INPUTS
    None
.OUTPUTS
    None
.NOTES
    Author: Michael Schönburg
    Version: 1.0
    Last Change: 23.12.2021
    GitHub Repository: https://github.com/MichaelSchoenburg/Oeffentliche-Ordner-Berechtigungen-Helfer
#>


Clear-Host

Write-Host "Verbinde zu online Exchange..."
Connect-ExchangeOnline

Write-Host ""
Write-Host ""

#Benutzer wählen
$mboxs = Get-Mailbox

$title = "Für welchen Benutzer wollen Sie die Berechtigungen ändern?"
$TargMBX = $mboxs | Select-Object Name, DisplayName, Alias, PrimarySMTPAddress | Out-GridView -PassThru -Title $title

#Benutzerauswahl: Berechtigungen von einem Benutzer kopieren oder neu zusammenstellen
$title    = 'Titel'
$question = 'Möchten Sie Ordner-Berechtigungen von einem Benutzer kopieren oder neu zusammenstellen?'
$choices  = '&Kopieren', '&Neu'

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
if ($decision -eq 0) {
    Write-Host 'Sie haben sich entschieden, die Berechtigungen von einem Benutzer zu kopieren.'
    
    #Benutzer auswählen
    $title = "Von welchem Benutzer möchten Sie die Berechtigungen kopieren?"
    $ChoiceMBXSource = $mboxs | Select-Object Name, DisplayName, Alias, PrimarySMTPAddress | Out-GridView -PassThru -Title $title
    
    #Auslesen, welche Berechtigungen der Benutzer hat
    Write-Host "Die Berechtigungen werden ausgelesen..."
    $PFs = Get-PublicFolder \ -recurse
    $PFPermissions = @()
    for ($i = 0; $i -le $PFs.count -1; $i++) {
        Write-Progress -Activity "Berechtigungen auslesen" -PercentComplete (($i*100)/$PFs.count) -Status "Verarbeite $($i) von $($PFs.count)"

        $PFPermission = $PFs[$i] | Get-PublicFolderClientPermission | Where-Object {$_.user -like $ChoiceMBXSource.DisplayName}
        $PFPermissions += $PFPermission
    }

    #Berechtigungen anzeigen und Bestätigung anfordern
    Write-Progress -Activity "Berechtigungen auslesen" -Status "Ready" -Completed
    Write-Host ""
    Write-Host "Der Benutzer hat folgende Berechtigungen:"
    $PFPermissions.Identity
    Pause
    
    #Berechtigungen setzen
    Write-Host "Die Berechtigungen werden gesetzt..."
    for ($i = 0; $i -le $PFPermissions.count -1; $i++) {
        Write-Progress -Activity "Berechtigungen setzen" -PercentComplete (($i*100)/$PFPermissions.count) -Status "Verarbeite $($i) von $($PFs.count)"

        if (Get-PublicFolderClientPermission $PFPermissions[$i].Identity -User $TargMBX.Alias -ErrorAction SilentlyContinue) {
            Write-Host "Berechtigung existiert bereits." -ForegroundColor Green
        } else {
            Write-Host "Berechtigung für $($PFPermissions[$i].FolderName) wird hinzugefügt." -ForegroundColor Gray
            try {
                $null = Add-PublicFolderClientPermission $PFPermissions[$i].Identity -User $TargMBX.PrimarySMTPAddress -AccessRight $PFPermissions[$i].AccessRights -ErrorAction Stop
                if (Get-PublicFolderClientPermission $PFPermissions[$i].Identity -User $TargMBX.PrimarySMTPAddress -ErrorAction SilentlyContinue) {
                    Write-Host "Die Berechtigung für $($PFPermissions[$i].FolderName) konnte erfolgreich hinzugefügt werden." -ForegroundColor Cyan
                } else {
                    Write-Host "Die Berechtigung für $($PFPermissions[$i].FolderName) konnte nicht hinzugefügt werden." -ForegroundColor Red
                }
            } catch {
                Write-Host "Es gibt ein Problem mit dem öffentlichen Ordner $($PFPermissions[$i].FolderName)" -ForegroundColor Red
                Write-Host "Error: "$_.Exception.Message -ForegroundColor Red
            }
        }
        Write-Progress -Activity "Berechtigungen setzen" -Status "Ready" -Completed
    }
} else {
    Write-Host 'Sie haben sich entschieden, die Berechtigungen neu zusammenzustellen.'
    #Ordner auswählen
    $pubfolders = get-publicfolder -Recurse
    
    $title = 'Wählen Sie alle öffentliche Ordner an, auf welche der Benutzter Zugriff haben soll (verwenden Sie ggf. Strg + Linksklick).'
    $choicepubf = $pubfolders.Identity | Out-GridView -PassThru
    
    Write-Host "Die Berechtigungen werden gesetzt..."

    for ($i=0; $i -le $choicepubf.Count - 1; $i++) {
        Write-Progress -Activity “Setze Berechtigungen” -status "Verarbeite Berechtigung $i von $($choicepubf.count)" -percentComplete ($i / $choicepubf.count*100)
        Write-Host "Setze Berechtigung für $($TargMBX.PrimarySMTPAddress) fuer $($choicepubf[$i].Identity):"
        $null = Add-PublicfolderclientPermission -Identity $choicepubf[$i] -User $TargMBX.Alias -AccessRights PublishingEditor
    }

    Write-Host "Die Berechtigungen wurden gesetzt. Das Skript wird beendet."
}

Get-PSSession | Remove-PSSession
