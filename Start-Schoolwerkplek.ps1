<#
    .NOTES

    Naam             : Start-Schoolwerkplek.ps1 
    Datum            : 15 nov 2018
    Laatst gewijzigd : 28 maart 2021 
    Auteur           : Paul Wiegmans (p.wiegmans@svok.nl)
    Github           : https://github.com/sikkepitje/Start-MagisterSWP

    .SYNOPSIS

    Starter voor Magister schoolwerkplek

    .DESCRIPTION

    Dit script opent een extern bureaublad-venster met een verbinding naar
    Magister SchoolWerkPlek (SWP) met extra functies.

    * Het SWP-venster is geopend met de meest optimale afmetingen: zo groot als
      het bureaublad het toelaat, maar zonder schuifbalken en houdt
      tegelijkertijd de Windows taakbalk zichtbaar zodat de gebruiker
      gemakkelijk kan schakelen tussen applicaties.
    * Toegang vanuit Magister SWP tot OneDrive via een zelf te kiezen
      schijfletter (standaard O:).
    * Toegang vanuit Magister SWP tot Teamsbestanden via een zelf te kiezen
      schijfletter (standaard T:).
    * Bij gebruik van meer dan één scherm wordt het SWP-venster geopend op het
      breedste scherm. Het is mogelijk om een voorkeursscherm naar keuze in te
      stellen. 

    Bekende tekortkomingen:
    * Wanneer SWP-venster wordt geopend op een scherm waarop een schaal ander
      dan 100% is gekozen, dan krijgt het venster niet de optimale afmetingen.
      De gebruiker moet dan zelf de zoom-factor van de terminalclient aanpassen.
      Anders gezegd: Als het scherm is geschaald naar 150%, zet dan handmatig
      extern bureaublad zoomfactor op 150%.

    .PARAMETER Remotehost
    
    Dit is de naam van de Magister SWP server waarmee verbinding wordt gemaakt. 
    Bijvoorbeeld: bonhoeffer.swp.nl
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [String]
    $Remotehost
)
Add-Type -AssemblyName System.Windows.Forms
Clear-Host 
$selfpath = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "****************************************"
Write-Host "** MAGISTER SWP STARTER               **"
Write-Host "** Paul Wiegmans (p.wiegmans@svok.nl) **"
Write-Host "****************************************"

# ============= BEGIN Aanpassen naar keuze ==============
# TeamFolder bevat de naam van de map in het gebruikerprofielmap waarin
# Microsoft Teams bestanden worden gesynchroniseerd. De naam van deze map wordt
# ingesteld in de Office 365 tenant instellingen.
$TeamFolder = "Stichting Voortgezet Onderwijs Kennemerland"
# Schermvoorkeuze bevat het schermnummer (0 of hoger) waarop het SWP-venster
# wordt geopend, of "auto" als het script dit zelf bepaalt.
$schermvoorkeuze = "auto"    # "auto" of het nummer van het gewenste scherm
# Hieronder worden de maximale afmetingen van het SWP-venster aangegeven. Te groot is niet fijn.
$maximale_venster_breedte = 1920 
$maximale_venster_hoogte = 1080
$Onedrive = "O:"
$TeamDrive="T:"
# ============= EINDE Aanpassen naar keuze ==============

$rdpbron = "$($selfpath)\template.rdp"
$rdptemp ="$($env:temp)\Magister SWP OneDrive.rdp"

# In geval van multimonitor configuratie, kies het breedste scherm
$minwidth = 0
$gekozenscherm = -1
$screens = [System.Windows.Forms.Screen]::AllScreens
if ($schermvoorkeuze -ge $screens.count) {
    $schermvoorkeuze = "auto"   # schermvoorkeuze was ongeldig nummer en vervalt naar automatische keuze.
}
Write-Host "Schermvoorkeuze : $schermvoorkeuze"
if ($schermvoorkeuze -eq "auto") {
    foreach ($t in (0..($screens.count - 1))) {
        $wa = $screens[$t].WorkingArea
        if ($minwidth -lt $wa.width) {
            $minwidth = $wa.width 
            $gekozenscherm = $t
            $winx = $wa.X
            $winy = $wa.Y
            $winwidth = $wa.Width
            $winheight = $wa.Height        
        }
    }
} else {
    $wa = $screens[$schermvoorkeuze].WorkingArea
    $gekozenscherm = $t
    $winx = $wa.X
    $winy = $wa.Y
    $winwidth = $wa.Width
    $winheight = $wa.Height        
}

# bepaal vensteroverhead (vensteroverheadx, vensteroverheady)
function Measure-WindowOverhead ($width, $height) {
    # Maak een venster om visueel de afmetingen te inspecteren
    $form = New-Object Windows.Forms.Form
    $font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
    $form.Text = "Bepaal Afmetingen"
    $form.Icon = [system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\calc.exe")
    $form.Width = $winwidth
    $form.Height = $winheight
    $script:vensteroverheadx = $form.Width - $form.ClientSize.Width # meestal 22
    $script:vensteroverheady = $form.Height - $form.ClientSize.Height # meestal 56

    $labelport = New-Object System.Windows.Forms.Label
    $labelport.font = $font
    $labelport.Text = 
    "Schermnummer      : " + $gekozenscherm + "`r`n" +
    "Scherm afmetingen : " + $desktopw + "," + $desktoph + "`r`n" +
    "Window positie    : " + $winx + "," + $winy + "`r`n" +
    "Window afmetingen : " + $winwidth + "," + $winheight + "`r`n"
    $labelport.top = 40
    $labelport.left = 50
    $labelport.AutoSize = $True
    $form.Controls.Add($labelport)

    #$Form.Show() | out-null   # laat het niet zien
    $form.left = $winx
    $form.Top = $winy
    $form.hide()
    #$Form.ShowDialog() | out-null
    $Form.Close()
}

# Begrenzen vensterafmetingen
if ($winwidth -gt $maximale_venster_breedte) {$winwidth = $maximale_venster_breedte}
if ($winheight -gt $maximale_venster_hoogte) {$winheight = $maximale_venster_hoogte}
$desktopw = $winwidth
$desktoph = $winheight

Measure-WindowOverhead -width $winwidth -height $winheight

# bepaal afmetingen voor remote desktop (niet het venster!)
$desktopw -= $vensteroverheadx
$desktoph -= $vensteroverheady
# RDP gebruikt niet alleen left, top, maar ook right, bottom positie
$winx2 = $winx + $winwidth
$winy2 = $winy + $winheight
Write-Host "RDP Parameters: " 
Write-Host "  Scherm            : $gekozenscherm"
Write-Host "  Schermafmetingen  : ($desktopw, $desktoph)"
Write-Host "  Vensterafmetingen : ($winwidth, $winheight)"
Write-Host "  Vensterpositie    : ($winx, $winy), ($winx2, $winy2)" 

# Pruts een RDP bestand voor mij
# uitgangspunt is 'template.rdp'. 
$rdp = get-content $rdpbron `
    | where {$_ -notlike "desktopwidth:*"} | where {$_ -notlike "desktopheight:*"} `
    | where {$_ -notlike "winposstr:*"} | where {$_ -notlike "screen mode id:*"} 
$rdp += ("desktopwidth:i:{0}" -f $desktopw)
$rdp += ("desktopheight:i:{0}" -f $desktoph)
$rdp +=  ("winposstr:s:0,1,{0},{1},{2},{3}" -f ($winx, $winy, $winx2, $winy2))
$rdp += "screen mode id:i:2"
# RDP adres instellen
$rdp = $rdp | where {$_ -notlike "full address:*"}
$rdp += "full address:s:$remotehost"
# drives instellen: drivestoredirect:s:O:\;T:\
$rdp = $rdp | where {$_ -notlike "drivestoredirect::*"}
$rdp += "drivestoredirect:s:$Onedrive\;$Teamdrive\"

$rdp | Out-File -FilePath $rdptemp -Force

# Koppel OneDrive indien aanwezig aan een schijfletter
# Koppel TeamDrive indien aanwezig aan een schijfletter 
if (!(Test-Path -Path $OneDrive)) {
    Write-Host "  Koppeling naar OneDrive wordt tot stand gebracht."
    if (Test-Path -Path "$env:OneDrive") {
        cmd.exe /c subst $OneDrive "$env:OneDrive"
    }
}
if (!(Test-Path -Path $TeamDrive)) {
    Write-Host "  Koppeling naar TeamDrive wordt tot stand gebracht."
    if (Test-Path -Path "$env:USERPROFILE\$TeamFolder") {
        cmd.exe /c subst $TeamDrive "$env:USERPROFILE\$TeamFolder"
    }
}
<##>

Write-Host 

&mstsc.exe "$rdptemp" /w $desktopw /h $desktoph
#$rdp | Out-host
#Read-Host "Druk op Enter om af te sluiten"