<#

    NAAM

    Start-Schoolwerkplek.ps1

    DATUM

    15 nov 2018
    
    AUTEUR

    Paul Wiegmans (p.wiegmans@bonhoeffer.nl)

    KORT BESCHRIJVING

    Magister schoolwerkplek startscript verbindt met OneDrive
    en toont windows taakbalk. 

    LANGE BESCHRIJVING

    Speciaal voor de cloud-werkplek, waarin OneDrive de locatie voor opslag 
    op een fileserver vervangt , zorgt dit script ervoor dat Magister 
    Schoolwerkplek toegang krijgt tot OneDrive voor de opslag van 
    exportbestanden, import van pasfoto's en dergelijke. Ten tweede wordt 
    Magister Schoolwerkplek gestart met een venster zo groot dat de Windows 
    taakbalk zichtbaar blijft en het de gebruiker mogelijk maakt om makkelijk 
    te schakelen tussen verschillende venster.

    TECHNISCHE BABBEL
    Hoeveel schermen heb ik ? Hou rekening met zoomfactor.
    Hoe groot is dit scherm?

winposstr:s:0,m,l,t,r,b
m = mode ( 1 = use coords for window position, 3 = open as a maximized window )
l = left
t = top
r = right  (ie Window width)
b = bottom (ie Window height)
    
screen mode id:i:x
Set x to 1 for Window mode and 2 for the RDP "Full Screen" mode. 
#>
Add-Type -AssemblyName System.Windows.Forms

Clear-Host 
$selfpath = Split-Path -Parent $MyInvocation.MyCommand.Path

# ============= zelf aanpassen ==============
$Onedrive = "X:"
$TeamDrive="Y:"
$TeamFolder = "Stichting Voortgezet Onderwijs Kennemerland"

$rdpbron = "$($selfpath)\template.rdp"
$rdptemp ="$($env:temp)\Magister SWP OneDrive.rdp"

# In geval van multimonitor configuratie, kies het breedste scherm
$minwidth = 0
$breedstescherm = -1
$screens = [System.Windows.Forms.Screen]::AllScreens
foreach ($t in (0..($screens.count - 1))) {
    $wa = $screens[$t].WorkingArea
    if ($minwidth -lt $wa.width) {
        $minwidth = $wa.width 
        $breedstescherm = $t
        $desktopw = $wa.Width
        $desktoph = $wa.Height
        $winx = $wa.Left
        $winy = 0
        $winwidth = $wa.Width
        $winheight = $wa.Height        
    }
}

# bepaal vensteroverhead (vensteroverheadx, vensteroverheady)
function Test-Window ($width, $height) {
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
    $labelport.Text = "Desktop afmetingen : " + $desktopw + "," + $desktoph
    $labelport.top = 40
    $labelport.left = 50
    $labelport.AutoSize = $True
    $form.Controls.Add($labelport)

    #$Form.Show() | out-null
    $form.left = $winx
    $form.Top = $winy
    $form.hide()
    #$Form.ShowDialog() | out-null
    $Form.Close()
}
Test-Window -width $winwidth -height $winheight

# bepaal nieuwe afmetingen voor remote desktop
$desktopw -= $vensteroverheadx
$desktoph -= $vensteroverheady
Write-Host "RDP Parameters: " $desktopw, $desktoph, $winx, $winy, $winwidth, $winheight

# Pruts een RDP bestand voor mij
# uitgangspunt is 'template.rdp'. 
$rdp = get-content $rdpbron `
    | where {$_ -notlike "desktopwidth:*"} | where {$_ -notlike "desktopheight:*"} `
    | where {$_ -notlike "winposstr:*"} | where {$_ -notlike "screen mode id:*"} 
$rdp += ("desktopwidth:i:{0}" -f $desktopw)
$rdp += ("desktopheight:i:{0}" -f $desktoph)
$rdp +=  ("winposstr:s:0,1,{0},{1},{2},{3}" -f ($winx, $winy, $winwidth, $winheight))
$rdp += "screen mode id:i:2"
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

mstsc.exe "$rdptemp" /w $desktopw /h $desktoph
