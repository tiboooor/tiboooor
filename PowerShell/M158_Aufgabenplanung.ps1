# .SYNOPSIS
#  Erstellt ein Log File zum Test der automatischen Aufgabenplanung unter Windows.
# 
# .NOTES 
# ===================================================================================================================
#  Version:                 1.0
#  Erstellt am:             Mittwoch 17.05.2023
#  Erstellt von:            Tibor Blasko
#  Firma:                   GBSSG - Gewerbliches Berufs- und Weiterbildungszentrum St. Gallen
#  Ort:                     St. Gallen
#  Dateiname:               M158_Aufgabenplanung.ps1
# ===================================================================================================================
# 
# .DESCRIPTION
#  Anhand einer Aufgabenplanung wird dieses Skript automatisch ausgeführt und ein Log File zur Überprüfung erstellt.
# 
# .INPUTS
#  Heutiges Datum
# 
# .OUTPUTS
#  Log File.

Param (
    # Today's date
    [Parameter(Mandatory = $false,
        Position = 0)]
    [String]$todaysDate = "$(Get-Date)",

    # Destination path
    [Parameter(Mandatory = $false,
        Position = 1,
        HelpMessage = "Bitte gib den Zielpfad ein, an dem deine Log-Datei abgespeichert werden soll.")]
    [String]$destPath = "C:\Users\tibor\Desktop\Logs"
)

If ( !(Test-Path -Path "$($destPath)\Log_$(Get-Date -Format dd/MM/yyyy/HH/mm).txt") ) {
    New-Item -ItemType File -Path "$($destPath)" -Name "Log_$(Get-Date -Format dd/MM/yyyy/HH/mm).txt"
}
Else {
    Write-Warning "Datei `"$($destPath)\Log_$(Get-Date -Format dd/MM/yyyy/HH/mm).txt`" existiert"
    Pause
}

[String]$fileName = (Get-ChildItem -Path $destPath | Sort-Object -Property CreationTime -Descending | Select-Object -First 1).Name

Write-Output "Task wurde ausgefuehrt am $($todaysDate)." | Out-File -FilePath "$($destPath)\$($fileName)" -Encoding utf8

return
