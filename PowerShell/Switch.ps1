Param (
    [Parameter(Mandatory = $false,
        Position = 0)]
    [String]$wochentag = "$(Get-Date -Format "dddd")"
)

[String[]]$wochenende = (
    "Samstag",
    "Sonntag"
    )
[String]$tag = $null

If ( $wochentag -in $wochenende ) {
    Write-Warning "Heute ist kein guter Wochentag. Fuer den Wochentag `"$($wochentag)`" ist leider keine Datei vorhanden."
    return
}
Else {
    Switch ( $wochentag ) {
        Montag {
            $tag = "Montag"
        }
        Dienstag {
            $tag = "Dienstag"
        }
        Mittwoch {
            $tag = "Mittwoch"
        }
        Donnerstag {
            $tag = "Donnerstag"
        }
        Freitag {
            $tag = "Freitag"
        }
    }
}

[String]$loesung = Get-Content -Path "D:\Files\$($day).txt" | ? -Property "Length" -EQ 6
[String]$token = "RAIF{$($loesung)}"
Write-Host $token -ForegroundColor Yellow
