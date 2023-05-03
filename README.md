```powershell
$envObjects = Import-Csv "Path/To/CSV" -Delimiter ";" | Where-Object { ($_.Name -ne "ALLURE_HOME") -and ($_.Name -ne "MAVEN_HOME") -and ($_.Name -ne "TNS_ADMIN") }
$envVars = Get-ChildItem -Path "Env:\"

ForEach ( $envObject in $envObjects ) {
    If ( $envObject.Name -notin $envVars.Name ) {
        $solution = $envObject.Name
    }
}

[String]$inputToken = "RAIF{$($solution)}"
Write-Host "$($inputToken)" -ForegroundColor Yellow

$auswertenReturn = Path/To/Auswerten.ps1 -InputToken $inputToken
If ( $auswertenReturn -eq $true ) {
    Write-Host "supi" -ForegroundColor Green
}
```
