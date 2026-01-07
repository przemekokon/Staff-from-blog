#CSV file should contain 'mail' header

$groups = Import-Csv .\groups.csv

foreach ($g in $groups) {

    $group = Get-ADGroup -Filter "mail -eq '$($g.mail)'" -Properties mail

    if ($group) {
        Set-ADGroup -Identity $group.DistinguishedName -Replace @{msExchExtensionAttribute20 = "Lorem Ipsum"}
        Write-Host "OK: $($g.mail) -> msExchExtensionAttribute20 has been set" -ForegroundColor Green
    }
    else {
        Write-Host "ERROR: DL not found $($g.mail)" -ForegroundColor Red
    }
}
