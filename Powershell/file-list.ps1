Get-Childitem -path  ".\*.pdf" -Recurse| Select-Object FullName, @{Name="KBytes";Expression={ "{0:N0}" -f ($_.Length / 1KB) }}| Export-Csv c:\temp\output.csv
