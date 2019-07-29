param(
  [int]$sheetIndex,
  [ValidateScript({ Test-Path $_ -Type Leaf })]
  [string]$csvPath
)

$data = Get-Content $csvPath | ConvertFrom-Csv

$data | `
  Group-Object { <# find the prefix #> } | `
  ForEach-Object { <# list things out sideways #> } | `
  Set-Content $csvPath.Replace(".csv", "-transposed.csv")

