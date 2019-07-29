Install-Module -Name ImportExcel -RequiredVersion 5.4.2

param(
  [int]$sheetIndex,
  [ValidateScript({ Test-Path $_ -Type Leaf })]
  [string]$xlPath
)

$data = Get-Content $xlPath 

"Here is the data:"
$data

$data | ForEach-Object { <# do excel stuff #> } 

# $data.Save()
