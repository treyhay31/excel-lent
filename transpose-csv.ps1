param(
  [int]$prefixLength,
  [ValidateScript({ Test-Path $_ -Type Leaf })]
  [string]$csvPath="data.csv",
  [string]$sep=","#separator

)

"`r`nStarting transpose-csv..."
"`r`n Parameters: "
"  prefix length: $prefixLength"
"  csv file: $csvPath"
"  separator: $sep"

" Getting csv data..."
$data = Get-Content $csvPath | ConvertFrom-Csv

$headers = (Get-Content $csvPath -First 1).Split($sep)

" Header Info..."
$infoColumn = $headers[0]
"  Info Column: $infoColumn"
$dataColumn = $headers[1]
"  Data Column: $dataColumn"

" 1. Grouping data on first $prefixLength characters..."
$grouped = $data.$dataColumn | Group-Object { $_.Substring(0, $prefixLength) }

$counts = @()

$grouped | ForEach-Object{
  $counts += $_.Count
}

$max = ($counts | Measure-Object -Max).Maximum

" 2. Transposing..."
$rows = ""
$grouped | ForEach-Object {
  
  $groups = $_.Group | ForEach-Object { 
    "$_|" 
  }
  
  $leftovers = 1..($max - $_.Count) | ForEach-Object { "|" }
  
  $rows += "`r$($_.Name)|$($groups)$($leftovers)".Replace(" ","")
}

$groupHeaders = 1..$max | ForEach-Object { "$dataColumn $_|" }

$transposedFile = $csvPath.Replace(".csv", "-transposed.csv")

" 3. Ouputting to $transposedFile..."
"$infoColumn|$($groupHeaders.Replace(' ', ''))$rows" | Set-Content $transposedFile

"`r`nFinishing transpose-csv...`r`n"
