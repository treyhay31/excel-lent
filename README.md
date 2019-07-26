# excel-lent
some stuff with excel... probably with powershell

## Problem \#1 - Transposing based on a prefix

Here is what we have:

|Location|Person|
|----|----|
|ZZQ|ZZQ-BOB|
|ZZQ|ZZQ-JON|
|ZZQ|ZZQ-TERRY|
|ZZQ|ZZQ-JONAS|
|ZZL|ZZL-BOB|
|ZZL|ZZL-JOE|
|ZZL|ZZL-GARY|
|FTR|FTR-BOB|
|FTR|FTR-BILL|
|FTR|FTR-BORLIN|
|FTR|FTR-TIM|
|FTR|FTR-TOOL|

Hers is what we want:

|Location|Person|Person|Person|Person|Person|
|----|----|----|----|----|----|
|ZZQ|ZZQ-BOB|ZZQ-JON|ZZQ-TERRY|ZZQ-JONAS||
|ZZL|ZZL-BOB|ZZL-JOE|ZZL-GARY|||
|FTR|FTR-BOB|FTR-BILL|FTR-BORLIN|FTR-TIM|FTR-TOOL|

### Solution \#1 - Powershell With EPPLUS

1. Run the powershell script `transpose-excel-sheet.ps1`
    - Provide the sheet index (example. `2`)
    - Provide the prefix number of characters (above example would be: `3`)
    - Provide the column alphabetical character (example: `A`)
    - Provide the path to the Excel file `.xls|.xlsx` file

```{powershell}
$sheetNumber = 2
$prefixLength = 3
$column = "A"
$file = "C:\path\to\file.xlsx"

PS> .\transpose-excel-sheet.ps1 $sheetNumber $prefixLength $column $file
```

### Solution \#2 - Powershell Without EPPLUS

1. Export sheet with data to `.csv`
1. Run the powershell script `transpose-csv.ps1`
    - Provide the prefix number of characters (above example would be: `3`)
    - Provide the path to the `.csv` file that was exported

```{powershell}
PS> .\transpose-csv.ps1 3 C:\path\to\file.csv
```
