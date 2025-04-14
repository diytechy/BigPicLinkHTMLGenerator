$scriptPathTst = $MyInvocation.MyCommand.Path
if ($scriptPathTst) {$scriptPath = $scriptPathTst}
$scriptDir = Split-Path $scriptPath
Set-Location $scriptDir
# Define input and output paths
$excelFilePath = "LinkDefinitions.xlsx"
$csvFileName = "LinkDefinitions.csv"
$outputHtmlPath = "LinkPage.html"

#Resolve csv
$sep = [System.IO.Path]::DirectorySeparatorChar
$csvFilePath = $scriptDir.ToString() + $sep.ToString() + $csvFileName

# Convert the Excel file
$Full2XLSX = Resolve-Path $excelFilePath
$ExcelFilePath = $Full2XLSX.Path.ToString()
$Excel = New-Object -ComObject Excel.Application

try{
    $Workbook = $Excel.Workbooks.Open($ExcelFilePath)
    foreach ($sheet in $Workbook.worksheets){
        if(Test-Path -Path $csvFilePath){
            remove-item $csvFilePath
        }
        $sheet.SaveAs($csvFilePath,6)
    }
} catch {Write-Error("Oh Darn!")}
$Excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)

# Import the converted csv
$data = Import-CSV $csvFilePath


if(Test-Path -Path $outputHtmlPath){
    remove-item $outputHtmlPath
}
# Start writing the HTML file
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Image Tiles</title>
    <style>
        body {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            margin: 0;
            padding: 0;
        }
        a {
            margin: 5px;
        }
        img {
            max-width: 300px;
            max-height: 300px;
            object-fit: cover;
        }
    </style>
</head>
<body>
"@

# Process each row and append to the HTML content
foreach ($row in $data) {
    $url = $row.URL
    $imagePath = "Assets" + $sep.ToString() + $row.Picture
    $label = $row.Label

    $htmlContent += @"
    <div class="item">
        <a href="$url" target="_blank">
            <img src="$imagePath" alt="$label">
            <p style = "font-size:20px" align="center">$label</p>
        </a>
    </div>
"@
}

# Close the HTML structure
$htmlContent += @"
</body>
</html>
"@

# Write the HTML content to a file
Set-Content -Path $outputHtmlPath -Value $htmlContent -Encoding UTF8

Write-Host "HTML file has been generated successfully at $outputHtmlPath"