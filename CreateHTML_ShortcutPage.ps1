# Define input and output paths
$excelFilePath = "LinkDefinitions.xlsx"
$outputHtmlPath = "LinkPage.html"

# Import the Excel file
$data = Import-Excel -Path $excelFilePath

# Start writing the HTML file
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Clickable Images and Labels</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.5;
        }
        .item {
            margin-bottom: 20px;
        }
        .item img {
            max-width: 300px;
            height: auto;
        }
        .item a {
            text-decoration: none;
            color: blue;
        }
    </style>
</head>
<body>
"@

# Process each row and append to the HTML content
foreach ($row in $data) {
    $url = $row.URL
    $imagePath = $row.Picture
    $label = $row.Label

    $htmlContent += @"
    <div class="item">
        <a href="$url" target="_blank">
            <img src="$imagePath" alt="$label">
            <p>$label</p>
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