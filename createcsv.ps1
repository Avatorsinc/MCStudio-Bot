# Define the path to the folder
$basePath = "E:\Jsons\Home\PDA\DK\foetex"

# Define the output CSV file path
$outputCsv = "E:\Jsons\Home\PDA\DK\foetex\FolderNames.csv"

# Get all folder names in the specified path
$folders = Get-ChildItem -Path $basePath -Directory

# Loop through each folder, extract the first 4 letters, and append 'Migration'
$modifiedNames = $folders | ForEach-Object {
    $folderName = $_.Name
    $firstFour = $folderName.Substring(0, 4) # Get the first 4 characters
    $newName = "$firstFour`Migration"        # Append 'Migration'
    [PSCustomObject]@{
        OriginalName = $folderName
        NewName = $newName
    }
}

# Export the results to a CSV file
$modifiedNames | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Folder names have been saved to $outputCsv"
