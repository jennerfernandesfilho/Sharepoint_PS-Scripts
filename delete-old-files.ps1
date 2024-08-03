# Parameters
$SiteURL = "Library_URL"
$ListName = "Library_Name"
$CSVFile = "C:\User\file.csv"
$DateThreshold = Get-Date -Year 2020 -Month 12 -Day 31
$PageSize = 500 # Number of items to process per batch

# Connect to SharePoint Online
Connect-PnPOnline -Url $SiteURL -Interactive

# Initialize an array to store the results
$results = @()

# Initialize variables for paging
$continuationToken = $null

do {
    # Get items in batches
    $items = Get-PnPListItem -List $ListName -Fields "Modified", "FileRef", "FileLeafRef" -PageSize $PageSize
    
    # Check if there are items to process
    if ($items.Count -eq 0) {
        break
    }
    
    # Filter items based on the last modified date
    $itemsToDelete = $items | Where-Object { $_["Modified"] -lt $DateThreshold }
    
    foreach ($item in $itemsToDelete) {
        $fileUrl = $item["FileRef"]
        $fileName = $item["FileLeafRef"]
        
        try {
            # Attempt to delete the file
            Remove-PnPFile -ServerRelativeUrl $fileUrl -Force
            
            # Log the successful deletion
            $results += [PSCustomObject]@{
                FileName = $fileName
                FileUrl  = $fileUrl
                Status   = "Deleted"
                ErrorMessage = ""
            }
            
            Write-Host "Successfully deleted: $fileName"
        }
        catch {
            # Capture and log the error message
            $errorMessage = $_.Exception.Message
            
            # Log the failed deletion
            $results += [PSCustomObject]@{
                FileName = $fileName
                FileUrl  = $fileUrl
                Status   = "Failed"
                ErrorMessage = $errorMessage
            }
            
            Write-Host "Failed to delete: $fileName. Error: $errorMessage"
        }
    }
    
    # Update continuation token for the next batch
    $continuationToken = $items.PagingToken

} while ($continuationToken) # Continue until no more items are returned

# Export results to CSV
$results | Export-Csv -Path $CSVFile -NoTypeInformation


Write-Host "Operation completed. Results saved to $CSVFile"