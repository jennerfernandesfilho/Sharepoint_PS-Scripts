# Parameters
$SiteURL = "Library_URL"
$ListName = "Library_Name"
$CSVFile = "C:\User\file.csv"

# Connect to SharePoint Online
Connect-PnPOnline -Url $SiteURL -Interactive

# Get the list
$List = Get-PnPList -Identity $ListName

# Get Folders from the Library - with progress bar
$global:counter = 0
$FolderItems = Get-PnPListItem -List $ListName -PageSize 500 -Fields FileLeafRef -ScriptBlock { Param($items) $global:counter += $items.Count; Write-Progress -PercentComplete ($global:Counter / ($List.ItemCount) * 100) -Activity "Getting Items from List:" -Status "Processing Items $global:Counter to $($List.ItemCount)";}  | Where {$_.FileSystemObjectType -eq "Folder"}
Write-Progress -Activity "Completed Retrieving Folders from List $ListName" -Completed

$FolderStats = @()
# Get Files and Subfolders count on each folder in the library
ForEach($FolderItem in $FolderItems)
{
    # Get Files and Folders of the Folder
    Get-PnPProperty -ClientObject $FolderItem.Folder -Property Files, Folders | Out-Null
    
    # Collect file details for the folder
    $FileDetails = @()
    foreach ($file in $FolderItem.Folder.Files)
    {
        $FileDetails += [PSCustomObject]@{
            FilePath = $file.ServerRelativeUrl
            LastModified = $file.TimeLastModified
            Type = "File"
            FileSize = $file.Length
        }
    }
    
    # Collect folder data
    $FolderData = [PSCustomObject][ordered]@{
        URL              = $FolderItem.FieldValues.FileRef
        FilesCount       = $FolderItem.Folder.Files.Count
        SubFolderCount   = $FolderItem.Folder.Folders.Count
        LastModified     = $FolderItem.Folder.TimeLastModified
        Type             = "Folder"
        FileSize         = 0
    }
    $FolderStats += $FolderData

    # Collect file data for each folder
    foreach ($file in $FileDetails)
    {
        $FolderStats += [PSCustomObject][ordered]@{
            URL              = $file.FilePath
            FilesCount       = 1
            SubFolderCount   = 0
            LastModified     = $file.LastModified
            Type             = "File"
            FileSize         = $file.FileSize
        }
    }
}

# Export the data to CSV with UTF-8 BOM
$CsvContent = $FolderStats | ConvertTo-Csv -NoTypeInformation -Delimiter "|"
[System.IO.File]::WriteAllLines($CSVFile, $CsvContent, [System.Text.Encoding]::UTF8)

Write-Output "File and folder details exported to $CSVFile"
