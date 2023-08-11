# Get current location and set the working to directory to it
$currentPath = split-path -parent $MyInvocation.MyCommand.Definition
Set-Location $currentPath

# Import config.json
$config = Get-Content ".\config.json" | ConvertFrom-Json

# Import source file
if($($config.formatting.source_file) -like "*.xlsx"){
    $listToWorkWith = Import-Excel ".\$($config.formatting.source_file)" -WorksheetName $($config.formatting.worksheet_name)
}
if($($config.formatting.source_file) -like "*.txt"){
    $listToWorkWith = Get-Content ".\$($config.formatting.source_file)"
}

if($listToWorkWith){
    $cats = $config.formatting.categories
    $fileName = $config.formatting.output_name
    $worksheetName = $config.formatting.output_worksheet
    
    $finalOutput = @()
    foreach($item in $listToWorkWith){
        $category = $null
        $location = $null
        $interested = $null
        $reason = $null
        $contactEmail = $null
        $attempted = $null
    
        foreach($cat in $cats){
            $stringFilters = $cat.stringFilter

            foreach($filter in $stringFilters){
                if ($item.ToLower().Contains($filter.ToLower())) {
                    $category = $cat.catID
                    $location = $cat.location
                    $interested = $cat.interested
                    $reason = $cat.reason
                    $contactEmail = $cat.contact
                    $attempted = $cat.attempted
                }
            }
        }
    
        $rowData = [PSCustomObject]@{
            Name = $item
            Type = $category
            Location = $location
            Interested = $interested
            Reason = $reason
            Contact = $contactEmail
            Attemtped = $attempted
        }
    
        $finalOutput += $rowData
    }
    
    if(Test-Path ".\$fileName"){
        Remove-Item ".\$fileName"
    }
    $finalOutput | Export-Excel ".\$fileName" -WorksheetName $worksheetName -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
    
}else{
    "Unable to work with source file"
}