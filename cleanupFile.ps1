# Get current location and set the working to directory to it
$currentPath = split-path -parent $MyInvocation.MyCommand.Definition
Set-Location $currentPath

# Import config.json
$config = Get-Content ".\config.json" | ConvertFrom-Json

# Import source file
if($($config.cleanup.source_file) -like "*.xlsx"){
    $import = Import-Excel ".\$($config.cleanup.source_file)" -WorksheetName $($config.cleanup.worksheet_name)
}
if($($config.cleanup.source_file) -like "*.txt"){
    $import = Get-Content ".\$($config.cleanup.source_file)"
}

# List of things to split on
$thingsToSplit = $config.cleanup.splitters
# List of things to replace, replace with in JSON config
$thingsToReplace = $config.cleanup.replacers
# Boolean if what to replace needs to be replaced strictly compared to case sensitivity
[bool]$caseSpecific = [System.Convert]::ToBoolean($config.cleanup.case_sensitive)

# List Variables for exporting/reporting
$originalList = $import.Name # Name = column to work with
$editedRows = @()
$newList = @()

# Create Empty Array Variables based on all things being split (for debugging/reporting)

foreach($splitWord in $thingsToSplit){

    # Firstly Remove all Variables if exists
    try {
        Remove-Variable -Name $("from" + $splitWord)        
    }
    catch {
        <#Do this if a terminating exception happens#>
    }
    # Create variables
    New-Variable -Name $("from" + $splitWord) -Value @()
}

# Loop through all items in the list
foreach($row in $originalList){
    $row # Output Row working with

    # replace each thing to replace based on how many inputted in $thingsToReplace
    $index = 0
    foreach($item in $thingsToReplace){
        if(!$caseSpecific){
            $row = $row -ireplace [regex]::Escape($($thingsToReplace[$index].toReplace)), $($thingsToReplace[$index].replaceWith) #.Replace("ry","PTY")
        }else{
            $row = $row.Replace($($thingsToReplace[$index].toReplace), $($thingsToReplace[$index].replaceWith))
        }
        $index++
    }

    foreach($item in $thingsToSplit){
        # split info on each delimiter but keeping the delimiter in place
        $split = $row -split "(?<=$item)"
            # if there is anything found after the split (incase of multiple info on one line)
            if($split.count -ge 2){
                # adds this row to edited object to see what rows were edited
                $editedRows += $row

                # adds all the info from the line to variable for specific delimiter and the new list individually
                foreach($index in $split){
                    if($index.length -gt 2){
                        (Get-Variable -Name $("from"+$item)).Value += $index
                        $newList += $index
                    }
                }

            }else{
                # adds to new list if nothing needs to be edited
                $newList += $row
            }
    }
}

# Trim each row to not have space before and after
$tempList = @()
foreach($messyString in $newList){
    
    $cleanString = [string]$messyString.Trim()
    $tempList += $cleanString
}
$tempList = $tempList | Select-Object -Unique

# loop through lists to ensure there are no duplicates
[System.Collections.ArrayList]$finalList = $tempList
foreach($thingo in $tempList){

    foreach($item in $thingsToSplit){

        # Check for and remove rows where delimiter is found and still contains info after (info after delmiter already on in its own row) 
        if($thingo.ToUpper().Contains("$item ".ToUpper())){
            $finalList.Remove($thingo)
        }

        # Check for and remove rows which have multiple delimiters in them
        $matched = Select-String -InputObject $thingo -Pattern ([regex]::Escape("$item")) -AllMatches
        if($matched.Matches.Count -gt 1){
            $finalList.Remove($thingo)
        }
    }

}

# Export final list to current directory
if(Test-Path ".\CleanedUpList.txt"){
    Remove-Item ".\CleanedUpList.txt"
}
$finalList | Out-File ".\CleanedUpList.txt"