# listCleanup #
Cleanup Lists of data

Edit config.json only

config.json is one config file containing objects that are used in each .ps1 script

## Script File "cleanupFile.ps1" ##
Takes a messy list and cleans it up by splitting rows by splitters (in config.json) and replaces words by replacers (in config.json)
Import file name = sourceFile in config.json
Output file name = "CleanedUpList.txt"

## Script File "formatNewFile.ps1" ##
Takes the newly created CleanedUpList.txt file and formats it to table form xlsx for easy use
Import file name = sourceFile in config.json
Output file name = output_name in config.json
