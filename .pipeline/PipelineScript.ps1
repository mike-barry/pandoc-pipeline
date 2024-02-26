### Script parameters
param (
    [string]$WikiRepoRoot,
    [string]$DocRepoRoot,
    [string]$TempLocation,
    [string]$Directories
)


### Set script error action
$ErrorActionPreference = 'Stop'


### Define script constants
# Paths to files used by the script
$LAST_RUN_FILE_PATH = ".pipeline/last_run"
$CUSTOM_REFERENCE_FILE_PATH = ".pipeline/custom-reference.docx"

# Metadata tags
$PANDOC_GENERATE_TAG = "pandoc_generate"
#$PANDOC_INCLUDE_CHILDREN_TAG = "pandoc_include_children" # Not yet implemented

# Metadata line text:  [_metadata_:<TAG_NAME>]:- "<TAG_VALUE"
$METADATA_TEXT_START = "[_metadata_:"
$METADATA_TEXT_SPLIT = "]:-"

# Invalid filename characters
$INVALID_FILENAME_CHARS = [System.IO.Path]::GetInvalidFileNameChars() -join ''

### Main function
function Main
{
    param (
        [string]$WikiRepoRoot,
        [string]$DocRepoRoot,
        [string]$TempLocation,
        [string]$Directories
    )

    try
    {
        # Split the Directories parameter on two pipe characters and keep only the non-empty results
        $monitoredDirectories = ($Directories -split '\|\|') | Where-Object { $_ -match '\S' }

        # Grab the wiki repo's commit from when the script last ran or set to null if its never been run
        PrintHeader "Getting last commit..."
        $lastCommit = $null
        if (Test-Path "$DocRepoRoot/$Script:LAST_RUN_FILE_PATH" -PathType Leaf)
        {
            $lastCommit = $(Get-Content "$DocRepoRoot/$Script:LAST_RUN_FILE_PATH")
            PrintMessage "Last commit = $lastCommit"
        }

        # Get the actions that the script needs to perform
        if ($lastCommit)
        {
            PrintHeader "Getting actions from Git history..."
            $actions = GetActionsFromGitHistory $WikiRepoRoot $monitoredDirectories $lastCommit
        }
        else
        {
            PrintMessage "Getting actions from repo files..."
            $actions = GetActionsFromRepoFiles $WikiRepoRoot $DocRepoRoot $monitoredDirectories
        }

        # Delete any docx files whose corresponding wiki page was deleted
        PrintHeader "Performing delete actions..."
        PerformDeletes $DocRepoRoot $actions.Deletes

        # Move any docx files that had their corresponding page renamed/moved
        PrintHeader "Performing move actions..."
        PerformMoves $DocRepoRoot $actions.Moves

        # Create docx files for any new or modified wiki pages
        PrintHeader "Performing create actions..."
        PerformCreates $WikiRepoRoot $DocRepoRoot $TempLocation $actions.Creates

        # Update the '.last_run' file with the current commit so we can use it
        # during the next run as our starting point
        PrintHeader "Updating last commit..."
        SetLastCommit $WikiRepoRoot $DocRepoRoot

        # Commit and push the new docx files + .last_run
        PrintHeader "Committing and Pushing changes..."
        #CommitAndPush $DocRepoRoot
    }
    catch 
    {
        PrintError $_.Exception.Message
        exit 1
    }
}

function PrintLine([string]$Text, [int]$Level, [System.ConsoleColor]$ForegroundColor)
{
    $indent = ""

    for ($i = 0; $i -lt $level; $i++)
    {
        $indent += "  "
    }
    
    Write-Host "$indent$Text" -ForegroundColor $ForegroundColor
}

function PrintHeader([string]$Header, [int]$Level = 0)
{
    PrintLine $Header $Level DarkCyan
}

function PrintError([string]$Text, [int]$Level = 1)
{
    PrintLine $Text $Level Red
}

function PrintMessage([string]$Text, [int]$Level = 1)
{
    PrintLine $Text $Level White
}

function PrintWarning([string]$Text, [int]$Level = 1)
{
    PrintLine $Text $Level DarkYellow
}

function CheckExitCode([int]$ExitCode, [string]$ErrorMessage)
{
    if ($ExitCode -ne 0)
    {
        throw "$ErrorMessage ($ExitCode)"
    }
}

function GetActionsFromRepoFiles([string]$WikiRepoRoot, [string]$DocRepoRoot, [string[]]$MonitoredDirectories)
{
    $actions = @{
        Deletes = @()
        Moves = @()
        Creates = @()
    }

    # Loop through each monitored directory
    foreach ($directory in $MonitoredDirectories)
    {
        PrintHeader "Processing '$directory' directory" 1

        # Set the paths we'll be working with
        $wikiDirectory = "$WikiRepoRoot/$directory"
        $docDirectory = "$DocRepoRoot/$directory"
            
        # Add create actions for all the *.md files in the directory in the Wiki repo
        $mdFiles = Get-ChildItem -LiteralPath "$wikiDirectory" -Include "*.md" -Recurse

        foreach ($mdFile in $mdFiles)
        {
            $mdFilePath = (Resolve-Path -LiteralPath $mdFile.FullName -RelativeBasePath $WikiRepoRoot -Relative).Substring(2)
            $actions.Creates += $mdFilePath
            PrintMessage "Create '$mdFilePath'" 2
        }
        
        # Add delete actions for all the *.docx files in the directory in the Doc repo
        $docxFiles = Get-ChildItem -LiteralPath "$docDirectory" -Include "*.docx" -Recurse

        foreach ($docxFile in $docxFiles)
        {
            $docxFilePath = (Resolve-Path -LiteralPath $docxFile.FullName -RelativeBasePath $DocRepoRoot -Relative).Substring(2)
            $actions.Deletes += $docxFilePath
            PrintMessage "Delete '$docxFilePath'" 2
        }
    }

    return $actions
}

function GetActionsFromGitHistory([string]$WikiRepoRoot, [string[]]$MonitoredDirectories, [string]$LastCommit)
{
    # TODO:  When implementing the PANDOC_INCLUDE_CHILDREN tag, need to also look for changes to .order files

    $actions = @{
        Deletes = @()
        Moves = @()
        Creates = @()
    }

    # Loop through each monitored directory
    foreach ($directory in $MonitoredDirectories)
    {
        PrintHeader "Processing '$directory' directory" 1

        # Set the paths we'll be working with
        $wikiDirectory = "$WikiRepoRoot/$directory"

        # Do not proceed if the directory does not exist in the Wiki repo
        if (-not (Test-Path -LiteralPath "$wikiDirectory" -PathType Container))
        {
            PrintWarning "Directory does not exist" 2
            continue
        }

        # Get the Git diff history since the last commit
        # https://git-scm.com/docs/git-diff
        $diffOutput = $(git -C "$WikiRepoRoot" diff --diff-filter=ACDMR --name-status $LastCommit $directory*.md)
        CheckExitCode $LASTEXITCODE "Git diff command failed"

        $lines = ($diffOutput -split '\r?\n')

        if (-not $lines)
        {
            PrintMessage "No changes found" 2
        }

        # Loop through each line of the 'git diff' output
        foreach ($line in $lines)
        {
            #PrintMessage $line 2

            # Split each line on the tab character to get the fields
            $fields = $line -split '\t'
    
            # Check for invalid field count
            if (($fields.Length -lt 2) -or ($fields.Length -gt 3))
            {
                throw "Invalid number of fields in 'git diff' output ($line)"
            }
            else
            {
                # Snag the code from the first character of the line
                $code = $fields[0].Substring(0, 1);
    
                # Evaluate the code to determine the corresponding action we need to do
                switch ($code)
                {
                    "A" # Added
                    {
                        $actions.Creates += $fields[1]
                        PrintMessage "Create '$($fields[1])'" 2
                    }
    
                    "C" # Copied - Probably won't ever see this one
                    {
                        if ($fields.Length -ne 3)
                        {
                            throw "Invalid number of fields for a copy in 'git diff' output ($line)"
                        }
    
                        $actions.Creates += $fields[2]
                        PrintMessage "Create '$($fields[2])'" 2
                    }
                    
                    "D" # Deleted
                    {
                        $actions.Deletes += $fields[1]
                        PrintMessage "Delete '$($fields[1])'" 2
                    }
    
                    "M" # Modified
                    {
                        $actions.Creates += $fields[1]
                        PrintMessage "Create '$($fields[1])'" 2
                    }
    
                    "R" # Renamed
                    {
                        if ($fields.Length -ne 3)
                        {
                            throw "Invalid number of fields for a rename in 'git diff' output ($line)"
                        }
                        
                        $actions.Moves += @{
                            Previous = $fields[1]
                            New      = $fields[2]
                        }
                        PrintMessage "Move '$($fields[1])' to '$($fields[2])'" 2
                    }
    
                    default # Shouldn't ever get here...
                    {
                        throw "Unknow code in 'git diff' output ($line)"
                    }
                }
            }
        }
    }

    return $actions
}

function PerformDeletes([string]$DocRepoRoot, [string[]]$Deletes)
{
    foreach ($item in $Deletes)
    {
        # Get the corresponding docx file path
        $docxPath = GetDocxFilename $item

        PrintMessage "Removing '$docxPath'"

        # Delete the file or print a warning if it does not exist
        if (Test-Path "$DocRepoRoot/$docxPath" -PathType Leaf)
        {
            Remove-Item "$DocRepoRoot/$docxPath" -Force
        }
        else
        {
            PrintWarning "File does not exist" 2
        }
    }
}

function UnmanglePath([string]$path)
{
    # Convert hyphens back to spaces
    $newPath = $path.Replace("-", " ")

    # Convert URL-encoded values back to their original values
    $newPath = [uri]::UnescapeDataString($newPath)

    # Remove any invalid filename characters
    $newPath = $newPath.Split($Script:INVALID_FILENAME_CHARS) -join ''

    return $newPath
}

function CreateParentDirectory([string]$filePath)
{
    # Get the path to the parent directory
    $directory = Split-Path -Path "$filePath" -Parent

    # Create the directory if it does not exist
    if (-not (Test-Path -Path $directory -PathType Container))
    {
        $null = New-Item -ItemType Directory -Path $directory -Force
    }
}

function PerformMoves([string]$DocRepoRoot, [object[]]$Moves)
{
    foreach ($item in $Moves)
    {
        # Get the corresponding docx file paths
        $previousPath = (UnmanglePath $item.Previous).Replace(".md", ".docx")
        $newPath = (UnmanglePath $item.New).Replace(".md", ".docx")

        PrintMessage "Moving '$previousPath' to '$newPath'"

        # Move the file or print a warning if it does not exist
        if (Test-Path "$DocRepoRoot/$previousPath" -PathType Leaf)
        {
            CreateParentDirectory "$DocRepoRoot/$newPath"
            Move-Item -Path "$DocRepoRoot/$previousPath" -Destination "$DocRepoRoot/$newPath" -Force
        }
        else
        {
            PrintWarning "File does not exist" 2
        }
    }
}

function GetMetadataTags([string]$filePath)
{
    # Create a hashtable to store the tag key/value pairs
    $tags = @{}

    # Read the lines of text from the file
    $lines = Get-Content "$WikiRepoRoot/$mdPath"

    # Loop through each line to look for metadata
    foreach ($line in $lines)
    {
        # Remove trailing and leading whitespace
        $trimmedLine = $line.Trim()

        # Look for metadata tags
        if ($trimmedLine -and ($trimmedLine.StartsWith($Script:METADATA_TEXT_START)))
        {
            # Get the tag's key and value
            $tagNameAndValue = $trimmedLine.Substring($Script:METADATA_TEXT_START.Length)
            $tagParts = $tagNameAndValue -split $Script:METADATA_TEXT_SPLIT | Where-Object { $_ -match '\S' }

            # Store the tag key and value (but make sure the key isn't blank)
            if ($tagParts.Count -eq 2)
            {
                # Trim the key string to remove any whitespace
                $key = $tagParts[0].Trim()

                # Make sure we've got a non-whitespace key
                if ($key)
                {
                    # Strip off whitespace and double quotes from the value
                    $value = $tagParts[1].Trim().Trim('"')
                    
                    # Convert true/false strings to Boolean values
                    if ($value.ToLower() -eq "true")
                    {
                        $value = $true
                    }
                    elseif ($value.ToLower() -eq "false")
                    {
                        $value = $false
                    }

                    # Store the tag and value
                    $tags[$key] = $value
                }
            }
            else
            {
                PrintWarning "Invalid metadata field found ($line)" 2
            }
        }
    }

    return $tags
}

function FixMarkdownIssues([string]$inputFilePath, [string]$outputFilePath)
{
    # Read the lines of text from the file
    $lines = Get-Content -Path "$inputFilePath"

    # Create a blank output file
    $null = New-Item -Path $outputFilePath -ItemType File -Force

    # Loop through each line of the input file
    foreach ($line in $lines)
    {
        # Remove the first slash on image links and suppress the text
        if ($line -like "*](/.attachments/*")
        {
            $line = $line -replace "\[.*\]\(/\.attachments/", "[](.attachments/"
        }

        # Write the line to the output file
        Add-Content -Path $outputFilePath -Value $line
    }
}

function PerformCreates([string]$WikiRepoRoot, [string]$DocRepoRoot, [string]$TempLocation, [string[]]$Creates)
{
    foreach ($mdPath in $Creates)
    {
        PrintMessage "Processing '$mdPath'"

        # Grab the metadata tags from the markdown file
        $tags = GetMetadataTags $mdPath

        if (-not ($tags[$Script:PANDOC_GENERATE_TAG] -eq $true))
        {
            PrintMessage "Skipped" 2
            continue
        }

        # Get the full paths that will be used
        $mdFullPath = $WikiRepoRoot + "/" + $mdPath
        $fixedFullPath = $TempLocation + "/fixed.md"
        $docxFullPath = $DocRepoRoot + "/" + (UnmanglePath $mdPath).Replace(".md", ".docx")
        $customReferenceFullPath = $DocRepoRoot + "/" + $Script:CUSTOM_REFERENCE_FILE_PATH

        # Make sure the parent directory exists
        CreateParentDirectory "$docxFullPath"

        #TODO:  Implement PANDOC_INCLUDE_CHILDREN tag

        # Fix markdown/pandoc issues
        FixMarkdownIssues $mdFullPath $fixedFullPath

        # Run pandoc
        #Set-Location "$WikiRepoRoot"
        pandoc "$fixedFullPath" -t docx -s --reference-doc="$customReferenceFullPath" --resource-path="$WikiRepoRoot" -o "$docxFullPath"
        CheckExitCode $LASTEXITCODE "pandoc failure"

        PrintMessage "Generated" 2
    }
}

function SetLastCommit([string]$WikiRepoRoot, [string]$DocRepoRoot)
{
    # Get the current commit ID of the Wiki repo
    $commitId = $(git -C $WikiRepoRoot rev-parse --verify HEAD)
    CheckExitCode $LASTEXITCODE "Git rev-parse command failed"

    # Write the commit ID to the 'last_run' file in the Doc repo
    Set-Content -Path "$DocRepoRoot/$Script:LAST_RUN_FILE_PATH" -Value "$commitId" -NoNewLine
}

# Call the Main function with the script parameters
Main @PSBoundParameters