$ErrorActionPreference = 'Stop'

# Main function
function Main()
{
    # Parameters
    $wikiRepoRoot = "C:/Repos/pandoc.wiki"
    $docRepoRoot = "C:/Repos/pandoc.docs"
    $directories = @("Procedures")

    # Grab the commit from when the script last ran
    $lastCommit = $(get-content "$docRepoRoot/.last_run")

    # Populate $actions with what needs to be done to apply the wiki repo history to the docs repo
    $actions = GetActions $wikiRepoRoot $directories $lastCommit

    # Delete any docx files whose corresponding wiki page was deleted
    PerformDeletes $actions.Deletes

    # Rename/move any docx files that had their corresponding page renamed/moved
    PerformRenames $actions.Renames

    # Create docx files for any new or modified wiki pages
    PerformCreates $actions.Creates

    # Update the '.last_run' file with the current commit so we can use it
    # during the next run as our starting point
    #SetLastCommit $docRepoRoot

    # Commit and push the new docx files + .last_run
    #CommitAndPush $docRepoRoot
}

# Populates $deletes and $creates with file paths based on the git history since the last time
# the script ran successfully
function GetActions($wikiRepoRoot, $directories, $lastCommit)
{
    $actions = [ScriptActionsStruct]::new()

    # Loop through each directory we are configured to monitor
    foreach ($directory in $directories)
    {
        PrintHeader "Processing '$directory'"

        # If this is our first time running, consider all *.md files as new
        if (-not $lastCommit)
        {
            $mdFiles = Get-ChildItem -LiteralPath $wikiRepoRoot -Include "*.md" -Recurse

            foreach ($mdFile in $mdFiles)
            {
                $actions.Creates += $mdFile.FullName
            }
            
            #TODO:  Delete existing docx files?

            continue
        }

        # https://git-scm.com/docs/git-diff
        $diffOutput = $(git -C $wikiRepoRoot diff --diff-filter=ACDMR --name-status $lastCommit $directory*.md)
    
        $diffs = $diffOutput -split '\r?\n'
    
        foreach ($line in $diffs)
        {
            PrintMessage $line

            $fields = $line -split '\t'
    
            if (($fields.Length -lt 2) -or ($fields.Length -gt 3))
            {
                PrintError "Invalid diff output:  $line"
                continue
            }
            else
            {
                $code = $fields[0].Substring(0, 1);
    
                switch ($code)
                {
                    "A" # Added
                    {
                        $actions.Creates += $fields[1]
                    }
    
                    "C" # Copied
                    {
                        if ($fields.Length -ne 3)
                        {
                            PrintError "Invalid diff output:  $line"
                            continue
                        }
    
                        $actions.Creates += $fields[2]
                    }
                    
                    "D" # Deleted
                    {
                        $actions.Deletes += $fields[1]
                    }
    
                    "M" # Modified
                    {
                        $actions.Creates += $fields[1]
                    }
    
                    "R" # Renamed
                    {
                        if ($fields.Length -ne 3)
                        {
                            PrintError "Invalid diff output:  $line"
                            continue
                        }
                        
                        $actions.Renames += [RenameStruct]::new($fields[1], $fields[2])
    
                        #$actions.Deletes += $fields[1]
                        #$actions.Creates += $fields[2]
                    }
    
                    default # Shouldn't ever get here...
                    {
                        PrintError "Unknown diff code:  $line"
                    }
                }
            }
        }
    }

    return $actions
}

function SetLastCommit()
{
    throw "Not implemented"
}

function PerformDeletes()
{
    foreach ($item in $deletes)
    {
        Write-Host "Removing DOCX for $item"
    }
}

function PerformRenames()
{

}

function PerformCreates()
{
    foreach ($item in $creates)
    {
        #TODO:
        # Open each file and look for a line containing:  [_metadata_:pandoc_generate]:- "true"
        #    If it is "true" then generate the DOCX.  Otherwise, ignore it.
        # Also check for [_metadata_:pandoc_include_children]:- "true"
        #    If it is "true" then incorporate child pages into the resulting DOCX (TBD if necessary).

        Write-Host "Generating DOCX for $path"
    }
}

function PrintError($message)
{
    Write-Host "  $message" -ForegroundColor DarkRed
}

function PrintHeader($header)
{
    Write-Host $header -ForegroundColor DarkCyan
}

function PrintMessage($message)
{
    Write-Host "  $message"
}


class RenameStruct
{
    [string]$Previous
    [string]$New

    RenameStruct($Previous, $New){
        $this.Previous = $Previous
        $this.New = $New
    }
}

class ScriptActionsStruct
{
    [RenameStruct[]]$Renames
    [string[]]$Deletes
    [string[]]$Creates
}


# Call the Main function
Main