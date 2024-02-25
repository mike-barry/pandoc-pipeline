$ErrorActionPreference = 'Stop'

# Main function
function Main()
{
    # Parameters
    $wikiRepoRoot = "C:/Repos/pandoc.wiki"
    $docRepoRoot = "C:/Repos/pandoc-test"
    $monitoredDirectories = @("Procedures")


    try {
        # Grab the commit from when the script last ran
        $lastCommit = $(get-content "$docRepoRoot/.pipeline/last_run")

        # Populate $actions with what needs to be done to apply the wiki repo history to the docs repo
        $actions = GetActionsFromGitHistory $wikiRepoRoot $monitoredDirectories $lastCommit

        # Delete any docx files whose corresponding wiki page was deleted
        PerformDeletes $actions.Deletes

        # Rename/move any docx files that had their corresponding page renamed/moved
        PerformRenames $actions.Renames

        # Create docx files for any new or modified wiki pages
        PerformCreates $actions.Creates

        # Update the '.last_run' file with the current commit so we can use it
        # during the next run as our starting point
        SetLastCommit $wikiRepoRoot $docRepoRoot

        # Commit and push the new docx files + .last_run
        #CommitAndPush $docRepoRoot
    }
    catch {
        <#Do this if a terminating exception happens#>
    }
}

function CheckExitCode($exitCode, $errorMessage)
{
    if ($exitCode -ne 0)
    {
        throw "$errorMessage ($exitCode)"
    }
}

function GetActionsFromGitHistory($wikiRepoRoot, $monitoredDirectories, $lastCommit)
{
    $actions = @{
        Deletes = @()
        Renames = @()
        Creates = @()
    }

    # Loop through each directory we are configured to monitor
    foreach ($directory in $monitoredDirectories)
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
            
            #TODO:  Delete existing docx files and folders?

            continue
        }

        # Get the Git diff history since the last commit
        # https://git-scm.com/docs/git-diff
        $diffOutput = $(git -C $wikiRepoRoot diff --diff-filter=ACDMR --name-status $lastCommit $directory*.md)
        CheckExitCode $LASTEXITCODE "Git diff command failed"

        # Loop through each line of the 'git diff' output
        foreach ($line in ($diffOutput -split '\r?\n'))
        {
            PrintMessage $line

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
                    }
    
                    "C" # Copied
                    {
                        if ($fields.Length -ne 3)
                        {
                            throw "Invalid number of fields for a copy in 'git diff' output ($line)"
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
                            throw "Invalid number of fields for a rename in 'git diff' output ($line)"
                        }
                        
                        $actions.Renames += @{
                            Previous = $fields[1]
                            New = $fields[2]
                        }
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


# Call the Main function
Main