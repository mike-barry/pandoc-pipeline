$repoRoot = "C:/Repos/pandoc.wiki"
$directories = @("Procedures")
$lastCommit =  $(get-content last_successful_run.txt)

$delete = @()
$create = @()

foreach ($directory in $directories)
{
    # https://git-scm.com/docs/git-diff
    $diffOutput = $(git -C $repoRoot diff --diff-filter=ACDMR --name-status $lastCommit $directory*.md)

    $diffs = $diffOutput -split '\r?\n'

    foreach ($line in $diffs)
    {
        $fields = $line -split '\t'

        if (($fields.Length -lt 2) -or ($fields.Length -gt 3))
        {
            Write-Error "Invalid diff output:  $line"
            continue
        }
        else
        {
            $code = $fields[0].Substring(0, 1);

            switch ($code)
            {
                "A" # Added
                {
                    $create += $fields[1]
                }

                "C" # Copied
                {
                    if ($fields.Length -ne 3)
                    {
                        Write-Error "Invalid diff output:  $line"
                        continue
                    }

                    $create += $fields[2]
                }
                
                "D" # Deleted
                {
                    $delete += $fields[1]
                }

                "M" # Modified
                {
                    $create += $fields[1]
                }

                "R" # Renamed
                {
                    if ($fields.Length -ne 3)
                    {
                        Write-Error "Invalid diff output:  $line"
                        continue
                    }

                    $delete += $fields[1]
                    $create += $fields[2]
                }

                default # Shouldn't ever get here...
                {
                    Write-Error "Unknown diff code:  $line"
                }
            }
        }
    }
}

# We do the deletes before creates in case a file was deleted and then another file
# created/renamed/copied with the same name.
foreach ($item in $delete)
{
    Write-Host "Removing DOCX for $item"
}

foreach ($item in $create)
{
    #TODO:
    # Open each file and look for a line containing:  [_metadata_:pandoc_generate]:- "true"
    #    If it is "true" then generate the DOCX.  Otherwise, ignore it.
    # Also check for [_metadata_:pandoc_include_children]:- "true"
    #    If it is "true" then incorporate child pages into the resulting DOCX (TBD if necessary).

    Write-Host "Generating DOCX for $path"
}