<#
.SYNOPSIS
    Publishes exactly one Excel Notes module per calendar day.

.DESCRIPTION
    What this script does, in order:
      1. Verifies the repository, staging library, and main branch.
      2. Sets the local Git identity to the GitHub noreply address.
      3. Fetches origin and pulls with rebase + autostash.
      4. Stops cleanly if a "docs: add Module" commit already exists today.
      5. Picks the next staging folder (01..35) that Git does not track yet.
      6. Validates the module: required files, links, images, workbook,
         local paths, encoding, and the standard lesson structure.
      7. Copies the module into the repository.
      8. Updates the README roadmap status and progress block.
      9. Commits with: docs: add Module NN - Topic Title
     10. Pushes to origin/main. If the push fails, the local commit is rolled
         back so the next run can retry cleanly.
     11. Writes a timestamped log to the workspace.

.PARAMETER DryRun
    Runs every check and shows what would be published, but does not copy,
    commit, or push anything.

.EXAMPLE
    .\scripts\post_daily.ps1 -DryRun

.EXAMPLE
    .\scripts\post_daily.ps1
#>
[CmdletBinding()]
param(
    [switch]$DryRun
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ======================= CONFIGURATION =======================
$RepoPath      = Split-Path -Parent $PSScriptRoot
$WorkspaceRoot = Split-Path -Parent $RepoPath
$LogFile       = Join-Path $WorkspaceRoot 'excel-notes-daily.log'
$Branch        = 'main'
$Remote        = 'origin'
$TotalModules  = 35

# GitHub noreply address: keeps the public commit email private.
$GitHubEmail = '230164411+AaryaAI-cyber@users.noreply.github.com'
$GitHubName  = 'AaryaAI-cyber'

# The staging library sits next to the repository. Override with an
# environment variable if you move it.
if ($env:EXCEL_NOTES_STAGING) {
    $StagingPath = $env:EXCEL_NOTES_STAGING
}
else {
    $StagingPath = Join-Path $WorkspaceRoot 'excel-staging\Excel-Notes-for-Beginners-Professionals-'
}

$TitleMap = @{
    '01' = 'Excel Basics and Interface'
    '02' = 'Data Entry and Number Formats'
    '03' = 'Formulas and Cell References'
    '04' = 'Logical Functions'
    '05' = 'Lookup Functions'
    '06' = 'Text Functions'
    '07' = 'Date and Time Functions'
    '08' = 'Math and Statistical Functions'
    '09' = 'Dynamic Arrays'
    '10' = 'LET, LAMBDA and Array Shaping'
    '11' = 'Error Handling'
    '12' = 'Data Cleaning and Validation'
    '13' = 'Sorting, Filtering and Tables'
    '14' = 'Conditional Formatting'
    '15' = 'Pivot Tables, Slicers and Timelines'
    '16' = 'Charts and Graphs'
    '17' = 'Dashboard Design'
    '18' = 'Power Query'
    '19' = 'Power Pivot, Data Model and DAX'
    '20' = 'What-If Analysis Tools'
    '21' = 'Statistics in Excel'
    '22' = 'Financial Functions and Models'
    '23' = 'Business Analytics Cases'
    '24' = 'Research and Survey Analysis'
    '25' = 'Excel for Software Engineers and Testers'
    '26' = 'Macros and VBA'
    '27' = 'Python in Excel and Office Scripts'
    '28' = 'Collaboration, Sharing, Protection and Security'
    '29' = 'Performance and Auditing'
    '30' = 'AI in Excel'
    '31' = 'Excel for Managers and CEOs'
    '32' = 'Excel vs Sheets, Power BI, SQL and Python'
    '33' = 'Capstone Projects'
    '34' = 'Interview Preparation'
    '35' = 'Cheat Sheet and Shortcuts'
}
# =============================================================

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    try {
        $logDir = Split-Path -Parent $LogFile
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::AppendAllText($LogFile, $line + [Environment]::NewLine, $utf8)
    }
    catch {
        # Logging must never hide the real work.
    }
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        default { Write-Host $line }
    }
}

function Fail {
    param([string]$Message, [int]$Code = 1)
    Write-Log $Message 'ERROR'
    exit $Code
}

function Invoke-Git {
    param([Parameter(Mandatory = $true)][string[]]$GitArgs)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & git -C $RepoPath @GitArgs 2>&1
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previous
    }
    $text = (($output | ForEach-Object { "$_" }) -join "`n").Trim()
    return [pscustomobject]@{ Code = $code; Output = $text }
}

function Assert-Git {
    param(
        [string[]]$GitArgs,
        [string]$Action
    )
    $result = Invoke-Git -GitArgs $GitArgs
    if ($result.Code -ne 0) {
        Fail ("{0} failed (exit {1}). Output: {2}" -f $Action, $result.Code, $result.Output)
    }
    return $result.Output
}

function Get-NextModule {
    $folders = Get-ChildItem -LiteralPath $StagingPath -Directory |
        Where-Object { $_.Name -match '^\d{2}_' } |
        Sort-Object Name

    foreach ($folder in $folders) {
        $tracked = Invoke-Git -GitArgs @('ls-files', '--', $folder.Name)
        if ($tracked.Code -ne 0) {
            Fail "git ls-files failed: $($tracked.Output)"
        }
        if ([string]::IsNullOrWhiteSpace($tracked.Output)) {
            return $folder
        }
    }
    return $null
}

function Test-Lesson {
    param([System.IO.DirectoryInfo]$Dir)

    $problems = New-Object System.Collections.Generic.List[string]
    $readme    = Join-Path $Dir.FullName 'README.md'
    $exercises = Join-Path $Dir.FullName 'exercises_and_solutions.md'
    $workbook  = Join-Path $Dir.FullName 'practice_workbook.xlsx'
    $images    = Join-Path $Dir.FullName 'images'

    foreach ($required in @($readme, $exercises, $workbook)) {
        if (-not (Test-Path -LiteralPath $required)) {
            $problems.Add("Missing file: $(Split-Path $required -Leaf)")
        }
    }
    if (-not (Test-Path -LiteralPath $images)) {
        $problems.Add('Missing images folder')
    }

    $localPathPattern = '(?i)([a-z]:\\|/Users/|/home/|file:///|C:/Users)'
    $mojibakePattern  = '(?i)(\u00C3.|\u00C2.|\u00E2\u20AC|dY[^\s])'
    $linkPattern      = '(?<!!)\[[^\]]*\]\(([^)\s#]+)'
    $imagePattern     = '!\[[^\]]*\]\(([^)\s]+)'

    foreach ($markdown in @($readme, $exercises)) {
        if (-not (Test-Path -LiteralPath $markdown)) { continue }
        $text = [System.IO.File]::ReadAllText($markdown, [System.Text.Encoding]::UTF8)

        if ($text -match $localPathPattern) {
            $problems.Add("Local or sandbox path found in $(Split-Path $markdown -Leaf)")
        }
        if ($text -match $mojibakePattern) {
            $problems.Add("Possible encoding damage in $(Split-Path $markdown -Leaf)")
        }

        foreach ($match in [regex]::Matches($text, $linkPattern)) {
            $link = $match.Groups[1].Value
            if ($link -match '^(https?:|mailto:|data:)') { continue }
            $target = Join-Path $Dir.FullName (($link -replace '/', '\'))
            if (-not (Test-Path -LiteralPath $target)) {
                $problems.Add("Broken link in $(Split-Path $markdown -Leaf): $link")
            }
        }

        foreach ($match in [regex]::Matches($text, $imagePattern)) {
            $link = $match.Groups[1].Value
            if ($link -match '^(https?:|data:)') { continue }
            $target = Join-Path $Dir.FullName (($link -replace '/', '\'))
            if (-not (Test-Path -LiteralPath $target)) {
                $problems.Add("Broken image in $(Split-Path $markdown -Leaf): $link")
            }
        }
    }

    if (Test-Path -LiteralPath $readme) {
        $readmeText = [System.IO.File]::ReadAllText($readme, [System.Text.Encoding]::UTF8)
        $requiredSections = @(
            'Objective & Business Use Case',
            'Core Concepts, Syntax & Jargon Buster',
            'Step-by-Step Practical Walkthrough',
            'Real-World Scenarios & Worked Examples',
            'Common Pitfalls & Performance Tips',
            'Interview Questions with Solutions',
            'Practice Exercises with Solutions'
        )
        foreach ($section in $requiredSections) {
            if ($readmeText -notmatch [regex]::Escape($section)) {
                $problems.Add("Missing lesson section: $section")
            }
        }
    }

    if (Test-Path -LiteralPath $exercises) {
        $exerciseText = [System.IO.File]::ReadAllText($exercises, [System.Text.Encoding]::UTF8)
        if ($exerciseText -notmatch '<summary>Show solution</summary>') {
            $problems.Add('Exercises file has no visible solutions')
        }
    }

    if (Test-Path -LiteralPath $workbook) {
        if ((Get-Item -LiteralPath $workbook).Length -gt 2MB) {
            $problems.Add('Workbook is larger than 2 MB')
        }
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [System.IO.Compression.ZipFile]::OpenRead($workbook)
            try {
                $names = @($zip.Entries | ForEach-Object { $_.FullName })
                if ($names -notcontains 'xl/workbook.xml') {
                    $problems.Add('Workbook is not a valid .xlsx file')
                }
            }
            finally {
                $zip.Dispose()
            }
        }
        catch {
            $problems.Add("Workbook could not be opened: $($_.Exception.Message)")
        }
    }

    if (Test-Path -LiteralPath $images) {
        foreach ($image in Get-ChildItem -LiteralPath $images -Recurse -File) {
            if ($image.Extension -ieq '.png') {
                $bytes = [System.IO.File]::ReadAllBytes($image.FullName)
                if ($bytes.Length -lt 8 -or $bytes[0] -ne 0x89 -or $bytes[1] -ne 0x50 -or $bytes[2] -ne 0x4E -or $bytes[3] -ne 0x47) {
                    $problems.Add("Invalid PNG image: $($image.Name)")
                }
            }
        }
    }

    return $problems
}

function Get-UpdatedReadmeText {
    param([string]$ModuleName)

    $number = $ModuleName.Substring(0, 2)
    $readmePath = Join-Path $RepoPath 'README.md'
    if (-not (Test-Path -LiteralPath $readmePath)) {
        Fail 'README.md is missing from the repository.'
    }

    $text = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)
    $text = $text.Replace("`r`n", "`n")

    $rowPattern = '(?m)^\|\s*' + [regex]::Escape($number) + '\s*\|.*$'
    $rowMatch = [regex]::Match($text, $rowPattern)
    if (-not $rowMatch.Success) {
        Fail "README.md has no roadmap row for module $number."
    }

    $row = $rowMatch.Value
    $folderPattern = '^(\|\s*' + [regex]::Escape($number) + '\s*\|)\s*(?:`[^`]+`|\[[^\]]+\]\([^)]+\))\s*\|'
    $replacement = '$1 [' + $ModuleName + '](' + $ModuleName + '/README.md) |'
    $newRow = [regex]::Replace($row, $folderPattern, $replacement)
    $publishedLabel = [char]0x2705
    $newRow = [regex]::Replace($newRow, '\|\s*\u23F3 Scheduled\s*\|\s*$', ('| ' + $publishedLabel + ' Published |'))

    if ($newRow -notmatch '\u2705 Published') {
        Fail "Could not mark module $number as published in README.md."
    }

    $text = $text.Substring(0, $rowMatch.Index) + $newRow + $text.Substring($rowMatch.Index + $rowMatch.Length)

    $published = ([regex]::Matches($text, '(?m)^\|\s*\d{2}\s*\|.*\|\s*\u2705 Published\s*\|\s*$')).Count
    $nextMatch = [regex]::Match($text, '(?m)^\|\s*(\d{2})\s*\|.*\|\s*\u23F3 Scheduled\s*\|\s*$')
    if ($nextMatch.Success) {
        $nextNumber = $nextMatch.Groups[1].Value
        $nextTitle = $TitleMap[$nextNumber]
        $nextLine = "Module $nextNumber - $nextTitle - 21:30 local time"
    }
    else {
        $nextLine = 'Curriculum complete - all 35 modules are live.'
    }

    $filled = [int][math]::Round(($published / $TotalModules) * 20)
    if ($filled -lt 0) { $filled = 0 }
    if ($filled -gt 20) { $filled = 20 }
    $fullBlock  = [char]0x2588
    $emptyBlock = [char]0x2591
    $bar = ([string]$fullBlock * $filled) + ([string]$emptyBlock * (20 - $filled))
    $percent = [int][math]::Round(($published / $TotalModules) * 100)

    $block = @"
<!-- PROGRESS:START -->
**Publishing progress:** ``$bar`` **$published / $TotalModules modules** ($percent%)

**Next automated publish:** $nextLine
<!-- PROGRESS:END -->
"@
    $text = [regex]::Replace($text, '(?s)<!-- PROGRESS:START -->.*?<!-- PROGRESS:END -->', $block)
    return $text
}

function Invoke-Publish {
    Write-Log "Repository: $RepoPath"
    Write-Log "Staging:    $StagingPath"

    if (-not (Test-Path -LiteralPath $RepoPath)) {
        Fail "Repository not found: $RepoPath"
    }
    if (-not (Test-Path -LiteralPath $StagingPath)) {
        Fail "Staging library not found: $StagingPath"
    }

    $currentBranch = (Invoke-Git -GitArgs @('rev-parse', '--abbrev-ref', 'HEAD')).Output
    if ($currentBranch -ne $Branch) {
        Fail "Repository is on '$currentBranch'; expected '$Branch'."
    }

    $null = Assert-Git -GitArgs @('config', 'user.email', $GitHubEmail) -Action 'git config user.email'
    $null = Assert-Git -GitArgs @('config', 'user.name', $GitHubName) -Action 'git config user.name'
    $identity = (Invoke-Git -GitArgs @('config', 'user.email')).Output
    Write-Log "Commit identity: $GitHubName <$identity>"

    Write-Log 'Fetching origin...'
    $null = Assert-Git -GitArgs @('fetch', $Remote) -Action 'git fetch origin'

    Write-Log 'Pulling with rebase and autostash...'
    $null = Assert-Git -GitArgs @('pull', '--rebase', '--autostash', $Remote, $Branch) -Action 'git pull --rebase --autostash'

    $today = Get-Date -Format 'yyyy-MM-dd'
    $lastDate = (Invoke-Git -GitArgs @('log', '-1', '--format=%cs', '--grep=^docs: add Module')).Output
    if ($lastDate -eq $today) {
        Write-Log "Already published a module today ($today). Nothing to do." 'OK'
        return
    }

    $next = Get-NextModule
    if ($null -eq $next) {
        Write-Log 'All 35 modules are already published. Nothing to do.' 'OK'
        return
    }

    $moduleName = $next.Name
    $number = $moduleName.Substring(0, 2)
    $title = $TitleMap[$number]
    if (-not $title) {
        Fail "No title configured for module $number."
    }
    Write-Log "Next module: $moduleName - $title"

    $problems = @(Test-Lesson -Dir $next)
    if ($problems.Count -gt 0) {
        foreach ($problem in $problems) {
            Write-Log "Validation: $problem" 'ERROR'
        }
        Fail "Module $moduleName failed validation. Nothing was published."
    }
    Write-Log 'Validation passed: files, links, images, workbook, paths and lesson structure.' 'OK'

    $commitMessage = "docs: add Module $number - $title"
    if ($DryRun) {
        $null = Get-UpdatedReadmeText -ModuleName $moduleName
        Write-Log "DRY RUN OK - would publish: $commitMessage" 'OK'
        return
    }

    $destination = Join-Path $RepoPath $moduleName
    if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
        Write-Log "Removed stale local copy of $moduleName."
    }
    Copy-Item -LiteralPath $next.FullName -Destination $destination -Recurse -Force
    Write-Log "Copied $moduleName into the repository."

    $updatedReadme = Get-UpdatedReadmeText -ModuleName $moduleName
    $readmePath = Join-Path $RepoPath 'README.md'
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($readmePath, $updatedReadme, $utf8)
    Write-Log 'Updated README roadmap status and progress.'

    $null = Assert-Git -GitArgs @('add', '--', $moduleName, 'README.md') -Action 'git add'
    $staged = (Invoke-Git -GitArgs @('diff', '--cached', '--name-only')).Output
    if ([string]::IsNullOrWhiteSpace($staged)) {
        Fail 'No changes were staged. Nothing to commit.'
    }

    $null = Assert-Git -GitArgs @('commit', '-m', $commitMessage) -Action 'git commit'
    Write-Log "Committed: $commitMessage"

    $push = Invoke-Git -GitArgs @('push', $Remote, $Branch)
    if ($push.Code -ne 0) {
        Write-Log "Push failed: $($push.Output)" 'WARN'
        Write-Log 'Retrying after one rebase...' 'WARN'
        $retryPull = Invoke-Git -GitArgs @('pull', '--rebase', '--autostash', $Remote, $Branch)
        if ($retryPull.Code -eq 0) {
            $push = Invoke-Git -GitArgs @('push', $Remote, $Branch)
        }
        else {
            Write-Log "Retry pull failed: $($retryPull.Output)" 'ERROR'
        }
    }

    if ($push.Code -ne 0) {
        Write-Log "Push failed: $($push.Output)" 'ERROR'
        $rollback = Invoke-Git -GitArgs @('reset', '--hard', 'HEAD~1')
        if ($rollback.Code -eq 0) {
            Write-Log 'Rolled back the local commit so the next run can retry cleanly.' 'WARN'
        }
        else {
            Write-Log "Rollback failed: $($rollback.Output)" 'ERROR'
        }
        Fail 'Push failed. Module was not published.'
    }

    Write-Log "Pushed to $Remote/$Branch." 'OK'
    Write-Log "Published: $commitMessage" 'OK'
}

try {
    Write-Log '=== Excel Notes daily publish run started ===' 'INFO'
    if ($DryRun) {
        Write-Log 'Mode: DRY RUN (no files will be changed)' 'WARN'
    }
    Invoke-Publish
    Write-Log '=== Run finished ===' 'OK'
    exit 0
}
catch {
    Write-Log "Unhandled error: $($_.Exception.Message)" 'ERROR'
    exit 1
}
