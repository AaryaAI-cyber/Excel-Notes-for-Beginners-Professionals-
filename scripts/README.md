# scripts

## post_daily.ps1

Publishes **exactly one** Excel Notes module per calendar day to `main`, so the repository shows steady, verifiable progress without manual work.

### What the publisher does

1. Verifies the repository, the staging library, and the `main` branch.
2. Sets the **local** Git identity to the GitHub noreply address, so no private email appears in public commits.
3. Runs `git fetch origin` and `git pull --rebase --autostash origin main`.
4. Stops cleanly if a `docs: add Module` commit already exists for the run date. The run date is captured at the start of the run, so a midnight rollover cannot cause a double publish.
5. Picks the next staging folder (`01` to `35`) that Git does not track yet.
6. Validates the module before publishing:
   - `README.md`, `exercises_and_solutions.md`, `practice_workbook.xlsx`, and `images/` exist.
   - Every relative link and image resolves, including links inside nested Markdown docs such as `images/screenshots/README.md`. Each link resolves from the folder that contains its file, and fenced code examples are ignored.
   - The workbook is a valid `.xlsx` file and is under 2 MB.
   - No local, sandbox, or personal file paths appear.
   - The seven standard lesson sections are present.
   - The exercises include visible solutions.
   - PNG images have a valid signature.
7. Copies the module into the repository.
8. Updates the roadmap status and progress block in `README.md`.
9. Commits with `docs: add Module NN - Topic Title`.
10. Pushes to `origin/main`.
11. If the push fails, retries once after a rebase; if it still fails, rolls the local commit back so the next run can retry cleanly.

A timestamped log is written to the workspace root as `excel-notes-daily.log`, outside the repository.

### Run it manually

```powershell
# Full validation, no changes
.\scripts\post_daily.ps1 -DryRun

# Publish the next module now
.\scripts\post_daily.ps1
```

### Schedule

The Windows Task Scheduler job is named **Excel Notes Daily Publish** and runs daily at **21:30 local time**. It is staggered away from the existing production publishing job.

The task is configured with:

- **StartWhenAvailable:** on — runs as soon as possible after a missed start.
- **AllowStartIfOnBatteries:** on — runs on battery power.
- **StopIfGoingOnBatteries:** off — a power change does not interrupt a publish.
- **MultipleInstances:** IgnoreNew — two runs never overlap.

Useful commands:

```powershell
Start-ScheduledTask -TaskName 'Excel Notes Daily Publish'
Get-ScheduledTaskInfo -TaskName 'Excel Notes Daily Publish'
Get-ScheduledTask -TaskName 'Excel Notes Daily Publish' | Select-Object State
```

### Change the schedule

```powershell
$trigger = New-ScheduledTaskTrigger -Daily -At '21:30'
Set-ScheduledTask -TaskName 'Excel Notes Daily Publish' -Trigger $trigger
```

### Move the staging library

The publisher expects the staging library next to the repository:

```text
workspace/
├── Excel-Notes-for-Beginners-Professionals-/   the Git repository
└── excel-staging/
    └── Excel-Notes-for-Beginners-Professionals-/   the 35 module folders
```

If you move it, set the `EXCEL_NOTES_STAGING` environment variable to the new path, or edit `$StagingPath` at the top of the script.

### Safety notes

- The script only touches this repository and the staging library.
- It never modifies the separate `production-data-science-ai` repository, script, or scheduled task.
- The local Git identity is set inside this repository only.
- A failed push is rolled back, so the next run starts from a clean state.
- The one-module-per-day guard uses the commit message and committer date, plus the run date captured at the start of the run, so a second run on the same day exits without publishing and a midnight rollover cannot cause a double publish.
