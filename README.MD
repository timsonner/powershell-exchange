# PowerShell for Exchange

`email-clawback.ps1` searches Exchange Online for messages with a subject you type, then hard-deletes the matches. The search-name prompt uses a spam search as its example (`Search for spam 01-02-2003`).

## Prerequisites

- PowerShell 7 (`pwsh`) or Windows PowerShell, with the `Install-Module` command.
- Network access to PSGallery. The script runs `Install-Module ExchangeOnlineManagement`. That command asks you to confirm the PSGallery repository. The default answer is No. Choose Yes to install the module.
- An account that can sign in when the script runs `Connect-IPPSSession`. The script has no username or password parameter. The signed-in account must be allowed to run `New-ComplianceSearch`, `Start-ComplianceSearch`, `Get-ComplianceSearch`, and `New-ComplianceSearchAction`.

## Install, build, run, and test

This repository has no dependency manifest, build script, or test suite. Module install, import, and sign-in all happen inside `email-clawback.ps1`.

From the repository directory:

```powershell
pwsh -NoProfile -File ./email-clawback.ps1
```

On Windows PowerShell:

```powershell
powershell -NoProfile -File .\email-clawback.ps1
```

## Usage

The script takes no command-line arguments. It runs these steps in order:

1. `Install-Module ExchangeOnlineManagement`
2. `Import-Module ExchangeOnlineManagement`
3. `Connect-IPPSSession`
4. Prompt: `Searchname (e.g Search for spam 01-02-2003)`
5. Prompt: `Subject`

It stores the answers in `$searchName` and `$subject`. The content query is the single string `Subject:` plus the subject you typed. The script then runs:

```powershell
New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery Subject:$subject
Start-ComplianceSearch -Identity $Search.Identity
```

`-ExchangeLocation All` is the search location passed by the script. The script prints `Waiting for the compliance search to complete...` and polls with `Get-ComplianceSearch -Identity $searchName` every 10 seconds. The loop ends when `.Status` equals `Completed`. Any other status keeps the loop running. The script does not stop on command errors, so a declined module install or a failed search still enters this loop.

When the status is `Completed`, the script runs:

```powershell
New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete
```

That action purges the messages matched by the search, with purge type `HardDelete`.

## Layout

```text
email-clawback.ps1
README.MD
```

## License

This repository has no license file.
