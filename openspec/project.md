# Project Context

## Purpose
 FindLinkedGPOs is a Windows PowerShell script that audits Active Directory Group Policy links across Organizational Units (OUs), the domain root, and Active Directory Sites, outputting machine-readable XML for later analysis. It exports full GPO reports by default, writes logs to both console and file, places all outputs in a timestamped folder, and zips the folder on completion while retaining the unzipped directory. The script operates with least privilege and performs read-only queries against AD and GPO metadata.

Goals:
- Enumerate OUs, the domain root, and Sites and their linked GPOs using least privilege; support multi-domain targets via `-Domain`.
- Capture for each link: GPO name and GUID; link location (OU DN, domain DN, or Site name); link attributes (`enabled`, `enforced`, `order`); security filtering principals (SID, display name if resolvable, and rights); WMI filter (name, GUID, and query text).
- Produce an aggregated index XML and per-OU XML files for downstream analysis; generate `validation.xsd` to validate outputs.
- Export per-GPO reports via `Get-GPOReport -ReportType Xml` by default (allow `-SkipGpoReports` to disable).
- Create an output folder named `Find-LinkedGPOs-YYYY-MM-DD-HH-MM`, log execution to `Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log`, then zip the folder and retain the unzipped directory.
- Default interactive prompts; support non-interactive parameters (`-OutputRoot`, `-Domain`, `-Credential`, `-SearchBase`).

## Tech Stack
- Windows PowerShell 5.1 (target/runtime)
- Modules
  - ActiveDirectory (via RSAT or on a domain controller)
  - GroupPolicy (for `Get-GPOReport` and link metadata)
- Windows OS: domain-joined host (DC or non-DC with RSAT) with network access to domain controllers
- Built-in utilities: `Compress-Archive` for zipping
- Optional tooling (TBD): Pester for tests, PSScriptAnalyzer for linting
- XML validation: `validation.xsd` produced alongside outputs
- Multi-domain support via `-Domain` (string array)

## Project Conventions

### Code Style
- PowerShell Verb-Noun functions; PascalCase for functions/parameters; camelCase for locals
- Indentation: 4 spaces; max line length: 120
- Use approved verbs (Get, Set, New, Write, Export)
- Prefer `Write-Information`/`Write-Verbose`/`Write-Warning` over `Write-Host`; tee to log
- Use `[CmdletBinding()]` and advanced functions; support `-Verbose`
- Error handling with try/catch; throw terminating errors for unrecoverable issues
- Naming
  - Output folder: `Find-LinkedGPOs-YYYY-MM-DD-HH-MM`
  - Log file: `Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log`
  - Zip: same as folder name with `.zip`

### Architecture Patterns
- Single-script initially with internal functions:
  - `Initialize-Logging`, `Get-TargetDomains`, `Get-TargetOUs`, `Get-DomainRootLinks`, `Get-SiteLinks`, `Get-LinkedGpos`, `Get-GpoSecurityFiltering`, `Get-GpoWmiFilter`, `Write-PerOuXml`, `Write-IndexXml`, `Write-Xsd`, `Export-GpoReport`, `Compress-Output`
- Read-only AD queries; no writes to AD
- Scope controls: `-SearchBase` (optional, not used by default), `-LDAPFilter`, `-IncludeDefaultContainers`
- Interactive by default; parameters support fully non-interactive operation

### Testing Strategy
- Unit tests with Pester (TBD):
  - Mock AD/GP cmdlets
  - Validate XML schema and file naming
  - Validate XML against `validation.xsd`
  - Verify logging behavior and zip retention
  - Verify fallback to SID-only when name resolution fails
  - Verify multi-domain enumeration and per-OU file generation
- Static analysis with PSScriptAnalyzer (ruleset TBD)

### Git Workflow
- Trunk-based on `main` with PRs
- Conventional commits (`feat:`, `fix:`, `docs:`)
- Use OpenSpec files under `openspec/` to drive changes

## Domain Context
- Active Directory objects addressed by LDAP distinguished names (DN), e.g., `OU=Sales,DC=contoso,DC=com`
- GPO links exist on OUs, the domain root, and Sites; include OU inheritance state (block inheritance)
- Security filtering includes rights evaluation (Read and Apply Group Policy)
- GPO WMI filter may include a query; include name, GUID, and query when present

## Important Constraints
- Must run with least privilege sufficient for read access to AD and GPO metadata
- PowerShell 5.1 compatibility; no third-party binaries required
- Performance: handle large domains gracefully; allow scoping and paging; multi-domain supported
- Default interactive UX; fully automatable via parameters (`-OutputRoot`, `-Domain`, `-Credential`, `-SearchBase`)
- Output and logs must be deterministic and timestamped; zip outputs by default and retain the unzipped directory
- No modifications to AD; read-only operations only

## External Dependencies
- RSAT ActiveDirectory and GroupPolicy modules installed locally or execution on a domain controller
- Network access to domain controllers/LDAP and SYSVOL for `Get-GPOReport`
- Zip creation via `Compress-Archive`

## Decisions
- Include links at OUs, the domain root, and Sites
- Multi-domain supported via `-Domain` (string array)
- Include link attributes: `enabled`, `enforced`, `order`; include OU inheritance state
- Security filtering includes rights; SID-only acceptable when name cannot be resolved
- WMI filter includes name, GUID, and query text
- Outputs: aggregated `linked-gpos.xml` and per-OU XML files; generate `validation.xsd`
- Export `Get-GPOReport` outputs by default; `-SkipGpoReports` disables
- `-SearchBase` available but not used by default
- Interactive by default; `-OutputRoot` to set path; support `-Credential`
- Zip by default and retain the unzipped folder
