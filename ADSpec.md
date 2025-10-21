Active Directory Application Specification  
Find-LinkedGPOs

The script will crawl active directory to note the different group policies that are linked in each OU. The script must output the data in an XML format for analysis later.

When representing the location in Active Directory such as the OU or organizational Unit use the full LDAP path to that location as a field in the XML rather than trying to construct an XML that conforms to different directory structures.

The script must use Active Directory PowerShell and PowerShell 5 at a minimum. The script must use the least permissions possible.

The data collected needs to be:

* The GPO Name  
* The GPO Location  
* The Security Settings for the GPO  
  * Member SID  
  * Member Name  
* Any WMI filter on the object

Also collect the following output file.  
Get-GPOReport \-Name "Name of Your GPO" \-ReportType Xml \-Path "C:\\Path\\To\\YourReport.xml"

The script also needs to have a log for execution of the script which logs to the console and a Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log file.

Files created by exporting data need to be in the following folder Find-LinkedGPOs-YYYY-MM-DD-HH-MM.

Prompt the user to store the logs in the documents folder, C:\\Temp or a custom folder.

Zip the folder and put the log file inside.

Document the required permissions, powershell module dependencies and how to use the script.

## Detailed Requirements

### Functional Requirements
- Enumerate all link locations in the target domain: Organizational Units (recursively), the domain root, and Active Directory Sites.
- For each link, capture:
  - GPO display name and GUID.
  - Link location as LDAP distinguished name (DN) for OUs/domain, or Site name for site links.
  - Link attributes: `enabled`, `enforced`, and link `order` (precedence).
  - Security filtering principals: SID, resolved display name (if available), and explicit rights (e.g., `Read`, `Apply Group Policy`).
  - Any associated WMI filter: name, GUID, and WMI query text.
  - OU-level inheritance state: whether inheritance is blocked at that OU.
- Produce outputs consisting of an aggregated index XML and per-OU XML files for detailed results.
- Export per-GPO XML reports via `Get-GPOReport -ReportType Xml` by default (allow disabling via parameter).
- Create a timestamped output folder and a matching `.log` file; zip the folder upon completion while retaining the unzipped folder.
- Support interactive prompts for output location, with non-interactive parameter equivalents for automation.

### Non-Functional Requirements
- Read-only: no modifications to AD or GPOs.
- Least privilege: requires only read access to directory and GPO metadata.
- Performance: must handle large domains; allow scoping to a search base.
- Compatibility: Windows PowerShell 5.1; GroupPolicy and ActiveDirectory modules.
- Observability: log to console and file; support `-Verbose` for additional detail.
- Error handling: clear error messages; non-zero exit code on fatal errors.

## Parameters (Proposed)
- `-OutputRoot [string]` default to user's Documents; prompt if not provided.
- `-SkipGpoReports [switch]` skip exporting per-GPO reports (exports by default to `reports/`).
- `-SearchBase [string]` LDAP DN to scope OU enumeration (optional; not used by default).
- `-Domain [string[]]` target domain(s); default current logon domain.
- `-Credential [pscredential]` optional alternate credentials.
- `-ExcludeDomainRoot [switch]` skip domain root GPO links (included by default).
- `-IncludeDefaultContainers [switch]` include default containers (default behavior is include).
- `-NoZip [switch]` skip zipping outputs (optional).
- `-Verbose` show verbose runtime details.

## Outputs
- Timestamped directory: `Find-LinkedGPOs-YYYY-MM-DD-HH-MM/`
  - Log: `Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log`
  - Aggregated index XML: `linked-gpos.xml`
  - Per-OU XML files: `<SanitizedOUName>-linked-gpos.xml` (OU names sanitized for filesystem safety)
  - Per-GPO reports (default): `reports/<gpo-name>.xml`
  - XML Schema: `validation.xsd` for validating XML outputs

## XML Schema (Illustrative)
```xml
<FindLinkedGPOs domain="contoso.com" generatedUtc="2025-01-01T12:00:00Z" host="HOSTNAME">
  <OU dn="OU=Sales,DC=contoso,DC=com" inheritanceBlocked="false">
    <GpoLink displayName="Baseline Workstations" guid="{00000000-0000-0000-0000-000000000000}" enabled="true" enforced="false" order="1">
      <SecurityFiltering>
        <Principal sid="S-1-5-32-545" name="Users" rights="Read" />
        <Principal sid="S-1-5-21-...-515" name="Domain Computers" rights="Read,Apply" />
      </SecurityFiltering>
      <WmiFilter name="Win10 and later" id="{11111111-1111-1111-1111-111111111111}" query="SELECT * FROM Win32_OperatingSystem WHERE Version >= '10.0'" />
    </GpoLink>
  </OU>
  <!-- Domain root and site examples would follow a similar structure -->
</FindLinkedGPOs>
```
Notes:
- Use full OU DN for location.
- Include WMI filter if present; omit element if none; include query text when present.
- Capture link attributes such as `enabled`, `enforced`, and `order`.
- Include explicit rights for security principals; if a name cannot be resolved, SID-only is acceptable.

## Dependencies
- PowerShell 5.1
- Modules: ActiveDirectory, GroupPolicy
- Built-in: `Compress-Archive`
- Network access to domain controllers/LDAP and SYSVOL (for `Get-GPOReport`).

## Permissions
- Directory read permissions to enumerate OUs and read GPO metadata.
- Read access to each GPO's delegation to retrieve security filtering (`Get-GPPermissions`).
- No administrative privileges required when read access is granted.

## Usage (Illustrative)
```powershell
# Full domain crawl (exports GPO reports by default)
./Find-LinkedGPOs.ps1 -Verbose

# Scoped crawl to an OU
./Find-LinkedGPOs.ps1 -SearchBase "OU=Desktops,DC=contoso,DC=com"

# Specify output root without prompts (non-interactive)
./Find-LinkedGPOs.ps1 -OutputRoot "C:\Temp\FindLinkedGPOs" -Verbose

```

## Acceptance Criteria
- Produces an aggregated index XML and per-OU XML files covering OUs, domain root, and sites.
- Each link entry includes GPO name and GUID, location (DN or site), link attributes (enabled, enforced, order), security principals (SID, display name if resolvable, and rights), and WMI filter (name, GUID, and query text when present).
- Creates `validation.xsd` and XML that validates against it.
- Exports GPO reports for all discovered GPOs by default; can be disabled with `-SkipGpoReports`.
- Creates timestamped output directory, logs to console and to a timestamped `.log` file, zips the folder by default, and retains the unzipped directory.
- Supports both interactive prompts and non-interactive parameters (e.g., `-OutputRoot`, `-Domain`, `-Credential`).
- Executes read-only without requiring elevated privileges beyond directory/GPO read access.

## Open Questions
- None. Decisions captured above.