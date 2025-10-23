# FindLinkedGPOs

A PowerShell script to audit Active Directory Group Policy Object (GPO) links and export comprehensive GPO reports in both XML and HTML formats.

## Features

- **Complete GPO Inventory**: Enumerates all GPOs in the domain (not just linked ones)
- **Dual-Format Reports**: Exports both XML and HTML for each GPO
- **Linked/Unlinked Classification**: Organizes GPOs by whether they're linked to OUs, domain root, or Sites
- **WMI Filter Query Extraction**: Captures complete WMI filter logic including query text
- **Granular Export Control**: Skip XML, HTML, or WMI query exports individually
- **Link Metadata**: Captures link attributes (enabled, enforced, order), security filtering, and WMI filters
- **Multi-Domain Support**: Audit multiple domains in a single run
- **Automatic ZIP Compression**: Creates compressed archives of output (optional)
- **Backward Compatible**: Existing `-SkipGpoReports` parameter still works

## Requirements

### Software Requirements
- **PowerShell**: 5.1 or later
- **Modules**: 
  - ActiveDirectory (RSAT)
  - GroupPolicy (RSAT)
- **Platform**: Windows domain-joined machine or Domain Controller

### Permissions Required

The script performs **read-only operations** and requires the following permissions:

#### Active Directory Permissions
- **Read access to Domain objects** - For domain enumeration and metadata
- **Read access to Organizational Unit (OU) objects** - For OU hierarchy traversal
- **Read access to WMI Filter objects** - Specifically:
  - Objects in `CN=SOM,CN=WMIPolicy,CN=System,DC=...`
  - Read the `msWMI-Parm2` attribute (contains WMI query text)

#### Group Policy Permissions
- **Read access to GPO objects** - Query GPO metadata and properties
- **Read GPO reports** - Generate XML and HTML reports

#### File System Permissions
- **Write access** to the output directory (Documents, C:\Temp, or custom path)
- **Create directories and files**

### Recommended Access Levels

**Minimum (works for most scenarios):**
- Standard **Domain User** account
  - Default AD read permissions are sufficient
  - All authenticated users can read GPOs by default
  - Write access to local output folder

**Recommended (for comprehensive access):**
- Member of **Group Policy Creator Owners** (domain group)
- Member of **Account Operators** (domain group)
- Or **Domain Admins** (for unrestricted access)

**Notes:**
- No local administrator privileges required (unless writing to protected folders)
- Use `-Credential` parameter if running under a different account
- The script does **NOT modify** any AD objects or GPOs

## Installation

1. Clone or download this repository
2. Ensure ActiveDirectory and GroupPolicy modules are installed
3. Run from a domain-joined machine with appropriate permissions

## Usage

### Basic Usage

```powershell
# Interactive mode (prompts for output location)
.\Find-LinkedGPOs.ps1

# Specify output location
.\Find-LinkedGPOs.ps1 -OutputRoot "C:\Temp"
```

### Export Control

```powershell
# XML reports only (skip HTML)
.\Find-LinkedGPOs.ps1 -SkipHtml

# HTML reports only (skip XML)
.\Find-LinkedGPOs.ps1 -SkipXml

# Skip all GPO exports (backward compatibility)
.\Find-LinkedGPOs.ps1 -SkipGpoReports
```

### Multi-Domain

```powershell
# Audit multiple domains
.\Find-LinkedGPOs.ps1 -Domain "contoso.com","fabrikam.com"
```

### Additional Options

```powershell
# Skip ZIP compression
.\Find-LinkedGPOs.ps1 -NoZip

# Use specific search base
.\Find-LinkedGPOs.ps1 -SearchBase "OU=Sales,DC=contoso,DC=com"

# Exclude domain root from enumeration
.\Find-LinkedGPOs.ps1 -ExcludeDomainRoot
```

## Output Structure

The script creates a timestamped folder with the following structure:

```
Find-LinkedGPOs-YYYY-MM-DD-HH-MM/
├── GPO/
│   ├── linked/                    # GPOs linked to OUs/domain/sites
│   │   ├── GPOName.xml           # XML report
│   │   ├── GPOName.html          # HTML report (human-readable)
│   │   └── ...
│   └── unlinked/                  # GPOs not linked anywhere
│       ├── GPOName.xml
│       ├── GPOName.html
│       └── ...
├── WMI/                           # WMI filter queries
│   ├── {FilterGUID1}.xml
│   ├── {FilterGUID2}.xml
│   └── ...
├── DomainRoot--domain-linked-gpos.xml    # Domain root links
├── OU_DN--domain-linked-gpos.xml         # Per-OU links
├── linked-gpos.xml                       # Aggregated index
├── validation.xsd                        # XML schema
└── Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log  # Execution log
```

### Output Files

- **GPO Reports**: Both XML (machine-parseable) and HTML (human-readable) formats
- **WMI Filters**: Individual XML files per WMI filter with Name, Id, and Query
- **Link Data**: Per-OU and domain root XML files with link metadata (includes query attribute when available)
- **Classification**: Separate folders for linked vs. unlinked GPOs
- **Log File**: Detailed execution log with timestamps
- **ZIP Archive**: Compressed version of the entire output (unless -NoZip)

## Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-OutputRoot` | String | Root directory for output (prompts if not specified) |
| `-SkipGpoReports` | Switch | Skip all GPO exports and WMI queries (backward compatibility) |
| `-SkipXml` | Switch | Skip XML report generation |
| `-SkipHtml` | Switch | Skip HTML report generation |
| `-SkipWmiQueries` | Switch | Skip WMI filter query extraction |
| `-Domain` | String[] | One or more domains to audit |
| `-Credential` | PSCredential | Alternate credentials for domain access |
| `-SearchBase` | String | LDAP search base for OU enumeration |
| `-ExcludeDomainRoot` | Switch | Exclude domain root from enumeration |
| `-NoZip` | Switch | Skip ZIP compression |

## Examples

### Example 1: Complete Audit

```powershell
.\Find-LinkedGPOs.ps1 -OutputRoot "C:\GPOAudit"
```

Performs a complete audit with both XML and HTML reports.

### Example 2: HTML Only for Documentation

```powershell
.\Find-LinkedGPOs.ps1 -OutputRoot "C:\GPODocs" -SkipXml
```

Creates human-readable HTML reports only.

### Example 3: Multi-Domain with Custom Base

```powershell
.\Find-LinkedGPOs.ps1 -Domain "corp.contoso.com","emea.contoso.com" -SearchBase "OU=Production,DC=corp,DC=contoso,DC=com"
```

Audits multiple domains starting from a specific OU.

## GPO Link Metadata

For each linked GPO, the script captures:

- **Display Name**: GPO name
- **GUID**: Unique identifier
- **Link Status**: Enabled/disabled
- **Enforcement**: Whether link is enforced
- **Order**: Application order
- **Security Filtering**: Principals with Apply Group Policy permissions
- **WMI Filter**: Name, ID, and Query text (if applied)

## Notes

- The script performs read-only operations (no modifications to AD)
- Large domains may take several minutes to process
- HTML reports are ideal for browsing GPO settings
- XML reports enable programmatic analysis
- Unlinked GPOs may represent unused policies that can be cleaned up

## Version History

### Version 2.1
- Added WMI filter query extraction with dedicated WMI folder
- Added `-SkipWmiQueries` parameter for granular control
- WMI filter query attribute included in link enumeration XMLs

### Version 2.0
- Added dual-format exports (XML + HTML)
- Introduced linked/unlinked GPO classification
- Added granular export control (-SkipXml, -SkipHtml)
- Removed temporary reports folder
- Enhanced logging with summary statistics

### Version 1.0
- Initial release
- XML-only exports for linked GPOs

## License

See LICENSE file for details.

## Support

For issues, questions, or contributions, please open an issue on GitHub.
