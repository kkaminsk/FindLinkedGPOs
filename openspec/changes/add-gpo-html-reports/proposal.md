# Proposal: Add GPO HTML Reports with Linked/Unlinked Organization

## Why

Currently, the script exports XML reports only for linked GPOs encountered during enumeration. Users need comprehensive GPO documentation that includes:
- **Both XML and HTML formats** for all GPOs (not just linked ones)
- **Human-readable HTML reports** for reviewing GPO settings in a browser
- **Organizational separation** between linked and unlinked GPOs for clarity
- **Granular control** over which report formats to generate

This change enables complete GPO inventory and analysis workflows by providing dual-format exports with clear linked/unlinked distinction.

## What Changes

### GPO Enumeration
- Enumerate **all GPOs in the domain** using `Get-GPO -All` (not just linked ones)
- Identify which GPOs are linked vs. unlinked by tracking GPO GUIDs during link enumeration

### Export Formats
- Export **both XML and HTML** for each GPO using `Get-GPOReport -ReportType Xml` and `Get-GPOReport -ReportType Html`
- Generate paired files: `{GPOName}.xml` and `{GPOName}.html`

### Folder Structure
```
Find-LinkedGPOs-YYYY-MM-DD-HH-MM/
├── GPO/                           # NEW: All GPO exports
│   ├── linked/                    # NEW: Linked GPOs
│   │   ├── {GPOName}.xml
│   │   └── {GPOName}.html
│   └── unlinked/                  # NEW: Unlinked GPOs
│       ├── {GPOName}.xml
│       └── {GPOName}.html
└── reports/                       # KEEP: Temporary XML for WMI filter extraction
```

### Parameter Changes
- **DEPRECATE** `-SkipGpoReports` (keep for backward compatibility, maps to both flags)
- **ADD** `-SkipXml` to skip XML report generation
- **ADD** `-SkipHtml` to skip HTML report generation
- Both skips must be true to skip GPO enumeration entirely

### Backward Compatibility
- `-SkipGpoReports` will set both `-SkipXml` and `-SkipHtml` to maintain existing behavior
- Existing `reports/` folder retained for WMI filter extraction workflow

## Impact

### Affected Specs
- **GPO Export** (new spec) - dual-format export with linked/unlinked organization
- **Output Structure** (new spec) - folder hierarchy for GPO reports
- **Script Parameters** (new spec) - granular control over export formats

### Affected Code
- `Initialize-Output` - create `GPO/linked/` and `GPO/unlinked/` subfolders
- `Export-GpoReport` - extend to support both XML and HTML, return both paths
- Main script logic - enumerate all GPOs, track linked GUIDs, export with proper organization
- Parameter block - add `-SkipXml` and `-SkipHtml`, handle backward compatibility for `-SkipGpoReports`
- Logging - report counts for linked/unlinked GPOs and successful/failed exports

### Breaking Changes
None - `-SkipGpoReports` continues to work as before

### Performance
- **Increase**: Exporting all GPOs (both linked and unlinked) and dual formats will increase execution time
- **Mitigation**: Granular skip flags allow users to optimize for their needs

### Benefits
- Complete GPO inventory (not just linked GPOs)
- Human-readable HTML reports alongside machine-parseable XML
- Clear organizational separation for analysis workflows
- Granular control over export formats
