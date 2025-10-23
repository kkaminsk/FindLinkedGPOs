# Proposal: Add WMI Filter Query Capture

## Why

The project.md specifies that WMI filters should include "name, GUID, and query text" (lines 8, 85). Currently, the script only captures WMI filter Name and ID but not the actual query text. This makes it impossible to:
- **Audit WMI filter logic** without manually opening Group Policy Management Console
- **Analyze filter complexity** or identify problematic queries
- **Document complete GPO configurations** for compliance and disaster recovery
- **Meet the original project requirement** for complete WMI filter capture

The query text contains critical logic (e.g., OS version checks, hardware criteria) that determines GPO targeting behavior.

## What Changes

### WMI Query Extraction
- Extract WMI filter query text by generating temporary XML reports using `Get-GPOReport -ReportType Xml`
- Parse XML to locate `<Query>` elements within WMI filter sections
- Cache results to avoid duplicate extractions across multiple link references

### Storage Format
- Create `WMI/` subfolder in output directory
- One XML file per unique WMI filter: `WMI/{FilterGUID}.xml`
- Each file contains:
  ```xml
  <WmiFilter>
    <Name>Filter Name</Name>
    <Id>{GUID}</Id>
    <Query>WQL query text</Query>
  </WmiFilter>
  ```

### Folder Structure
```
Find-LinkedGPOs-YYYY-MM-DD-HH-MM/
├── GPO/
│   ├── linked/
│   └── unlinked/
├── WMI/                           # NEW: WMI filter queries
│   ├── {FilterGUID1}.xml
│   ├── {FilterGUID2}.xml
│   └── ...
├── *.xml files
└── Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log
```

### Parameter Changes
- **ADD** `-SkipWmiQueries` to skip WMI query extraction
- **Behavior**: When both `-SkipXml` AND `-SkipHtml` are specified, automatically set `-SkipWmiQueries` to true (no GPO reports = no query extraction possible)
- **Independence**: `-SkipWmiQueries` can be used standalone without affecting XML/HTML exports

### XML Output Updates
- Link enumeration XMLs continue to capture Name and ID
- Add `<query>` attribute to `<WmiFilter>` elements in link XMLs when query is available
- Gracefully omit query attribute if extraction fails

## Impact

### Affected Specs
- **WMI Filter Extraction** (new spec) - query extraction and storage
- **Output Structure** (modified) - add WMI folder
- **Script Parameters** (modified) - add -SkipWmiQueries

### Affected Code
- `Initialize-Output` - create `WMI/` subfolder
- `Get-GpoWmiInfo` - restore query extraction logic with temporary XML generation
- `Write-DomainRootXml` and `Write-PerOuXml` - include query in WMI filter attributes
- Parameter block - add `-SkipWmiQueries` switch
- Main script logic - handle SkipWmiQueries when both SkipXml and SkipHtml are true

### Breaking Changes
None - this is additive functionality restoring original project requirements

### Performance
- **Slight increase**: Generates temporary XML for each unique GPO with a WMI filter
- **Mitigation**: 
  - Cache filter queries by GUID to avoid duplicate extractions
  - Skip extraction entirely with `-SkipWmiQueries`
  - Only generate XML when query is not already cached

### Benefits
- Completes original project requirements (project.md lines 8, 85)
- Enables complete GPO configuration documentation
- Supports WMI filter auditing and analysis workflows
- Provides machine-readable query data for automation

## Design Considerations

### Temporary XML Management
- Generate temporary XML in memory or temp location
- Parse immediately and discard
- Do NOT write to output folder (avoid clutter)

### Caching Strategy
- Use script-scoped hashtable: `$script:WmiFilterQueries[@{FilterGuid -> Query}]`
- Check cache before generating XML
- Populate cache during link enumeration
- Write WMI folder files after all links enumerated

### Error Handling
- If XML generation fails: log warning, continue without query
- If parsing fails: log warning, continue without query
- Never fail script execution due to WMI query extraction errors
