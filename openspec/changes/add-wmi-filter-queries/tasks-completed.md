# Implementation Tasks

## 1. Add New Parameter

- [x] 1.1 Add `-SkipWmiQueries` switch parameter to script parameter block
- [x] 1.2 Add logic to set `-SkipWmiQueries` to true when both `-SkipXml` and `-SkipHtml` are true
- [x] 1.3 Update `-SkipGpoReports` logic to also set `-SkipWmiQueries` to true
- [x] 1.4 Update script help/comments to document `-SkipWmiQueries` parameter

## 2. Update Output Folder Structure

- [x] 2.1 Modify `Initialize-Output` to create `WMI/` subfolder
- [x] 2.2 Update returned object to include `WmiFolder` property
- [x] 2.3 Ensure WMI folder creation is conditional (only if queries will be extracted)

## 3. Create WMI Filter Query Cache

- [x] 3.1 Create script-scoped hashtable `$script:WmiFilterQueries` to cache queries by GUID
- [x] 3.2 Initialize cache at script start (before domain enumeration)

## 4. Restore Query Extraction in Get-GpoWmiInfo

- [x] 4.1 Add `-SkipWmiQueries` parameter to `Get-GpoWmiInfo` function
- [x] 4.2 Check cache for existing query before generating XML
- [x] 4.3 Generate temporary XML report using `Get-GPOReport -ReportType Xml` when not cached
- [x] 4.4 Parse XML to extract WMI filter query text
- [x] 4.5 Store query in cache hashtable keyed by filter GUID
- [x] 4.6 Clean up temporary XML file after parsing
- [x] 4.7 Return object with Name, Id, and Query (if available)
- [x] 4.8 Handle errors gracefully with warning logs

## 5. Create WMI Query Parser Function

- [x] 5.1 Create `Get-WmiFilterQueryFromXml` function accepting XML content or path
- [x] 5.2 Parse XML and locate `<Query>` element within WMI filter section
- [x] 5.3 Extract and return query text
- [x] 5.4 Return $null if query not found or parsing fails
- [x] 5.5 Handle XML parsing errors gracefully

## 6. Update Link XML Writers

- [x] 6.1 Update `Write-DomainRootXml` to pass `-SkipWmiQueries` to `Get-GpoWmiInfo`
- [x] 6.2 Update `Write-DomainRootXml` to include `query` attribute when available
- [x] 6.3 Update `Write-PerOuXml` to pass `-SkipWmiQueries` to `Get-GpoWmiInfo`
- [x] 6.4 Update `Write-PerOuXml` to include `query` attribute when available
- [x] 6.5 Ensure query attribute is omitted (not empty string) when unavailable

## 7. Write WMI Filter Files

- [x] 7.1 Create `Write-WmiFilterFiles` function accepting WmiFolder path and cache hashtable
- [x] 7.2 Iterate through cached filter queries
- [x] 7.3 For each filter, create XML file named `{FilterGUID}.xml`
- [x] 7.4 Write XML structure with Name, Id, and Query elements
- [x] 7.5 Handle file write errors gracefully with warning logs

## 8. Integrate WMI File Writing into Main Script

- [x] 8.1 Call `Write-WmiFilterFiles` after link enumeration completes
- [x] 8.2 Only call if `-SkipWmiQueries` is false and cache is not empty
- [x] 8.3 Log summary of WMI filters written

## 9. Update Logging

- [x] 9.1 Log when WMI query extraction begins for a filter
- [x] 9.2 Log warnings when XML generation or parsing fails
- [x] 9.3 Log summary at end: total WMI filters found, queries extracted, queries failed
- [x] 9.4 Log WMI folder creation

## 10. Testing and Validation

- [ ] 10.1 Test with GPOs that have WMI filters: verify query extraction
- [ ] 10.2 Test with GPOs without WMI filters: verify no errors
- [ ] 10.3 Test with `-SkipWmiQueries`: verify no WMI folder created
- [ ] 10.4 Test with `-SkipXml -SkipHtml`: verify auto-skip of WMI queries
- [ ] 10.5 Test with `-SkipGpoReports`: verify all exports skipped including WMI
- [ ] 10.6 Test cache functionality: verify same filter only extracted once
- [ ] 10.7 Test multiple GPOs with same WMI filter: verify single WMI file
- [ ] 10.8 Test XML file structure: validate WMI filter XML files
- [ ] 10.9 Test link XML output: verify query attribute present when extracted
- [ ] 10.10 Test ZIP archive: verify WMI folder included

## 11. Documentation

- [x] 11.1 Update script help text to document `-SkipWmiQueries` parameter
- [x] 11.2 Update README.md to mention WMI filter query capture
- [x] 11.3 Update output structure diagram in README and script help
- [x] 11.4 Add example showing WMI query extraction
