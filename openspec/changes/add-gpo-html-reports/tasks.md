# Implementation Tasks

## 1. Add New Parameters

- [x] 1.1 Add `-SkipXml` switch parameter to script parameter block
- [x] 1.2 Add `-SkipHtml` switch parameter to script parameter block
- [x] 1.3 Add logic to set both flags when `-SkipGpoReports` is specified (backward compatibility)
- [x] 1.4 Update script help/comments to document new parameters

## 2. Update Output Folder Structure

- [x] 2.1 Modify `Initialize-Output` to create `GPO/linked/` subfolder
- [x] 2.2 Modify `Initialize-Output` to create `GPO/unlinked/` subfolder
- [x] 2.3 Update returned object to include `GpoLinkedFolder` and `GpoUnlinkedFolder` properties
- [x] 2.4 Keep existing `ReportsFolder` for WMI filter extraction

## 3. Track Linked GPO GUIDs

- [x] 3.1 Create a script-scoped hashtable `$script:LinkedGpoGuids` to track linked GPOs
- [x] 3.2 Populate hashtable during domain root link enumeration (`Get-DomainRootLinks`)
- [x] 3.3 Populate hashtable during OU link enumeration (`Get-GpoLinksForOu`)
- [x] 3.4 Ensure GUID format normalization (with or without braces) for consistent lookup

## 4. Create Dual-Format Export Function

- [x] 4.1 Create `Export-GpoDualReport` function accepting GPO ID, target folder, domain, and format flags
- [x] 4.2 Implement XML export when `-SkipXml` is not set
- [x] 4.3 Implement HTML export when `-SkipHtml` is not set
- [x] 4.4 Use `Sanitize-FileName` for consistent naming between formats
- [x] 4.5 Implement filename collision detection with hashtable
- [x] 4.6 Append GUID to filename when collision detected
- [x] 4.7 Return result object with success/failure status per format
- [x] 4.8 Log individual GPO processing (name, GUID, classification)

## 5. Enumerate All GPOs

- [x] 5.1 Create `Get-AllGpos` function to enumerate all GPOs per domain using `Get-GPO -All`
- [x] 5.2 Return array of GPO objects with Id and DisplayName properties
- [x] 5.3 Handle errors gracefully with logging and continue

## 6. Add GPO Export Phase to Main Script

- [x] 6.1 Add GPO export phase after link enumeration completes
- [x] 6.2 Check if both `-SkipXml` and `-SkipHtml` are true; skip phase entirely if so
- [x] 6.3 Initialize tracking variables (counters for linked/unlinked, successes/failures)
- [x] 6.4 Iterate through each target domain
- [x] 6.5 Call `Get-AllGpos` for each domain
- [x] 6.6 For each GPO, check if GUID exists in `$script:LinkedGpoGuids`
- [x] 6.7 Determine target folder (`GpoLinkedFolder` or `GpoUnlinkedFolder`)
- [x] 6.8 Call `Export-GpoDualReport` with appropriate parameters
- [x] 6.9 Track results and increment counters

## 7. Update Logging

- [x] 7.1 Log start of GPO export phase
- [x] 7.2 Log total GPO count per domain
- [x] 7.3 Log each GPO: display name, GUID, linked/unlinked classification
- [x] 7.4 Log per-format export success or warning on failure
- [x] 7.5 Log summary after phase completes: total, linked count, unlinked count, XML successes, HTML successes, failures

## 8. Preserve Existing WMI Filter Extraction

- [x] 8.1 Verify `Get-GpoWmiInfo` continues to use `ReportsFolder` for temporary XML
- [x] 8.2 Ensure `Export-GpoReport` (existing function) still works for WMI extraction
- [x] 8.3 Confirm `-SkipGpoReports` flag properly controls WMI extraction behavior (existing logic)

## 9. Testing and Validation

- [x] 9.1 Test with small domain: verify both XML and HTML created in correct folders
- [x] 9.2 Test with domain having unlinked GPOs: verify unlinked subfolder populated
- [x] 9.3 Test with `-SkipXml`: verify only HTML files created
- [x] 9.4 Test with `-SkipHtml`: verify only XML files created
- [x] 9.5 Test with `-SkipGpoReports`: verify backward compatibility (no exports in GPO folder)
- [x] 9.6 Test with both `-SkipXml` and `-SkipHtml`: verify no GPO export phase runs
- [x] 9.7 Test filename collision handling with duplicate GPO names
- [x] 9.8 Test multi-domain scenario
- [x] 9.9 Verify WMI filter extraction still works with new folder structure
- [x] 9.10 Verify ZIP archive includes new GPO subfolders

## 10. Documentation

- [x] 10.1 Update script help text and comments
- [x] 10.2 Document new parameters in README (if exists)
- [x] 10.3 Update any examples or usage guides
