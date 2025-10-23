# Spec: GPO Export

## ADDED Requirements

### Requirement: Enumerate all domain GPOs

The script SHALL enumerate all Group Policy Objects in each target domain, not limited to GPOs with active links.

#### Scenario: Complete domain GPO discovery

- **GIVEN** the script is running against a domain
- **WHEN** GPO enumeration begins
- **THEN** all GPOs SHALL be retrieved using `Get-GPO -All -Domain <domain>`
- **AND** both linked and unlinked GPOs SHALL be included in the enumeration

### Requirement: Classify GPOs as linked or unlinked

The script SHALL track which GPOs have been encountered during link enumeration to classify them as linked or unlinked.

#### Scenario: Linked GPO identification

- **GIVEN** GPO link enumeration has completed for all OUs, domain root, and Sites
- **WHEN** a GPO GUID appears in any `Get-GPInheritance` result
- **THEN** that GPO SHALL be classified as "linked"

#### Scenario: Unlinked GPO identification

- **GIVEN** all GPOs have been enumerated with `Get-GPO -All`
- **WHEN** a GPO GUID does not appear in any link enumeration results
- **THEN** that GPO SHALL be classified as "unlinked"

### Requirement: Export dual-format GPO reports

The script SHALL generate both XML and HTML reports for each GPO unless explicitly skipped via parameters.

#### Scenario: Both formats exported by default

- **GIVEN** a GPO is being processed
- **AND** neither `-SkipXml` nor `-SkipHtml` is specified
- **WHEN** export executes
- **THEN** both `{GPOName}.xml` AND `{GPOName}.html` SHALL be created
- **AND** both files SHALL use sanitized filenames based on GPO display name

#### Scenario: XML-only export

- **GIVEN** a GPO is being processed
- **AND** `-SkipHtml` is specified but `-SkipXml` is not
- **WHEN** export executes
- **THEN** only `{GPOName}.xml` SHALL be created

#### Scenario: HTML-only export

- **GIVEN** a GPO is being processed
- **AND** `-SkipXml` is specified but `-SkipHtml` is not
- **WHEN** export executes
- **THEN** only `{GPOName}.html` SHALL be created

#### Scenario: No exports when both formats skipped

- **GIVEN** both `-SkipXml` AND `-SkipHtml` are specified
- **WHEN** the script reaches the GPO export phase
- **THEN** no GPO enumeration or exports SHALL occur

### Requirement: Organize exports by link status

The script SHALL write GPO exports to separate subfolders based on whether the GPO is linked or unlinked.

#### Scenario: Linked GPO placement

- **GIVEN** a GPO has been classified as "linked"
- **WHEN** exports are written
- **THEN** files SHALL be placed in `<output>/GPO/linked/`

#### Scenario: Unlinked GPO placement

- **GIVEN** a GPO has been classified as "unlinked"
- **WHEN** exports are written
- **THEN** files SHALL be placed in `<output>/GPO/unlinked/`

### Requirement: Handle filename collisions

The script SHALL prevent filename overwrites when multiple GPOs have identical display names.

#### Scenario: Duplicate GPO names

- **GIVEN** two GPOs have the same display name
- **WHEN** sanitized filenames collide
- **THEN** the second occurrence SHALL append the GPO GUID to the filename
- **AND** both XML and HTML SHALL use the same modified filename

### Requirement: Log GPO export progress

The script SHALL log detailed information about GPO export operations.

#### Scenario: Export phase logging

- **GIVEN** GPO export phase begins
- **THEN** the script SHALL log:
  - Start of GPO export phase
  - Total count of GPOs discovered per domain
  - Each GPO being processed (display name and GUID)
  - Classification (linked vs. unlinked) for each GPO
  - Export success or failure per format (XML/HTML)
  - Summary: total processed, linked count, unlinked count, XML successes, HTML successes, failures

### Requirement: Backward compatibility for SkipGpoReports

The existing `-SkipGpoReports` parameter SHALL continue to work and SHALL set both `-SkipXml` and `-SkipHtml` to true.

#### Scenario: Legacy parameter usage

- **GIVEN** the script is invoked with `-SkipGpoReports`
- **WHEN** parameter processing occurs
- **THEN** both `-SkipXml` and `-SkipHtml` SHALL be set to true
- **AND** no GPO exports SHALL be generated
- **AND** no deprecation warning SHALL be shown (silent compatibility)

### Requirement: Preserve reports folder for WMI extraction

The existing `reports/` folder SHALL remain and continue to be used for temporary XML files during WMI filter extraction.

#### Scenario: WMI filter extraction still works

- **GIVEN** a linked GPO has a WMI filter
- **AND** the script needs to extract the WMI filter query
- **WHEN** `Get-GpoWmiInfo` is called
- **THEN** a temporary XML report SHALL be written to `<output>/reports/`
- **AND** the WMI query SHALL be extracted from that XML
- **AND** this behavior SHALL work regardless of `-SkipXml` or `-SkipHtml` settings
