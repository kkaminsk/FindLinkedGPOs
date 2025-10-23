# Spec: WMI Filter Extraction

## ADDED Requirements

### Requirement: Extract WMI filter query text

The script SHALL extract the WQL query text from WMI filters by parsing temporary GPO XML reports.

#### Scenario: Query extraction from GPO XML

- **GIVEN** a GPO has a WMI filter applied
- **WHEN** `Get-GpoWmiInfo` is called for that GPO
- **THEN** the script SHALL generate a temporary XML report using `Get-GPOReport -ReportType Xml`
- **AND** parse the XML to locate the `<Query>` element
- **AND** extract the query text
- **AND** return Name, Id, and Query

#### Scenario: Query extraction failure

- **GIVEN** a GPO has a WMI filter
- **AND** XML generation or parsing fails
- **WHEN** query extraction is attempted
- **THEN** the script SHALL log a warning
- **AND** return Name and Id without Query
- **AND** continue execution without error

### Requirement: Cache WMI filter queries

The script SHALL cache extracted WMI filter queries to avoid duplicate XML generation for the same filter.

#### Scenario: First query extraction

- **GIVEN** a WMI filter GUID has not been cached
- **WHEN** `Get-GpoWmiInfo` encounters the filter
- **THEN** the script SHALL extract the query
- **AND** store it in `$script:WmiFilterQueries` hashtable keyed by GUID

#### Scenario: Cached query retrieval

- **GIVEN** a WMI filter GUID exists in the cache
- **WHEN** `Get-GpoWmiInfo` encounters the filter again
- **THEN** the script SHALL retrieve the query from cache
- **AND** NOT generate a new temporary XML report

### Requirement: Store WMI filter queries in dedicated folder

The script SHALL write WMI filter query data to individual XML files in a `WMI/` subfolder.

#### Scenario: WMI folder creation

- **GIVEN** at least one WMI filter with a query was extracted
- **WHEN** the script completes link enumeration
- **THEN** a `WMI/` folder SHALL exist in the output directory

#### Scenario: One file per filter

- **GIVEN** WMI filter queries have been extracted
- **WHEN** writing WMI filter files
- **THEN** each unique filter SHALL have one XML file named `{FilterGUID}.xml`
- **AND** the file SHALL contain Name, Id, and Query elements

#### Scenario: XML file structure

- **GIVEN** a WMI filter file is being written
- **THEN** the file SHALL use this structure:
  ```xml
  <?xml version="1.0" encoding="UTF-8"?>
  <WmiFilter>
    <Name>FilterName</Name>
    <Id>{GUID}</Id>
    <Query>WQL query text</Query>
  </WmiFilter>
  ```

### Requirement: Include query in link enumeration XML

The script SHALL include the WMI filter query as an attribute in link enumeration XML outputs when available.

#### Scenario: Query attribute in domain root XML

- **GIVEN** a domain root link has a WMI filter
- **AND** the query was successfully extracted
- **WHEN** `Write-DomainRootXml` creates the XML
- **THEN** the `<WmiFilter>` element SHALL include a `query` attribute with the query text

#### Scenario: Query attribute in per-OU XML

- **GIVEN** an OU link has a WMI filter
- **AND** the query was successfully extracted
- **WHEN** `Write-PerOuXml` creates the XML
- **THEN** the `<WmiFilter>` element SHALL include a `query` attribute with the query text

#### Scenario: Omit query attribute when unavailable

- **GIVEN** a WMI filter query could not be extracted
- **WHEN** writing link enumeration XML
- **THEN** the `<WmiFilter>` element SHALL include `name` and `id` attributes
- **AND** the `query` attribute SHALL be omitted

### Requirement: Skip WMI query extraction with parameter

The script SHALL support skipping WMI filter query extraction via the `-SkipWmiQueries` parameter.

#### Scenario: Skip with explicit parameter

- **GIVEN** the script is invoked with `-SkipWmiQueries`
- **WHEN** `Get-GpoWmiInfo` is called
- **THEN** no temporary XML SHALL be generated
- **AND** only Name and Id SHALL be returned
- **AND** the WMI folder SHALL NOT be created

#### Scenario: Auto-skip when both XML and HTML skipped

- **GIVEN** the script is invoked with both `-SkipXml` and `-SkipHtml`
- **WHEN** parameter processing occurs
- **THEN** `-SkipWmiQueries` SHALL be automatically set to true
- **AND** no WMI query extraction SHALL occur

#### Scenario: Independent skip control

- **GIVEN** the script is invoked with `-SkipWmiQueries` but NOT `-SkipXml` or `-SkipHtml`
- **WHEN** GPO export phase executes
- **THEN** GPO XML and HTML reports SHALL still be generated
- **AND** WMI query extraction SHALL be skipped
- **AND** link XMLs SHALL only contain WMI filter Name and Id

### Requirement: Temporary XML cleanup

The script SHALL NOT persist temporary XML files used for query extraction in the output folder.

#### Scenario: Temporary file handling

- **GIVEN** a temporary XML is generated for query extraction
- **WHEN** parsing completes
- **THEN** the temporary file SHALL be discarded or deleted
- **AND** SHALL NOT appear in the output directory
