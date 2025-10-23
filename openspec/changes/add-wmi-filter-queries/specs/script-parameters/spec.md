# Spec: Script Parameters

## ADDED Requirements

### Requirement: SkipWmiQueries parameter

The script SHALL provide a `-SkipWmiQueries` parameter to skip WMI filter query extraction.

#### Scenario: SkipWmiQueries parameter behavior

- **GIVEN** the script is invoked with `-SkipWmiQueries`
- **WHEN** WMI filter processing occurs
- **THEN** no temporary XML reports SHALL be generated for query extraction
- **AND** `Get-GpoWmiInfo` SHALL return only Name and Id
- **AND** link enumeration XMLs SHALL omit the `query` attribute
- **AND** the `WMI/` folder SHALL NOT be created

#### Scenario: Auto-enable when GPO reports skipped

- **GIVEN** the script is invoked with both `-SkipXml` and `-SkipHtml`
- **WHEN** parameter binding completes
- **THEN** `-SkipWmiQueries` SHALL be automatically set to true
- **NOTE**: Query extraction requires generating temporary XML reports, so skipping all GPO reports implicitly skips queries

#### Scenario: Independent from GPO report parameters

- **GIVEN** the script is invoked with `-SkipWmiQueries` alone
- **WHEN** GPO export phase executes
- **THEN** XML and HTML GPO reports SHALL still be generated (unless explicitly skipped)
- **AND** WMI query extraction SHALL be skipped
- **AND** link XMLs SHALL only contain WMI filter Name and Id

#### Scenario: Combined with SkipGpoReports

- **GIVEN** the script is invoked with `-SkipGpoReports`
- **WHEN** parameter processing occurs
- **THEN** `-SkipXml`, `-SkipHtml`, AND `-SkipWmiQueries` SHALL all be set to true
- **AND** no GPO exports or WMI query extraction SHALL occur
