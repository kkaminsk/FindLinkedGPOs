# Spec: Output Structure

## MODIFIED Requirements

### Requirement: Output folder hierarchy includes WMI subfolder

The script SHALL create a WMI subfolder when WMI filter queries are extracted.

#### Scenario: WMI folder creation

- **GIVEN** at least one WMI filter with a query has been extracted
- **AND** `-SkipWmiQueries` is not specified
- **WHEN** `Initialize-Output` executes
- **THEN** a `WMI/` directory SHALL be created within the output folder

#### Scenario: WMI folder structure

- **GIVEN** the script has completed execution with WMI query extraction
- **THEN** the output structure SHALL be:
  ```
  Find-LinkedGPOs-YYYY-MM-DD-HH-MM/
  ├── GPO/
  │   ├── linked/
  │   └── unlinked/
  ├── WMI/
  │   ├── {FilterGUID1}.xml
  │   ├── {FilterGUID2}.xml
  │   └── ...
  ├── DomainRoot--domain-linked-gpos.xml
  ├── OU_DN--domain-linked-gpos.xml
  ├── linked-gpos.xml
  ├── validation.xsd
  └── Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log
  ```

#### Scenario: WMI folder omitted when skipped

- **GIVEN** `-SkipWmiQueries` is specified
- **OR** both `-SkipXml` and `-SkipHtml` are specified
- **WHEN** the script completes
- **THEN** the `WMI/` folder SHALL NOT be created

#### Scenario: ZIP archive includes WMI folder

- **GIVEN** WMI filter queries have been extracted
- **AND** `-NoZip` is not specified
- **WHEN** `Compress-Output` executes
- **THEN** the ZIP SHALL include all files in the `WMI/` folder
- **AND** preserve the folder hierarchy
