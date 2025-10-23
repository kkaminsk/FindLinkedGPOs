# Spec: Output Structure

## MODIFIED Requirements

### Requirement: Output folder hierarchy

The script SHALL create a structured output directory with dedicated subfolders for GPO exports organized by link status.

#### Scenario: Standard folder creation

- **GIVEN** the script initializes output
- **WHEN** `Initialize-Output` executes
- **THEN** the following structure SHALL be created:
  ```
  Find-LinkedGPOs-YYYY-MM-DD-HH-MM/
  ├── GPO/
  │   ├── linked/
  │   └── unlinked/
  ├── reports/
  ├── linked-gpos.xml
  ├── validation.xsd
  └── Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log
  ```

#### Scenario: GPO subfolders exist

- **GIVEN** `Initialize-Output` has completed
- **THEN** the `GPO/linked/` directory SHALL exist
- **AND** the `GPO/unlinked/` directory SHALL exist
- **AND** the `reports/` directory SHALL exist for WMI filter extraction

### Requirement: Zip archive includes GPO folders

The script SHALL include the new GPO folder structure when creating the output ZIP archive.

#### Scenario: Complete archive

- **GIVEN** GPO exports have been generated
- **AND** `-NoZip` is not specified
- **WHEN** `Compress-Output` executes
- **THEN** the ZIP SHALL contain all files in `GPO/linked/` and `GPO/unlinked/`
- **AND** the ZIP SHALL preserve the folder hierarchy
