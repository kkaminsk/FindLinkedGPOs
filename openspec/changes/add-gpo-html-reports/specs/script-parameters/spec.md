# Spec: Script Parameters

## ADDED Requirements

### Requirement: Granular export control parameters

The script SHALL provide separate parameters to control XML and HTML report generation independently.

#### Scenario: SkipXml parameter

- **GIVEN** the script is invoked with `-SkipXml`
- **WHEN** GPO export phase executes
- **THEN** no XML files SHALL be generated in `GPO/linked/` or `GPO/unlinked/`
- **AND** HTML files MAY still be generated unless `-SkipHtml` is also specified

#### Scenario: SkipHtml parameter

- **GIVEN** the script is invoked with `-SkipHtml`
- **WHEN** GPO export phase executes
- **THEN** no HTML files SHALL be generated in `GPO/linked/` or `GPO/unlinked/`
- **AND** XML files MAY still be generated unless `-SkipXml` is also specified

#### Scenario: Both skip parameters

- **GIVEN** the script is invoked with both `-SkipXml` and `-SkipHtml`
- **WHEN** the script reaches the GPO export phase
- **THEN** GPO enumeration SHALL be skipped entirely
- **AND** the `GPO/linked/` and `GPO/unlinked/` folders SHALL remain empty

## MODIFIED Requirements

### Requirement: Backward-compatible SkipGpoReports parameter

The existing `-SkipGpoReports` parameter SHALL remain functional and SHALL implicitly set both new skip flags.

#### Scenario: Legacy parameter behavior

- **GIVEN** the script is invoked with `-SkipGpoReports`
- **WHEN** parameter binding occurs
- **THEN** `-SkipXml` SHALL be set to true
- **AND** `-SkipHtml` SHALL be set to true
- **AND** behavior SHALL match pre-change behavior (no GPO exports)

#### Scenario: WMI extraction with SkipGpoReports

- **GIVEN** the script is invoked with `-SkipGpoReports`
- **WHEN** WMI filter information is needed during link enumeration
- **THEN** temporary XML SHALL still be written to `reports/` folder
- **AND** WMI filter extraction SHALL succeed
- **NOTE**: This preserves existing behavior where `-SkipGpoReports` skips permanent exports but not temporary extraction files
