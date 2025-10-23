#requires -Version 5.1
<#
.SYNOPSIS
    Audits Active Directory Group Policy Object (GPO) links and exports detailed GPO reports.

.DESCRIPTION
    Find-LinkedGPOs enumerates all GPOs in an Active Directory domain and identifies which are linked to OUs, 
    the domain root, or Sites. For each GPO, it exports both XML and HTML reports and organizes them by 
    linked/unlinked status. The script also captures GPO link attributes, security filtering, and WMI filters.

.PARAMETER OutputRoot
    Specifies the root directory for output. If not provided, the script will prompt interactively.
    Default locations: Documents, C:\Temp, or custom path.

.PARAMETER SkipGpoReports
    Skips all GPO report exports (both XML and HTML). Maintained for backward compatibility.
    Equivalent to specifying both -SkipXml and -SkipHtml.

.PARAMETER SkipXml
    Skips XML report generation. HTML reports will still be created unless -SkipHtml is also specified.

.PARAMETER SkipHtml
    Skips HTML report generation. XML reports will still be created unless -SkipXml is also specified.

.PARAMETER Domain
    Specifies one or more domains to enumerate. If not provided, uses the current domain.

.PARAMETER Credential
    Specifies alternate credentials for domain access.

.PARAMETER SearchBase
    Specifies the LDAP search base for OU enumeration. If not provided, searches the entire domain.

.PARAMETER ExcludeDomainRoot
    Excludes the domain root from link enumeration.

.PARAMETER NoZip
    Prevents automatic ZIP compression of the output folder.

.EXAMPLE
    .\Find-LinkedGPOs.ps1
    
    Runs interactively, prompting for output location. Exports all GPOs with both XML and HTML reports.

.EXAMPLE
    .\Find-LinkedGPOs.ps1 -OutputRoot "C:\Temp" -SkipHtml
    
    Exports to C:\Temp with only XML reports (no HTML).

.EXAMPLE
    .\Find-LinkedGPOs.ps1 -Domain "contoso.com","fabrikam.com" -NoZip
    
    Audits multiple domains and skips ZIP compression.

.EXAMPLE
    .\Find-LinkedGPOs.ps1 -SkipGpoReports
    
    Audits GPO links but skips all report generation (backward compatibility mode).

.OUTPUTS
    Creates a timestamped folder structure:
    Find-LinkedGPOs-YYYY-MM-DD-HH-MM/
    ├── GPO/
    │   ├── linked/       (GPOs linked to OUs/domain/sites)
    │   │   ├── GPOName.xml
    │   │   └── GPOName.html
    │   └── unlinked/     (GPOs not linked anywhere)
    │       ├── GPOName.xml
    │       └── GPOName.html
    ├── *.xml files       (Per-OU and domain root link data)
    ├── linked-gpos.xml   (Aggregated index)
    ├── validation.xsd    (XML schema)
    └── Find-LinkedGPOs-YYYY-MM-DD-HH-MM.log

.NOTES
    Version: 2.0
    Requires: ActiveDirectory and GroupPolicy PowerShell modules
    Requires: PowerShell 5.1 or later
    Requires: Domain member or RSAT with appropriate permissions

.LINK
    https://github.com/yourusername/FindLinkedGPOs
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputRoot,

    [Parameter()]
    [switch]$SkipGpoReports,

    [Parameter()]
    [switch]$SkipXml,

    [Parameter()]
    [switch]$SkipHtml,

    [Parameter()]
    [string[]]$Domain,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [string]$SearchBase,

    [Parameter()]
    [switch]$ExcludeDomainRoot,

    [Parameter()]
    [switch]$NoZip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Backward compatibility: -SkipGpoReports sets both new skip flags
if ($SkipGpoReports) {
    $SkipXml = $true
    $SkipHtml = $true
}

# Track linked GPO GUIDs for classification
$script:LinkedGpoGuids = @{}

function New-TimestampString {
    (Get-Date).ToString('yyyy-MM-dd-HH-mm')
}

function Normalize-GpoGuid {
    param([Parameter(Mandatory)][string]$GuidStr)
    $normalized = $GuidStr.Trim()
    if ($normalized -notmatch '^\{.*\}$') {
        $normalized = '{' + $normalized.Trim('{}') + '}'
    }
    return $normalized.ToUpper()
}

function Get-DomainRootLinks {
    param([Parameter(Mandatory)][string]$DomainDn)
    try {
        $inherit = Get-GPInheritance -Target $DomainDn -ErrorAction Stop
    } catch {
        Write-Log "Get-GPInheritance failed for domain root ${DomainDn}: $_" 'WARN'
        return @()
    }
    if ($inherit -and $inherit.GpoLinks) { return @($inherit.GpoLinks) }
    return @()
}

function Write-DomainRootXml {
    param(
        [Parameter(Mandatory)][string]$OutFolder,
        [Parameter(Mandatory)][string]$Domain,
        [Parameter(Mandatory)][string]$DomainDn,
        [Parameter()][array]$GpoLinks
    )
    $xml = New-Object System.Xml.XmlDocument
    $decl = $xml.CreateXmlDeclaration('1.0','UTF-8',$null)
    $xml.AppendChild($decl) | Out-Null

    $root = $xml.CreateElement('FindLinkedGPOs')
    $null = $root.SetAttribute('domain', $Domain)
    $null = $root.SetAttribute('generatedUtc', (Get-Date).ToUniversalTime().ToString('o'))
    $null = $root.SetAttribute('host', $env:COMPUTERNAME)
    $xml.AppendChild($root) | Out-Null

    $dr = $xml.CreateElement('DomainRoot')
    $null = $dr.SetAttribute('dn', $DomainDn)
    $root.AppendChild($dr) | Out-Null

    foreach ($link in ($GpoLinks | Sort-Object Order)) {
        $g = $xml.CreateElement('GpoLink')
        $null = $g.SetAttribute('displayName', [string]$link.DisplayName)
        $guidStr = [string]$link.GPOId
        if (-not $guidStr) { $guidStr = [string]$link.GpoId }
        if ($guidStr -and $guidStr -notmatch '^\{.*\}$') { $guidStr = '{' + $guidStr.Trim('{}') + '}' }
        if ($guidStr) { 
            $null = $g.SetAttribute('guid', $guidStr)
            $normalizedGuid = Normalize-GpoGuid -GuidStr $guidStr
            $script:LinkedGpoGuids[$normalizedGuid] = $true
        }
        $null = $g.SetAttribute('enabled', ([bool]$link.Enabled).ToString().ToLower())
        $null = $g.SetAttribute('enforced', ([bool]$link.Enforced).ToString().ToLower())
        if ($null -ne $link.Order) { $null = $g.SetAttribute('order', [string]$link.Order) }

        $principals = Normalize-GpoPermissions -GpoId $guidStr -Domain $Domain
        if ($principals -and $principals.Count -gt 0) {
            $sec = $xml.CreateElement('SecurityFiltering')
            foreach ($pr in $principals) {
                $pnode = $xml.CreateElement('Principal')
                $null = $pnode.SetAttribute('sid', [string]$pr.Sid)
                if ($pr.Name) { $null = $pnode.SetAttribute('name', [string]$pr.Name) }
                if ($pr.Rights -and $pr.Rights.Count -gt 0) { $null = $pnode.SetAttribute('rights', ($pr.Rights -join ',')) }
                $sec.AppendChild($pnode) | Out-Null
            }
            $g.AppendChild($sec) | Out-Null
        }

        $wmi = Get-GpoWmiInfo -GpoId $guidStr -Domain $Domain
        if ($wmi.Name -or $wmi.Id) {
            $w = $xml.CreateElement('WmiFilter')
            if ($wmi.Name) { $null = $w.SetAttribute('name', [string]$wmi.Name) }
            if ($wmi.Id) { $null = $w.SetAttribute('id', [string]$wmi.Id) }
            $g.AppendChild($w) | Out-Null
        }

        $dr.AppendChild($g) | Out-Null
    }

    $file = "DomainRoot--$Domain-linked-gpos.xml"
    $path = Join-Path $OutFolder $file
    $xml.Save($path)
    Write-Log "Wrote domain root XML: $path"
}
function Sanitize-FileName {
    param([Parameter(Mandatory)][string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $pattern = "[" + [Regex]::Escape($invalid) + "]"
    return ([Regex]::Replace($Name, $pattern, '_'))
}

function Initialize-Output {
    param(
        [string]$OutputRootParam
    )
    $root = $OutputRootParam
    if (-not $root) {
        Write-Host "Select output root:" -ForegroundColor Cyan
        Write-Host "  1) Documents" -ForegroundColor DarkCyan
        Write-Host "  2) C:\\Temp" -ForegroundColor DarkCyan
        Write-Host "  3) Custom" -ForegroundColor DarkCyan
        $choice = Read-Host "Enter 1, 2, or 3"
        switch ($choice) {
            '1' { $root = [Environment]::GetFolderPath('MyDocuments') }
            '2' { $root = 'C:\\Temp' }
            '3' { $root = Read-Host 'Enter full folder path' }
            default { $root = [Environment]::GetFolderPath('MyDocuments') }
        }
    }
    if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root | Out-Null }
    $stamp = New-TimestampString
    $folderName = "Find-LinkedGPOs-$stamp"
    $folder = Join-Path $root $folderName
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $logFile = Join-Path $folder ("$folderName.log")
    $gpoFolder = Join-Path $folder 'GPO'
    New-Item -ItemType Directory -Path $gpoFolder -Force | Out-Null
    $gpoLinkedFolder = Join-Path $gpoFolder 'linked'
    New-Item -ItemType Directory -Path $gpoLinkedFolder -Force | Out-Null
    $gpoUnlinkedFolder = Join-Path $gpoFolder 'unlinked'
    New-Item -ItemType Directory -Path $gpoUnlinkedFolder -Force | Out-Null
    return [pscustomobject]@{
        Root               = $root
        Folder             = $folder
        LogFile            = $logFile
        GpoLinkedFolder    = $gpoLinkedFolder
        GpoUnlinkedFolder  = $gpoUnlinkedFolder
        IndexXml           = (Join-Path $folder 'linked-gpos.xml')
        XsdPath            = (Join-Path $folder 'validation.xsd')
    }
}

$script:LogFile = $null
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = (Get-Date).ToString('s')
    $line = "[$timestamp][$Level] $Message"
    Write-Host $line
    if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value $line }
}

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) { Write-Log "Module $Name not available; continuing with limited functionality" 'WARN'; return $false }
    Import-Module -Name $Name -ErrorAction SilentlyContinue | Out-Null
    return $true
}

function Write-Xsd {
    param([Parameter(Mandatory)][string]$Destination)
    $source = Join-Path $PSScriptRoot 'schema/validation.xsd'
    if (Test-Path -LiteralPath $source) {
        Copy-Item -LiteralPath $source -Destination $Destination -Force
        Write-Log "Copied validation.xsd to output folder"
    } else {
        Write-Log "schema/validation.xsd not found next to script; skipping copy" 'WARN'
    }
}

function Get-CurrentHostInfo {
    $hostName = $env:COMPUTERNAME
    $domainName = $null
    try {
        if (Ensure-Module -Name ActiveDirectory) {
            $domainName = (Get-ADDomain).DNSRoot
        }
    } catch {}
    if (-not $domainName) { $domainName = $env:USERDNSDOMAIN }
    if (-not $domainName) { $domainName = 'unknown.local' }
    return [pscustomobject]@{ Host=$hostName; Domain=$domainName }
}

function Write-IndexXml {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Domain
    )
    $xml = New-Object System.Xml.XmlDocument
    $decl = $xml.CreateXmlDeclaration('1.0','UTF-8',$null)
    $xml.AppendChild($decl) | Out-Null

    $root = $xml.CreateElement('FindLinkedGPOs')
    $null = $root.SetAttribute('domain', $Domain)
    $null = $root.SetAttribute('generatedUtc', (Get-Date).ToUniversalTime().ToString('o'))
    $null = $root.SetAttribute('host', $env:COMPUTERNAME)
    $xml.AppendChild($root) | Out-Null

    $xml.Save($Path)
    Write-Log "Wrote aggregated index XML: $Path"
}

function Compress-Output {
    param(
        [Parameter(Mandatory)][string]$Folder
    )
    $zipPath = "$Folder.zip"
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    Compress-Archive -Path $Folder -DestinationPath $zipPath -Force
    Write-Log "Zipped output to: $zipPath"
}

function Get-TargetDomains {
    param(
        [string[]]$DomainParam,
        [string]$Fallback
    )
    if ($DomainParam -and $DomainParam.Count -gt 0) { return $DomainParam }
    return @($Fallback)
}

function Get-TargetOUs {
    param(
        [Parameter(Mandatory)][string]$Server,
        [string]$SearchBase
    )
    $base = $SearchBase
    if (-not $base) {
        $base = (Get-ADDomain -Server $Server).DistinguishedName
    }
    Get-ADOrganizationalUnit -Server $Server -SearchBase $base -SearchScope Subtree -Filter * -Properties gPOptions | Sort-Object DistinguishedName
}

function Normalize-GpoPermissions {
    param(
        [Parameter(Mandatory)][string]$GpoId,
        [Parameter(Mandatory)][string]$Domain
    )
    $results = @{}
    try {
        $perms = Get-GPPermissions -Guid $GpoId -All -Domain $Domain -ErrorAction Stop
    } catch {
        return @()
    }
    foreach ($p in $perms) {
        $sid = $null; $name = $null
        if ($p.Trustee -and $p.Trustee.PSObject.Properties.Match('Sid').Count -gt 0) {
            $sid = [string]$p.Trustee.Sid
            $name = [string]$p.Trustee.Name
        } else {
            $name = [string]$p.Trustee
            try { $sid = ([System.Security.Principal.NTAccount]$name).Translate([System.Security.Principal.SecurityIdentifier]).Value } catch {}
        }
        if (-not $sid) { continue }
        if (-not $results.ContainsKey($sid)) {
            $results[$sid] = [pscustomobject]@{ Sid=$sid; Name=$name; Rights=@() }
        }
        $rights = @($results[$sid].Rights)
        if ([string]$p.Permission -match 'Apply') { if ('Apply' -notin $rights) { $rights += 'Apply' } }
        if ([string]$p.Permission -match 'Read')  { if ('Read'  -notin $rights) { $rights += 'Read'  } }
        $results[$sid].Rights = $rights
    }
    return $results.Values
}

function Export-GpoDualReport {
    param(
        [Parameter(Mandatory)][string]$GpoId,
        [Parameter(Mandatory)][string]$TargetFolder,
        [Parameter(Mandatory)][string]$Domain,
        [Parameter(Mandatory)][hashtable]$UsedFilenames,
        [Parameter()][bool]$SkipXml,
        [Parameter()][bool]$SkipHtml
    )
    $result = [pscustomobject]@{
        DisplayName = $null
        Guid = $GpoId
        XmlSuccess = $false
        HtmlSuccess = $false
        XmlPath = $null
        HtmlPath = $null
        Error = $null
    }
    
    try {
        $gpo = Get-GPO -Guid $GpoId -Domain $Domain -ErrorAction Stop
        $result.DisplayName = $gpo.DisplayName
        
        $baseFileName = Sanitize-FileName -Name $gpo.DisplayName
        if ($UsedFilenames.ContainsKey($baseFileName)) {
            $baseFileName = "${baseFileName}_${GpoId}"
        }
        $UsedFilenames[$baseFileName] = $true
        
        if (-not $SkipXml) {
            $xmlPath = Join-Path $TargetFolder "${baseFileName}.xml"
            try {
                Get-GPOReport -Guid $GpoId -Domain $Domain -ReportType Xml -Path $xmlPath -ErrorAction Stop | Out-Null
                $result.XmlSuccess = $true
                $result.XmlPath = $xmlPath
            } catch {
                Write-Log "Failed to export XML for GPO $($gpo.DisplayName): $_" 'WARN'
            }
        }
        
        if (-not $SkipHtml) {
            $htmlPath = Join-Path $TargetFolder "${baseFileName}.html"
            try {
                Get-GPOReport -Guid $GpoId -Domain $Domain -ReportType Html -Path $htmlPath -ErrorAction Stop | Out-Null
                $result.HtmlSuccess = $true
                $result.HtmlPath = $htmlPath
            } catch {
                Write-Log "Failed to export HTML for GPO $($gpo.DisplayName): $_" 'WARN'
            }
        }
        
    } catch {
        $result.Error = $_.Exception.Message
        Write-Log "Failed to retrieve GPO ${GpoId}: $_" 'WARN'
    }
    
    return $result
}

function Get-AllGpos {
    param(
        [Parameter(Mandatory)][string]$Domain
    )
    try {
        $gpos = Get-GPO -All -Domain $Domain -ErrorAction Stop
        Write-Log "Found $($gpos.Count) GPOs in domain: $Domain"
        return @($gpos)
    } catch {
        Write-Log "Failed to enumerate GPOs in ${Domain}: $_" 'WARN'
        return @()
    }
}

function Get-GpoWmiInfo {
    param(
        [Parameter(Mandatory)][string]$GpoId,
        [Parameter(Mandatory)][string]$Domain
    )
    $name=$null; $id=$null
    try {
        $gpoObj = Get-GPO -Guid $GpoId -Domain $Domain -ErrorAction Stop
        if ($gpoObj.WmiFilter) { $name = [string]$gpoObj.WmiFilter.Name; $id = [string]$gpoObj.WmiFilter.Id }
    } catch {}
    return [pscustomobject]@{ Name=$name; Id=$id }
}

function Get-GpoLinksForOu {
    param([Parameter(Mandatory)][string]$OuDn)
    try {
        $inherit = Get-GPInheritance -Target $OuDn -ErrorAction Stop
    } catch {
        Write-Log "Get-GPInheritance failed for ${OuDn}: $_" 'WARN'
        return @()
    }
    if ($inherit -and $inherit.GpoLinks) { return @($inherit.GpoLinks) }
    return @()
}

function Write-PerOuXml {
    param(
        [Parameter(Mandatory)][string]$OutFolder,
        [Parameter(Mandatory)][string]$OuDn,
        [Parameter(Mandatory)][string]$Domain,
        [Parameter()][bool]$InheritanceBlocked,
        [Parameter()][array]$GpoLinks
    )
    $xml = New-Object System.Xml.XmlDocument
    $decl = $xml.CreateXmlDeclaration('1.0','UTF-8',$null)
    $xml.AppendChild($decl) | Out-Null

    $root = $xml.CreateElement('FindLinkedGPOs')
    $null = $root.SetAttribute('domain', $Domain)
    $null = $root.SetAttribute('generatedUtc', (Get-Date).ToUniversalTime().ToString('o'))
    $null = $root.SetAttribute('host', $env:COMPUTERNAME)
    $xml.AppendChild($root) | Out-Null

    $ou = $xml.CreateElement('OU')
    $null = $ou.SetAttribute('dn', $OuDn)
    if ($InheritanceBlocked) { $null = $ou.SetAttribute('inheritanceBlocked','true') }
    $root.AppendChild($ou) | Out-Null

    foreach ($link in ($GpoLinks | Sort-Object Order)) {
        $g = $xml.CreateElement('GpoLink')
        $null = $g.SetAttribute('displayName', [string]$link.DisplayName)
        $guidStr = [string]$link.GPOId
        if (-not $guidStr) { $guidStr = [string]$link.GpoId }
        if ($guidStr -and $guidStr -notmatch '^\{.*\}$') { $guidStr = '{' + $guidStr.Trim('{}') + '}' }
        if ($guidStr) { 
            $null = $g.SetAttribute('guid', $guidStr)
            $normalizedGuid = Normalize-GpoGuid -GuidStr $guidStr
            $script:LinkedGpoGuids[$normalizedGuid] = $true
        }
        $null = $g.SetAttribute('enabled', ([bool]$link.Enabled).ToString().ToLower())
        $null = $g.SetAttribute('enforced', ([bool]$link.Enforced).ToString().ToLower())
        if ($null -ne $link.Order) { $null = $g.SetAttribute('order', [string]$link.Order) }

        $principals = Normalize-GpoPermissions -GpoId $guidStr -Domain $Domain
        if ($principals -and $principals.Count -gt 0) {
            $sec = $xml.CreateElement('SecurityFiltering')
            foreach ($pr in $principals) {
                $pnode = $xml.CreateElement('Principal')
                $null = $pnode.SetAttribute('sid', [string]$pr.Sid)
                if ($pr.Name) { $null = $pnode.SetAttribute('name', [string]$pr.Name) }
                if ($pr.Rights -and $pr.Rights.Count -gt 0) { $null = $pnode.SetAttribute('rights', ($pr.Rights -join ',')) }
                $sec.AppendChild($pnode) | Out-Null
            }
            $g.AppendChild($sec) | Out-Null
        }

        $wmi = Get-GpoWmiInfo -GpoId $guidStr -Domain $Domain
        if ($wmi.Name -or $wmi.Id) {
            $w = $xml.CreateElement('WmiFilter')
            if ($wmi.Name) { $null = $w.SetAttribute('name', [string]$wmi.Name) }
            if ($wmi.Id) { $null = $w.SetAttribute('id', [string]$wmi.Id) }
            $g.AppendChild($w) | Out-Null
        }

        $ou.AppendChild($g) | Out-Null
    }

    $file = (Sanitize-FileName -Name $OuDn) + "--$Domain-linked-gpos.xml"
    $path = Join-Path $OutFolder $file
    $xml.Save($path)
    Write-Log "Wrote per-OU XML: $path"
}

# --- Main ---
$out = Initialize-Output -OutputRootParam $OutputRoot
$script:LogFile = $out.LogFile
Write-Log "Output folder: $($out.Folder)"

$hostInfo = Get-CurrentHostInfo
Write-Xsd -Destination $out.XsdPath

$adAvailable = Ensure-Module -Name ActiveDirectory
$gpAvailable = Ensure-Module -Name GroupPolicy
if (-not $gpAvailable) {
    Write-Log "GroupPolicy module not available; cannot enumerate links." 'ERROR'
    if (-not $NoZip) { Compress-Output -Folder $out.Folder }
    Write-Log 'Completed.'
    return
}

$domains = Get-TargetDomains -DomainParam $Domain -Fallback $hostInfo.Domain
foreach ($dom in $domains) {
    Write-Log "Enumerating domain root and OUs in domain: $dom"
    $base = $null
    $domainDn = $null
    try {
        $domainDn = (Get-ADDomain -Server $dom).DistinguishedName
        $base = if ($SearchBase) { $SearchBase } else { $domainDn }
    } catch { Write-Log "Failed to resolve domain ${dom}: $_" 'WARN'; continue }

    if (-not $ExcludeDomainRoot) {
        $rootLinks = Get-DomainRootLinks -DomainDn $domainDn
        Write-DomainRootXml -OutFolder $out.Folder -Domain $dom -DomainDn $domainDn -GpoLinks $rootLinks
    }
    $ous = @()
    if ($adAvailable) {
        try {
            $ous = Get-ADOrganizationalUnit -Server $dom -SearchBase $base -SearchScope Subtree -Filter * -Properties gPOptions | Sort-Object DistinguishedName
        } catch { Write-Log "Get-ADOrganizationalUnit failed for ${dom}: $_" 'WARN' }
    }
    foreach ($ou in $ous) {
        $inheritBlocked = $false
        if ($ou.gPOptions -ne $null) { try { $inheritBlocked = ( ($ou.gPOptions -band 1) -eq 1 ) } catch {} }
        $gpoLinks = Get-GpoLinksForOu -OuDn $ou.DistinguishedName
        Write-PerOuXml -OutFolder $out.Folder -OuDn $ou.DistinguishedName -Domain $dom -InheritanceBlocked:$inheritBlocked -GpoLinks $gpoLinks
    }
}

if (-not ($SkipXml -and $SkipHtml)) {
    Write-Log "=== Starting GPO Export Phase ==="
    $usedFilenames = @{}
    $totalLinked = 0
    $totalUnlinked = 0
    $totalXmlSuccess = 0
    $totalHtmlSuccess = 0
    $totalFailed = 0
    
    foreach ($dom in $domains) {
        Write-Log "Exporting all GPOs for domain: $dom"
        $allGpos = Get-AllGpos -Domain $dom
        
        if ($allGpos.Count -eq 0) {
            Write-Log "No GPOs found in domain $dom" 'WARN'
            continue
        }
        
        foreach ($gpo in $allGpos) {
            $normalizedGuid = Normalize-GpoGuid -GuidStr $gpo.Id
            $isLinked = $script:LinkedGpoGuids.ContainsKey($normalizedGuid)
            $targetFolder = if ($isLinked) { $out.GpoLinkedFolder } else { $out.GpoUnlinkedFolder }
            $classification = if ($isLinked) { "linked" } else { "unlinked" }
            
            Write-Log "Processing GPO: $($gpo.DisplayName) ($($gpo.Id)) - $classification"
            
            $result = Export-GpoDualReport -GpoId $gpo.Id -TargetFolder $targetFolder -Domain $dom -UsedFilenames $usedFilenames -SkipXml $SkipXml -SkipHtml $SkipHtml
            
            if ($isLinked) { $totalLinked++ } else { $totalUnlinked++ }
            if ($result.XmlSuccess) { $totalXmlSuccess++ }
            if ($result.HtmlSuccess) { $totalHtmlSuccess++ }
            if (-not $result.XmlSuccess -and -not $result.HtmlSuccess) { $totalFailed++ }
        }
    }
    
    Write-Log "=== GPO Export Summary ==="
    Write-Log "Total GPOs processed: $($totalLinked + $totalUnlinked)"
    Write-Log "Linked GPOs: $totalLinked"
    Write-Log "Unlinked GPOs: $totalUnlinked"
    Write-Log "Successful XML exports: $totalXmlSuccess"
    Write-Log "Successful HTML exports: $totalHtmlSuccess"
    Write-Log "Failed exports: $totalFailed"
} else {
    Write-Log "Skipping GPO export phase (both -SkipXml and -SkipHtml specified)"
}

Write-IndexXml -Path $out.IndexXml -Domain $hostInfo.Domain

if (-not $NoZip) { Compress-Output -Folder $out.Folder }

Write-Log 'Completed.'
