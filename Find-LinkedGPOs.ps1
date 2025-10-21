#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputRoot,

    [Parameter()]
    [switch]$SkipGpoReports,

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

function New-TimestampString {
    (Get-Date).ToString('yyyy-MM-dd-HH-mm')
}

function Get-DomainRootLinks {
    param([Parameter(Mandatory)][string]$DomainDn)
    try {
        $inherit = Get-GPInheritance -Target $DomainDn -ErrorAction Stop
    } catch {
        Write-Log "Get-GPInheritance failed for domain root $DomainDn: $_" 'WARN'
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
        [Parameter()][array]$GpoLinks,
        [Parameter(Mandatory)][string]$ReportsFolder,
        [Parameter()][switch]$SkipGpoReports
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
        if ($guidStr) { $null = $g.SetAttribute('guid', $guidStr) }
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

        $wmi = Get-GpoWmiInfo -GpoId $guidStr -ReportsFolder $ReportsFolder -Domain $Domain -SkipReports:$SkipGpoReports
        if ($wmi.Name -or $wmi.Id -or $wmi.Query) {
            $w = $xml.CreateElement('WmiFilter')
            if ($wmi.Name) { $null = $w.SetAttribute('name', [string]$wmi.Name) }
            if ($wmi.Id) { $null = $w.SetAttribute('id', [string]$wmi.Id) }
            if ($wmi.Query) { $null = $w.SetAttribute('query', [string]$wmi.Query) }
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
    $reports = Join-Path $folder 'reports'
    New-Item -ItemType Directory -Path $reports -Force | Out-Null
    return [pscustomobject]@{
        Root          = $root
        Folder        = $folder
        LogFile       = $logFile
        ReportsFolder = $reports
        IndexXml      = (Join-Path $folder 'linked-gpos.xml')
        XsdPath       = (Join-Path $folder 'validation.xsd')
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

function Export-GpoReport {
    param(
        [Parameter(Mandatory)][string]$GpoId,
        [Parameter(Mandatory)][string]$ReportsFolder,
        [Parameter(Mandatory)][string]$Domain
    )
    try {
        $gpo = Get-GPO -Guid $GpoId -Domain $Domain -ErrorAction Stop
    } catch {
        return $null
    }
    $file = (Sanitize-FileName -Name $gpo.DisplayName) + '.xml'
    $path = Join-Path $ReportsFolder $file
    try { Get-GPOReport -Guid $GpoId -Domain $Domain -ReportType Xml -Path $path | Out-Null } catch {}
    return [pscustomobject]@{ Path=$path; Gpo=$gpo }
}

function Get-WmiFilterQueryFromReport {
    param([Parameter(Mandatory)][string]$ReportPath)
    try { [xml]$doc = Get-Content -LiteralPath $ReportPath -Raw } catch { return $null }
    $node = $doc.SelectSingleNode("//*[translate(local-name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='QUERY']")
    if ($node -and $node.InnerText) { return $node.InnerText.Trim() }
    return $null
}

function Get-GpoWmiInfo {
    param(
        [Parameter(Mandatory)][string]$GpoId,
        [Parameter(Mandatory)][string]$ReportsFolder,
        [Parameter(Mandatory)][string]$Domain,
        [switch]$SkipReports
    )
    $name=$null; $id=$null; $query=$null
    try {
        $gpoObj = Get-GPO -Guid $GpoId -Domain $Domain -ErrorAction Stop
        if ($gpoObj.WmiFilter) { $name = [string]$gpoObj.WmiFilter.Name; $id = [string]$gpoObj.WmiFilter.Id }
    } catch {}
    if (-not $SkipReports) {
        $rep = Export-GpoReport -GpoId $GpoId -ReportsFolder $ReportsFolder -Domain $Domain
        if ($rep -and (Test-Path -LiteralPath $rep.Path)) {
            $q = Get-WmiFilterQueryFromReport -ReportPath $rep.Path
            if ($q) { $query = $q }
        }
    }
    return [pscustomobject]@{ Name=$name; Id=$id; Query=$query }
}

function Get-GpoLinksForOu {
    param([Parameter(Mandatory)][string]$OuDn)
    try {
        $inherit = Get-GPInheritance -Target $OuDn -ErrorAction Stop
    } catch {
        Write-Log "Get-GPInheritance failed for $OuDn: $_" 'WARN'
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
        [Parameter()][array]$GpoLinks,
        [Parameter(Mandatory)][string]$ReportsFolder,
        [Parameter()][switch]$SkipGpoReports
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
        if ($guidStr) { $null = $g.SetAttribute('guid', $guidStr) }
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

        $wmi = Get-GpoWmiInfo -GpoId $guidStr -ReportsFolder $ReportsFolder -Domain $Domain -SkipReports:$SkipGpoReports
        if ($wmi.Name -or $wmi.Id -or $wmi.Query) {
            $w = $xml.CreateElement('WmiFilter')
            if ($wmi.Name) { $null = $w.SetAttribute('name', [string]$wmi.Name) }
            if ($wmi.Id) { $null = $w.SetAttribute('id', [string]$wmi.Id) }
            if ($wmi.Query) { $null = $w.SetAttribute('query', [string]$wmi.Query) }
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
    } catch { Write-Log "Failed to resolve domain $dom: $_" 'WARN'; continue }

    if (-not $ExcludeDomainRoot) {
        $rootLinks = Get-DomainRootLinks -DomainDn $domainDn
        Write-DomainRootXml -OutFolder $out.Folder -Domain $dom -DomainDn $domainDn -GpoLinks $rootLinks -ReportsFolder $out.ReportsFolder -SkipGpoReports:$SkipGpoReports
    }
    $ous = @()
    if ($adAvailable) {
        try {
            $ous = Get-ADOrganizationalUnit -Server $dom -SearchBase $base -SearchScope Subtree -Filter * -Properties gPOptions | Sort-Object DistinguishedName
        } catch { Write-Log "Get-ADOrganizationalUnit failed for $dom: $_" 'WARN' }
    }
    foreach ($ou in $ous) {
        $inheritBlocked = $false
        if ($ou.gPOptions -ne $null) { try { $inheritBlocked = ( ($ou.gPOptions -band 1) -eq 1 ) } catch {} }
        $gpoLinks = Get-GpoLinksForOu -OuDn $ou.DistinguishedName
        Write-PerOuXml -OutFolder $out.Folder -OuDn $ou.DistinguishedName -Domain $dom -InheritanceBlocked:$inheritBlocked -GpoLinks $gpoLinks -ReportsFolder $out.ReportsFolder -SkipGpoReports:$SkipGpoReports
    }
}

Write-IndexXml -Path $out.IndexXml -Domain $hostInfo.Domain

if (-not $NoZip) { Compress-Output -Folder $out.Folder }

Write-Log 'Completed.'
