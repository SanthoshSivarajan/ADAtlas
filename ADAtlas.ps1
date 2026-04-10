<#
================================================================================
 ADAtlas -- Map Every Corner of Your Active Directory
 Version : 1.0
 Author  : Santhosh Sivarajan, Microsoft MVP
 LinkedIn: https://www.linkedin.com/in/sivarajan/
 GitHub  : https://github.com/SanthoshSivarajan/ADAtlas
 License : MIT -- Free to use, modify, and distribute.
--------------------------------------------------------------------------------
 Purpose : Picture-only documentation of an Active Directory forest topology.
           Collects forest, domains, DCs, sites, subnets, site links, trusts,
           replication connection objects, Entra Connect, DNS, NTP hierarchy,
           Exchange, Certificate Services, and Authentication configuration.
           Generates a self-contained interactive HTML map.
           No health checks, no analysis -- just the current configuration.
================================================================================
#>

#Requires -Modules ActiveDirectory

$ReportDate = Get-Date -Format "yyyy-MM-dd_HHmmss"
$OutputHtml = Join-Path $PSScriptRoot "ADAtlas_$ReportDate.html"

Write-Host ""
Write-Host " +============================================================+" -ForegroundColor Cyan
Write-Host " |                                                            |" -ForegroundColor Cyan
Write-Host " |   ADAtlas -- Map Every Corner of Your Active Directory     |" -ForegroundColor Cyan
Write-Host " |   Version 1.0                                              |" -ForegroundColor Cyan
Write-Host " |                                                            |" -ForegroundColor Cyan
Write-Host " |   Author   : Santhosh Sivarajan, Microsoft MVP             |" -ForegroundColor Cyan
Write-Host " |   LinkedIn : https://www.linkedin.com/in/sivarajan/        |" -ForegroundColor Cyan
Write-Host " |   GitHub   : https://github.com/SanthoshSivarajan/ADAtlas  |" -ForegroundColor Cyan
Write-Host " |                                                            |" -ForegroundColor Cyan
Write-Host " +============================================================+" -ForegroundColor Cyan
Write-Host ""
Write-Host " [*] Running As  : $($env:USERDOMAIN)\$($env:USERNAME)" -ForegroundColor White
Write-Host " [*] Host        : $env:COMPUTERNAME" -ForegroundColor White
Write-Host " [*] PowerShell  : $($PSVersionTable.PSVersion)" -ForegroundColor White
Write-Host " [*] Timestamp   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host ""
Write-Host " Starting AD topology collection ..." -ForegroundColor Yellow
Write-Host ""

# ==============================================================================
# HELPERS
# ==============================================================================

$FuncLevelMap = @{
    '0'='Windows 2000'; '1'='Windows Server 2003 Interim'; '2'='Windows Server 2003'
    '3'='Windows Server 2008'; '4'='Windows Server 2008 R2'; '5'='Windows Server 2012'
    '6'='Windows Server 2012 R2'; '7'='Windows Server 2016'; '8'='Windows Server 2016'
    '9'='Windows Server 2016'; '10'='Windows Server 2025'
    'Windows2000Domain'='Windows 2000'; 'Windows2000Forest'='Windows 2000'
    'Windows2003Domain'='Windows Server 2003'; 'Windows2003Forest'='Windows Server 2003'
    'Windows2003InterimDomain'='Windows Server 2003 Interim'
    'Windows2008Domain'='Windows Server 2008'; 'Windows2008Forest'='Windows Server 2008'
    'Windows2008R2Domain'='Windows Server 2008 R2'; 'Windows2008R2Forest'='Windows Server 2008 R2'
    'Windows2012Domain'='Windows Server 2012'; 'Windows2012Forest'='Windows Server 2012'
    'Windows2012R2Domain'='Windows Server 2012 R2'; 'Windows2012R2Forest'='Windows Server 2012 R2'
    'Windows2016Domain'='Windows Server 2016'; 'Windows2016Forest'='Windows Server 2016'
}

function Get-FriendlyFuncLevel($level) {
    $s = [string]$level
    if ($FuncLevelMap.ContainsKey($s)) { return $FuncLevelMap[$s] }
    return $s
}

$SchemaVersionMap = @{
    13='Windows 2000'; 30='Windows Server 2003'; 31='Windows Server 2003 R2'
    44='Windows Server 2008'; 47='Windows Server 2008 R2'; 56='Windows Server 2012'
    69='Windows Server 2012 R2'; 87='Windows Server 2016'; 88='Windows Server 2019/2022'
    89='Windows Server 2022 (23H2)'; 90='Windows Server 2025'; 91='Windows Server 2025'
}

$ExchangeSchemaMap = @{
    4397='Exchange 2000 SP3'; 6870='Exchange 2003 SP3'
    14625='Exchange 2007 SP3'; 14734='Exchange 2010 SP3'
    15137='Exchange 2013 RTM'; 15254='Exchange 2013 CU1'
    15281='Exchange 2013 CU2'; 15283='Exchange 2013 CU3'
    15292='Exchange 2013 SP1/CU4'; 15300='Exchange 2013 CU5'
    15303='Exchange 2013 CU6'; 15312='Exchange 2013 CU7-CU23'
    15317='Exchange 2016 CU1'; 15323='Exchange 2016 CU2'
    15325='Exchange 2016 CU3-CU4'; 15326='Exchange 2016 CU5-CU6'
    15330='Exchange 2016 CU7-CU9'; 15332='Exchange 2016 CU10-CU18'
    15333='Exchange 2016 CU19'; 15334='Exchange 2016 CU20+'
    15349='Exchange 2016 CU23'
    17000='Exchange 2019 CU1'; 17001='Exchange 2019 CU2-CU11'
    17002='Exchange 2019 CU12-CU13'; 17003='Exchange 2019 CU14+'
    17005='Exchange Server SE CU1'
}

$Atlas = [ordered]@{
    GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    GeneratedBy = "$($env:USERDOMAIN)\$($env:USERNAME)"
    GeneratedOn = $env:COMPUTERNAME
    Forest = $null; Domains = @(); DCs = @(); Sites = @(); Subnets = @()
    SiteLinks = @(); SiteLinkBridges = @(); Trusts = @(); Connections = @()
    EntraConnect = $null; DNS = $null; Exchange = $null; PKI = $null; Authentication = $null; NTP = $null
}

# ==============================================================================
# FOREST-LEVEL COLLECTION
# ==============================================================================

try {
    $Forest = Get-ADForest -ErrorAction Stop
    Write-Host " [+] Forest: $($Forest.Name)" -ForegroundColor Green
    $RootDSE = Get-ADRootDSE
    $ConfigDN = $RootDSE.configurationNamingContext
    $SchemaDN = $RootDSE.schemaNamingContext
    $SchemaObj = Get-ADObject $SchemaDN -Property objectVersion -ErrorAction SilentlyContinue
    $SchemaVersion = if ($SchemaObj) { $SchemaObj.objectVersion } else { 0 }
    $SchemaOS = if ($SchemaVersionMap.ContainsKey([int]$SchemaVersion)) { $SchemaVersionMap[[int]$SchemaVersion] } else { "Version $SchemaVersion" }
    $DSConfig = Get-ADObject "CN=Directory Service,CN=Windows NT,CN=Services,$ConfigDN" -Properties tombstoneLifetime, garbageCollPeriod -ErrorAction SilentlyContinue
    $TombstoneLife = if ($DSConfig -and $DSConfig.tombstoneLifetime) { $DSConfig.tombstoneLifetime } else { 60 }

    $Atlas.Forest = [ordered]@{
        Name=$Forest.Name; RootDomain=$Forest.RootDomain
        ForestMode=[string]$Forest.ForestMode
        ForestModeDisplay=Get-FriendlyFuncLevel $Forest.ForestMode
        SchemaMaster=$Forest.SchemaMaster; DomainNamingMaster=$Forest.DomainNamingMaster
        SchemaVersion=$SchemaVersion; SchemaOS=$SchemaOS
        TombstoneLifetime=$TombstoneLife
        GlobalCatalogs=@($Forest.GlobalCatalogs | ForEach-Object { [string]$_ })
        UPNSuffixes=@($Forest.UPNSuffixes | ForEach-Object { [string]$_ })
        SPNSuffixes=@($Forest.SPNSuffixes | ForEach-Object { [string]$_ })
        DomainCount=@($Forest.Domains).Count
    }
    Write-Host " [+] Schema $SchemaVersion ($SchemaOS), forest level $($Atlas.Forest.ForestModeDisplay)" -ForegroundColor Green
}
catch {
    Write-Host " [!] Failed to read forest: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ==============================================================================
# ENTRA CONNECT DETECTION (improved description parsing for server + tenant)
# ==============================================================================

Write-Host ""
Write-Host " Checking for Entra Connect ..." -ForegroundColor Yellow

$entraDetected = $false
$entraAccounts = @()
$entraTenant = ''

try {
    $msol = Get-ADUser -Filter 'SamAccountName -like "MSOL_*"' -Properties WhenCreated, Description, Enabled -ErrorAction SilentlyContinue
    foreach ($m in $msol) {
        $serverName = ''; $tenant = ''
        if ($m.Description -match 'running on computer (\S+)') { $serverName = $Matches[1].TrimEnd('.,') }
        elseif ($m.Description -match 'installed on (\S+)') { $serverName = $Matches[1].TrimEnd('.,') }
        if ($m.Description -match 'synchronize to tenant (\S+)') { $tenant = $Matches[1].TrimEnd('.,') }
        elseif ($m.Description -match 'tenant (\S+\.onmicrosoft\.com)') { $tenant = $Matches[1].TrimEnd('.,') }
        if ($tenant -and -not $entraTenant) { $entraTenant = $tenant }
        $entraAccounts += [ordered]@{
            DetectionMethod='MSOL Service Account'; ServiceAccount=[string]$m.SamAccountName
            ServerName=$serverName; Tenant=$tenant; Enabled=[bool]$m.Enabled
            Created=if($m.WhenCreated){$m.WhenCreated.ToString('yyyy-MM-dd')}else{''}
            Description=[string]$m.Description
        }
    }
    $sync = Get-ADUser -Filter 'SamAccountName -like "AAD_*" -or SamAccountName -like "Sync_*"' -Properties WhenCreated, Description, Enabled -ErrorAction SilentlyContinue
    foreach ($s in $sync) {
        $serverName = ''; $tenant = ''
        if ($s.SamAccountName -match 'Sync_(.+?)_') { $serverName = $Matches[1] }
        elseif ($s.Description -match 'running on computer (\S+)') { $serverName = $Matches[1].TrimEnd('.,') }
        elseif ($s.Description -match 'installed on (\S+)') { $serverName = $Matches[1].TrimEnd('.,') }
        if ($s.Description -match 'synchronize to tenant (\S+)') { $tenant = $Matches[1].TrimEnd('.,') }
        if ($tenant -and -not $entraTenant) { $entraTenant = $tenant }
        $entraAccounts += [ordered]@{
            DetectionMethod='Sync Service Account'; ServiceAccount=[string]$s.SamAccountName
            ServerName=$serverName; Tenant=$tenant; Enabled=[bool]$s.Enabled
            Created=if($s.WhenCreated){$s.WhenCreated.ToString('yyyy-MM-dd')}else{''}
            Description=[string]$s.Description
        }
    }
    if ($entraAccounts.Count -gt 0) {
        $entraDetected = $true
        Write-Host " [+] Entra Connect detected ($($entraAccounts.Count) service account(s))" -ForegroundColor Green
        if ($entraTenant) { Write-Host "     Tenant: $entraTenant" -ForegroundColor Gray }
    } else {
        Write-Host " [i] No Entra Connect detected" -ForegroundColor Gray
    }
}
catch {
    Write-Host " [i] Could not check for Entra Connect: $($_.Exception.Message)" -ForegroundColor Gray
}

$entraServerName = ''
foreach ($a in $entraAccounts) { if ($a.ServerName) { $entraServerName = $a.ServerName; break } }

$Atlas.EntraConnect = [ordered]@{
    Detected=$entraDetected; ServerName=$entraServerName; Tenant=$entraTenant; Accounts=$entraAccounts
}

# ==============================================================================
# SITES, SUBNETS, SITE LINKS, BRIDGES
# ==============================================================================

Write-Host ""
Write-Host " Collecting sites, subnets, site links ..." -ForegroundColor Yellow

try {
    $Sites = Get-ADReplicationSite -Filter * -Properties Description, Location, WhenCreated -ErrorAction Stop
    foreach ($s in $Sites) {
        $Atlas.Sites += [ordered]@{
            Name=$s.Name; Description=[string]$s.Description; Location=[string]$s.Location
            DN=$s.DistinguishedName
            WhenCreated=if($s.WhenCreated){$s.WhenCreated.ToString('yyyy-MM-dd')}else{''}
        }
    }
    Write-Host " [+] $($Atlas.Sites.Count) site(s)" -ForegroundColor Green
} catch { Write-Host " [i] Sites: $($_.Exception.Message)" -ForegroundColor Gray }

try {
    $Subnets = Get-ADReplicationSubnet -Filter * -Properties Description, Location -ErrorAction Stop
    foreach ($sn in $Subnets) {
        $siteName = ''
        if ($sn.Site) { $siteName = ($sn.Site -split ',')[0] -replace '^CN=','' }
        $Atlas.Subnets += [ordered]@{
            Name=$sn.Name; Site=$siteName; Location=[string]$sn.Location; Description=[string]$sn.Description
        }
    }
    Write-Host " [+] $($Atlas.Subnets.Count) subnet(s)" -ForegroundColor Green
} catch { Write-Host " [i] Subnets: $($_.Exception.Message)" -ForegroundColor Gray }

try {
    $SiteLinks = Get-ADReplicationSiteLink -Filter * -Properties Description, Options, InterSiteTransportProtocol -ErrorAction Stop
    foreach ($sl in $SiteLinks) {
        $siteList = @()
        foreach ($sdn in $sl.SitesIncluded) { $siteList += (($sdn -split ',')[0] -replace '^CN=','') }
        $Atlas.SiteLinks += [ordered]@{
            Name=$sl.Name; Cost=[int]$sl.Cost; Frequency=[int]$sl.ReplicationFrequencyInMinutes
            Sites=$siteList; Transport=[string]$sl.InterSiteTransportProtocol
            Options=[string]$sl.Options; Description=[string]$sl.Description
        }
    }
    Write-Host " [+] $($Atlas.SiteLinks.Count) site link(s)" -ForegroundColor Green
} catch { Write-Host " [i] Site links: $($_.Exception.Message)" -ForegroundColor Gray }

try {
    $Bridges = Get-ADReplicationSiteLinkBridge -Filter * -ErrorAction SilentlyContinue
    foreach ($b in $Bridges) {
        $linkList = @()
        foreach ($ldn in $b.SiteLinksIncluded) { $linkList += (($ldn -split ',')[0] -replace '^CN=','') }
        $Atlas.SiteLinkBridges += [ordered]@{ Name=$b.Name; SiteLinks=$linkList }
    }
    Write-Host " [+] $($Atlas.SiteLinkBridges.Count) site link bridge(s)" -ForegroundColor Green
} catch {}

# ==============================================================================
# REPLICATION CONNECTIONS
# ==============================================================================

Write-Host ""
Write-Host " Collecting replication connections ..." -ForegroundColor Yellow

try {
    $Connections = Get-ADReplicationConnection -Filter * -ErrorAction Stop
    foreach ($c in $Connections) {
        $fromName = ''; $toName = ''
        if ($c.ReplicateFromDirectoryServer) {
            $parts = $c.ReplicateFromDirectoryServer -split ','
            if ($parts[0] -match '^CN=NTDS Settings') { $fromName = $parts[1] -replace '^CN=','' }
            else { $fromName = $parts[0] -replace '^CN=','' }
        }
        if ($c.ReplicateToDirectoryServer) {
            $parts = $c.ReplicateToDirectoryServer -split ','
            if ($parts[0] -match '^CN=NTDS Settings') { $toName = $parts[1] -replace '^CN=','' }
            else { $toName = $parts[0] -replace '^CN=','' }
        }
        $siteName = ''
        if ($c.DistinguishedName -match 'CN=Sites,') {
            $dnParts = $c.DistinguishedName -split ','
            for ($i = 0; $i -lt $dnParts.Count; $i++) {
                if ($dnParts[$i] -match '^CN=Servers$' -and $i -gt 0) {
                    $siteName = $dnParts[$i + 1] -replace '^CN=',''; break
                }
            }
        }
        $Atlas.Connections += [ordered]@{
            Name=$c.Name; From=$fromName; To=$toName; Site=$siteName
            AutoGenerated=[bool]$c.AutoGenerated; Enabled=[bool]$c.Enabled
        }
    }
    Write-Host " [+] $($Atlas.Connections.Count) connection object(s)" -ForegroundColor Green
} catch { Write-Host " [i] Connections: $($_.Exception.Message)" -ForegroundColor Gray }

# ==============================================================================
# PER-DOMAIN COLLECTION
# ==============================================================================

Write-Host ""
Write-Host " Enumerating $($Forest.Domains.Count) domain(s) ..." -ForegroundColor Yellow

foreach ($domName in $Forest.Domains) {
    Write-Host ""
    Write-Host " [*] Domain: $domName" -ForegroundColor Yellow
    try {
        $dom = Get-ADDomain -Identity $domName -Server $domName -ErrorAction Stop
        $domainObj = [ordered]@{
            DNSRoot=$dom.DNSRoot; NetBIOSName=$dom.NetBIOSName
            DomainMode=[string]$dom.DomainMode
            DomainModeDisplay=Get-FriendlyFuncLevel $dom.DomainMode
            DistinguishedName=$dom.DistinguishedName
            ParentDomain=if($dom.ParentDomain){[string]$dom.ParentDomain}else{''}
            ChildDomains=@($dom.ChildDomains | ForEach-Object { [string]$_ })
            PDCEmulator=$dom.PDCEmulator; RIDMaster=$dom.RIDMaster
            InfrastructureMaster=$dom.InfrastructureMaster
            DCCount=0; IsForestRoot=($dom.DNSRoot -eq $Forest.RootDomain)
        }
        $dcCount = 0
        try {
            $domDCs = Get-ADDomainController -Filter * -Server $domName -ErrorAction Stop
            foreach ($dc in $domDCs) {
                $Atlas.DCs += [ordered]@{
                    Name=$dc.Name; HostName=$dc.HostName; Domain=$domName
                    IPv4Address=[string]$dc.IPv4Address
                    OperatingSystem=[string]$dc.OperatingSystem
                    OSVersion=[string]$dc.OperatingSystemVersion
                    Site=[string]$dc.Site
                    Type=if($dc.IsReadOnly){'RODC'}else{'RWDC'}
                    IsGlobalCatalog=[bool]$dc.IsGlobalCatalog
                    FSMORoles=@($dc.OperationMasterRoles | ForEach-Object { [string]$_ })
                    Enabled=[bool]$dc.Enabled
                }
                $dcCount++
            }
            Write-Host "     DCs: $dcCount" -ForegroundColor Gray
        } catch { Write-Host "     [i] DCs: $($_.Exception.Message)" -ForegroundColor Gray }
        $domainObj.DCCount = $dcCount

        try {
            $domTrusts = Get-ADTrust -Filter * -Server $domName -ErrorAction SilentlyContinue
            foreach ($t in $domTrusts) {
                $Atlas.Trusts += [ordered]@{
                    SourceDomain=$domName; TargetDomain=[string]$t.Name
                    TargetDN=[string]$t.Target; Direction=[string]$t.Direction
                    TrustType=[string]$t.TrustType; Transitive=(-not $t.DisallowTransivity)
                    SelectiveAuth=[bool]$t.SelectiveAuthentication
                    SIDFilter=[bool]$t.SIDFilteringQuarantined
                    IntraForest=[bool]$t.IntraForest
                    ForestTransitive=[bool]$t.ForestTransitive
                    UplevelOnly=[bool]$t.UplevelOnly
                    WhenCreated=if($t.WhenCreated){$t.WhenCreated.ToString('yyyy-MM-dd')}else{''}
                }
            }
            Write-Host "     Trusts: $(@($domTrusts).Count)" -ForegroundColor Gray
        } catch { Write-Host "     [i] Trusts: $($_.Exception.Message)" -ForegroundColor Gray }

        $Atlas.Domains += $domainObj
        Write-Host " [+] $domName ($($domainObj.DomainModeDisplay)) -- $dcCount DC(s)" -ForegroundColor Green
    } catch {
        Write-Host " [!] Could not reach domain $domName : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Dedupe trusts
$seenTrusts = @{}
$uniqueTrusts = @()
foreach ($t in $Atlas.Trusts) {
    $k1 = "$($t.SourceDomain)=>$($t.TargetDomain)"
    $k2 = "$($t.TargetDomain)=>$($t.SourceDomain)"
    if (-not $seenTrusts.ContainsKey($k1) -and -not $seenTrusts.ContainsKey($k2)) {
        $seenTrusts[$k1] = $true
        $uniqueTrusts += $t
    }
}
$Atlas.Trusts = $uniqueTrusts

# ==============================================================================
# DNS COLLECTION (query DNS from PDC fallback chain)
# ==============================================================================

Write-Host ""
Write-Host " Collecting DNS configuration ..." -ForegroundColor Yellow

$dnsZones = @()
$dnsForwarders = @()
$dnsServerName = ''
$dnsHosts = @()

try {
    if (Get-Module -ListAvailable -Name DnsServer) {
        Import-Module DnsServer -ErrorAction SilentlyContinue
        # Build fallback chain: forest root PDC first, then all DCs
        $dnsTargets = @()
        if ($Atlas.Forest.SchemaMaster) { $dnsTargets += $Atlas.Forest.SchemaMaster }
        foreach ($d in $Atlas.Domains) { if ($d.PDCEmulator -and $dnsTargets -notcontains $d.PDCEmulator) { $dnsTargets += $d.PDCEmulator } }
        foreach ($dc in $Atlas.DCs) { if ($dc.HostName -and $dnsTargets -notcontains $dc.HostName) { $dnsTargets += $dc.HostName } }

        foreach ($target in $dnsTargets) {
            try {
                $zones = Get-DnsServerZone -ComputerName $target -ErrorAction Stop
                $dnsServerName = $target
                foreach ($z in $zones) {
                    $dnsZones += [ordered]@{
                        ZoneName=[string]$z.ZoneName
                        ZoneType=[string]$z.ZoneType
                        IsReverseLookupZone=[bool]$z.IsReverseLookupZone
                        IsDsIntegrated=[bool]$z.IsDsIntegrated
                        IsAutoCreated=[bool]$z.IsAutoCreated
                        ReplicationScope=[string]$z.ReplicationScope
                        DynamicUpdate=[string]$z.DynamicUpdate
                    }
                }
                Write-Host " [+] $($zones.Count) DNS zone(s) collected from $target" -ForegroundColor Green
                # Forwarders
                try {
                    $fwd = Get-DnsServerForwarder -ComputerName $target -ErrorAction SilentlyContinue
                    if ($fwd -and $fwd.IPAddress) {
                        $dnsForwarders = @($fwd.IPAddress | ForEach-Object { [string]$_ })
                        Write-Host " [+] $($dnsForwarders.Count) forwarder(s)" -ForegroundColor Green
                    }
                } catch {}
                break
            } catch { continue }
        }
        if (-not $dnsServerName) { Write-Host " [i] No DNS server reachable" -ForegroundColor Gray }
    } else {
        Write-Host " [i] DnsServer module not available" -ForegroundColor Gray
    }
} catch { Write-Host " [i] DNS error: $($_.Exception.Message)" -ForegroundColor Gray }

# DCs that ARE DNS servers (heuristic: any DC, since DNS is usually co-located on AD-integrated zones)
# We can't probe each DC for the role without remoting, so we mark all DCs as potential DNS hosts
# and indicate the data was collected from one of them
foreach ($dc in $Atlas.DCs) {
    $dnsHosts += [ordered]@{
        Name=$dc.Name; HostName=$dc.HostName; Domain=$dc.Domain; Site=$dc.Site
        IPv4Address=$dc.IPv4Address; IsQueriedSource=($dc.HostName -eq $dnsServerName)
    }
}

$Atlas.DNS = [ordered]@{
    Available=($dnsZones.Count -gt 0)
    QueriedFrom=$dnsServerName
    Zones=$dnsZones
    Forwarders=$dnsForwarders
    DCHosts=$dnsHosts
}

# ==============================================================================
# EXCHANGE DETECTION (read-only from AD config partition)
# ==============================================================================

Write-Host ""
Write-Host " Checking for Exchange ..." -ForegroundColor Yellow

$exchangeDetected = $false
$exchangeOrgName = ''
$exchangeSchemaVer = 0
$exchangeServers = @()
$exchangeDAGs = @()
$exchangeAcceptedDomains = @()
$exchangeHybrid = $false

try {
    # Schema version
    $exSchemaObj = Get-ADObject -Identity "CN=ms-Exch-Schema-Version-Pt,$SchemaDN" -Properties rangeUpper -ErrorAction SilentlyContinue
    if ($exSchemaObj) {
        $exchangeSchemaVer = [int]$exSchemaObj.rangeUpper
        $exchangeDetected = $true
    }

    # Exchange organization
    $exOrgContainer = Get-ADObject -SearchBase "CN=Services,$ConfigDN" -LDAPFilter '(objectClass=msExchOrganizationContainer)' -ErrorAction SilentlyContinue
    if ($exOrgContainer) {
        $exchangeDetected = $true
        $exchangeOrgName = $exOrgContainer.Name
    }

    if ($exchangeDetected) {
        Write-Host " [+] Exchange detected (schema $exchangeSchemaVer)" -ForegroundColor Green

        # Exchange servers
        try {
            $exServers = Get-ADObject -SearchBase "CN=Services,$ConfigDN" -LDAPFilter '(objectClass=msExchExchangeServer)' -Properties serialNumber, msExchCurrentServerRoles, networkAddress, msExchProductID, WhenCreated -ErrorAction SilentlyContinue
            foreach ($s in $exServers) {
                $serverSite = ''
                if ($s.DistinguishedName -match 'CN=([^,]+),CN=Servers,CN=Exchange Administrative Group') {
                    # name only
                }
                # Decode roles bitmask
                $roles = @()
                if ($s.msExchCurrentServerRoles) {
                    $r = [int]$s.msExchCurrentServerRoles
                    if ($r -band 2)  { $roles += 'Mailbox' }
                    if ($r -band 4)  { $roles += 'ClientAccess' }
                    if ($r -band 16) { $roles += 'UnifiedMessaging' }
                    if ($r -band 32) { $roles += 'HubTransport' }
                    if ($r -band 64) { $roles += 'EdgeTransport' }
                }
                $exchangeServers += [ordered]@{
                    Name=$s.Name
                    Version=if($s.serialNumber){[string]($s.serialNumber | Select-Object -First 1)}else{''}
                    Roles=$roles
                    Created=if($s.WhenCreated){$s.WhenCreated.ToString('yyyy-MM-dd')}else{''}
                }
            }
            Write-Host " [+] $($exchangeServers.Count) Exchange server(s)" -ForegroundColor Green
        } catch { Write-Host " [i] Exchange servers: $($_.Exception.Message)" -ForegroundColor Gray }

        # DAGs
        try {
            $dags = Get-ADObject -SearchBase "CN=Services,$ConfigDN" -LDAPFilter '(objectClass=msExchMDBAvailabilityGroup)' -ErrorAction SilentlyContinue
            foreach ($dag in $dags) {
                $exchangeDAGs += [ordered]@{ Name=$dag.Name }
            }
            if ($exchangeDAGs.Count -gt 0) { Write-Host " [+] $($exchangeDAGs.Count) DAG(s)" -ForegroundColor Green }
        } catch {}

        # Accepted domains
        try {
            $accepted = Get-ADObject -SearchBase "CN=Services,$ConfigDN" -LDAPFilter '(objectClass=msExchAcceptedDomain)' -Properties msExchAcceptedDomainName, msExchAcceptedDomainFlags -ErrorAction SilentlyContinue
            foreach ($a in $accepted) {
                $exchangeAcceptedDomains += [ordered]@{
                    Name=$a.Name
                    Domain=if($a.msExchAcceptedDomainName){[string]$a.msExchAcceptedDomainName}else{$a.Name}
                }
            }
        } catch {}

        # Hybrid
        try {
            $hybrid = Get-ADObject -SearchBase "CN=Services,$ConfigDN" -LDAPFilter '(objectClass=msExchCoexistenceRelationship)' -ErrorAction SilentlyContinue
            if ($hybrid) { $exchangeHybrid = $true; Write-Host " [+] Hybrid Exchange configuration detected" -ForegroundColor Green }
        } catch {}
    } else {
        Write-Host " [i] No Exchange detected in AD" -ForegroundColor Gray
    }
} catch { Write-Host " [i] Exchange check: $($_.Exception.Message)" -ForegroundColor Gray }

$exchangeVersionDisplay = if ($ExchangeSchemaMap.ContainsKey($exchangeSchemaVer)) { $ExchangeSchemaMap[$exchangeSchemaVer] } else { if ($exchangeSchemaVer -gt 0) { "Schema $exchangeSchemaVer" } else { '' } }

$Atlas.Exchange = [ordered]@{
    Detected=$exchangeDetected; OrganizationName=$exchangeOrgName
    SchemaVersion=$exchangeSchemaVer; VersionDisplay=$exchangeVersionDisplay
    Servers=$exchangeServers; DAGs=$exchangeDAGs
    AcceptedDomains=$exchangeAcceptedDomains; Hybrid=$exchangeHybrid
}

# ==============================================================================
# CERTIFICATE SERVICES (PKI)
# ==============================================================================

Write-Host ""
Write-Host " Checking for Certificate Services ..." -ForegroundColor Yellow

$pkiDetected = $false
$enterpriseCAs = @()
$rootCAs = @()
$ntauthCAs = @()
$certTemplates = @()

try {
    # Enterprise CAs (Issuing)
    $entCAs = Get-ADObject -SearchBase "CN=Enrollment Services,CN=Public Key Services,CN=Services,$ConfigDN" -Filter "objectClass -eq 'pKIEnrollmentService'" -Properties dNSHostName, certificateTemplates, WhenCreated -ErrorAction SilentlyContinue
    foreach ($ca in $entCAs) {
        $pkiDetected = $true
        $enterpriseCAs += [ordered]@{
            Name=$ca.Name; Server=[string]$ca.dNSHostName
            PublishedTemplates=@($ca.certificateTemplates | ForEach-Object { [string]$_ })
            Created=if($ca.WhenCreated){$ca.WhenCreated.ToString('yyyy-MM-dd')}else{''}
            Type='Enterprise / Issuing CA'
        }
    }

    # Root CAs (in Certification Authorities container)
    $roots = Get-ADObject -SearchBase "CN=Certification Authorities,CN=Public Key Services,CN=Services,$ConfigDN" -Filter "objectClass -eq 'certificationAuthority'" -Properties WhenCreated -ErrorAction SilentlyContinue
    foreach ($r in $roots) {
        $pkiDetected = $true
        $rootCAs += [ordered]@{
            Name=$r.Name
            Created=if($r.WhenCreated){$r.WhenCreated.ToString('yyyy-MM-dd')}else{''}
            Type='Trusted Root CA'
        }
    }

    # NTAuth CAs (smart card / cert auth)
    try {
        $ntauthObj = Get-ADObject -Identity "CN=NTAuthCertificates,CN=Public Key Services,CN=Services,$ConfigDN" -Properties cACertificate -ErrorAction SilentlyContinue
        if ($ntauthObj -and $ntauthObj.cACertificate) {
            $ntauthCount = @($ntauthObj.cACertificate).Count
            $ntauthCAs += [ordered]@{ Name='NTAuth Store'; CertCount=$ntauthCount }
        }
    } catch {}

    # Certificate Templates (count only - listing all templates is huge)
    try {
        $tpls = Get-ADObject -SearchBase "CN=Certificate Templates,CN=Public Key Services,CN=Services,$ConfigDN" -Filter "objectClass -eq 'pKICertificateTemplate'" -Properties displayName, msPKI-Template-Schema-Version -ErrorAction SilentlyContinue
        foreach ($t in $tpls) {
            $certTemplates += [ordered]@{
                Name=$t.Name
                DisplayName=if($t.displayName){[string]$t.displayName}else{$t.Name}
                SchemaVersion=if($t.'msPKI-Template-Schema-Version'){[int]$t.'msPKI-Template-Schema-Version'}else{0}
            }
        }
    } catch {}

    if ($pkiDetected) {
        Write-Host " [+] PKI detected: $($enterpriseCAs.Count) Enterprise CA(s), $($rootCAs.Count) Root CA(s), $($certTemplates.Count) template(s)" -ForegroundColor Green
    } else {
        Write-Host " [i] No PKI detected" -ForegroundColor Gray
    }
} catch { Write-Host " [i] PKI: $($_.Exception.Message)" -ForegroundColor Gray }

$Atlas.PKI = [ordered]@{
    Detected=$pkiDetected
    EnterpriseCAs=$enterpriseCAs
    RootCAs=$rootCAs
    NTAuthCAs=$ntauthCAs
    Templates=$certTemplates
}

# ==============================================================================
# AUTHENTICATION (per-domain: KRBTGT, encryption types, password/Kerberos policy, FGPP, protected groups)
# ==============================================================================

Write-Host ""
Write-Host " Collecting authentication configuration ..." -ForegroundColor Yellow

$authDomains = @()
$authPolicies = @()
$authSilos = @()

foreach ($d in $Atlas.Domains) {
    $domName = $d.DNSRoot
    $authData = [ordered]@{
        Domain=$domName; NetBIOS=$d.NetBIOSName
        # KRBTGT
        KrbtgtLastReset=''; KrbtgtEnabled=$null; KrbtgtCreated=''
        # Encryption (read from KRBTGT, where it actually lives)
        SupportedEncryptionTypes=''; EncryptionTypesDisplay=@(); EncryptionSource=''
        # Domain-wide
        MachineAccountQuota=$null
        # Password Policy (from Get-ADDefaultDomainPasswordPolicy)
        PwdPolicy=$null
        # Kerberos Policy (parsed from Default Domain Policy GPT.ini in SYSVOL)
        KerbPolicy=$null
        # Fine-Grained Password Policies
        FGPPs=@()
        # Protected groups
        ProtectedUsersCount=$null
        PreWin2kCompatMembers=@()
    }

    # Helper to decode encryption type bitmask
    $DecodeEnc = {
        param($bits)
        $list = @()
        if ($bits -band 1)  { $list += 'DES-CBC-CRC' }
        if ($bits -band 2)  { $list += 'DES-CBC-MD5' }
        if ($bits -band 4)  { $list += 'RC4-HMAC' }
        if ($bits -band 8)  { $list += 'AES128-CTS-HMAC-SHA1-96' }
        if ($bits -band 16) { $list += 'AES256-CTS-HMAC-SHA1-96' }
        if ($bits -band 32) { $list += 'AES128-CTS-HMAC-SHA256-128' }
        if ($bits -band 64) { $list += 'AES256-CTS-HMAC-SHA384-192' }
        return $list
    }

    # ---- KRBTGT account (also reads encryption types from it) ----
    try {
        $krbtgt = Get-ADUser -Identity 'krbtgt' -Server $domName -Properties PasswordLastSet, Enabled, WhenCreated, 'msDS-SupportedEncryptionTypes' -ErrorAction SilentlyContinue
        if ($krbtgt) {
            if ($krbtgt.PasswordLastSet) { $authData.KrbtgtLastReset = $krbtgt.PasswordLastSet.ToString('yyyy-MM-dd') }
            $authData.KrbtgtEnabled = [bool]$krbtgt.Enabled
            if ($krbtgt.WhenCreated) { $authData.KrbtgtCreated = $krbtgt.WhenCreated.ToString('yyyy-MM-dd') }
            $enc = $krbtgt.'msDS-SupportedEncryptionTypes'
            if ($enc) {
                $authData.SupportedEncryptionTypes = [int]$enc
                $authData.EncryptionTypesDisplay = & $DecodeEnc ([int]$enc)
                $authData.EncryptionSource = 'krbtgt account (msDS-SupportedEncryptionTypes)'
            }
        }
    } catch {}

    # ---- Fallback: read encryption types from PDC registry via WMI ----
    # This is the actual KDC-wide setting; krbtgt attribute is often unset.
    if (-not $authData.EncryptionTypesDisplay -or $authData.EncryptionTypesDisplay.Count -eq 0) {
        $pdc = $d.PDCEmulator
        if ($pdc) {
            try {
                $HKLM = 2147483650
                $reg = [wmiclass]"\\$pdc\root\default:StdRegProv"

                # Path 1: KDC service config (HKLM\SYSTEM\...\Kdc\Parameters\DefaultDomainSupportedEncTypes)
                $val1 = $reg.GetDWORDValue($HKLM, 'SYSTEM\CurrentControlSet\Services\Kdc\Parameters', 'DefaultDomainSupportedEncTypes')
                if ($val1.ReturnValue -eq 0 -and $null -ne $val1.uValue) {
                    $authData.SupportedEncryptionTypes = [int]$val1.uValue
                    $authData.EncryptionTypesDisplay   = & $DecodeEnc ([int]$val1.uValue)
                    $authData.EncryptionSource         = "PDC registry: HKLM\SYSTEM\...\Kdc\Parameters\DefaultDomainSupportedEncTypes (on $pdc)"
                }

                # Path 2: GPO-applied per-machine Kerberos client setting
                # (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters\SupportedEncryptionTypes)
                if (-not $authData.EncryptionTypesDisplay -or $authData.EncryptionTypesDisplay.Count -eq 0) {
                    $val2 = $reg.GetDWORDValue($HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters', 'SupportedEncryptionTypes')
                    if ($val2.ReturnValue -eq 0 -and $null -ne $val2.uValue) {
                        $authData.SupportedEncryptionTypes = [int]$val2.uValue
                        $authData.EncryptionTypesDisplay   = & $DecodeEnc ([int]$val2.uValue)
                        $authData.EncryptionSource         = "PDC registry: HKLM\SOFTWARE\...\Kerberos\Parameters\SupportedEncryptionTypes (GPO, on $pdc)"
                    }
                }

                # Still nothing? Make it explicit.
                if (-not $authData.EncryptionTypesDisplay -or $authData.EncryptionTypesDisplay.Count -eq 0) {
                    $authData.EncryptionSource = "Not explicitly configured -- OS default applies (varies by Windows patch level: pre-Nov 2022 = RC4+AES128+AES256; post-Nov 2022 KB5021131 = AES128+AES256 only)"
                }
            } catch {
                $authData.EncryptionSource = "(could not query PDC $pdc via WMI: $($_.Exception.Message))"
            }
        }
    }

    # ---- Machine Account Quota (from domain head) ----
    try {
        $domObj = Get-ADObject -Identity $d.DistinguishedName -Server $domName -Properties 'ms-DS-MachineAccountQuota' -ErrorAction SilentlyContinue
        if ($domObj -and $null -ne $domObj.'ms-DS-MachineAccountQuota') {
            $authData.MachineAccountQuota = [int]$domObj.'ms-DS-MachineAccountQuota'
        }
    } catch {}

    # ---- Default Domain Password Policy ----
    try {
        $pwd = Get-ADDefaultDomainPasswordPolicy -Server $domName -ErrorAction SilentlyContinue
        if ($pwd) {
            $authData.PwdPolicy = [ordered]@{
                MinPasswordLength            = [int]$pwd.MinPasswordLength
                PasswordHistoryCount         = [int]$pwd.PasswordHistoryCount
                MaxPasswordAgeDays           = if ($pwd.MaxPasswordAge -and $pwd.MaxPasswordAge.Days) { [int]$pwd.MaxPasswordAge.Days } else { 0 }
                MinPasswordAgeDays           = if ($pwd.MinPasswordAge -and $pwd.MinPasswordAge.Days) { [int]$pwd.MinPasswordAge.Days } else { 0 }
                ComplexityEnabled            = [bool]$pwd.ComplexityEnabled
                ReversibleEncryptionEnabled  = [bool]$pwd.ReversibleEncryptionEnabled
                LockoutThreshold             = [int]$pwd.LockoutThreshold
                LockoutDurationMinutes       = if ($pwd.LockoutDuration) { [int]$pwd.LockoutDuration.TotalMinutes } else { 0 }
                LockoutObservationMinutes    = if ($pwd.LockoutObservationWindow) { [int]$pwd.LockoutObservationWindow.TotalMinutes } else { 0 }
            }
        }
    } catch {}

    # ---- Kerberos Policy (parse Default Domain Policy GPT.ini from SYSVOL) ----
    try {
        $defaultDomainPolicyGuid = '{31B2F340-016D-11D2-945F-00C04FB984F9}'
        $gptPath = "\\$domName\SYSVOL\$domName\Policies\$defaultDomainPolicyGuid\MACHINE\Microsoft\Windows NT\SecEdit\GptTmpl.inf"
        if (Test-Path -LiteralPath $gptPath -ErrorAction SilentlyContinue) {
            $gptContent = Get-Content -LiteralPath $gptPath -ErrorAction SilentlyContinue
            $inKerb = $false
            $kerb = [ordered]@{
                MaxTicketAgeHours    = $null
                MaxRenewAgeDays      = $null
                MaxServiceAgeMinutes = $null
                MaxClockSkewMinutes  = $null
                TicketValidateClient = $null
            }
            foreach ($line in $gptContent) {
                if ($line -match '^\s*\[Kerberos Policy\]') { $inKerb = $true; continue }
                if ($line -match '^\s*\[') { $inKerb = $false; continue }
                if (-not $inKerb) { continue }
                if ($line -match '^\s*MaxTicketAge\s*=\s*(\d+)')        { $kerb.MaxTicketAgeHours    = [int]$Matches[1] }
                elseif ($line -match '^\s*MaxRenewAge\s*=\s*(\d+)')      { $kerb.MaxRenewAgeDays      = [int]$Matches[1] }
                elseif ($line -match '^\s*MaxServiceAge\s*=\s*(\d+)')    { $kerb.MaxServiceAgeMinutes = [int]$Matches[1] }
                elseif ($line -match '^\s*MaxClockSkew\s*=\s*(\d+)')     { $kerb.MaxClockSkewMinutes  = [int]$Matches[1] }
                elseif ($line -match '^\s*TicketValidateClient\s*=\s*(\d+)') { $kerb.TicketValidateClient = [int]$Matches[1] }
            }
            if ($null -ne $kerb.MaxTicketAgeHours -or $null -ne $kerb.MaxServiceAgeMinutes) {
                $authData.KerbPolicy = $kerb
            }
        }
    } catch {}

    # ---- Fine-Grained Password Policies ----
    try {
        $fgpps = Get-ADFineGrainedPasswordPolicy -Filter * -Server $domName -ErrorAction SilentlyContinue
        foreach ($f in $fgpps) {
            $appliesTo = @()
            try {
                $appliesObjs = Get-ADFineGrainedPasswordPolicySubject -Identity $f -Server $domName -ErrorAction SilentlyContinue
                $appliesTo = @($appliesObjs | ForEach-Object { [string]$_.Name })
            } catch {}
            $authData.FGPPs += [ordered]@{
                Name              = [string]$f.Name
                Precedence        = [int]$f.Precedence
                MinPasswordLength = [int]$f.MinPasswordLength
                PasswordHistoryCount = [int]$f.PasswordHistoryCount
                MaxPasswordAgeDays = if ($f.MaxPasswordAge -and $f.MaxPasswordAge.Days) { [int]$f.MaxPasswordAge.Days } else { 0 }
                ComplexityEnabled = [bool]$f.ComplexityEnabled
                LockoutThreshold  = [int]$f.LockoutThreshold
                AppliesTo         = $appliesTo
            }
        }
    } catch {}

    # ---- Protected Users group member count ----
    try {
        $pu = Get-ADGroup -Identity 'Protected Users' -Server $domName -Properties Members -ErrorAction SilentlyContinue
        if ($pu) { $authData.ProtectedUsersCount = @($pu.Members).Count }
    } catch {}

    # ---- Pre-Windows 2000 Compatible Access membership ----
    try {
        $pw2k = Get-ADGroup -Identity 'Pre-Windows 2000 Compatible Access' -Server $domName -Properties Members -ErrorAction SilentlyContinue
        if ($pw2k -and $pw2k.Members) {
            foreach ($mDn in $pw2k.Members) {
                $resolved = ''
                # FSP detection FIRST: Get-ADObject succeeds for FSP DNs but returns Name as raw SID,
                # so we have to detect the FSP DN pattern up front and translate the SID directly.
                if ($mDn -match 'CN=(S-1-[\d\-]+),CN=ForeignSecurityPrincipals') {
                    $sidStr = $Matches[1]
                    try {
                        $sid = New-Object System.Security.Principal.SecurityIdentifier($sidStr)
                        $resolved = $sid.Translate([System.Security.Principal.NTAccount]).Value
                    } catch { $resolved = $sidStr }
                } else {
                    # Regular AD object - look up its Name
                    try {
                        $mObj = Get-ADObject -Identity $mDn -Server $domName -ErrorAction SilentlyContinue
                        if ($mObj) { $resolved = [string]$mObj.Name }
                    } catch {}
                    if (-not $resolved) {
                        $resolved = ($mDn -split ',')[0] -replace '^CN=',''
                    }
                }
                $authData.PreWin2kCompatMembers += $resolved
            }
        }
    } catch {}

    $authDomains += $authData
    Write-Host "     Auth: $domName -- pwd policy=$($null -ne $authData.PwdPolicy), kerb policy=$($null -ne $authData.KerbPolicy), FGPPs=$($authData.FGPPs.Count), enc types=$($authData.EncryptionTypesDisplay.Count)" -ForegroundColor Gray
}

# Authentication Policies & Silos (forest-wide)
try {
    $polContainer = "CN=AuthN Policy Configuration,CN=Services,$ConfigDN"
    $policies = Get-ADObject -SearchBase $polContainer -Filter "objectClass -eq 'msDS-AuthNPolicy'" -ErrorAction SilentlyContinue
    foreach ($p in $policies) {
        $authPolicies += [ordered]@{ Name=$p.Name }
    }
    $silos = Get-ADObject -SearchBase $polContainer -Filter "objectClass -eq 'msDS-AuthNPolicySilo'" -ErrorAction SilentlyContinue
    foreach ($s in $silos) {
        $authSilos += [ordered]@{ Name=$s.Name }
    }
    if ($authPolicies.Count -gt 0 -or $authSilos.Count -gt 0) {
        Write-Host " [+] $($authPolicies.Count) auth policies, $($authSilos.Count) silos" -ForegroundColor Green
    }
} catch {}

Write-Host " [+] Authentication data collected for $($authDomains.Count) domain(s)" -ForegroundColor Green

$Atlas.Authentication = [ordered]@{
    Domains=$authDomains
    Policies=$authPolicies
    Silos=$authSilos
}

# ==============================================================================
# NTP / TIME CONFIGURATION (parallel runspaces, w32tm + WMI registry fallback)
# ==============================================================================

Write-Host ""
Write-Host " Collecting NTP configuration in parallel (w32tm with WMI registry fallback) ..." -ForegroundColor Yellow

$NtpMaxConcurrent = 20
$NtpTimeoutSec = 8

# Worker scriptblock - runs once per DC inside a runspace.
# Tries w32tm first; falls back to WMI StdRegProv reading W32Time registry directly.
$ntpWorker = {
    param($Computer, $TimeoutSec)

    $r = [ordered]@{
        Reachable    = $false
        Source       = ''
        ConfigType   = ''
        NtpServerCfg = ''
        Method       = ''
    }

    # ---- Method 1: w32tm /query /computer ----
    function _Run-W32tm($cmp, $args, $tms) {
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = 'w32tm.exe'
            $psi.Arguments = "/query /computer:$cmp $args"
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $psi.CreateNoWindow = $true
            $p = [System.Diagnostics.Process]::Start($psi)
            if (-not $p.WaitForExit($tms)) { try { $p.Kill() } catch {}; return $null }
            if ($p.ExitCode -ne 0) { return $null }
            return $p.StandardOutput.ReadToEnd()
        } catch { return $null }
    }

    $srcOut = _Run-W32tm $Computer '/source' ($TimeoutSec * 1000)
    if ($srcOut) {
        $first = ($srcOut -split "`r?`n" | Where-Object { $_ -and $_.Trim() -ne '' } | Select-Object -First 1)
        if ($first) {
            $r.Reachable = $true
            $r.Source    = $first.Trim()
            $r.Method    = 'w32tm'
        }
        $cfgOut = _Run-W32tm $Computer '/configuration' ($TimeoutSec * 1000)
        if ($cfgOut) {
            foreach ($line in ($cfgOut -split "`r?`n")) {
                if     ($line -match '^\s*Type:\s*(\S+)')                              { $r.ConfigType   = $Matches[1] }
                elseif ($line -match '^\s*NtpServer:\s*(.+?)(\s+\(Local\))?\s*$')       { $r.NtpServerCfg = $Matches[1].Trim() }
            }
        }
    }

    # ---- Method 2: WMI StdRegProv (DCOM) - fallback if w32tm RPC was blocked ----
    if (-not $r.Reachable) {
        try {
            $HKLM = 2147483650
            $reg = [wmiclass]"\\$Computer\root\default:StdRegProv"
            $base = 'SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
            $ntpVal  = $reg.GetStringValue($HKLM, $base, 'NtpServer')
            $typeVal = $reg.GetStringValue($HKLM, $base, 'Type')
            if ($ntpVal.ReturnValue -eq 0 -or $typeVal.ReturnValue -eq 0) {
                $r.Reachable = $true
                $r.Method    = 'WMI registry'
                if ($ntpVal.ReturnValue  -eq 0 -and $ntpVal.sValue)  { $r.NtpServerCfg = $ntpVal.sValue }
                if ($typeVal.ReturnValue -eq 0 -and $typeVal.sValue) { $r.ConfigType   = $typeVal.sValue }
                # Synthesize a Source string from the registry data so the diagram has something to work with
                if ($r.ConfigType -eq 'NTP' -and $r.NtpServerCfg) {
                    $r.Source = ($r.NtpServerCfg -split '\s+' | Select-Object -First 1)
                } elseif ($r.ConfigType -eq 'NT5DS') {
                    $r.Source = '(domain hierarchy - NT5DS)'
                } elseif ($r.ConfigType -eq 'AllSync') {
                    $r.Source = '(AllSync)'
                } elseif ($r.NtpServerCfg) {
                    $r.Source = ($r.NtpServerCfg -split '\s+' | Select-Object -First 1)
                }
            }
        } catch {}
    }

    return $r
}

# ---- Spin up runspace pool ----
$rsPool = [runspacefactory]::CreateRunspacePool(1, $NtpMaxConcurrent)
$rsPool.Open()
$rsJobs = @()

foreach ($dc in $Atlas.DCs) {
    $hn = $dc.HostName
    if (-not $hn) { $hn = $dc.Name }
    $ps = [powershell]::Create()
    $ps.RunspacePool = $rsPool
    [void]$ps.AddScript($ntpWorker).AddArgument($hn).AddArgument($NtpTimeoutSec)
    $rsJobs += [pscustomobject]@{
        DC       = $dc
        HostName = $hn
        PS       = $ps
        Handle   = $ps.BeginInvoke()
    }
}

# Wait with progress
$totalJobs = $rsJobs.Count
do {
    $done = @($rsJobs | Where-Object { $_.Handle.IsCompleted }).Count
    Write-Progress -Activity "Querying NTP (parallel, $NtpMaxConcurrent at a time)" -Status "$done of $totalJobs DCs complete" -PercentComplete (($done / [math]::Max($totalJobs,1)) * 100)
    if ($done -lt $totalJobs) { Start-Sleep -Milliseconds 250 }
} while ($done -lt $totalJobs)
Write-Progress -Activity "Querying NTP (parallel, $NtpMaxConcurrent at a time)" -Completed

# Collect results
$ntpDCs = @()
$ntpReachable = 0
$ntpUnreachable = 0
$methodCounts = @{ 'w32tm' = 0; 'WMI registry' = 0 }

foreach ($j in $rsJobs) {
    $result = $null
    try {
        $out = $j.PS.EndInvoke($j.Handle)
        if ($out -and $out.Count -gt 0) { $result = $out[0] }
    } catch {}
    finally { $j.PS.Dispose() }

    if (-not $result) { $result = @{ Reachable=$false; Source=''; ConfigType=''; NtpServerCfg=''; Method='' } }

    if ($result.Reachable) {
        $ntpReachable++
        if ($result.Method -and $methodCounts.ContainsKey($result.Method)) { $methodCounts[$result.Method]++ }
    } else {
        $ntpUnreachable++
    }

    $dc = $j.DC
    $ntpDCs += [ordered]@{
        Name            = $dc.Name
        HostName        = $j.HostName
        Domain          = $dc.Domain
        Site            = $dc.Site
        Type            = $dc.Type
        IsPDC           = ($dc.FSMORoles -contains 'PDCEmulator')
        IsForestRootPDC = (($dc.FSMORoles -contains 'PDCEmulator') -and ($dc.Domain -eq $Atlas.Forest.RootDomain))
        Reachable       = [bool]$result.Reachable
        Source          = [string]$result.Source
        ConfigType      = [string]$result.ConfigType
        NtpServerCfg    = [string]$result.NtpServerCfg
        Method          = [string]$result.Method
    }
}
$rsPool.Close()
$rsPool.Dispose()

Write-Host " [+] NTP queried: $ntpReachable reachable, $ntpUnreachable unreachable" -ForegroundColor Green
Write-Host "     Methods used: w32tm=$($methodCounts['w32tm']), WMI registry fallback=$($methodCounts['WMI registry'])" -ForegroundColor Gray

$Atlas.NTP = [ordered]@{
    Collected   = $true
    Reachable   = $ntpReachable
    Unreachable = $ntpUnreachable
    DCs         = $ntpDCs
}

# ==============================================================================
# COLLECTION SUMMARY
# ==============================================================================

Write-Host ""
Write-Host " +============================================================+" -ForegroundColor Cyan
Write-Host " |                  Collection Summary                       |" -ForegroundColor Cyan
Write-Host " +============================================================+" -ForegroundColor Cyan
Write-Host ""
Write-Host ("   Forest              : {0}" -f $Atlas.Forest.Name) -ForegroundColor White
Write-Host ("   Forest Mode         : {0}" -f $Atlas.Forest.ForestModeDisplay) -ForegroundColor White
Write-Host ("   Schema Version      : {0} ({1})" -f $Atlas.Forest.SchemaVersion, $Atlas.Forest.SchemaOS) -ForegroundColor White
Write-Host ("   Domains             : {0}" -f $Atlas.Domains.Count) -ForegroundColor White
Write-Host ("   Domain Controllers  : {0}" -f $Atlas.DCs.Count) -ForegroundColor White
Write-Host ("   Sites               : {0}" -f $Atlas.Sites.Count) -ForegroundColor White
Write-Host ("   Subnets             : {0}" -f $Atlas.Subnets.Count) -ForegroundColor White
Write-Host ("   Site Links          : {0}" -f $Atlas.SiteLinks.Count) -ForegroundColor White
Write-Host ("   Trusts (unique)     : {0}" -f $Atlas.Trusts.Count) -ForegroundColor White
Write-Host ("   Replication Conns   : {0}" -f $Atlas.Connections.Count) -ForegroundColor White
Write-Host ("   Entra Connect       : {0}" -f $(if($Atlas.EntraConnect.Detected){"Detected" + $(if($Atlas.EntraConnect.Tenant){" ($($Atlas.EntraConnect.Tenant))"}else{""})}else{"Not detected"})) -ForegroundColor White
Write-Host ("   DNS Zones           : {0}" -f $Atlas.DNS.Zones.Count) -ForegroundColor White
Write-Host ("   Exchange            : {0}" -f $(if($Atlas.Exchange.Detected){"$($Atlas.Exchange.Servers.Count) server(s)"}else{"Not detected"})) -ForegroundColor White
Write-Host ("   PKI / CAs           : {0}" -f $(if($Atlas.PKI.Detected){"$($Atlas.PKI.EnterpriseCAs.Count) Enterprise CA(s)"}else{"Not detected"})) -ForegroundColor White
Write-Host ("   NTP (w32tm queried) : {0} reachable / {1} unreachable" -f $Atlas.NTP.Reachable, $Atlas.NTP.Unreachable) -ForegroundColor White
Write-Host ""

# ==============================================================================
# HTML GENERATION
# ==============================================================================

Write-Host ""
Write-Host " Building HTML report ..." -ForegroundColor Yellow

Add-Type -AssemblyName System.Web

function HtmlEnc($s) {
    if ($null -eq $s) { return '--' }
    return [System.Web.HttpUtility]::HtmlEncode([string]$s)
}

$AtlasJson = $Atlas | ConvertTo-Json -Depth 12 -Compress
$AtlasJson = $AtlasJson -replace '</script>', '<\/script>'

$fw = $Atlas.Forest
$rwdcCount = @($Atlas.DCs | Where-Object { $_.Type -eq 'RWDC' }).Count
$rodcCount = @($Atlas.DCs | Where-Object { $_.Type -eq 'RODC' }).Count
$gcCount   = @($Atlas.DCs | Where-Object { $_.IsGlobalCatalog }).Count
$intraTrusts = @($Atlas.Trusts | Where-Object { $_.IntraForest }).Count
$extTrusts   = @($Atlas.Trusts | Where-Object { -not $_.IntraForest }).Count

# ---- Domain summary table ----
$domainRowsHtml = ''
foreach ($d in ($Atlas.Domains | Sort-Object -Property @{E={-not $_.IsForestRoot}}, DNSRoot)) {
    $rootBadge = if ($d.IsForestRoot) { '<span class="badge badge-accent">Forest Root</span>' } else { '' }
    $parent = if ($d.ParentDomain) { HtmlEnc $d.ParentDomain } else { '<span class="dim">(none)</span>' }
    $children = if ($d.ChildDomains.Count -gt 0) { HtmlEnc (($d.ChildDomains) -join ', ') } else { '<span class="dim">(none)</span>' }
    $domainRowsHtml += "<tr><td><strong>$(HtmlEnc $d.DNSRoot)</strong> $rootBadge</td><td>$(HtmlEnc $d.NetBIOSName)</td><td>$(HtmlEnc $d.DomainModeDisplay)</td><td>$parent</td><td>$children</td><td style='text-align:center'>$($d.DCCount)</td><td style='font-size:.72rem'>$(HtmlEnc $d.PDCEmulator)</td></tr>`n"
}

$fsmoHtml = "<div class='fsmo-grid'><div class='fsmo-card'><div class='role'>Schema Master</div><div class='holder'>$(HtmlEnc $fw.SchemaMaster)</div></div><div class='fsmo-card'><div class='role'>Domain Naming Master</div><div class='holder'>$(HtmlEnc $fw.DomainNamingMaster)</div></div></div><p class='section-desc'>Per-domain FSMO roles (PDC Emulator, RID Master, Infrastructure Master) are shown in the Domain Summary table above and on the Forest Map.</p>"

$domainNavHtml = ''
foreach ($d in ($Atlas.Domains | Sort-Object -Property @{E={-not $_.IsForestRoot}}, DNSRoot)) {
    $domainNavHtml += "    <a href=`"#`" data-tab=`"tab-forest`" data-focus=`"$(HtmlEnc $d.DNSRoot)`">$(HtmlEnc $d.DNSRoot)</a>`n"
}

# ---- Entra overview card ----
$entraOverviewHtml = ''
if ($Atlas.EntraConnect.Detected) {
    $svrName = if ($Atlas.EntraConnect.ServerName) { HtmlEnc $Atlas.EntraConnect.ServerName } else { '<span class="dim">(unknown)</span>' }
    $tenantName = if ($Atlas.EntraConnect.Tenant) { HtmlEnc $Atlas.EntraConnect.Tenant } else { '<span class="dim">(unknown)</span>' }
    $entraOverviewHtml = "<div class='info-card' style='border-color:var(--accent2)'><span class='info-label' style='color:var(--accent2)'>Entra Connect (Hybrid Identity)</span><span class='info-value'>Detected</span><div style='font-size:.7rem;color:var(--text-dim);margin-top:4px'>Server: $svrName &middot; Tenant: $tenantName</div></div>"
} else {
    $entraOverviewHtml = "<div class='info-card'><span class='info-label'>Entra Connect (Hybrid Identity)</span><span class='info-value'><span class='dim'>Not detected</span></span></div>"
}

# ---- Trust matrix ----
$allMatrixDomains = @{}
foreach ($d in $Atlas.Domains) { $allMatrixDomains[$d.DNSRoot] = $true }
foreach ($t in $Atlas.Trusts) {
    $allMatrixDomains[$t.SourceDomain] = $true
    $allMatrixDomains[$t.TargetDomain] = $true
}
$matrixDomainList = @($allMatrixDomains.Keys | Sort-Object)
$trustLookup = @{}
foreach ($t in $Atlas.Trusts) {
    $trustLookup["$($t.SourceDomain)|$($t.TargetDomain)"] = $t
    if ($t.Direction -eq 'BiDirectional') { $trustLookup["$($t.TargetDomain)|$($t.SourceDomain)"] = $t }
}
$forestDomainSet = @{}
foreach ($d in $Atlas.Domains) { $forestDomainSet[$d.DNSRoot] = $true }

$matrixHtml = ''
if ($matrixDomainList.Count -eq 0) {
    $matrixHtml = '<p class="empty-note">No trusts or domains to display.</p>'
} else {
    $matrixHtml = '<div class="matrix-wrap"><table class="matrix-table"><thead><tr><th class="corner">Source \ Target</th>'
    foreach ($col in $matrixDomainList) {
        $extBadge = if (-not $forestDomainSet.ContainsKey($col)) { ' <span class="badge badge-amber" style="font-size:.55rem">EXT</span>' } else { '' }
        $matrixHtml += '<th class="col-head">' + (HtmlEnc $col) + $extBadge + '</th>'
    }
    $matrixHtml += '</tr></thead><tbody>'
    foreach ($row in $matrixDomainList) {
        $extBadge = if (-not $forestDomainSet.ContainsKey($row)) { ' <span class="badge badge-amber" style="font-size:.55rem">EXT</span>' } else { '' }
        $matrixHtml += '<tr><th class="row-head">' + (HtmlEnc $row) + $extBadge + '</th>'
        foreach ($col in $matrixDomainList) {
            if ($row -eq $col) {
                $matrixHtml += '<td class="matrix-cell-self">&mdash;</td>'
            } else {
                $t = $trustLookup["$row|$col"]
                if ($t) {
                    $cellClass = if ($t.IntraForest) { 'matrix-cell-intra' } else { 'matrix-cell-external' }
                    $direction = switch ($t.Direction) { 'BiDirectional'{'&harr;'} 'Inbound'{'&larr;'} 'Outbound'{'&rarr;'} default{'?'} }
                    $typeLabel = if ($t.IntraForest) { 'Intra' } else { 'Ext' }
                    $tooltip = "$($t.TrustType) | $($t.Direction) | Transitive: $($t.Transitive) | SID Filter: $($t.SIDFilter)"
                    $matrixHtml += '<td class="' + $cellClass + '" title="' + (HtmlEnc $tooltip) + '">' + $direction + ' ' + $typeLabel + '</td>'
                } else {
                    $matrixHtml += '<td class="matrix-cell-none">&middot;</td>'
                }
            }
        }
        $matrixHtml += '</tr>'
    }
    $matrixHtml += '</tbody></table></div>'
}

# ---- Site links table ----
$siteLinksHtml = ''
if ($Atlas.SiteLinks.Count -eq 0) {
    $siteLinksHtml = '<p class="empty-note">No site links found.</p>'
} else {
    $siteLinksHtml = '<div class="table-wrap"><table><thead><tr><th>Name</th><th>Cost</th><th>Frequency (min)</th><th>Transport</th><th>Sites Included</th><th>Description</th></tr></thead><tbody>'
    foreach ($sl in ($Atlas.SiteLinks | Sort-Object Cost, Name)) {
        $sitesStr = if ($sl.Sites.Count -gt 0) { HtmlEnc ($sl.Sites -join ', ') } else { '<span class="dim">(none)</span>' }
        $desc = if ($sl.Description) { HtmlEnc $sl.Description } else { '<span class="dim">--</span>' }
        $siteLinksHtml += "<tr><td><strong>$(HtmlEnc $sl.Name)</strong></td><td style='text-align:center'>$($sl.Cost)</td><td style='text-align:center'>$($sl.Frequency)</td><td>$(HtmlEnc $sl.Transport)</td><td>$sitesStr</td><td>$desc</td></tr>"
    }
    $siteLinksHtml += '</tbody></table></div>'
}

# ---- Sites & subnets combined table ----
$subnetsBySite = @{}
foreach ($sn in $Atlas.Subnets) {
    $key = if ($sn.Site) { $sn.Site } else { '(unassigned)' }
    if (-not $subnetsBySite.ContainsKey($key)) { $subnetsBySite[$key] = @() }
    $subnetsBySite[$key] += $sn
}
$sitesSubnetsHtml = ''
if ($Atlas.Sites.Count -eq 0) {
    $sitesSubnetsHtml = '<p class="empty-note">No sites found.</p>'
} else {
    $sitesSubnetsHtml = '<div class="table-wrap"><table><thead><tr><th>Site</th><th>Description</th><th>Location</th><th>DCs</th><th>Subnets</th></tr></thead><tbody>'
    foreach ($s in ($Atlas.Sites | Sort-Object Name)) {
        $siteDcs = @($Atlas.DCs | Where-Object { $_.Site -eq $s.Name })
        $dcCnt = $siteDcs.Count
        $siteSubnets = if ($subnetsBySite.ContainsKey($s.Name)) { $subnetsBySite[$s.Name] } else { @() }
        $subnetCellHtml = if ($siteSubnets.Count -eq 0) { '<span class="dim">(none)</span>' } else {
            ($siteSubnets | ForEach-Object {
                $loc = if ($_.Location) { " <span class='dim'>($(HtmlEnc $_.Location))</span>" } else { '' }
                "<div style='font-size:.74rem'><code>$(HtmlEnc $_.Name)</code>$loc</div>"
            }) -join ''
        }
        $desc = if ($s.Description) { HtmlEnc $s.Description } else { '<span class="dim">--</span>' }
        $loc = if ($s.Location) { HtmlEnc $s.Location } else { '<span class="dim">--</span>' }
        $sitesSubnetsHtml += "<tr><td><strong>$(HtmlEnc $s.Name)</strong></td><td>$desc</td><td>$loc</td><td style='text-align:center'>$dcCnt</td><td>$subnetCellHtml</td></tr>"
    }
    if ($subnetsBySite.ContainsKey('(unassigned)')) {
        $unassigned = $subnetsBySite['(unassigned)']
        $subnetCellHtml = ($unassigned | ForEach-Object { "<div style='font-size:.74rem'><code>$(HtmlEnc $_.Name)</code></div>" }) -join ''
        $sitesSubnetsHtml += "<tr><td><strong><span class='dim'>(unassigned)</span></strong></td><td><span class='dim'>--</span></td><td><span class='dim'>--</span></td><td style='text-align:center'>0</td><td>$subnetCellHtml</td></tr>"
    }
    $sitesSubnetsHtml += '</tbody></table></div>'
}

# ---- DC inventory ----
$dcInventoryHtml = ''
if ($Atlas.DCs.Count -eq 0) {
    $dcInventoryHtml = '<p class="empty-note">No domain controllers found.</p>'
} else {
    $dcInventoryHtml = "<div class='dc-controls'><input type='text' id='dc-search' class='dc-search' placeholder='Search by name, domain, site, IP, OS...' /><select id='dc-filter-domain' class='dc-filter'><option value=''>All Domains</option>"
    foreach ($d in ($Atlas.Domains | Sort-Object DNSRoot)) {
        $dcInventoryHtml += "<option value=`"$(HtmlEnc $d.DNSRoot)`">$(HtmlEnc $d.DNSRoot)</option>"
    }
    $dcInventoryHtml += "</select><select id='dc-filter-site' class='dc-filter'><option value=''>All Sites</option>"
    foreach ($s in ($Atlas.Sites | Sort-Object Name)) {
        $dcInventoryHtml += "<option value=`"$(HtmlEnc $s.Name)`">$(HtmlEnc $s.Name)</option>"
    }
    $dcInventoryHtml += "</select><select id='dc-filter-type' class='dc-filter'><option value=''>All Types</option><option value='RWDC'>RWDC only</option><option value='RODC'>RODC only</option></select><span class='dc-count' id='dc-count'></span></div><div class='table-wrap'><table id='dc-table'><thead><tr><th data-sort='Name'>Name</th><th data-sort='Domain'>Domain</th><th data-sort='Site'>Site</th><th data-sort='IPv4Address'>IP Address</th><th data-sort='OperatingSystem'>Operating System</th><th data-sort='OSVersion'>OS Version</th><th data-sort='Type'>Type</th><th>GC</th><th>FSMO Roles</th></tr></thead><tbody id='dc-tbody'></tbody></table></div>"
}

# ---- DNS table ----
$dnsTableHtml = ''
if ($Atlas.DNS.Zones.Count -eq 0) {
    $dnsTableHtml = '<p class="empty-note">No DNS zone data collected. Either the DnsServer module is unavailable or no DC responded.</p>'
} else {
    $dnsTableHtml = '<div class="table-wrap"><table><thead><tr><th>Zone Name</th><th>Type</th><th>AD-Integrated</th><th>Replication Scope</th><th>Reverse</th><th>Dynamic Update</th><th>Auto-Created</th></tr></thead><tbody>'
    foreach ($z in ($Atlas.DNS.Zones | Sort-Object @{E={$_.IsAutoCreated}}, ZoneName)) {
        $intBadge = if ($z.IsDsIntegrated) { '<span class="badge badge-green">Yes</span>' } else { '<span class="badge badge-amber">No</span>' }
        $revBadge = if ($z.IsReverseLookupZone) { '<span class="dim">Reverse</span>' } else { '' }
        $autoBadge = if ($z.IsAutoCreated) { '<span class="dim">Auto</span>' } else { '' }
        $dnsTableHtml += "<tr><td><strong>$(HtmlEnc $z.ZoneName)</strong></td><td>$(HtmlEnc $z.ZoneType)</td><td style='text-align:center'>$intBadge</td><td>$(HtmlEnc $z.ReplicationScope)</td><td>$revBadge</td><td>$(HtmlEnc $z.DynamicUpdate)</td><td>$autoBadge</td></tr>"
    }
    $dnsTableHtml += '</tbody></table></div>'
    if ($Atlas.DNS.Forwarders.Count -gt 0) {
        $dnsTableHtml += "<h3 class='sub-header'>Configured Forwarders ($($Atlas.DNS.Forwarders.Count))</h3><div class='info-grid'>"
        foreach ($f in $Atlas.DNS.Forwarders) {
            $dnsTableHtml += "<div class='info-card'><span class='info-label'>Forwarder</span><span class='info-value'><code>$(HtmlEnc $f)</code></span></div>"
        }
        $dnsTableHtml += '</div>'
    }
    $dnsTableHtml += "<p class='section-desc' style='margin-top:14px'>DNS data was collected from <strong>$(HtmlEnc $Atlas.DNS.QueriedFrom)</strong>. Other DCs may host the same AD-integrated zones via replication.</p>"
}

# ---- Exchange table ----
$exchangeTableHtml = ''
if (-not $Atlas.Exchange.Detected) {
    $exchangeTableHtml = '<p class="empty-note">No Exchange detected in the AD configuration partition.</p>'
} else {
    $exchangeTableHtml = "<div class='info-grid'><div class='info-card'><span class='info-label'>Organization Name</span><span class='info-value'>$(HtmlEnc $Atlas.Exchange.OrganizationName)</span></div><div class='info-card'><span class='info-label'>Schema Version</span><span class='info-value'>$($Atlas.Exchange.SchemaVersion)</span></div><div class='info-card'><span class='info-label'>Version</span><span class='info-value'>$(HtmlEnc $Atlas.Exchange.VersionDisplay)</span></div><div class='info-card'><span class='info-label'>Hybrid Configuration</span><span class='info-value'>$(if($Atlas.Exchange.Hybrid){'<span style=color:var(--accent2)>Yes (Exchange Online connected)</span>'}else{'<span class=dim>No</span>'})</span></div><div class='info-card'><span class='info-label'>Servers</span><span class='info-value'>$($Atlas.Exchange.Servers.Count)</span></div><div class='info-card'><span class='info-label'>DAGs</span><span class='info-value'>$($Atlas.Exchange.DAGs.Count)</span></div></div>"
    if ($Atlas.Exchange.Servers.Count -gt 0) {
        $exchangeTableHtml += "<h3 class='sub-header'>Exchange Servers</h3><div class='table-wrap'><table><thead><tr><th>Name</th><th>Version</th><th>Roles</th><th>Created</th></tr></thead><tbody>"
        foreach ($s in $Atlas.Exchange.Servers) {
            $roles = if ($s.Roles.Count -gt 0) { HtmlEnc ($s.Roles -join ', ') } else { '<span class="dim">--</span>' }
            $exchangeTableHtml += "<tr><td><strong>$(HtmlEnc $s.Name)</strong></td><td>$(HtmlEnc $s.Version)</td><td>$roles</td><td>$(HtmlEnc $s.Created)</td></tr>"
        }
        $exchangeTableHtml += '</tbody></table></div>'
    }
    if ($Atlas.Exchange.AcceptedDomains.Count -gt 0) {
        $exchangeTableHtml += "<h3 class='sub-header'>Accepted Domains ($($Atlas.Exchange.AcceptedDomains.Count))</h3><div class='info-grid'>"
        foreach ($a in $Atlas.Exchange.AcceptedDomains) {
            $exchangeTableHtml += "<div class='info-card'><span class='info-label'>Domain</span><span class='info-value'><code>$(HtmlEnc $a.Domain)</code></span></div>"
        }
        $exchangeTableHtml += '</div>'
    }
}

# ---- PKI table ----
$pkiTableHtml = ''
if (-not $Atlas.PKI.Detected) {
    $pkiTableHtml = '<p class="empty-note">No Certificate Services detected in AD.</p>'
} else {
    $pkiTableHtml = "<div class='info-grid'><div class='info-card'><span class='info-label'>Enterprise / Issuing CAs</span><span class='info-value'>$($Atlas.PKI.EnterpriseCAs.Count)</span></div><div class='info-card'><span class='info-label'>Trusted Root CAs</span><span class='info-value'>$($Atlas.PKI.RootCAs.Count)</span></div><div class='info-card'><span class='info-label'>Certificate Templates</span><span class='info-value'>$($Atlas.PKI.Templates.Count)</span></div></div>"
    if ($Atlas.PKI.EnterpriseCAs.Count -gt 0) {
        $pkiTableHtml += "<h3 class='sub-header'>Enterprise CAs</h3><div class='table-wrap'><table><thead><tr><th>CA Name</th><th>Server</th><th>Templates Published</th><th>Created</th></tr></thead><tbody>"
        foreach ($ca in $Atlas.PKI.EnterpriseCAs) {
            $tplCount = $ca.PublishedTemplates.Count
            $pkiTableHtml += "<tr><td><strong>$(HtmlEnc $ca.Name)</strong></td><td>$(HtmlEnc $ca.Server)</td><td style='text-align:center'>$tplCount</td><td>$(HtmlEnc $ca.Created)</td></tr>"
        }
        $pkiTableHtml += '</tbody></table></div>'
    }
    if ($Atlas.PKI.RootCAs.Count -gt 0) {
        $pkiTableHtml += "<h3 class='sub-header'>Trusted Root CAs</h3><div class='table-wrap'><table><thead><tr><th>Name</th><th>Created</th></tr></thead><tbody>"
        foreach ($r in $Atlas.PKI.RootCAs) {
            $pkiTableHtml += "<tr><td><strong>$(HtmlEnc $r.Name)</strong></td><td>$(HtmlEnc $r.Created)</td></tr>"
        }
        $pkiTableHtml += '</tbody></table></div>'
    }
    if ($Atlas.PKI.Templates.Count -gt 0) {
        $pkiTableHtml += "<h3 class='sub-header'>Certificate Templates ($($Atlas.PKI.Templates.Count))</h3><div class='table-wrap' style='max-height:400px'><table><thead><tr><th>Display Name</th><th>Template Name</th><th>Schema Version</th></tr></thead><tbody>"
        foreach ($t in ($Atlas.PKI.Templates | Sort-Object DisplayName)) {
            $pkiTableHtml += "<tr><td><strong>$(HtmlEnc $t.DisplayName)</strong></td><td><code style='font-size:.7rem'>$(HtmlEnc $t.Name)</code></td><td style='text-align:center'>$($t.SchemaVersion)</td></tr>"
        }
        $pkiTableHtml += '</tbody></table></div>'
    }
}

# ---- Authentication table ----
$authTableHtml = ''
if ($Atlas.Authentication.Domains.Count -eq 0) {
    $authTableHtml = '<p class="empty-note">No authentication data collected.</p>'
} else {
    foreach ($a in $Atlas.Authentication.Domains) {
        if ($a.EncryptionTypesDisplay.Count -gt 0) {
            $encDisplay = HtmlEnc ($a.EncryptionTypesDisplay -join ', ')
        } elseif ($a.EncryptionSource) {
            $encDisplay = "<span class='dim' style='font-size:.74rem'>$(HtmlEnc $a.EncryptionSource)</span>"
        } else {
            $encDisplay = '<span class="dim">(not configured)</span>'
        }
        $encRaw = if ($a.SupportedEncryptionTypes) { " <span class='dim' style='font-size:.7rem'>(0x$([Convert]::ToString([int]$a.SupportedEncryptionTypes,16)))</span>" } else { '' }
        $encSource = if ($a.EncryptionTypesDisplay.Count -gt 0 -and $a.EncryptionSource) { "<div style='font-size:.66rem;color:var(--text-dim);margin-top:3px;font-style:italic'>Source: $(HtmlEnc $a.EncryptionSource)</div>" } else { '' }
        $maq = if ($null -ne $a.MachineAccountQuota) { $a.MachineAccountQuota.ToString() } else { '<span class="dim">(default 10)</span>' }
        $krbReset = if ($a.KrbtgtLastReset) { HtmlEnc $a.KrbtgtLastReset } else { '<span class="dim">--</span>' }
        $krbEnabled = if ($null -eq $a.KrbtgtEnabled) { '<span class="dim">--</span>' } elseif ($a.KrbtgtEnabled) { '<span class="badge badge-green">Enabled</span>' } else { '<span class="badge badge-amber">Disabled</span>' }
        $puCount = if ($null -ne $a.ProtectedUsersCount) { [string]$a.ProtectedUsersCount } else { '<span class="dim">--</span>' }
        $pwMembers = if ($a.PreWin2kCompatMembers.Count -gt 0) { HtmlEnc ($a.PreWin2kCompatMembers -join ', ') } else { '<span class="dim">(empty)</span>' }

        $authTableHtml += "<div class='section'><h3 class='sub-header'>Domain: $(HtmlEnc $a.Domain) <span class='dim' style='font-weight:normal'>($(HtmlEnc $a.NetBIOS))</span></h3>"

        # KRBTGT + encryption + MAQ
        $authTableHtml += "<h4 style='font-size:.82rem;color:var(--accent2);margin:12px 0 6px;text-transform:uppercase;letter-spacing:.05em'>Kerberos / KRBTGT</h4><div class='info-grid'>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>KRBTGT Password Last Reset</span><span class='info-value'>$krbReset</span></div>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>KRBTGT Account Created</span><span class='info-value'>$(HtmlEnc $a.KrbtgtCreated)</span></div>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>KRBTGT Account Status</span><span class='info-value'>$krbEnabled</span></div>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>Supported Encryption Types</span><span class='info-value' style='font-size:.78rem'>$encDisplay$encRaw</span>$encSource</div>"
        $authTableHtml += '</div>'

        # Kerberos Policy
        if ($a.KerbPolicy) {
            $kp = $a.KerbPolicy
            $authTableHtml += "<h4 style='font-size:.82rem;color:var(--accent2);margin:14px 0 6px;text-transform:uppercase;letter-spacing:.05em'>Kerberos Policy <span class='dim' style='font-size:.7rem;text-transform:none'>(from Default Domain Policy GPO)</span></h4><div class='info-grid'>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Max User Ticket Lifetime</span><span class='info-value'>$($kp.MaxTicketAgeHours) hours</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Max Service Ticket Lifetime</span><span class='info-value'>$($kp.MaxServiceAgeMinutes) min</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Max Renewal Age</span><span class='info-value'>$($kp.MaxRenewAgeDays) days</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Max Clock Skew</span><span class='info-value'>$($kp.MaxClockSkewMinutes) min</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Validate Client Tickets</span><span class='info-value'>$(if($kp.TicketValidateClient -eq 1){'Yes'}else{'No'})</span></div>"
            $authTableHtml += '</div>'
        } else {
            $authTableHtml += "<h4 style='font-size:.82rem;color:var(--accent2);margin:14px 0 6px;text-transform:uppercase;letter-spacing:.05em'>Kerberos Policy</h4><p class='section-desc'><span class='dim'>Could not parse Default Domain Policy GPT.inf from SYSVOL for this domain (file unreachable or section missing).</span></p>"
        }

        # Password Policy
        if ($a.PwdPolicy) {
            $pp = $a.PwdPolicy
            $cmplx = if ($pp.ComplexityEnabled) { '<span class="badge badge-green">Yes</span>' } else { '<span class="badge badge-amber">No</span>' }
            $revEnc = if ($pp.ReversibleEncryptionEnabled) { '<span class="badge badge-amber">Yes (insecure)</span>' } else { '<span class="badge badge-green">No</span>' }
            $authTableHtml += "<h4 style='font-size:.82rem;color:var(--accent2);margin:14px 0 6px;text-transform:uppercase;letter-spacing:.05em'>Default Domain Password Policy</h4><div class='info-grid'>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Min Password Length</span><span class='info-value'>$($pp.MinPasswordLength) chars</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Password History</span><span class='info-value'>$($pp.PasswordHistoryCount) remembered</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Max Password Age</span><span class='info-value'>$($pp.MaxPasswordAgeDays) days</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Min Password Age</span><span class='info-value'>$($pp.MinPasswordAgeDays) days</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Complexity</span><span class='info-value'>$cmplx</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Reversible Encryption</span><span class='info-value'>$revEnc</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Lockout Threshold</span><span class='info-value'>$(if($pp.LockoutThreshold -gt 0){"$($pp.LockoutThreshold) bad attempts"}else{'<span class=dim>No lockout</span>'})</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Lockout Duration</span><span class='info-value'>$($pp.LockoutDurationMinutes) min</span></div>"
            $authTableHtml += "<div class='info-card'><span class='info-label'>Lockout Observation</span><span class='info-value'>$($pp.LockoutObservationMinutes) min</span></div>"
            $authTableHtml += '</div>'
        }

        # Domain-wide settings
        $authTableHtml += "<h4 style='font-size:.82rem;color:var(--accent2);margin:14px 0 6px;text-transform:uppercase;letter-spacing:.05em'>Domain-Wide Settings</h4><div class='info-grid'>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>Machine Account Quota</span><span class='info-value'>$maq</span></div>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>Protected Users Group</span><span class='info-value'>$puCount member(s)</span></div>"
        $authTableHtml += "<div class='info-card' style='grid-column:span 2'><span class='info-label'>Pre-Windows 2000 Compatible Access Members</span><span class='info-value' style='font-size:.78rem'>$pwMembers</span></div>"
        $authTableHtml += '</div>'

        # FGPPs
        if ($a.FGPPs.Count -gt 0) {
            $authTableHtml += "<h4 style='font-size:.82rem;color:var(--accent2);margin:14px 0 6px;text-transform:uppercase;letter-spacing:.05em'>Fine-Grained Password Policies ($($a.FGPPs.Count))</h4><div class='table-wrap'><table><thead><tr><th>Name</th><th>Precedence</th><th>Min Length</th><th>History</th><th>Max Age (days)</th><th>Complexity</th><th>Lockout</th><th>Applies To</th></tr></thead><tbody>"
            foreach ($f in ($a.FGPPs | Sort-Object Precedence)) {
                $fApplies = if ($f.AppliesTo.Count -gt 0) { HtmlEnc ($f.AppliesTo -join ', ') } else { '<span class="dim">(none)</span>' }
                $fCmplx = if ($f.ComplexityEnabled) { 'Yes' } else { 'No' }
                $authTableHtml += "<tr><td><strong>$(HtmlEnc $f.Name)</strong></td><td style='text-align:center'>$($f.Precedence)</td><td style='text-align:center'>$($f.MinPasswordLength)</td><td style='text-align:center'>$($f.PasswordHistoryCount)</td><td style='text-align:center'>$($f.MaxPasswordAgeDays)</td><td style='text-align:center'>$fCmplx</td><td style='text-align:center'>$($f.LockoutThreshold)</td><td>$fApplies</td></tr>"
            }
            $authTableHtml += '</tbody></table></div>'
        } else {
            $authTableHtml += "<h4 style='font-size:.82rem;color:var(--accent2);margin:14px 0 6px;text-transform:uppercase;letter-spacing:.05em'>Fine-Grained Password Policies</h4><p class='section-desc'><span class='dim'>None defined in this domain.</span></p>"
        }

        $authTableHtml += '</div>'
    }

    if ($Atlas.Authentication.Policies.Count -gt 0 -or $Atlas.Authentication.Silos.Count -gt 0) {
        $authTableHtml += "<div class='section'><h3 class='sub-header'>Authentication Policies &amp; Silos (Forest-Wide)</h3><div class='info-grid'>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>Authentication Policies</span><span class='info-value'>$($Atlas.Authentication.Policies.Count)</span></div>"
        $authTableHtml += "<div class='info-card'><span class='info-label'>Authentication Policy Silos</span><span class='info-value'>$($Atlas.Authentication.Silos.Count)</span></div>"
        $authTableHtml += '</div>'
        if ($Atlas.Authentication.Policies.Count -gt 0) {
            $authTableHtml += "<h4 style='font-size:.85rem;color:var(--text);margin:14px 0 8px'>Policies</h4><ul style='margin-left:20px;color:var(--text-dim);font-size:.82rem'>"
            foreach ($p in $Atlas.Authentication.Policies) { $authTableHtml += "<li>$(HtmlEnc $p.Name)</li>" }
            $authTableHtml += '</ul>'
        }
        if ($Atlas.Authentication.Silos.Count -gt 0) {
            $authTableHtml += "<h4 style='font-size:.85rem;color:var(--text);margin:14px 0 8px'>Silos</h4><ul style='margin-left:20px;color:var(--text-dim);font-size:.82rem'>"
            foreach ($s in $Atlas.Authentication.Silos) { $authTableHtml += "<li>$(HtmlEnc $s.Name)</li>" }
            $authTableHtml += '</ul>'
        }
        $authTableHtml += '</div>'
    }
}

# ==============================================================================
# HTML DOCUMENT - HEAD + STYLES
# ==============================================================================

$Html = @"
<!--
================================================================================
 ADAtlas -- Map Every Corner of Your Active Directory
 Version  : 1.0
 Generated: $($Atlas.GeneratedAt)
 By       : $($Atlas.GeneratedBy) on $($Atlas.GeneratedOn)
 Author   : Santhosh Sivarajan, Microsoft MVP
 LinkedIn : https://www.linkedin.com/in/sivarajan/
 GitHub   : https://github.com/SanthoshSivarajan/ADAtlas
 License  : MIT
================================================================================
-->
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<meta name="author" content="Santhosh Sivarajan, Microsoft MVP"/>
<title>ADAtlas -- $(HtmlEnc $fw.Name)</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0f172a;--surface:#1e293b;--surface2:#273548;--border:#334155;
  --text:#e2e8f0;--text-dim:#94a3b8;--accent:#60a5fa;--accent2:#22d3ee;
  --green:#34d399;--red:#f87171;--amber:#fbbf24;--purple:#a78bfa;
  --pink:#f472b6;--orange:#fb923c;--accent-bg:rgba(96,165,250,.1);
  --radius:8px;--shadow:0 1px 3px rgba(0,0,0,.3);
  --font-body:'Segoe UI',system-ui,-apple-system,sans-serif;
}
html{scroll-behavior:smooth;font-size:15px}
body{font-family:var(--font-body);background:var(--bg);color:var(--text);line-height:1.65;min-height:100vh}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}
.dim{color:var(--text-dim)}
code{font-family:'Consolas','Monaco',monospace;font-size:.85em;color:var(--accent2)}
.wrapper{display:flex;min-height:100vh}

.sidebar{position:fixed;top:0;left:0;width:260px;height:100vh;background:var(--surface);border-right:1px solid var(--border);overflow-y:auto;padding:20px 0;z-index:100;box-shadow:2px 0 12px rgba(0,0,0,.3)}
.sidebar::-webkit-scrollbar{width:4px}
.sidebar::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px}
.sidebar .logo{padding:0 18px 14px;border-bottom:1px solid var(--border);margin-bottom:8px}
.sidebar .logo h2{font-size:1.1rem;color:var(--accent);font-weight:800;letter-spacing:.02em}
.sidebar .logo .tagline{font-size:.66rem;color:var(--text-dim);margin-top:3px;font-style:italic}
.sidebar .logo .by{font-size:.66rem;color:var(--text-dim);margin-top:6px}
.sidebar .logo .forest{font-size:.72rem;color:var(--text);margin-top:8px;word-break:break-all}
.sidebar .logo .forest strong{color:var(--accent2)}
.sidebar nav a{display:block;padding:6px 18px 6px 22px;font-size:.8rem;color:var(--text-dim);border-left:3px solid transparent;transition:all .15s;cursor:pointer}
.sidebar nav a:hover,.sidebar nav a.active{color:var(--accent);background:rgba(96,165,250,.08);border-left-color:var(--accent);text-decoration:none}
.sidebar nav .nav-group{font-size:.62rem;text-transform:uppercase;letter-spacing:.08em;color:var(--accent2);padding:12px 18px 3px;font-weight:700}

.main{margin-left:260px;flex:1;padding:24px 32px 60px;max-width:1500px}
.tab-content{display:none}
.tab-content.active{display:block;animation:fadeIn .2s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}

.section{margin-bottom:36px}
.section-title{font-size:1.28rem;font-weight:700;color:var(--text);margin-bottom:4px;padding-bottom:8px;border-bottom:2px solid var(--border);display:flex;align-items:center;gap:10px}
.section-title .icon{width:26px;height:26px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:.85rem;flex-shrink:0}
.sub-header{font-size:.95rem;color:var(--text);margin:18px 0 10px;padding-bottom:4px;border-bottom:1px solid var(--border);font-weight:600}
.section-desc{color:var(--text-dim);font-size:.82rem;margin-bottom:14px}
.empty-note{color:var(--text-dim);font-style:italic;padding:10px 0}

.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:16px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;box-shadow:var(--shadow);transition:border-color .15s}
.card:hover{border-color:var(--accent)}
.card .card-val{font-size:1.55rem;font-weight:800;line-height:1.1}
.card .card-label{font-size:.68rem;color:var(--text-dim);margin-top:3px;text-transform:uppercase;letter-spacing:.05em}

.info-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:8px}
.info-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:10px 14px;box-shadow:var(--shadow)}
.info-label{display:block;font-size:.68rem;color:var(--text-dim);text-transform:uppercase;letter-spacing:.05em;margin-bottom:3px}
.info-value{font-size:.92rem;font-weight:600;color:var(--text);word-break:break-word}

.table-wrap{overflow-x:auto;margin-bottom:8px;border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);max-height:78vh}
table{width:100%;border-collapse:collapse;font-size:.8rem}
thead{background:var(--accent-bg)}
th{text-align:left;padding:9px 11px;font-weight:600;color:var(--accent);white-space:nowrap;border-bottom:2px solid var(--border);position:sticky;top:0;background:#1e2a44;z-index:2}
th[data-sort]{cursor:pointer;user-select:none}
th[data-sort]:hover{color:var(--accent2)}
th[data-sort].sorted-asc::after{content:' \25B2';font-size:.7em;color:var(--accent2)}
th[data-sort].sorted-desc::after{content:' \25BC';font-size:.7em;color:var(--accent2)}
td{padding:8px 11px;border-bottom:1px solid var(--border);color:var(--text-dim);vertical-align:top}
tbody tr:hover{background:rgba(96,165,250,.06)}
tbody tr:nth-child(even){background:var(--surface2)}

.exec-summary{background:linear-gradient(135deg,#1e293b 0%,#1e3a5f 100%);border:1px solid #334155;border-radius:var(--radius);padding:22px 26px;margin-bottom:28px;box-shadow:var(--shadow)}
.exec-summary h2{font-size:1.15rem;color:var(--accent);margin-bottom:8px;font-weight:700}
.exec-summary p{color:var(--text-dim);font-size:.86rem;line-height:1.7;margin-bottom:6px}
.exec-kv{display:inline-block;background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:3px 9px;margin:3px 2px;font-size:.76rem;color:var(--text)}
.exec-kv strong{color:var(--accent2)}

.fsmo-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin-bottom:14px}
.fsmo-card{background:var(--surface2);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px;text-align:center}
.fsmo-card .role{font-size:.66rem;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);margin-bottom:5px;font-weight:700}
.fsmo-card .holder{font-size:.84rem;color:var(--text);font-weight:600;word-break:break-all}

.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:.66rem;font-weight:700;text-transform:uppercase;letter-spacing:.04em;vertical-align:middle;margin-left:4px}
.badge-accent{background:rgba(96,165,250,.15);color:var(--accent);border:1px solid rgba(96,165,250,.4)}
.badge-green{background:rgba(52,211,153,.15);color:var(--green);border:1px solid rgba(52,211,153,.4)}
.badge-amber{background:rgba(251,191,36,.15);color:var(--amber);border:1px solid rgba(251,191,36,.4)}
.badge-purple{background:rgba(167,139,250,.15);color:var(--purple);border:1px solid rgba(167,139,250,.4)}

.diagram-wrap{display:grid;grid-template-columns:1fr 320px;gap:16px;margin-top:10px}
@media(max-width:1200px){.diagram-wrap{grid-template-columns:1fr}}
.diagram-canvas-wrap{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px;overflow:auto;max-height:82vh;box-shadow:var(--shadow)}
.diagram-canvas-wrap svg{width:100%;height:auto;display:block;min-width:600px}
.diagram-legend{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:14px;padding:12px 16px;background:var(--surface2);border:1px solid var(--border);border-radius:var(--radius);font-size:.76rem;color:var(--text-dim)}
.legend-item{display:flex;align-items:center;gap:7px}
.legend-swatch{display:inline-block;width:26px;height:3px;border-radius:2px;background:var(--text)}
.legend-swatch.dashed{background:transparent;height:0;border-top:2px dashed var(--text)}
.legend-node{display:inline-block;width:14px;height:14px;border-radius:3px;border:2px solid;background:transparent}
.legend-node.circle{border-radius:50%}
.legend-node.triangle{border:none;width:0;height:0;border-left:9px solid transparent;border-right:9px solid transparent;border-bottom:14px solid var(--accent)}
.legend-node.diamond{width:16px;height:16px;background:linear-gradient(135deg,#22d3ee 0%,#0ea5e9 50%,#0369a1 100%);transform:rotate(45deg);border:none;border-radius:2px}

.side-panel{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:16px 18px;box-shadow:var(--shadow);height:fit-content;position:sticky;top:20px;max-height:82vh;overflow-y:auto}
.side-panel::-webkit-scrollbar{width:6px}
.side-panel::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px}
.side-panel h3{font-size:.95rem;color:var(--accent);margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid var(--border);word-break:break-all;font-weight:700}
.side-panel .field{margin-bottom:9px}
.side-panel .field label{display:block;font-size:.64rem;text-transform:uppercase;letter-spacing:.05em;color:var(--text-dim);margin-bottom:2px;font-weight:600}
.side-panel .field .val{font-size:.8rem;color:var(--text);word-break:break-word}
.side-panel .empty{color:var(--text-dim);font-style:italic;font-size:.8rem;padding:30px 0;text-align:center}
.side-panel .dc-list{margin-top:4px}
.side-panel .dc-list .dc-item{padding:4px 0;font-size:.76rem;color:var(--text);border-bottom:1px dashed var(--border)}
.side-panel .dc-list .dc-item:last-child{border-bottom:none}
.side-panel .dc-list .dc-item .dc-meta{color:var(--text-dim);font-size:.7rem}
.side-panel .pill{display:inline-block;padding:1px 7px;border-radius:9px;font-size:.62rem;font-weight:700;margin-left:3px;text-transform:uppercase;letter-spacing:.04em}
.side-panel .pill-fsmo{background:rgba(167,139,250,.18);color:var(--purple);border:1px solid rgba(167,139,250,.5)}
.side-panel .pill-gc{background:rgba(34,211,238,.15);color:var(--accent2);border:1px solid rgba(34,211,238,.45)}
.side-panel .pill-rodc{background:rgba(251,191,36,.15);color:var(--amber);border:1px solid rgba(251,191,36,.45)}

.node-box{cursor:pointer;transition:filter .15s}
.node-box:hover{filter:brightness(1.25)}
.node-box.selected polygon,.node-box.selected circle,.node-box.selected rect,.node-box.selected path{stroke-width:4 !important}
.trust-edge,.repl-edge{cursor:pointer;transition:stroke-width .15s}
.trust-edge:hover,.repl-edge:hover{stroke-width:4}
.trust-edge.selected,.repl-edge.selected{stroke-width:4}

.matrix-wrap{overflow:auto;max-width:100%;border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);max-height:80vh}
.matrix-table{border-collapse:collapse;font-size:.74rem;background:var(--bg)}
.matrix-table th,.matrix-table td{border:1px solid var(--border);padding:8px 10px;text-align:center;min-width:130px}
.matrix-table th.corner{background:#0f172a;color:var(--accent2);font-weight:700;position:sticky;top:0;left:0;z-index:5}
.matrix-table th.col-head{background:var(--surface2);color:var(--accent);font-weight:600;position:sticky;top:0;z-index:3;white-space:nowrap}
.matrix-table th.row-head{background:var(--surface2);color:var(--accent);font-weight:600;position:sticky;left:0;z-index:3;text-align:left;white-space:nowrap}
.matrix-table td{background:var(--surface);color:var(--text-dim)}
.matrix-cell-intra{background:rgba(52,211,153,.12) !important;color:var(--green) !important;font-weight:700}
.matrix-cell-external{background:rgba(251,191,36,.12) !important;color:var(--amber) !important;font-weight:700}
.matrix-cell-self{background:#0f172a !important;color:var(--text-dim) !important}
.matrix-cell-none{background:var(--surface) !important;color:#475569 !important}

.dc-controls{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:14px;align-items:center}
.dc-search{flex:1;min-width:240px;background:var(--surface);border:1px solid var(--border);color:var(--text);padding:9px 14px;border-radius:var(--radius);font-family:var(--font-body);font-size:.84rem;outline:none}
.dc-search:focus{border-color:var(--accent)}
.dc-filter{background:var(--surface);border:1px solid var(--border);color:var(--text);padding:9px 12px;border-radius:var(--radius);font-family:var(--font-body);font-size:.8rem;outline:none;cursor:pointer;min-width:140px}
.dc-filter:focus{border-color:var(--accent)}
.dc-count{font-size:.78rem;color:var(--text-dim);margin-left:auto;font-weight:600}
.dc-count strong{color:var(--accent2)}

.repl-controls{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:14px;align-items:center}

.footer{margin-top:50px;padding:22px 0 8px;border-top:1px solid var(--border);text-align:center;color:var(--text-dim);font-size:.76rem;line-height:1.7}
.footer a{color:var(--accent)}
.footer .brand{font-weight:700;color:var(--accent);font-size:.88rem;margin-bottom:4px;display:block}

@media print{
  .sidebar{display:none}
  .main{margin-left:0}
  body{background:#fff;color:#222}
  .card,.info-card,.exec-summary,.fsmo-card,.diagram-canvas-wrap,.side-panel,.diagram-legend{background:#f9f9f9;border-color:#ccc;color:#222}
  .card-val,.info-value,.section-title,.sub-header{color:#222}
  .card-label,.info-label,.section-desc{color:#555}
  th{color:#333;background:#eee}
  td{color:#444}
  .tab-content{display:block !important;page-break-before:always}
}
@media(max-width:900px){.sidebar{display:none}.main{margin-left:0;padding:14px}}
</style>
</head>
<body>
<div class="wrapper">

<aside class="sidebar">
  <div class="logo">
    <h2>ADAtlas</h2>
    <div class="tagline">Map Every Corner of Your Active Directory</div>
    <div class="by">Developed by Santhosh Sivarajan</div>
    <div class="forest">Forest: <strong>$(HtmlEnc $fw.Name)</strong></div>
  </div>
  <nav>
    <div class="nav-group">Overview</div>
    <a href="#" data-tab="tab-overview" class="active">Executive Summary</a>

    <div class="nav-group">Forest &amp; Trusts</div>
    <a href="#" data-tab="tab-forest">Forest &amp; Domains Map</a>
    <a href="#" data-tab="tab-trusts">Trust Map</a>
    <a href="#" data-tab="tab-trust-matrix">Trust Matrix</a>

    <div class="nav-group">Sites &amp; Replication</div>
    <a href="#" data-tab="tab-sites">Site Topology</a>
    <a href="#" data-tab="tab-sitelinks">Site Links</a>
    <a href="#" data-tab="tab-replication">Replication Topology</a>
    <a href="#" data-tab="tab-subnets">Sites &amp; Subnets</a>

    <div class="nav-group">Domain Controllers</div>
    <a href="#" data-tab="tab-dcs">DC Inventory</a>

    <div class="nav-group">Supporting Services</div>
    <a href="#" data-tab="tab-dns">DNS Architecture</a>
    <a href="#" data-tab="tab-ntp">NTP Hierarchy</a>
    <a href="#" data-tab="tab-exchange">Exchange</a>
    <a href="#" data-tab="tab-pki">Certificate Services</a>
    <a href="#" data-tab="tab-auth">Authentication</a>

    <div class="nav-group">Jump to Domain</div>
$domainNavHtml
  </nav>
</aside>

<main class="main">

<!-- ============================ OVERVIEW TAB ============================ -->
<div id="tab-overview" class="tab-content active">
  <div class="exec-summary">
    <h2>Executive Summary &mdash; $(HtmlEnc $fw.Name)</h2>
    <p>Point-in-time topology map of the Active Directory forest <strong>$(HtmlEnc $fw.Name)</strong>, generated on <strong>$(Get-Date -Date $Atlas.GeneratedAt -Format 'MMMM dd, yyyy HH:mm')</strong>. This report documents the current configuration of forest structure, domains, DCs, sites, trusts, replication, Entra Connect, DNS, NTP hierarchy, Exchange, Certificate Services, and authentication. No health checks or analysis &mdash; this is a picture of what exists today.</p>
    <p>
      <span class="exec-kv"><strong>Forest:</strong> $(HtmlEnc $fw.Name)</span>
      <span class="exec-kv"><strong>Forest Level:</strong> $(HtmlEnc $fw.ForestModeDisplay)</span>
      <span class="exec-kv"><strong>Schema:</strong> $($fw.SchemaVersion) ($(HtmlEnc $fw.SchemaOS))</span>
      <span class="exec-kv"><strong>Domains:</strong> $($Atlas.Domains.Count)</span>
      <span class="exec-kv"><strong>DCs:</strong> $($Atlas.DCs.Count)</span>
      <span class="exec-kv"><strong>Sites:</strong> $($Atlas.Sites.Count)</span>
      <span class="exec-kv"><strong>Subnets:</strong> $($Atlas.Subnets.Count)</span>
      <span class="exec-kv"><strong>Site Links:</strong> $($Atlas.SiteLinks.Count)</span>
      <span class="exec-kv"><strong>Trusts:</strong> $($Atlas.Trusts.Count)</span>
      <span class="exec-kv"><strong>DNS Zones:</strong> $($Atlas.DNS.Zones.Count)</span>
      <span class="exec-kv"><strong>Entra Connect:</strong> $(if($Atlas.EntraConnect.Detected){'Detected'}else{'Not detected'})</span>
      <span class="exec-kv"><strong>Exchange:</strong> $(if($Atlas.Exchange.Detected){"$($Atlas.Exchange.Servers.Count) server(s)"}else{'Not detected'})</span>
      <span class="exec-kv"><strong>PKI:</strong> $(if($Atlas.PKI.Detected){"$($Atlas.PKI.EnterpriseCAs.Count) CA(s)"}else{'Not detected'})</span>
    </p>
  </div>

  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(96,165,250,.15);color:var(--accent)">&#128200;</span> Topology at a Glance</h2>
    <div class="cards">
      <div class="card"><div class="card-val" style="color:var(--accent)">$($Atlas.Domains.Count)</div><div class="card-label">Domains</div></div>
      <div class="card"><div class="card-val" style="color:var(--accent2)">$($Atlas.DCs.Count)</div><div class="card-label">Domain Controllers</div></div>
      <div class="card"><div class="card-val" style="color:var(--green)">$rwdcCount</div><div class="card-label">RWDC</div></div>
      <div class="card"><div class="card-val" style="color:var(--amber)">$rodcCount</div><div class="card-label">RODC</div></div>
      <div class="card"><div class="card-val" style="color:var(--purple)">$gcCount</div><div class="card-label">Global Catalogs</div></div>
      <div class="card"><div class="card-val" style="color:var(--accent)">$($Atlas.Sites.Count)</div><div class="card-label">Sites</div></div>
      <div class="card"><div class="card-val" style="color:var(--accent2)">$($Atlas.Subnets.Count)</div><div class="card-label">Subnets</div></div>
      <div class="card"><div class="card-val" style="color:var(--pink)">$($Atlas.SiteLinks.Count)</div><div class="card-label">Site Links</div></div>
      <div class="card"><div class="card-val" style="color:var(--green)">$intraTrusts</div><div class="card-label">Intra-Forest Trusts</div></div>
      <div class="card"><div class="card-val" style="color:var(--amber)">$extTrusts</div><div class="card-label">External Trusts</div></div>
      <div class="card"><div class="card-val" style="color:var(--purple)">$($Atlas.Connections.Count)</div><div class="card-label">Replication Conns</div></div>
      <div class="card"><div class="card-val" style="color:var(--accent2)">$($Atlas.DNS.Zones.Count)</div><div class="card-label">DNS Zones</div></div>
    </div>
  </div>

  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(34,211,238,.15);color:var(--accent2)">&#127794;</span> Forest Configuration</h2>
    <div class="info-grid">
      <div class="info-card"><span class="info-label">Forest Name</span><span class="info-value">$(HtmlEnc $fw.Name)</span></div>
      <div class="info-card"><span class="info-label">Root Domain</span><span class="info-value">$(HtmlEnc $fw.RootDomain)</span></div>
      <div class="info-card"><span class="info-label">Forest Functional Level</span><span class="info-value">$(HtmlEnc $fw.ForestModeDisplay)</span></div>
      <div class="info-card"><span class="info-label">Schema Version</span><span class="info-value">$($fw.SchemaVersion) &mdash; $(HtmlEnc $fw.SchemaOS)</span></div>
      <div class="info-card"><span class="info-label">Tombstone Lifetime</span><span class="info-value">$($fw.TombstoneLifetime) days</span></div>
      <div class="info-card"><span class="info-label">Total Domains</span><span class="info-value">$($Atlas.Domains.Count)</span></div>
      <div class="info-card"><span class="info-label">Global Catalogs</span><span class="info-value">$($fw.GlobalCatalogs.Count) GC(s)</span></div>
      <div class="info-card"><span class="info-label">UPN Suffixes</span><span class="info-value">$(if($fw.UPNSuffixes.Count -gt 0){HtmlEnc ($fw.UPNSuffixes -join ', ')}else{'<span class=dim>(default only)</span>'})</span></div>
      <div class="info-card"><span class="info-label">SPN Suffixes</span><span class="info-value">$(if($fw.SPNSuffixes.Count -gt 0){HtmlEnc ($fw.SPNSuffixes -join ', ')}else{'<span class=dim>(none)</span>'})</span></div>
      $entraOverviewHtml
    </div>
  </div>

  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(167,139,250,.15);color:var(--purple)">&#9733;</span> Forest-Wide FSMO Roles</h2>
    $fsmoHtml
  </div>

  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(96,165,250,.15);color:var(--accent)">&#127760;</span> Domain Summary</h2>
    <div class="table-wrap"><table>
      <thead><tr><th>Domain</th><th>NetBIOS</th><th>Functional Level</th><th>Parent</th><th>Children</th><th>DCs</th><th>PDC Emulator</th></tr></thead>
      <tbody>$domainRowsHtml</tbody>
    </table></div>
  </div>
</div>

<!-- ============================ FOREST MAP ============================ -->
<div id="tab-forest" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(96,165,250,.15);color:var(--accent)">&#127794;</span> Forest &amp; Domains Map</h2>
    <p class="section-desc">Hierarchical view of the forest using the classic AD triangle symbol for domains. Click any domain triangle to see its details. Entra ID (if detected) appears as a hexagonal cloud node attached to the forest root.</p>
    <div class="diagram-legend">
      <div class="legend-item"><span class="legend-node triangle" style="border-bottom-color:var(--accent)"></span> Forest Root Domain</div>
      <div class="legend-item"><span class="legend-node triangle" style="border-bottom-color:var(--accent2)"></span> Tree Root / Parent</div>
      <div class="legend-item"><span class="legend-node triangle" style="border-bottom-color:var(--purple)"></span> Child Domain</div>
      <div class="legend-item"><span class="legend-swatch" style="background:#475569"></span> Parent &rarr; Child</div>
      <div class="legend-item"><span class="legend-node diamond"></span> Entra ID (if detected)</div>
    </div>
    <div class="diagram-wrap">
      <div class="diagram-canvas-wrap"><div id="forest-canvas"></div></div>
      <div class="side-panel" id="forest-panel"><div class="empty">Click a domain triangle to see details.</div></div>
    </div>
  </div>
</div>

<!-- ============================ TRUST MAP (REDESIGNED) ============================ -->
<div id="tab-trusts" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(251,191,36,.15);color:var(--amber)">&#128279;</span> Trust Map</h2>
    <p class="section-desc">Professional hierarchical view of all trust relationships. Forest domains are arranged as a tree on the left; external trust partners are grouped on the right. Trust lines are routed as smooth curves to keep the diagram readable.</p>
    <div class="diagram-legend">
      <div class="legend-item"><span class="legend-node triangle" style="border-bottom-color:var(--accent)"></span> Forest Root</div>
      <div class="legend-item"><span class="legend-node triangle" style="border-bottom-color:var(--accent2)"></span> Forest Domain (parent)</div>
      <div class="legend-item"><span class="legend-node triangle" style="border-bottom-color:var(--purple)"></span> Forest Domain (child)</div>
      <div class="legend-item"><span class="legend-node triangle" style="border-bottom-color:var(--amber)"></span> External Domain</div>
      <div class="legend-item"><span class="legend-swatch" style="background:var(--green)"></span> Intra-Forest Trust</div>
      <div class="legend-item"><span class="legend-swatch dashed" style="border-top-color:var(--amber)"></span> External / Forest Trust</div>
    </div>
    <div class="diagram-wrap">
      <div class="diagram-canvas-wrap"><div id="trust-canvas"></div></div>
      <div class="side-panel" id="trust-panel"><div class="empty">Click a trust line or a domain to see details.</div></div>
    </div>
  </div>
</div>

<!-- ============================ TRUST MATRIX ============================ -->
<div id="tab-trust-matrix" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(52,211,153,.15);color:var(--green)">&#9783;</span> Trust Matrix</h2>
    <p class="section-desc">Grid view of all trust relationships. Rows are source, columns are target. Hover any cell for details.</p>
    $matrixHtml
  </div>
</div>

<!-- ============================ SITE TOPOLOGY ============================ -->
<div id="tab-sites" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(34,211,238,.15);color:var(--accent2)">&#9737;</span> Site Topology</h2>
    <p class="section-desc">All AD sites shown as circles, sized by DC count, connected by site links. Click a site to see its DCs, subnets, and site link membership.</p>
    <div class="diagram-legend">
      <div class="legend-item"><span class="legend-node circle" style="border-color:var(--accent2);background:rgba(34,211,238,.18)"></span> Site (size = DC count)</div>
      <div class="legend-item"><span class="legend-node circle" style="border-color:var(--text-dim)"></span> Empty Site</div>
      <div class="legend-item"><span class="legend-swatch" style="background:var(--accent)"></span> Site Link</div>
    </div>
    <div class="diagram-wrap">
      <div class="diagram-canvas-wrap"><div id="site-canvas"></div></div>
      <div class="side-panel" id="site-panel"><div class="empty">Click a site circle to see details.</div></div>
    </div>
  </div>
</div>

<!-- ============================ SITE LINKS ============================ -->
<div id="tab-sitelinks" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(244,114,182,.15);color:var(--pink)">&#128279;</span> Site Links</h2>
    <p class="section-desc">All AD site links sorted by cost. Lower cost = preferred path.</p>
    $siteLinksHtml
  </div>
</div>

<!-- ============================ REPLICATION TOPOLOGY ============================ -->
<div id="tab-replication" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(52,211,153,.15);color:var(--green)">&#8645;</span> Replication Topology</h2>
    <p class="section-desc">Static KCC-generated and manual replication connection objects. Use the dropdown to scope by site.</p>
    <div class="repl-controls">
      <label style="font-size:.78rem;color:var(--text-dim);font-weight:600">Scope:</label>
      <select id="repl-site-select" class="dc-filter"></select>
      <span class="dc-count" id="repl-count"></span>
    </div>
    <div class="diagram-legend">
      <div class="legend-item"><span class="legend-node" style="border-color:var(--accent);background:rgba(96,165,250,.15)"></span> RWDC</div>
      <div class="legend-item"><span class="legend-node" style="border-color:var(--amber);background:rgba(251,191,36,.15)"></span> RODC</div>
      <div class="legend-item"><span class="legend-swatch" style="background:var(--green)"></span> KCC auto</div>
      <div class="legend-item"><span class="legend-swatch dashed" style="border-top-color:var(--accent2)"></span> Manual</div>
    </div>
    <div class="diagram-wrap">
      <div class="diagram-canvas-wrap"><div id="repl-canvas"></div></div>
      <div class="side-panel" id="repl-panel"><div class="empty">Click a connection or DC to see details.</div></div>
    </div>
  </div>
</div>

<!-- ============================ SITES & SUBNETS ============================ -->
<div id="tab-subnets" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(96,165,250,.15);color:var(--accent)">&#127760;</span> Sites &amp; Subnets</h2>
    <p class="section-desc">Combined view of all sites with their associated subnets and DC counts.</p>
    $sitesSubnetsHtml
  </div>
</div>

<!-- ============================ DC INVENTORY ============================ -->
<div id="tab-dcs" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(34,211,238,.15);color:var(--accent2)">&#128187;</span> Domain Controller Inventory</h2>
    <p class="section-desc">All domain controllers across the forest. Use the search box and filters to narrow down. Click any column header to sort.</p>
    $dcInventoryHtml
  </div>
</div>

<!-- ============================ DNS ARCHITECTURE ============================ -->
<div id="tab-dns" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(96,165,250,.15);color:var(--accent)">&#127760;</span> DNS Architecture</h2>
    <p class="section-desc">DNS zones hosted in the AD forest, AD-integration status, and replication scope. The diagram shows DCs (which typically also run DNS) and the zones they host. AD-integrated zones replicate via AD replication.</p>
    <div class="diagram-legend">
      <div class="legend-item"><span class="legend-node" style="border-color:var(--accent);background:rgba(96,165,250,.15)"></span> Domain Controller / DNS Server</div>
      <div class="legend-item"><span class="legend-node" style="border-color:var(--green);background:rgba(52,211,153,.15)"></span> AD-Integrated Zone</div>
      <div class="legend-item"><span class="legend-node" style="border-color:var(--amber);background:rgba(251,191,36,.15)"></span> Standalone Zone</div>
    </div>
    <div class="diagram-canvas-wrap" style="margin-bottom:18px"><div id="dns-canvas"></div></div>
    <h3 class="sub-header">DNS Zones</h3>
    $dnsTableHtml
  </div>
</div>

<!-- ============================ NTP HIERARCHY ============================ -->
<div id="tab-ntp" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(251,191,36,.15);color:var(--amber)">&#9201;</span> NTP Hierarchy</h2>
    <p class="section-desc">Actual time configuration on each domain controller, collected via <code>w32tm /query</code>. Each DC node shows its real configured time source. DCs that pointed to the same source are grouped, and the forest root PDC's external source (if reachable) appears at the top. Unreachable DCs are shown in grey.</p>
    <div class="diagram-legend">
      <div class="legend-item"><span class="legend-node" style="border-color:var(--accent2);background:rgba(34,211,238,.15)"></span> External Source (forest root PDC's upstream)</div>
      <div class="legend-item"><span class="legend-node" style="border-color:var(--accent);background:rgba(96,165,250,.18)"></span> Forest Root PDC</div>
      <div class="legend-item"><span class="legend-node" style="border-color:var(--purple);background:rgba(167,139,250,.12)"></span> Domain PDC</div>
      <div class="legend-item"><span class="legend-node" style="border-color:var(--green);background:rgba(52,211,153,.12)"></span> DC (NT5DS / domain hierarchy)</div>
      <div class="legend-item"><span class="legend-node" style="border-color:var(--text-dim);background:rgba(148,163,184,.05)"></span> Unreachable</div>
    </div>
    <div class="diagram-canvas-wrap"><div id="ntp-canvas"></div></div>
    <h3 class="sub-header">Per-DC Time Configuration</h3>
    <div id="ntp-table-wrap"></div>
  </div>
</div>

<!-- ============================ EXCHANGE ============================ -->
<div id="tab-exchange" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(34,211,238,.15);color:var(--accent2)">&#9993;</span> Exchange Architecture</h2>
    <p class="section-desc">Exchange data extracted from the AD configuration partition (read-only). Includes Exchange organization, servers, DAGs, accepted domains, and hybrid configuration detection.</p>
    <div class="diagram-canvas-wrap" style="margin-bottom:18px"><div id="exchange-canvas"></div></div>
    $exchangeTableHtml
  </div>
</div>

<!-- ============================ PKI ============================ -->
<div id="tab-pki" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(167,139,250,.15);color:var(--purple)">&#128272;</span> Certificate Services (PKI)</h2>
    <p class="section-desc">Active Directory Certificate Services (ADCS) hierarchy. Includes Trusted Root CAs, Enterprise Issuing CAs, and the Certificate Templates published in the forest.</p>
    <div class="diagram-canvas-wrap" style="margin-bottom:18px"><div id="pki-canvas"></div></div>
    $pkiTableHtml
  </div>
</div>

<!-- ============================ AUTHENTICATION ============================ -->
<div id="tab-auth" class="tab-content">
  <div class="section">
    <h2 class="section-title"><span class="icon" style="background:rgba(244,114,182,.15);color:var(--pink)">&#128274;</span> Authentication</h2>
    <p class="section-desc">Per-domain Kerberos and authentication configuration: KRBTGT account state, supported encryption types, machine account quota, and authentication policies / silos. Read directly from AD &mdash; no GPO scanning or runtime probing.</p>
    $authTableHtml
  </div>
</div>

<div class="footer">
  <span class="brand">ADAtlas &mdash; Map Every Corner of Your Active Directory</span>
  Developed by <strong>Santhosh Sivarajan</strong>, Microsoft MVP &nbsp;|&nbsp;
  <a href="https://www.linkedin.com/in/sivarajan/" target="_blank">LinkedIn</a> &nbsp;|&nbsp;
  <a href="https://github.com/SanthoshSivarajan/ADAtlas" target="_blank">GitHub</a> &nbsp;|&nbsp;
  MIT License<br/>
  Generated $(HtmlEnc $Atlas.GeneratedAt) by $(HtmlEnc $Atlas.GeneratedBy) on $(HtmlEnc $Atlas.GeneratedOn)
</div>

</main>
</div>

<script>
window.ADAtlasData = $AtlasJson;
</script>

<script>
(function(){

  function escapeXml(s){
    return String(s == null ? '' : s)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  }
  function escapeHtml(s){
    return String(s == null ? '' : s)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  }
  function field(label, value){
    var v = (value == null || value === '') ? '<span class="dim">(none)</span>' : escapeHtml(value);
    return '<div class="field"><label>' + escapeHtml(label) + '</label><div class="val">' + v + '</div></div>';
  }

  // ==================== DOMAIN TRIANGLE BUILDER ====================
  function buildDomainTriangle(opts){
    var halfW = opts.w / 2;
    var topY = opts.cy - opts.h / 2;
    var botY = opts.cy + opts.h / 2;
    var tipX = opts.cx;
    var leftX = opts.cx - halfW;
    var rightX = opts.cx + halfW;
    var points = tipX + ',' + topY + ' ' + leftX + ',' + botY + ' ' + rightX + ',' + botY;
    var parts = [];
    parts.push('<g class="node-box" ' + opts.dataAttr + '>');
    parts.push('<polygon points="' + points + '" fill="' + opts.bg + '" stroke="' + opts.color + '" stroke-width="2.5"/>');
    var textY1 = botY - 38;
    var textY2 = botY - 22;
    var textY3 = botY - 8;
    parts.push('<text x="' + opts.cx + '" y="' + textY1 + '" text-anchor="middle" font-family="Segoe UI" font-size="13" font-weight="700" fill="#e2e8f0">' + escapeXml(opts.label) + '</text>');
    if (opts.sublabel){
      parts.push('<text x="' + opts.cx + '" y="' + textY2 + '" text-anchor="middle" font-family="Segoe UI" font-size="10" fill="#94a3b8">' + escapeXml(opts.sublabel) + '</text>');
    }
    if (opts.dcText){
      parts.push('<text x="' + opts.cx + '" y="' + textY3 + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(opts.dcText) + '</text>');
    }
    parts.push('</g>');
    return { svg: parts.join(''), topX: tipX, topY: topY, botY: botY, leftX: leftX, rightX: rightX };
  }

  // ==================== ENTRA ID DIAMOND (stylized 3D cube) ====================
  function buildEntraLogo(cx, cy, tenantName){
    var parts = [];
    var s = 48; // half-size
    var top    = { x: cx,     y: cy - s };
    var right  = { x: cx + s, y: cy };
    var bottom = { x: cx,     y: cy + s };
    var left   = { x: cx - s, y: cy };
    var inner  = { x: cx, y: cy - s * 0.18 };
    var midL   = { x: cx - s * 0.55, y: cy + s * 0.28 };
    var midR   = { x: cx + s * 0.55, y: cy + s * 0.28 };

    parts.push('<defs>');
    parts.push('<linearGradient id="entra-face-light" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" stop-color="#7dd3fc"/><stop offset="100%" stop-color="#22d3ee"/></linearGradient>');
    parts.push('<linearGradient id="entra-face-mid" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" stop-color="#22d3ee"/><stop offset="100%" stop-color="#0ea5e9"/></linearGradient>');
    parts.push('<linearGradient id="entra-face-dark" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" stop-color="#0ea5e9"/><stop offset="100%" stop-color="#0c4a6e"/></linearGradient>');
    parts.push('</defs>');

    parts.push('<g class="node-box" data-entra="cloud">');
    parts.push('<polygon points="' + top.x + ',' + top.y + ' ' + inner.x + ',' + inner.y + ' ' + midL.x + ',' + midL.y + ' ' + left.x + ',' + left.y + '" fill="url(#entra-face-light)" stroke="#0c4a6e" stroke-width="1.5" stroke-linejoin="round"/>');
    parts.push('<polygon points="' + top.x + ',' + top.y + ' ' + right.x + ',' + right.y + ' ' + midR.x + ',' + midR.y + ' ' + inner.x + ',' + inner.y + '" fill="url(#entra-face-mid)" stroke="#0c4a6e" stroke-width="1.5" stroke-linejoin="round"/>');
    parts.push('<polygon points="' + inner.x + ',' + inner.y + ' ' + midR.x + ',' + midR.y + ' ' + bottom.x + ',' + bottom.y + ' ' + midL.x + ',' + midL.y + '" fill="url(#entra-face-dark)" stroke="#0c4a6e" stroke-width="1.5" stroke-linejoin="round"/>');
    parts.push('<polygon points="' + top.x + ',' + top.y + ' ' + right.x + ',' + right.y + ' ' + bottom.x + ',' + bottom.y + ' ' + left.x + ',' + left.y + '" fill="none" stroke="#0c4a6e" stroke-width="2.5" stroke-linejoin="round"/>');

    parts.push('<text x="' + cx + '" y="' + (cy + s + 22) + '" text-anchor="middle" font-family="Segoe UI" font-size="14" font-weight="700" fill="#22d3ee">Microsoft Entra ID</text>');
    if (tenantName){
      parts.push('<text x="' + cx + '" y="' + (cy + s + 38) + '" text-anchor="middle" font-family="Segoe UI" font-size="10" fill="#94a3b8">' + escapeXml(tenantName) + '</text>');
    } else {
      parts.push('<text x="' + cx + '" y="' + (cy + s + 38) + '" text-anchor="middle" font-family="Segoe UI" font-size="10" fill="#94a3b8">Cloud Identity</text>');
    }
    parts.push('</g>');
    return parts.join('');
  }

  function buildSyncServer(cx, cy, name){
    var parts = [];
    var w = 160, h = 46;
    var x = cx - w/2, y = cy - h/2;
    parts.push('<g class="node-box" data-entra="sync">');
    parts.push('<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" rx="6" fill="rgba(34,211,238,0.1)" stroke="#22d3ee" stroke-width="2" stroke-dasharray="4,3"/>');
    parts.push('<text x="' + cx + '" y="' + (cy - 4) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#22d3ee">Entra Connect Sync</text>');
    parts.push('<text x="' + cx + '" y="' + (cy + 11) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(name || 'Sync server') + '</text>');
    parts.push('</g>');
    return parts.join('');
  }

  // ==================== FOREST & DOMAINS MAP ====================
  function drawForestMap(){
    var container = document.getElementById('forest-canvas');
    if (!container) return;
    var domains = (window.ADAtlasData.Domains || []).map(function(d){
      return Object.assign({}, d, { children: [] });
    });
    if (domains.length === 0){
      container.innerHTML = '<svg viewBox="0 0 600 100" xmlns="http://www.w3.org/2000/svg"><text x="20" y="50" fill="#94a3b8" font-family="Segoe UI" font-size="14">No domains found.</text></svg>';
      return;
    }
    var byName = {};
    domains.forEach(function(d){ byName[d.DNSRoot] = d; });
    var root = null;
    domains.forEach(function(d){
      if (d.IsForestRoot) root = d;
      else if (byName[d.ParentDomain]) byName[d.ParentDomain].children.push(d);
    });
    if (!root) root = domains[0];
    function sortTree(n){
      n.children.sort(function(a,b){ return a.DNSRoot.localeCompare(b.DNSRoot); });
      n.children.forEach(sortTree);
    }
    sortTree(root);
    var nodeW = 220, nodeH = 130, vGap = 60, hGap = 40;
    var cursorX = 30;
    function layout(n, depth){
      n.depth = depth;
      if (n.children.length === 0){
        n.cx = cursorX + nodeW/2;
        cursorX += nodeW + hGap;
      } else {
        n.children.forEach(function(c){ layout(c, depth+1); });
        var first = n.children[0];
        var last = n.children[n.children.length-1];
        n.cx = (first.cx + last.cx) / 2;
      }
      n.cy = 80 + depth * (nodeH + vGap) + nodeH/2;
    }
    layout(root, 0);
    var allNodes = [];
    function collect(n){ allNodes.push(n); n.children.forEach(collect); }
    collect(root);
    domains.forEach(function(d){
      if (allNodes.indexOf(d) === -1){
        d.depth = 0;
        d.cx = cursorX + nodeW/2;
        d.cy = 80 + nodeH/2;
        cursorX += nodeW + hGap;
        allNodes.push(d);
      }
    });
    var entra = window.ADAtlasData.EntraConnect;
    var hasEntra = entra && entra.Detected;
    var entraOffset = hasEntra ? 230 : 0;
    if (hasEntra){
      allNodes.forEach(function(n){ n.cy += entraOffset; });
    }
    var maxX = 0, maxY = 0;
    allNodes.forEach(function(n){
      if (n.cx + nodeW/2 > maxX) maxX = n.cx + nodeW/2;
      if (n.cy + nodeH/2 > maxY) maxY = n.cy + nodeH/2;
    });
    var viewW = Math.max(maxX + 30, 700);
    var viewH = maxY + 40;
    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMin meet" style="min-width:' + Math.min(viewW, 1300) + 'px">');

    if (hasEntra){
      var entraCx = root.cx;
      var entraCy = 70;
      var syncCy = 165;
      svgParts.push(buildEntraLogo(entraCx, entraCy, entra.Tenant || ''));
      svgParts.push(buildSyncServer(entraCx, syncCy, entra.ServerName || ''));
      svgParts.push('<line x1="' + entraCx + '" y1="' + (entraCy + 55) + '" x2="' + entraCx + '" y2="' + (syncCy - 23) + '" stroke="#22d3ee" stroke-width="2" stroke-dasharray="4,3"/>');
      svgParts.push('<line x1="' + entraCx + '" y1="' + (syncCy + 23) + '" x2="' + root.cx + '" y2="' + (root.cy - nodeH/2) + '" stroke="#22d3ee" stroke-width="2" stroke-dasharray="4,3"/>');
    }

    function drawEdges(n){
      n.children.forEach(function(c){
        var x1 = n.cx, y1 = n.cy + nodeH/2;
        var x2 = c.cx, y2 = c.cy - nodeH/2;
        var midY = (y1 + y2) / 2;
        svgParts.push('<path d="M ' + x1 + ' ' + y1 + ' L ' + x1 + ' ' + midY + ' L ' + x2 + ' ' + midY + ' L ' + x2 + ' ' + y2 + '" stroke="#475569" stroke-width="2" fill="none"/>');
        drawEdges(c);
      });
    }
    drawEdges(root);

    allNodes.forEach(function(n){
      var color, bg;
      if (n.IsForestRoot){ color = '#60a5fa'; bg = 'rgba(96,165,250,0.18)'; }
      else if (n.children.length > 0){ color = '#22d3ee'; bg = 'rgba(34,211,238,0.12)'; }
      else { color = '#a78bfa'; bg = 'rgba(167,139,250,0.12)'; }
      var dcText = (n.DCCount || 0) + ' DC' + ((n.DCCount === 1) ? '' : 's');
      var t = buildDomainTriangle({
        cx: n.cx, cy: n.cy, w: nodeW, h: nodeH,
        label: n.DNSRoot, sublabel: n.DomainModeDisplay || '', dcText: dcText,
        color: color, bg: bg,
        dataAttr: 'data-domain="' + escapeXml(n.DNSRoot) + '"'
      });
      svgParts.push(t.svg);
    });

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');

    container.querySelectorAll('.node-box[data-domain]').forEach(function(g){
      g.addEventListener('click', function(){ selectDomain(g.getAttribute('data-domain')); });
    });
    container.querySelectorAll('.node-box[data-entra]').forEach(function(g){
      g.addEventListener('click', function(){ showEntraPanel(g.getAttribute('data-entra')); });
    });
  }

  function showEntraPanel(which){
    var panel = document.getElementById('forest-panel');
    if (!panel) return;
    var e = window.ADAtlasData.EntraConnect || {};
    var html = '';
    if (which === 'cloud'){
      html += '<h3>Entra ID (Cloud Identity)</h3>';
      html += field('Status', e.Detected ? 'Hybrid identity detected via Entra Connect' : 'Not detected');
      html += field('Tenant', e.Tenant || '(unknown)');
      html += '<div class="field"><label>Note</label><div class="val dim">ADAtlas detects Entra Connect via MSOL_/Sync_ service accounts in AD. It does not query Entra ID itself.</div></div>';
    } else {
      html += '<h3>Entra Connect Sync Server</h3>';
      html += field('Detected', e.Detected ? 'Yes' : 'No');
      html += field('Server Name', e.ServerName || '(unknown)');
      html += field('Tenant', e.Tenant || '(unknown)');
      var accts = e.Accounts || [];
      html += '<div class="field"><label>Service Accounts (' + accts.length + ')</label><div class="val dc-list">';
      if (accts.length === 0){ html += '<span class="dim">(none)</span>'; }
      else { accts.forEach(function(a){
        html += '<div class="dc-item"><strong>' + escapeHtml(a.ServiceAccount) + '</strong><div class="dc-meta">' + escapeHtml(a.DetectionMethod) + ' &middot; Created ' + escapeHtml(a.Created || '') + ' &middot; ' + (a.Enabled ? 'Enabled' : 'Disabled') + '</div></div>';
      });}
      html += '</div></div>';
    }
    panel.innerHTML = html;
  }

  function selectDomain(dnsRoot){
    var d = (window.ADAtlasData.Domains || []).find(function(x){ return x.DNSRoot === dnsRoot; });
    var panel = document.getElementById('forest-panel');
    if (!panel) return;
    if (!d){ panel.innerHTML = '<div class="empty">Domain not found.</div>'; return; }
    var dcs = (window.ADAtlasData.DCs || []).filter(function(x){ return x.Domain === dnsRoot; });
    var html = '<h3>' + escapeHtml(d.DNSRoot) + '</h3>';
    html += field('NetBIOS Name', d.NetBIOSName);
    html += field('Functional Level', d.DomainModeDisplay);
    html += field('Forest Root', d.IsForestRoot ? 'Yes' : 'No');
    html += field('Parent Domain', d.ParentDomain || '(none)');
    html += field('Child Domains', (d.ChildDomains && d.ChildDomains.length > 0) ? d.ChildDomains.join(', ') : '(none)');
    html += field('Distinguished Name', d.DistinguishedName);
    html += '<div class="field"><label>PDC Emulator</label><div class="val">' + escapeHtml(d.PDCEmulator || '') + '<span class="pill pill-fsmo">FSMO</span></div></div>';
    html += '<div class="field"><label>RID Master</label><div class="val">' + escapeHtml(d.RIDMaster || '') + '<span class="pill pill-fsmo">FSMO</span></div></div>';
    html += '<div class="field"><label>Infrastructure Master</label><div class="val">' + escapeHtml(d.InfrastructureMaster || '') + '<span class="pill pill-fsmo">FSMO</span></div></div>';
    html += '<div class="field"><label>Domain Controllers (' + dcs.length + ')</label><div class="val dc-list">';
    if (dcs.length === 0){ html += '<span class="dim">(none)</span>'; }
    else { dcs.forEach(function(dc){
      var pills = '';
      if (dc.Type === 'RODC') pills += '<span class="pill pill-rodc">RODC</span>';
      if (dc.IsGlobalCatalog) pills += '<span class="pill pill-gc">GC</span>';
      if (dc.FSMORoles && dc.FSMORoles.length > 0) pills += '<span class="pill pill-fsmo">FSMO</span>';
      html += '<div class="dc-item"><strong>' + escapeHtml(dc.Name) + '</strong>' + pills + '<div class="dc-meta">' + escapeHtml(dc.Site || '(no site)') + ' &middot; ' + escapeHtml(dc.IPv4Address || '') + ' &middot; ' + escapeHtml(dc.OperatingSystem || '') + '</div></div>';
    });}
    html += '</div></div>';
    panel.innerHTML = html;
    document.querySelectorAll('#forest-canvas .node-box[data-domain]').forEach(function(g){
      g.classList.toggle('selected', g.getAttribute('data-domain') === dnsRoot);
    });
  }

  // ==================== TRUST MAP - REDESIGNED HIERARCHICAL ====================
  function drawTrustMap(){
    var container = document.getElementById('trust-canvas');
    if (!container) return;
    var trusts = window.ADAtlasData.Trusts || [];
    var forestDomains = (window.ADAtlasData.Domains || []).map(function(d){
      return Object.assign({}, d, { children: [] });
    });

    // Build forest tree (same as forest map)
    var byName = {};
    forestDomains.forEach(function(d){ byName[d.DNSRoot] = d; });
    var root = null;
    forestDomains.forEach(function(d){
      if (d.IsForestRoot) root = d;
      else if (byName[d.ParentDomain]) byName[d.ParentDomain].children.push(d);
    });
    if (!root && forestDomains.length > 0) root = forestDomains[0];

    function sortTree(n){
      if (!n) return;
      n.children.sort(function(a,b){ return a.DNSRoot.localeCompare(b.DNSRoot); });
      n.children.forEach(sortTree);
    }
    sortTree(root);

    // Find external domains (in trusts but not in forest)
    var forestSet = {};
    forestDomains.forEach(function(d){ forestSet[d.DNSRoot] = true; });
    var externalSet = {};
    trusts.forEach(function(t){
      if (!forestSet[t.SourceDomain]) externalSet[t.SourceDomain] = true;
      if (!forestSet[t.TargetDomain]) externalSet[t.TargetDomain] = true;
    });
    var externals = Object.keys(externalSet).sort();

    if (forestDomains.length === 0 && externals.length === 0){
      container.innerHTML = '<svg viewBox="0 0 600 100" xmlns="http://www.w3.org/2000/svg"><text x="20" y="50" fill="#94a3b8" font-family="Segoe UI" font-size="14">No trusts or domains found.</text></svg>';
      return;
    }

    // Layout forest tree on the LEFT
    var nodeW = 200, nodeH = 110, vGap = 50, hGap = 30;
    var leftPad = 40;
    var cursorX = leftPad;
    function layout(n, depth){
      if (!n) return;
      n.depth = depth;
      if (n.children.length === 0){
        n.cx = cursorX + nodeW/2;
        cursorX += nodeW + hGap;
      } else {
        n.children.forEach(function(c){ layout(c, depth+1); });
        var first = n.children[0];
        var last = n.children[n.children.length-1];
        n.cx = (first.cx + last.cx) / 2;
      }
      n.cy = 80 + depth * (nodeH + vGap) + nodeH/2;
    }
    if (root) layout(root, 0);

    var allForestNodes = [];
    function collect(n){ if(n){ allForestNodes.push(n); n.children.forEach(collect); } }
    collect(root);
    forestDomains.forEach(function(d){
      if (allForestNodes.indexOf(d) === -1){
        d.depth = 0;
        d.cx = cursorX + nodeW/2;
        d.cy = 80 + nodeH/2;
        cursorX += nodeW + hGap;
        allForestNodes.push(d);
      }
    });

    var forestMaxX = cursorX;
    var forestMaxY = 0;
    allForestNodes.forEach(function(n){
      if (n.cy + nodeH/2 > forestMaxY) forestMaxY = n.cy + nodeH/2;
    });

    // Layout external domains on the RIGHT in a vertical stack
    var externalGroupX = forestMaxX + 220;
    var externalGroupY = 80;
    var externalNodes = externals.map(function(name, i){
      return {
        DNSRoot: name,
        external: true,
        cx: externalGroupX + nodeW/2,
        cy: externalGroupY + i * (nodeH + 20) + nodeH/2
      };
    });

    var totalMaxY = Math.max(forestMaxY, externalGroupY + externals.length * (nodeH + 20));
    var viewW = externalGroupX + nodeW + 60;
    var viewH = totalMaxY + 60;

    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMid meet" style="min-width:' + Math.min(viewW, 1300) + 'px">');

    // Defs for markers
    svgParts.push('<defs>');
    svgParts.push('<marker id="t-arrow-intra" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#34d399"/></marker>');
    svgParts.push('<marker id="t-arrow-ext" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#fbbf24"/></marker>');
    svgParts.push('</defs>');

    // Group label for external domains (if any)
    if (externals.length > 0){
      svgParts.push('<rect x="' + (externalGroupX - 20) + '" y="50" width="' + (nodeW + 40) + '" height="' + (externals.length * (nodeH + 20) + 30) + '" rx="8" fill="rgba(251,191,36,0.04)" stroke="#fbbf24" stroke-width="1.5" stroke-dasharray="5,4"/>');
      svgParts.push('<text x="' + (externalGroupX + nodeW/2) + '" y="40" text-anchor="middle" font-family="Segoe UI" font-size="12" font-weight="700" fill="#fbbf24">EXTERNAL DOMAINS</text>');
    }

    // Forest tree connector lines (parent to child)
    function drawForestEdges(n){
      if (!n) return;
      n.children.forEach(function(c){
        var x1 = n.cx, y1 = n.cy + nodeH/2;
        var x2 = c.cx, y2 = c.cy - nodeH/2;
        var midY = (y1 + y2) / 2;
        svgParts.push('<path d="M ' + x1 + ' ' + y1 + ' L ' + x1 + ' ' + midY + ' L ' + x2 + ' ' + midY + ' L ' + x2 + ' ' + y2 + '" stroke="#475569" stroke-width="2" fill="none"/>');
        drawForestEdges(c);
      });
    }
    drawForestEdges(root);

    // Build node lookup for trust line drawing
    var nodeLookup = {};
    allForestNodes.forEach(function(n){ nodeLookup[n.DNSRoot] = n; });
    externalNodes.forEach(function(n){ nodeLookup[n.DNSRoot] = n; });

    // Draw trust edges as smooth curves
    trusts.forEach(function(t, idx){
      var src = nodeLookup[t.SourceDomain];
      var tgt = nodeLookup[t.TargetDomain];
      if (!src || !tgt) return;
      // Skip parent-child intra-forest trusts (they're already shown by the forest tree connector)
      // We only skip them if it's exactly the parent-child relationship
      var srcD = forestDomains.find(function(x){ return x.DNSRoot === t.SourceDomain; });
      var tgtD = forestDomains.find(function(x){ return x.DNSRoot === t.TargetDomain; });
      if (t.IntraForest && srcD && tgtD){
        if (srcD.ParentDomain === tgtD.DNSRoot || tgtD.ParentDomain === srcD.DNSRoot){
          // It's a parent-child trust - skip the line, the tree edge represents it
          return;
        }
      }

      // Choose start/end edge points (right side of source if going to external, otherwise nearest)
      var sx, sy, ex, ey;
      if (tgt.external){
        sx = src.cx + nodeW/2 - 5;
        sy = src.cy;
        ex = tgt.cx - nodeW/2 + 5;
        ey = tgt.cy;
      } else if (src.external){
        sx = src.cx - nodeW/2 + 5;
        sy = src.cy;
        ex = tgt.cx + nodeW/2 - 5;
        ey = tgt.cy;
      } else {
        // Both forest, intra-forest non-parent-child (shortcut trust)
        sx = src.cx;
        sy = src.cy + nodeH/2 - 5;
        ex = tgt.cx;
        ey = tgt.cy + nodeH/2 - 5;
      }

      // Cubic bezier curve
      var dx = ex - sx, dy = ey - sy;
      var ctrlOffset = Math.max(60, Math.abs(dx) * 0.4);
      var c1x, c1y, c2x, c2y;
      if (Math.abs(dx) > Math.abs(dy)){
        c1x = sx + ctrlOffset; c1y = sy;
        c2x = ex - ctrlOffset; c2y = ey;
      } else {
        c1x = sx; c1y = sy + ctrlOffset;
        c2x = ex; c2y = ey - ctrlOffset;
      }
      var path = 'M ' + sx + ' ' + sy + ' C ' + c1x + ' ' + c1y + ', ' + c2x + ' ' + c2y + ', ' + ex + ' ' + ey;
      var color = t.IntraForest ? '#34d399' : '#fbbf24';
      var dash = t.IntraForest ? '' : ' stroke-dasharray="7,5"';
      var marker = t.IntraForest ? 't-arrow-intra' : 't-arrow-ext';
      var markers = '';
      if (t.Direction === 'BiDirectional'){
        markers = ' marker-start="url(#' + marker + ')" marker-end="url(#' + marker + ')"';
      } else if (t.Direction === 'Inbound'){
        markers = ' marker-start="url(#' + marker + ')"';
      } else {
        markers = ' marker-end="url(#' + marker + ')"';
      }
      svgParts.push('<path class="trust-edge" data-trust="' + idx + '" d="' + path + '" stroke="' + color + '" stroke-width="2.5" fill="none"' + dash + markers + '/>');
    });

    // Draw forest domain triangles
    allForestNodes.forEach(function(n){
      var color, bg;
      if (n.IsForestRoot){ color = '#60a5fa'; bg = 'rgba(96,165,250,0.18)'; }
      else if (n.children && n.children.length > 0){ color = '#22d3ee'; bg = 'rgba(34,211,238,0.12)'; }
      else { color = '#a78bfa'; bg = 'rgba(167,139,250,0.12)'; }
      var t = buildDomainTriangle({
        cx: n.cx, cy: n.cy, w: nodeW, h: nodeH,
        label: n.DNSRoot, sublabel: n.DomainModeDisplay || '',
        dcText: (n.DCCount || 0) + ' DC' + (n.DCCount === 1 ? '' : 's'),
        color: color, bg: bg,
        dataAttr: 'data-trust-node="' + escapeXml(n.DNSRoot) + '"'
      });
      svgParts.push(t.svg);
    });

    // Draw external domain triangles
    externalNodes.forEach(function(n){
      var t = buildDomainTriangle({
        cx: n.cx, cy: n.cy, w: nodeW, h: nodeH,
        label: n.DNSRoot, sublabel: 'External', dcText: '',
        color: '#fbbf24', bg: 'rgba(251,191,36,0.15)',
        dataAttr: 'data-trust-node="' + escapeXml(n.DNSRoot) + '"'
      });
      svgParts.push(t.svg);
    });

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');

    container.querySelectorAll('.trust-edge').forEach(function(line){
      line.addEventListener('click', function(){
        selectTrust(parseInt(line.getAttribute('data-trust'), 10));
      });
    });
    container.querySelectorAll('.node-box[data-trust-node]').forEach(function(g){
      g.addEventListener('click', function(){
        selectTrustNode(g.getAttribute('data-trust-node'));
      });
    });
  }

  function selectTrust(idx){
    var trusts = window.ADAtlasData.Trusts || [];
    var t = trusts[idx];
    var panel = document.getElementById('trust-panel');
    if (!panel || !t) return;
    var arrow = t.Direction === 'BiDirectional' ? '\u2194' : (t.Direction === 'Inbound' ? '\u2190' : '\u2192');
    var html = '<h3>Trust: ' + escapeHtml(t.SourceDomain) + ' ' + arrow + ' ' + escapeHtml(t.TargetDomain) + '</h3>';
    html += field('Source Domain', t.SourceDomain);
    html += field('Target Domain', t.TargetDomain);
    html += field('Direction', t.Direction);
    html += field('Trust Type', t.TrustType);
    html += field('Scope', t.IntraForest ? 'Intra-Forest' : 'External / Cross-Forest');
    html += field('Transitive', t.Transitive ? 'Yes' : 'No');
    html += field('Forest Transitive', t.ForestTransitive ? 'Yes' : 'No');
    html += field('SID Filtering Quarantined', t.SIDFilter ? 'Yes' : 'No');
    html += field('Selective Authentication', t.SelectiveAuth ? 'Yes' : 'No');
    html += field('Uplevel Only', t.UplevelOnly ? 'Yes' : 'No');
    html += field('When Created', t.WhenCreated);
    panel.innerHTML = html;
    document.querySelectorAll('#trust-canvas .trust-edge').forEach(function(l){
      l.classList.toggle('selected', parseInt(l.getAttribute('data-trust'), 10) === idx);
    });
    document.querySelectorAll('#trust-canvas .node-box').forEach(function(g){ g.classList.remove('selected'); });
  }

  function selectTrustNode(name){
    var panel = document.getElementById('trust-panel');
    if (!panel) return;
    var forestDomains = window.ADAtlasData.Domains || [];
    var d = forestDomains.find(function(x){ return x.DNSRoot === name; });
    var trusts = (window.ADAtlasData.Trusts || []).filter(function(t){
      return t.SourceDomain === name || t.TargetDomain === name;
    });
    var html = '<h3>' + escapeHtml(name) + '</h3>';
    if (d){
      html += field('NetBIOS', d.NetBIOSName);
      html += field('Functional Level', d.DomainModeDisplay);
      html += field('Scope', 'Forest Domain' + (d.IsForestRoot ? ' (Root)' : ''));
      html += field('DC Count', d.DCCount);
    } else {
      html += field('Scope', 'External Domain');
      html += '<div class="field"><label>Note</label><div class="val dim">This domain is not part of the forest. It appears here because an external trust exists.</div></div>';
    }
    html += '<div class="field"><label>Trusts Involving This Domain (' + trusts.length + ')</label><div class="val dc-list">';
    if (trusts.length === 0){ html += '<span class="dim">(none)</span>'; }
    else { trusts.forEach(function(t){
      var arrow = t.Direction === 'BiDirectional' ? '\u2194' : (t.Direction === 'Inbound' ? '\u2190' : '\u2192');
      var other = (t.SourceDomain === name) ? t.TargetDomain : t.SourceDomain;
      var badge = t.IntraForest ? '<span class="pill pill-gc">Intra</span>' : '<span class="pill pill-rodc">Ext</span>';
      html += '<div class="dc-item"><strong>' + escapeHtml(name) + '</strong> ' + arrow + ' <strong>' + escapeHtml(other) + '</strong>' + badge + '<div class="dc-meta">' + escapeHtml(t.TrustType) + ' &middot; ' + escapeHtml(t.Direction) + '</div></div>';
    });}
    html += '</div></div>';
    panel.innerHTML = html;
    document.querySelectorAll('#trust-canvas .node-box').forEach(function(g){
      g.classList.toggle('selected', g.getAttribute('data-trust-node') === name);
    });
    document.querySelectorAll('#trust-canvas .trust-edge').forEach(function(l){ l.classList.remove('selected'); });
  }

  // ==================== SITE TOPOLOGY ====================
  function drawSiteTopology(){
    var container = document.getElementById('site-canvas');
    if (!container) return;
    var sites = window.ADAtlasData.Sites || [];
    var dcs = window.ADAtlasData.DCs || [];
    var siteLinks = window.ADAtlasData.SiteLinks || [];
    if (sites.length === 0){
      container.innerHTML = '<svg viewBox="0 0 600 100" xmlns="http://www.w3.org/2000/svg"><text x="20" y="50" fill="#94a3b8" font-family="Segoe UI" font-size="14">No sites found.</text></svg>';
      return;
    }
    // Compute DC count per site
    var siteNodes = sites.map(function(s){
      var count = dcs.filter(function(d){ return d.Site === s.Name; }).length;
      return { name: s.Name, dcCount: count, description: s.Description, location: s.Location };
    });
    // Layout in circular arrangement
    var n = siteNodes.length;
    var cx = 500, cy = 380, R = Math.min(280, 60 + n * 28);
    siteNodes.forEach(function(node, i){
      if (n === 1){ node.x = cx; node.y = cy; }
      else {
        var ang = (i / n) * 2 * Math.PI - Math.PI/2;
        node.x = cx + R * Math.cos(ang);
        node.y = cy + R * Math.sin(ang);
      }
      // Radius scaled by dc count
      node.r = Math.max(28, Math.min(60, 28 + node.dcCount * 4));
    });

    var viewW = 1000, viewH = 760;
    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMid meet" style="min-width:700px">');

    // Draw site links as lines between sites that share a link
    var nodeByName = {};
    siteNodes.forEach(function(s){ nodeByName[s.name] = s; });
    siteLinks.forEach(function(sl, idx){
      var sites = sl.Sites || [];
      for (var i = 0; i < sites.length; i++){
        for (var j = i+1; j < sites.length; j++){
          var a = nodeByName[sites[i]], b = nodeByName[sites[j]];
          if (a && b){
            svgParts.push('<line class="repl-edge" data-sitelink="' + idx + '" x1="' + a.x + '" y1="' + a.y + '" x2="' + b.x + '" y2="' + b.y + '" stroke="#60a5fa" stroke-width="2" opacity="0.6"/>');
          }
        }
      }
    });

    // Draw site circles
    siteNodes.forEach(function(node){
      var color = node.dcCount > 0 ? '#22d3ee' : '#94a3b8';
      var bg = node.dcCount > 0 ? 'rgba(34,211,238,0.18)' : 'rgba(148,163,184,0.1)';
      svgParts.push('<g class="node-box" data-site="' + escapeXml(node.name) + '">');
      svgParts.push('<circle cx="' + node.x + '" cy="' + node.y + '" r="' + node.r + '" fill="' + bg + '" stroke="' + color + '" stroke-width="2.5"/>');
      svgParts.push('<text x="' + node.x + '" y="' + (node.y - 4) + '" text-anchor="middle" font-family="Segoe UI" font-size="12" font-weight="700" fill="#e2e8f0">' + escapeXml(node.name) + '</text>');
      svgParts.push('<text x="' + node.x + '" y="' + (node.y + 12) + '" text-anchor="middle" font-family="Segoe UI" font-size="10" fill="#94a3b8">' + node.dcCount + ' DC' + (node.dcCount === 1 ? '' : 's') + '</text>');
      svgParts.push('</g>');
    });

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');

    container.querySelectorAll('.node-box[data-site]').forEach(function(g){
      g.addEventListener('click', function(){ selectSite(g.getAttribute('data-site')); });
    });
    container.querySelectorAll('.repl-edge[data-sitelink]').forEach(function(l){
      l.addEventListener('click', function(){ selectSiteLink(parseInt(l.getAttribute('data-sitelink'),10)); });
    });
  }

  function selectSite(name){
    var panel = document.getElementById('site-panel');
    if (!panel) return;
    var sites = window.ADAtlasData.Sites || [];
    var s = sites.find(function(x){ return x.Name === name; });
    if (!s){ panel.innerHTML = '<div class="empty">Site not found.</div>'; return; }
    var dcs = (window.ADAtlasData.DCs || []).filter(function(d){ return d.Site === name; });
    var subnets = (window.ADAtlasData.Subnets || []).filter(function(sn){ return sn.Site === name; });
    var siteLinks = (window.ADAtlasData.SiteLinks || []).filter(function(sl){ return (sl.Sites || []).indexOf(name) >= 0; });
    var html = '<h3>' + escapeHtml(name) + '</h3>';
    html += field('Description', s.Description);
    html += field('Location', s.Location);
    html += field('When Created', s.WhenCreated);
    html += '<div class="field"><label>Domain Controllers (' + dcs.length + ')</label><div class="val dc-list">';
    if (dcs.length === 0){ html += '<span class="dim">(none)</span>'; }
    else { dcs.forEach(function(dc){
      var pills = '';
      if (dc.Type === 'RODC') pills += '<span class="pill pill-rodc">RODC</span>';
      if (dc.IsGlobalCatalog) pills += '<span class="pill pill-gc">GC</span>';
      html += '<div class="dc-item"><strong>' + escapeHtml(dc.Name) + '</strong>' + pills + '<div class="dc-meta">' + escapeHtml(dc.Domain) + ' &middot; ' + escapeHtml(dc.IPv4Address || '') + '</div></div>';
    });}
    html += '</div></div>';
    html += '<div class="field"><label>Subnets (' + subnets.length + ')</label><div class="val dc-list">';
    if (subnets.length === 0){ html += '<span class="dim">(none)</span>'; }
    else { subnets.forEach(function(sn){
      html += '<div class="dc-item"><code>' + escapeHtml(sn.Name) + '</code><div class="dc-meta">' + escapeHtml(sn.Location || '') + '</div></div>';
    });}
    html += '</div></div>';
    html += '<div class="field"><label>Site Links (' + siteLinks.length + ')</label><div class="val dc-list">';
    if (siteLinks.length === 0){ html += '<span class="dim">(none)</span>'; }
    else { siteLinks.forEach(function(sl){
      html += '<div class="dc-item"><strong>' + escapeHtml(sl.Name) + '</strong><div class="dc-meta">Cost ' + sl.Cost + ' &middot; ' + sl.Frequency + ' min</div></div>';
    });}
    html += '</div></div>';
    panel.innerHTML = html;
    document.querySelectorAll('#site-canvas .node-box[data-site]').forEach(function(g){
      g.classList.toggle('selected', g.getAttribute('data-site') === name);
    });
  }

  function selectSiteLink(idx){
    var panel = document.getElementById('site-panel');
    var sl = (window.ADAtlasData.SiteLinks || [])[idx];
    if (!panel || !sl) return;
    var html = '<h3>Site Link: ' + escapeHtml(sl.Name) + '</h3>';
    html += field('Cost', sl.Cost);
    html += field('Replication Frequency (min)', sl.Frequency);
    html += field('Transport', sl.Transport);
    html += field('Sites Included', (sl.Sites || []).join(', '));
    html += field('Description', sl.Description);
    panel.innerHTML = html;
  }

  // ==================== REPLICATION TOPOLOGY ====================
  function populateReplSiteSelector(){
    var sel = document.getElementById('repl-site-select');
    if (!sel) return;
    var sites = (window.ADAtlasData.Sites || []).slice().sort(function(a,b){ return a.Name.localeCompare(b.Name); });
    var html = '<option value="__ALL__">All Sites (entire forest)</option>';
    sites.forEach(function(s){
      html += '<option value="' + escapeHtml(s.Name) + '">' + escapeHtml(s.Name) + '</option>';
    });
    sel.innerHTML = html;
    sel.addEventListener('change', function(){ drawReplicationTopology(sel.value); });
  }

  function drawReplicationTopology(scope){
    var container = document.getElementById('repl-canvas');
    if (!container) return;
    if (!scope) scope = '__ALL__';
    var allConns = window.ADAtlasData.Connections || [];
    var allDcs = window.ADAtlasData.DCs || [];
    var conns = (scope === '__ALL__') ? allConns : allConns.filter(function(c){ return c.Site === scope; });
    // Get DC names involved
    var dcSet = {};
    conns.forEach(function(c){
      if (c.From) dcSet[c.From] = true;
      if (c.To) dcSet[c.To] = true;
    });
    if (scope !== '__ALL__'){
      // Add all DCs in the site even if no connections
      allDcs.forEach(function(d){ if (d.Site === scope) dcSet[d.Name] = true; });
    }
    var dcNames = Object.keys(dcSet).sort();
    var countEl = document.getElementById('repl-count');
    if (countEl) countEl.innerHTML = '<strong>' + conns.length + '</strong> connection(s) &middot; <strong>' + dcNames.length + '</strong> DC(s)';

    if (dcNames.length === 0){
      container.innerHTML = '<svg viewBox="0 0 600 100" xmlns="http://www.w3.org/2000/svg"><text x="20" y="50" fill="#94a3b8" font-family="Segoe UI" font-size="14">No connections in this scope.</text></svg>';
      return;
    }

    // Layout DCs in circle
    var dcLookup = {};
    allDcs.forEach(function(d){ dcLookup[d.Name] = d; });
    var n = dcNames.length;
    var cx = 500, cy = 360, R = Math.min(280, 80 + n * 18);
    var nodes = dcNames.map(function(name, i){
      var d = dcLookup[name] || { Name: name, Type: 'RWDC' };
      var ang = (i / n) * 2 * Math.PI - Math.PI/2;
      return {
        name: name,
        x: n === 1 ? cx : cx + R * Math.cos(ang),
        y: n === 1 ? cy : cy + R * Math.sin(ang),
        type: d.Type || 'RWDC',
        site: d.Site || '',
        domain: d.Domain || ''
      };
    });
    var nodeByName = {};
    nodes.forEach(function(node){ nodeByName[node.name] = node; });

    var viewW = 1000, viewH = 720;
    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMid meet" style="min-width:700px">');
    svgParts.push('<defs>');
    svgParts.push('<marker id="repl-arrow-auto" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#34d399"/></marker>');
    svgParts.push('<marker id="repl-arrow-manual" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#22d3ee"/></marker>');
    svgParts.push('</defs>');

    // Draw connections (curved)
    conns.forEach(function(c, idx){
      var src = nodeByName[c.From];
      var tgt = nodeByName[c.To];
      if (!src || !tgt) return;
      var dx = tgt.x - src.x, dy = tgt.y - src.y;
      var midX = (src.x + tgt.x)/2 + dy * 0.12;
      var midY = (src.y + tgt.y)/2 - dx * 0.12;
      var color = c.AutoGenerated ? '#34d399' : '#22d3ee';
      var dash = c.AutoGenerated ? '' : ' stroke-dasharray="6,4"';
      var marker = c.AutoGenerated ? 'repl-arrow-auto' : 'repl-arrow-manual';
      svgParts.push('<path class="repl-edge" data-conn="' + idx + '" d="M ' + src.x + ' ' + src.y + ' Q ' + midX + ' ' + midY + ' ' + tgt.x + ' ' + tgt.y + '" stroke="' + color + '" stroke-width="2" fill="none" opacity="0.75"' + dash + ' marker-end="url(#' + marker + ')"/>');
    });

    // Draw DC boxes
    nodes.forEach(function(node){
      var color = node.type === 'RODC' ? '#fbbf24' : '#60a5fa';
      var bg = node.type === 'RODC' ? 'rgba(251,191,36,0.15)' : 'rgba(96,165,250,0.15)';
      var w = 130, h = 38;
      svgParts.push('<g class="node-box" data-repl-dc="' + escapeXml(node.name) + '">');
      svgParts.push('<rect x="' + (node.x - w/2) + '" y="' + (node.y - h/2) + '" width="' + w + '" height="' + h + '" rx="6" fill="' + bg + '" stroke="' + color + '" stroke-width="2.5"/>');
      svgParts.push('<text x="' + node.x + '" y="' + (node.y - 3) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#e2e8f0">' + escapeXml(node.name) + '</text>');
      svgParts.push('<text x="' + node.x + '" y="' + (node.y + 12) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(node.site || '') + '</text>');
      svgParts.push('</g>');
    });

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');

    container.querySelectorAll('.node-box[data-repl-dc]').forEach(function(g){
      g.addEventListener('click', function(){ selectReplDC(g.getAttribute('data-repl-dc')); });
    });
    container.querySelectorAll('.repl-edge[data-conn]').forEach(function(l){
      l.addEventListener('click', function(){ selectReplConn(parseInt(l.getAttribute('data-conn'),10), conns); });
    });
  }

  function selectReplDC(name){
    var panel = document.getElementById('repl-panel');
    if (!panel) return;
    var dc = (window.ADAtlasData.DCs || []).find(function(x){ return x.Name === name; });
    var conns = (window.ADAtlasData.Connections || []).filter(function(c){ return c.From === name || c.To === name; });
    var html = '<h3>' + escapeHtml(name) + '</h3>';
    if (dc){
      html += field('Domain', dc.Domain);
      html += field('Site', dc.Site);
      html += field('Type', dc.Type);
      html += field('IP', dc.IPv4Address);
      html += field('OS', dc.OperatingSystem);
      html += field('Global Catalog', dc.IsGlobalCatalog ? 'Yes' : 'No');
    }
    html += '<div class="field"><label>Connections (' + conns.length + ')</label><div class="val dc-list">';
    if (conns.length === 0){ html += '<span class="dim">(none)</span>'; }
    else { conns.forEach(function(c){
      var dir = (c.From === name) ? '\u2192 ' + c.To : '\u2190 ' + c.From;
      var badge = c.AutoGenerated ? '<span class="pill pill-gc">KCC</span>' : '<span class="pill pill-rodc">Manual</span>';
      html += '<div class="dc-item">' + escapeHtml(dir) + badge + '<div class="dc-meta">' + escapeHtml(c.Site || '') + '</div></div>';
    });}
    html += '</div></div>';
    panel.innerHTML = html;
  }

  function selectReplConn(idx, conns){
    var panel = document.getElementById('repl-panel');
    var c = conns[idx];
    if (!panel || !c) return;
    var html = '<h3>Replication Connection</h3>';
    html += field('Name', c.Name);
    html += field('From', c.From);
    html += field('To', c.To);
    html += field('Site', c.Site);
    html += field('Auto-Generated (KCC)', c.AutoGenerated ? 'Yes' : 'No (manual)');
    html += field('Enabled', c.Enabled ? 'Yes' : 'No');
    panel.innerHTML = html;
  }

  // ==================== DC INVENTORY (sortable, filterable) ====================
  var dcSortKey = 'Name';
  var dcSortAsc = true;

  function renderDcTable(){
    var tbody = document.getElementById('dc-tbody');
    if (!tbody) return;
    var dcs = (window.ADAtlasData.DCs || []).slice();
    var search = (document.getElementById('dc-search').value || '').toLowerCase();
    var fDom = document.getElementById('dc-filter-domain').value;
    var fSite = document.getElementById('dc-filter-site').value;
    var fType = document.getElementById('dc-filter-type').value;
    dcs = dcs.filter(function(d){
      if (fDom && d.Domain !== fDom) return false;
      if (fSite && d.Site !== fSite) return false;
      if (fType && d.Type !== fType) return false;
      if (search){
        var hay = (d.Name + ' ' + d.HostName + ' ' + d.Domain + ' ' + d.Site + ' ' + d.IPv4Address + ' ' + d.OperatingSystem + ' ' + d.OSVersion).toLowerCase();
        if (hay.indexOf(search) < 0) return false;
      }
      return true;
    });
    dcs.sort(function(a, b){
      var av = a[dcSortKey] || '', bv = b[dcSortKey] || '';
      if (av < bv) return dcSortAsc ? -1 : 1;
      if (av > bv) return dcSortAsc ? 1 : -1;
      return 0;
    });
    var html = '';
    dcs.forEach(function(d){
      var typeBadge = d.Type === 'RODC' ? '<span class="badge badge-amber">RODC</span>' : '<span class="badge badge-green">RWDC</span>';
      var gcBadge = d.IsGlobalCatalog ? '<span class="badge badge-purple">GC</span>' : '<span class="dim">--</span>';
      var fsmoText = (d.FSMORoles && d.FSMORoles.length > 0) ? d.FSMORoles.join(', ') : '<span class="dim">--</span>';
      html += '<tr><td><strong>' + escapeHtml(d.Name) + '</strong></td>';
      html += '<td>' + escapeHtml(d.Domain) + '</td>';
      html += '<td>' + escapeHtml(d.Site || '') + '</td>';
      html += '<td><code>' + escapeHtml(d.IPv4Address || '') + '</code></td>';
      html += '<td>' + escapeHtml(d.OperatingSystem || '') + '</td>';
      html += '<td>' + escapeHtml(d.OSVersion || '') + '</td>';
      html += '<td>' + typeBadge + '</td>';
      html += '<td>' + gcBadge + '</td>';
      html += '<td style="font-size:.7rem">' + fsmoText + '</td></tr>';
    });
    tbody.innerHTML = html;
    var countEl = document.getElementById('dc-count');
    if (countEl) countEl.innerHTML = 'Showing <strong>' + dcs.length + '</strong> of <strong>' + (window.ADAtlasData.DCs || []).length + '</strong>';
    document.querySelectorAll('#dc-table th[data-sort]').forEach(function(th){
      th.classList.remove('sorted-asc','sorted-desc');
      if (th.getAttribute('data-sort') === dcSortKey){
        th.classList.add(dcSortAsc ? 'sorted-asc' : 'sorted-desc');
      }
    });
  }

  function initDcInventory(){
    var search = document.getElementById('dc-search');
    if (!search) return;
    search.addEventListener('input', renderDcTable);
    document.getElementById('dc-filter-domain').addEventListener('change', renderDcTable);
    document.getElementById('dc-filter-site').addEventListener('change', renderDcTable);
    document.getElementById('dc-filter-type').addEventListener('change', renderDcTable);
    document.querySelectorAll('#dc-table th[data-sort]').forEach(function(th){
      th.addEventListener('click', function(){
        var key = th.getAttribute('data-sort');
        if (dcSortKey === key) dcSortAsc = !dcSortAsc;
        else { dcSortKey = key; dcSortAsc = true; }
        renderDcTable();
      });
    });
    renderDcTable();
  }

  // ==================== DNS ARCHITECTURE ====================
  function drawDnsArchitecture(){
    var container = document.getElementById('dns-canvas');
    if (!container) return;
    var dns = window.ADAtlasData.DNS || {};
    var zones = (dns.Zones || []).filter(function(z){ return !z.IsAutoCreated && !z.IsReverseLookupZone; });
    var dcs = (window.ADAtlasData.DCs || []);
    if (zones.length === 0 && dcs.length === 0){
      container.innerHTML = '<svg viewBox="0 0 600 100" xmlns="http://www.w3.org/2000/svg"><text x="20" y="50" fill="#94a3b8" font-family="Segoe UI" font-size="14">No DNS data to display.</text></svg>';
      return;
    }

    // DCs on the left, zones on the right
    var dcW = 200, dcH = 38, dcGap = 14;
    var zoneW = 280, zoneH = 38, zoneGap = 14;
    var leftX = 80, rightX = 600;

    // Limit to displaying first ~15 DCs and ~20 zones to keep diagram readable
    var displayDcs = dcs.slice(0, 15);
    var displayZones = zones.slice(0, 20);

    var dcStartY = 60;
    var zoneStartY = 60;

    var viewH = Math.max(
      dcStartY + displayDcs.length * (dcH + dcGap),
      zoneStartY + displayZones.length * (zoneH + zoneGap)
    ) + 60;
    var viewW = 980;

    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMid meet" style="min-width:780px">');

    // Headers
    svgParts.push('<text x="' + (leftX + dcW/2) + '" y="35" text-anchor="middle" font-family="Segoe UI" font-size="13" font-weight="700" fill="#60a5fa">DOMAIN CONTROLLERS / DNS SERVERS</text>');
    svgParts.push('<text x="' + (rightX + zoneW/2) + '" y="35" text-anchor="middle" font-family="Segoe UI" font-size="13" font-weight="700" fill="#34d399">DNS ZONES</text>');

    // Draw connecting lines (each DC hosts AD-integrated zones; we draw faint lines from each DC to each AD-integrated zone)
    var adIntegratedZones = displayZones.filter(function(z){ return z.IsDsIntegrated; });
    displayDcs.forEach(function(dc, di){
      var dcCy = dcStartY + di * (dcH + dcGap) + dcH/2;
      adIntegratedZones.forEach(function(z){
        var zi = displayZones.indexOf(z);
        var zCy = zoneStartY + zi * (zoneH + zoneGap) + zoneH/2;
        svgParts.push('<line x1="' + (leftX + dcW) + '" y1="' + dcCy + '" x2="' + rightX + '" y2="' + zCy + '" stroke="#475569" stroke-width="1" opacity="0.25"/>');
      });
    });

    // Draw DCs
    displayDcs.forEach(function(dc, di){
      var y = dcStartY + di * (dcH + dcGap);
      svgParts.push('<g class="node-box">');
      svgParts.push('<rect x="' + leftX + '" y="' + y + '" width="' + dcW + '" height="' + dcH + '" rx="6" fill="rgba(96,165,250,0.15)" stroke="#60a5fa" stroke-width="2"/>');
      svgParts.push('<text x="' + (leftX + 10) + '" y="' + (y + 16) + '" font-family="Segoe UI" font-size="11" font-weight="700" fill="#e2e8f0">' + escapeXml(dc.Name) + '</text>');
      svgParts.push('<text x="' + (leftX + 10) + '" y="' + (y + 30) + '" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(dc.Domain) + '</text>');
      svgParts.push('</g>');
    });
    if (dcs.length > displayDcs.length){
      var moreY = dcStartY + displayDcs.length * (dcH + dcGap) + 8;
      svgParts.push('<text x="' + (leftX + dcW/2) + '" y="' + moreY + '" text-anchor="middle" font-family="Segoe UI" font-size="10" font-style="italic" fill="#94a3b8">... and ' + (dcs.length - displayDcs.length) + ' more</text>');
    }

    // Draw zones
    displayZones.forEach(function(z, zi){
      var y = zoneStartY + zi * (zoneH + zoneGap);
      var color = z.IsDsIntegrated ? '#34d399' : '#fbbf24';
      var bg = z.IsDsIntegrated ? 'rgba(52,211,153,0.15)' : 'rgba(251,191,36,0.15)';
      svgParts.push('<g class="node-box">');
      svgParts.push('<rect x="' + rightX + '" y="' + y + '" width="' + zoneW + '" height="' + zoneH + '" rx="6" fill="' + bg + '" stroke="' + color + '" stroke-width="2"/>');
      svgParts.push('<text x="' + (rightX + 10) + '" y="' + (y + 16) + '" font-family="Segoe UI" font-size="11" font-weight="700" fill="#e2e8f0">' + escapeXml(z.ZoneName) + '</text>');
      var zType = z.IsDsIntegrated ? 'AD-Integrated' : 'Standalone';
      svgParts.push('<text x="' + (rightX + 10) + '" y="' + (y + 30) + '" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(zType + ' &middot; ' + (z.ReplicationScope || '')) + '</text>');
      svgParts.push('</g>');
    });
    if (zones.length > displayZones.length){
      var moreY2 = zoneStartY + displayZones.length * (zoneH + zoneGap) + 8;
      svgParts.push('<text x="' + (rightX + zoneW/2) + '" y="' + moreY2 + '" text-anchor="middle" font-family="Segoe UI" font-size="10" font-style="italic" fill="#94a3b8">... and ' + (zones.length - displayZones.length) + ' more (see table below)</text>');
    }

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');
  }

  // ==================== NTP HIERARCHY (real w32tm data) ====================
  function normalizeNtpSource(src){
    if (!src) return '';
    var s = String(src).trim();
    // Strip trailing flag annotations like ",0x9" or " (Local)"
    s = s.replace(/,0x[0-9a-fA-F]+/g, '');
    s = s.replace(/\s*\(Local\)\s*$/i, '');
    return s.trim();
  }

  function classifyNtpSource(src){
    if (!src) return 'unknown';
    var low = src.toLowerCase();
    if (low.indexOf('local cmos') >= 0 || low.indexOf('vm ic time') >= 0 || low.indexOf('free-running') >= 0 || low === 'local') return 'local';
    if (low.indexOf('byname') >= 0 || low.indexOf('flags') >= 0) return 'unresolved';
    return 'host';
  }

  function drawNtpHierarchy(){
    var container = document.getElementById('ntp-canvas');
    if (!container) return;
    var ntp = window.ADAtlasData.NTP || { DCs: [] };
    var ntpDcs = ntp.DCs || [];
    if (ntpDcs.length === 0){
      container.innerHTML = '<svg viewBox="0 0 600 100" xmlns="http://www.w3.org/2000/svg"><text x="20" y="50" fill="#94a3b8" font-family="Segoe UI" font-size="14">No NTP data collected.</text></svg>';
      return;
    }

    // Build host -> dc lookup so we can resolve "VM-EUDC01.landd.lab" -> dc node
    var dcByHost = {};
    ntpDcs.forEach(function(d){
      var hn = (d.HostName || '').toLowerCase();
      var nm = (d.Name || '').toLowerCase();
      if (hn) dcByHost[hn] = d;
      if (nm) dcByHost[nm] = d;
    });

    // For each DC, normalize its source string and figure out where it points
    ntpDcs.forEach(function(d){
      d._source = normalizeNtpSource(d.Source);
      d._kind = classifyNtpSource(d._source);
      d._pointsTo = null; // dc object if it points to another DC in our list
      if (d._kind === 'host'){
        var key = d._source.toLowerCase().split(',')[0].split(' ')[0].trim();
        if (dcByHost[key]) d._pointsTo = dcByHost[key];
        else {
          // try short-name match
          var shortName = key.split('.')[0];
          if (dcByHost[shortName]) d._pointsTo = dcByHost[shortName];
        }
      }
    });

    // ---- NT5DS hierarchy inference ----
    // For DCs configured Type=NT5DS (use AD hierarchy), the actual upstream is determined by AD itself:
    //   - non-PDC DC  -> its own domain's PDC
    //   - domain PDC  -> parent domain's PDC (or forest root PDC if no parent)
    //   - forest root PDC -> top of tree (uses local CMOS in pure NT5DS mode)
    // We have all this data from FSMO + domain parent/child info, so we can render the actual tree.
    var domainPdcMap = {};
    ntpDcs.forEach(function(d){ if (d.IsPDC) domainPdcMap[d.Domain] = d; });
    var allDomainsList = window.ADAtlasData.Domains || [];
    var parentByDomain = {};
    allDomainsList.forEach(function(d){ parentByDomain[d.DNSRoot] = d.ParentDomain || ''; });
    var rootDomainName = '';
    var fr = allDomainsList.find(function(d){ return d.IsForestRoot; });
    if (fr) rootDomainName = fr.DNSRoot;

    ntpDcs.forEach(function(d){
      if (d._pointsTo) return; // already resolved by source string
      if (d.ConfigType !== 'NT5DS') return; // only infer for NT5DS
      var upstream = null;
      if (d.IsForestRootPDC){
        upstream = null; // top of tree
      } else if (d.IsPDC){
        // Domain PDC syncs from parent domain's PDC, or forest root PDC if direct child of forest
        var parent = parentByDomain[d.Domain];
        if (parent && domainPdcMap[parent]){
          upstream = domainPdcMap[parent];
        } else if (rootDomainName && domainPdcMap[rootDomainName] && d.Domain !== rootDomainName){
          upstream = domainPdcMap[rootDomainName];
        }
      } else {
        // Non-PDC DC syncs from its own domain's PDC
        if (domainPdcMap[d.Domain]) upstream = domainPdcMap[d.Domain];
      }
      if (upstream && upstream !== d){
        d._pointsTo = upstream;
        d._inferred = true; // mark as AD-hierarchy-inferred
      }
    });

    // Identify the forest root PDC
    var rootPdc = ntpDcs.find(function(d){ return d.IsForestRootPDC; });

    // External sources: anything from a host-source DC where _pointsTo is null AND src is not local
    // Group DCs by their effective upstream
    // We will lay out as: External sources at top, DCs grouped beneath their upstream
    var externalSources = {};   // sourceString -> [dcs that point to it directly]
    var groupedByUpstreamDC = {}; // upstreamHostName -> [dcs that point to it]
    var rootlessDcs = [];       // DCs that are local/unresolved/unreachable
    var topOfTreeDcs = [];      // Forest root PDC (or other top-of-tree DCs with no upstream)

    ntpDcs.forEach(function(d){
      if (!d.Reachable){
        rootlessDcs.push(d);
        return;
      }
      if (d._pointsTo){
        var key = (d._pointsTo.HostName || d._pointsTo.Name).toLowerCase();
        if (!groupedByUpstreamDC[key]) groupedByUpstreamDC[key] = [];
        groupedByUpstreamDC[key].push(d);
      } else if (d.IsForestRootPDC && d.ConfigType === 'NT5DS'){
        // Forest root PDC in NT5DS mode = top of the AD time tree, uses local CMOS
        topOfTreeDcs.push(d);
      } else if (d._kind === 'host' && d._source){
        // Points to an external host (not in our forest)
        var key2 = d._source;
        if (!externalSources[key2]) externalSources[key2] = [];
        externalSources[key2].push(d);
      } else if (d.IsForestRootPDC){
        // Forest root PDC with unknown source - still put it at the top
        topOfTreeDcs.push(d);
      } else if (d._kind === 'local'){
        rootlessDcs.push(d);
      } else {
        rootlessDcs.push(d);
      }
    });

    // Layout
    var nodeW = 220, nodeH = 56, hGap = 24, vGap = 90;
    var pad = 40;
    var topY = 60;
    var pdcY = topY + 110;
    var dcY  = pdcY + 130;
    var orphanY = dcY + 130;

    // Top tier: external sources used by ANY DC. Spread horizontally.
    var extSourceList = Object.keys(externalSources);

    // Compute width estimate
    var allLeafDcs = ntpDcs.length;
    var viewW = Math.max(1100, pad * 2 + allLeafDcs * 110);
    var viewH = orphanY + (rootlessDcs.length > 0 ? 120 : 60);

    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMin meet" style="min-width:900px">');
    svgParts.push('<defs>');
    svgParts.push('<marker id="ntp-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#94a3b8"/></marker>');
    svgParts.push('</defs>');

    // Helper to draw a node and return its center coords
    function drawNode(x, y, w, h, color, bg, title, sub1, sub2){
      svgParts.push('<g class="node-box">');
      svgParts.push('<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" rx="8" fill="' + bg + '" stroke="' + color + '" stroke-width="2.5"/>');
      svgParts.push('<text x="' + (x + w/2) + '" y="' + (y + 18) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="' + color + '">' + escapeXml(title) + '</text>');
      if (sub1) svgParts.push('<text x="' + (x + w/2) + '" y="' + (y + 33) + '" text-anchor="middle" font-family="Segoe UI" font-size="10" fill="#e2e8f0">' + escapeXml(sub1) + '</text>');
      if (sub2) svgParts.push('<text x="' + (x + w/2) + '" y="' + (y + 47) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(sub2) + '</text>');
      svgParts.push('</g>');
      return { cx: x + w/2, cy: y + h/2, top: y, bot: y + h };
    }

    // Layout external sources spaced across the top
    var extPositions = {};
    if (extSourceList.length > 0){
      var perCol = Math.max(280, viewW / (extSourceList.length + 0.5));
      extSourceList.forEach(function(srcStr, i){
        var x = pad + i * perCol + (perCol - nodeW) / 2;
        var node = drawNode(x, topY, nodeW + 40, 56, '#22d3ee', 'rgba(34,211,238,0.12)', 'EXTERNAL SOURCE', srcStr, externalSources[srcStr].length + ' DC(s) sync from this');
        extPositions[srcStr] = node;
      });
    } else if (topOfTreeDcs.length > 0){
      // No external sources but we have a forest root PDC at NT5DS top-of-tree.
      // Show a synthetic "Local CMOS / Top of NT5DS hierarchy" node.
      var topNode = drawNode(pad + (viewW - pad*2)/2 - (nodeW+40)/2, topY, nodeW + 40, 56, '#fbbf24', 'rgba(251,191,36,0.10)', 'TOP OF AD TIME HIERARCHY', 'Local CMOS clock', 'Forest root PDC in NT5DS mode (no AD upstream)');
      extPositions['__topoftree__'] = topNode;
    }

    // Determine which DCs go in the "PDC tier" (those that have other DCs pointing to them, OR forest root PDC)
    var pdcTierDcs = [];
    var seen = {};
    Object.keys(groupedByUpstreamDC).forEach(function(k){
      var upstreamDc = ntpDcs.find(function(d){
        return (d.HostName || d.Name).toLowerCase() === k;
      });
      if (upstreamDc && !seen[k]){ pdcTierDcs.push(upstreamDc); seen[k] = true; }
    });
    // Make sure forest root PDC is in the tier even if nothing points to it
    if (rootPdc){
      var rk = (rootPdc.HostName || rootPdc.Name).toLowerCase();
      if (!seen[rk]){ pdcTierDcs.unshift(rootPdc); seen[rk] = true; }
    }

    // Layout PDC tier centered
    var pdcPositions = {};
    if (pdcTierDcs.length > 0){
      var pdcSpacing = Math.min(320, (viewW - pad * 2) / pdcTierDcs.length);
      var pdcStartX = (viewW - (pdcTierDcs.length - 1) * pdcSpacing) / 2 - nodeW/2;
      pdcTierDcs.forEach(function(dc, i){
        var x = pdcStartX + i * pdcSpacing;
        var isRoot = dc.IsForestRootPDC;
        var color = isRoot ? '#60a5fa' : (dc.IsPDC ? '#a78bfa' : '#34d399');
        var bg = isRoot ? 'rgba(96,165,250,0.18)' : (dc.IsPDC ? 'rgba(167,139,250,0.12)' : 'rgba(52,211,153,0.10)');
        var title = isRoot ? 'FOREST ROOT PDC' : (dc.IsPDC ? (dc.Domain + ' PDC') : dc.Domain + ' DC');
        // Build a meaningful source label
        var src;
        if (!dc.Reachable) {
          color = '#94a3b8'; bg = 'rgba(148,163,184,0.06)'; src = '(unreachable)';
        } else if (dc.ConfigType === 'NT5DS') {
          if (isRoot) src = 'Local CMOS (NT5DS, top of tree)';
          else if (dc._pointsTo) src = 'NT5DS hierarchy -> ' + (dc._pointsTo.HostName || dc._pointsTo.Name);
          else src = 'NT5DS (no upstream resolved)';
        } else if (dc._source) {
          src = dc._source;
        } else {
          src = '(' + (dc.ConfigType || 'unknown type') + ')';
        }
        var node = drawNode(x, pdcY, nodeW, nodeH, color, bg, title, dc.HostName, 'src: ' + src);
        pdcPositions[(dc.HostName || dc.Name).toLowerCase()] = node;

        // Draw arrow from external source if applicable
        if (dc.Reachable && dc._kind === 'host' && !dc._pointsTo && dc._source && extPositions[dc._source]){
          var ext = extPositions[dc._source];
          svgParts.push('<line x1="' + ext.cx + '" y1="' + ext.bot + '" x2="' + node.cx + '" y2="' + node.top + '" stroke="#94a3b8" stroke-width="2" marker-end="url(#ntp-arrow)"/>');
        }
        // Draw arrow from synthetic "top of tree" node to forest root PDC if applicable
        if (isRoot && dc.ConfigType === 'NT5DS' && extPositions['__topoftree__']){
          var topN = extPositions['__topoftree__'];
          svgParts.push('<line x1="' + topN.cx + '" y1="' + topN.bot + '" x2="' + node.cx + '" y2="' + node.top + '" stroke="#94a3b8" stroke-width="2" marker-end="url(#ntp-arrow)"/>');
        }
      });
    }

    // Layout child DCs underneath their upstream PDC
    Object.keys(groupedByUpstreamDC).forEach(function(upstreamKey){
      var children = groupedByUpstreamDC[upstreamKey];
      var parentNode = pdcPositions[upstreamKey];
      if (!parentNode) return;
      var spacing = 240;
      var totalW = children.length * spacing;
      var startX = parentNode.cx - totalW/2 + (spacing - nodeW) / 2;
      children.forEach(function(dc, i){
        if ((dc.HostName || dc.Name).toLowerCase() === upstreamKey) return; // skip self
        var x = startX + i * spacing;
        var color, bg;
        if (!dc.Reachable){ color = '#94a3b8'; bg = 'rgba(148,163,184,0.06)'; }
        else if (dc.IsPDC){ color = '#a78bfa'; bg = 'rgba(167,139,250,0.12)'; }
        else { color = '#34d399'; bg = 'rgba(52,211,153,0.10)'; }
        var title = dc.IsPDC ? (dc.Domain + ' PDC') : dc.Domain + ' DC';
        var sub2 = 'type: ' + (dc.ConfigType || '?') + (dc.Site ? ' \u00b7 ' + dc.Site : '');
        var node = drawNode(x, dcY, nodeW, nodeH, color, bg, title, dc.HostName, sub2);
        svgParts.push('<line x1="' + parentNode.cx + '" y1="' + parentNode.bot + '" x2="' + node.cx + '" y2="' + node.top + '" stroke="#94a3b8" stroke-width="1.5" marker-end="url(#ntp-arrow)"/>');
      });
    });

    // Orphan / local / unreachable row
    if (rootlessDcs.length > 0){
      svgParts.push('<text x="' + pad + '" y="' + (orphanY - 14) + '" font-family="Segoe UI" font-size="11" font-weight="700" fill="#94a3b8">Local clock / unreachable / unresolved (' + rootlessDcs.length + '):</text>');
      var spacing2 = 240;
      rootlessDcs.forEach(function(dc, i){
        var perRow = Math.floor((viewW - pad * 2) / spacing2);
        if (perRow < 1) perRow = 1;
        var row = Math.floor(i / perRow);
        var col = i % perRow;
        var x = pad + col * spacing2;
        var y = orphanY + row * (nodeH + 14);
        var color, bg, label;
        if (!dc.Reachable){ color = '#94a3b8'; bg = 'rgba(148,163,184,0.06)'; label = '(unreachable)'; }
        else if (dc._kind === 'local'){ color = '#fbbf24'; bg = 'rgba(251,191,36,0.10)'; label = 'LOCAL CMOS clock'; }
        else { color = '#fbbf24'; bg = 'rgba(251,191,36,0.10)'; label = dc._source || '(unresolved)'; }
        var title = dc.IsForestRootPDC ? 'FOREST ROOT PDC' : (dc.IsPDC ? (dc.Domain + ' PDC') : dc.Domain + ' DC');
        drawNode(x, y, nodeW, nodeH, color, bg, title, dc.HostName, label);
      });
    }

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');

    // Render per-DC table
    var tw = document.getElementById('ntp-table-wrap');
    if (tw){
      var rows = '';
      ntpDcs.forEach(function(d){
        var statusBadge = d.Reachable ? '<span class="badge badge-green">Reachable</span>' : '<span class="badge badge-amber">Unreachable</span>';
        var roleBadge = d.IsForestRootPDC ? '<span class="badge badge-accent">Forest Root PDC</span>' : (d.IsPDC ? '<span class="badge badge-purple">PDC</span>' : '');
        var srcDisplay = d.Reachable ? (d._source || '<span class="dim">--</span>') : '<span class="dim">--</span>';
        var typeDisplay = d.Reachable ? (d.ConfigType || '<span class="dim">--</span>') : '<span class="dim">--</span>';
        var ntpCfgDisplay = d.Reachable ? (d.NtpServerCfg || '<span class="dim">--</span>') : '<span class="dim">--</span>';
        var methodDisplay = d.Method ? '<span class="dim" style="font-size:.7rem">' + escapeHtml(d.Method) + '</span>' : '<span class="dim">--</span>';
        rows += '<tr><td><strong>' + escapeHtml(d.HostName || d.Name) + '</strong> ' + roleBadge + '</td>';
        rows += '<td>' + escapeHtml(d.Domain) + '</td>';
        rows += '<td>' + escapeHtml(d.Site || '') + '</td>';
        rows += '<td>' + statusBadge + '</td>';
        rows += '<td><code>' + srcDisplay + '</code></td>';
        rows += '<td>' + typeDisplay + '</td>';
        rows += '<td style="font-size:.72rem"><code>' + ntpCfgDisplay + '</code></td>';
        rows += '<td>' + methodDisplay + '</td></tr>';
      });
      tw.innerHTML = '<div class="table-wrap"><table><thead><tr><th>DC</th><th>Domain</th><th>Site</th><th>Status</th><th>Active Source</th><th>Type</th><th>Configured NtpServer</th><th>Collected Via</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
    }
  }

  // ==================== EXCHANGE ARCHITECTURE ====================
  function drawExchangeArchitecture(){
    var container = document.getElementById('exchange-canvas');
    if (!container) return;
    var ex = window.ADAtlasData.Exchange || {};
    if (!ex.Detected){
      container.innerHTML = '<svg viewBox="0 0 600 80" xmlns="http://www.w3.org/2000/svg"><text x="20" y="45" fill="#94a3b8" font-family="Segoe UI" font-size="14">No Exchange detected.</text></svg>';
      return;
    }
    var servers = ex.Servers || [];
    var dags = ex.DAGs || [];

    var viewW = 1100, viewH = Math.max(400, 220 + Math.ceil(servers.length / 4) * 100);
    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMid meet" style="min-width:800px">');
    svgParts.push('<defs>');
    svgParts.push('<marker id="ex-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#22d3ee"/></marker>');
    svgParts.push('</defs>');

    // Exchange Org node at top
    var orgCx = viewW / 2, orgCy = 60;
    svgParts.push('<g>');
    svgParts.push('<rect x="' + (orgCx - 200) + '" y="' + (orgCy - 32) + '" width="400" height="64" rx="8" fill="rgba(34,211,238,0.18)" stroke="#22d3ee" stroke-width="3"/>');
    svgParts.push('<text x="' + orgCx + '" y="' + (orgCy - 8) + '" text-anchor="middle" font-family="Segoe UI" font-size="13" font-weight="700" fill="#22d3ee">EXCHANGE ORGANIZATION</text>');
    svgParts.push('<text x="' + orgCx + '" y="' + (orgCy + 8) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#e2e8f0">' + escapeXml(ex.OrganizationName || '(unnamed)') + '</text>');
    svgParts.push('<text x="' + orgCx + '" y="' + (orgCy + 22) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(ex.VersionDisplay || ('Schema ' + ex.SchemaVersion)) + '</text>');
    svgParts.push('</g>');

    // Hybrid Cloud node (if hybrid)
    if (ex.Hybrid){
      var hbCx = viewW - 140, hbCy = 60;
      svgParts.push('<g>');
      svgParts.push('<rect x="' + (hbCx - 100) + '" y="' + (hbCy - 28) + '" width="200" height="56" rx="8" fill="rgba(96,165,250,0.15)" stroke="#60a5fa" stroke-width="2.5" stroke-dasharray="5,4"/>');
      svgParts.push('<text x="' + hbCx + '" y="' + (hbCy - 4) + '" text-anchor="middle" font-family="Segoe UI" font-size="12" font-weight="700" fill="#60a5fa">Exchange Online</text>');
      svgParts.push('<text x="' + hbCx + '" y="' + (hbCy + 14) + '" text-anchor="middle" font-family="Segoe UI" font-size="10" fill="#94a3b8">Hybrid configured</text>');
      svgParts.push('</g>');
      svgParts.push('<line x1="' + (orgCx + 200) + '" y1="' + orgCy + '" x2="' + (hbCx - 100) + '" y2="' + hbCy + '" stroke="#60a5fa" stroke-width="2" stroke-dasharray="5,4" marker-end="url(#ex-arrow)"/>');
    }

    // DAGs row (if any)
    var dagY = 170;
    if (dags.length > 0){
      var dagSpacing = Math.min(220, (viewW - 100) / dags.length);
      var dagStartX = (viewW - (dags.length - 1) * dagSpacing) / 2;
      dags.forEach(function(dag, i){
        var dx = dagStartX + i * dagSpacing;
        svgParts.push('<line x1="' + orgCx + '" y1="' + (orgCy + 32) + '" x2="' + dx + '" y2="' + (dagY - 24) + '" stroke="#475569" stroke-width="1.5"/>');
        svgParts.push('<g>');
        svgParts.push('<rect x="' + (dx - 90) + '" y="' + (dagY - 24) + '" width="180" height="48" rx="6" fill="rgba(167,139,250,0.12)" stroke="#a78bfa" stroke-width="2"/>');
        svgParts.push('<text x="' + dx + '" y="' + (dagY - 4) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#a78bfa">DAG</text>');
        svgParts.push('<text x="' + dx + '" y="' + (dagY + 12) + '" text-anchor="middle" font-family="Segoe UI" font-size="10" fill="#e2e8f0">' + escapeXml(dag.Name) + '</text>');
        svgParts.push('</g>');
      });
    }

    // Servers grid
    var srvStartY = dags.length > 0 ? 270 : 180;
    if (servers.length > 0){
      svgParts.push('<text x="' + orgCx + '" y="' + (srvStartY - 10) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#94a3b8">EXCHANGE SERVERS</text>');
      var perRow = 4;
      var srvW = 200, srvH = 60, gapX = 24, gapY = 18;
      var rowCount = Math.ceil(servers.length / perRow);
      servers.forEach(function(s, i){
        var row = Math.floor(i / perRow);
        var col = i % perRow;
        var rowItems = Math.min(perRow, servers.length - row * perRow);
        var rowW = rowItems * srvW + (rowItems - 1) * gapX;
        var startX = (viewW - rowW) / 2;
        var sx = startX + col * (srvW + gapX);
        var sy = srvStartY + row * (srvH + gapY);
        svgParts.push('<g class="node-box">');
        svgParts.push('<rect x="' + sx + '" y="' + sy + '" width="' + srvW + '" height="' + srvH + '" rx="6" fill="rgba(34,211,238,0.1)" stroke="#22d3ee" stroke-width="2"/>');
        svgParts.push('<text x="' + (sx + srvW/2) + '" y="' + (sy + 18) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#e2e8f0">' + escapeXml(s.Name) + '</text>');
        var roleStr = (s.Roles && s.Roles.length > 0) ? s.Roles.join(', ') : '(unknown roles)';
        svgParts.push('<text x="' + (sx + srvW/2) + '" y="' + (sy + 34) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(roleStr) + '</text>');
        svgParts.push('<text x="' + (sx + srvW/2) + '" y="' + (sy + 48) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(s.Created || '') + '</text>');
        svgParts.push('</g>');
      });
    }

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');
  }

  // ==================== PKI HIERARCHY ====================
  function drawPkiHierarchy(){
    var container = document.getElementById('pki-canvas');
    if (!container) return;
    var pki = window.ADAtlasData.PKI || {};
    if (!pki.Detected){
      container.innerHTML = '<svg viewBox="0 0 600 80" xmlns="http://www.w3.org/2000/svg"><text x="20" y="45" fill="#94a3b8" font-family="Segoe UI" font-size="14">No PKI detected.</text></svg>';
      return;
    }
    var roots = pki.RootCAs || [];
    var ents = pki.EnterpriseCAs || [];

    var viewW = 1100;
    var viewH = 100 + (roots.length > 0 ? 100 : 0) + (ents.length > 0 ? 120 : 0);
    var svgParts = [];
    svgParts.push('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + viewW + ' ' + viewH + '" preserveAspectRatio="xMidYMid meet" style="min-width:800px">');

    // Root CAs (top tier)
    var rootY = 50;
    if (roots.length > 0){
      var rootSpacing = Math.min(280, (viewW - 80) / Math.max(roots.length, 1));
      var rootStartX = (viewW - (roots.length - 1) * rootSpacing) / 2;
      roots.forEach(function(r, i){
        var rx = rootStartX + i * rootSpacing;
        svgParts.push('<g class="node-box">');
        svgParts.push('<rect x="' + (rx - 130) + '" y="' + (rootY - 32) + '" width="260" height="64" rx="8" fill="rgba(167,139,250,0.18)" stroke="#a78bfa" stroke-width="3"/>');
        svgParts.push('<text x="' + rx + '" y="' + (rootY - 10) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#a78bfa">TRUSTED ROOT CA</text>');
        svgParts.push('<text x="' + rx + '" y="' + (rootY + 8) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#e2e8f0">' + escapeXml(r.Name) + '</text>');
        svgParts.push('<text x="' + rx + '" y="' + (rootY + 22) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(r.Created || '') + '</text>');
        svgParts.push('</g>');
      });
    }

    // Enterprise CAs (next tier)
    if (ents.length > 0){
      var entY = roots.length > 0 ? 200 : 80;
      var entSpacing = Math.min(280, (viewW - 80) / Math.max(ents.length, 1));
      var entStartX = (viewW - (ents.length - 1) * entSpacing) / 2;
      svgParts.push('<text x="' + (viewW/2) + '" y="' + (entY - 50) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#94a3b8">ENTERPRISE / ISSUING CAs</text>');
      ents.forEach(function(ca, i){
        var cx = entStartX + i * entSpacing;
        // Connect to nearest root
        if (roots.length > 0){
          var nearestRoot = Math.min(Math.floor(i * roots.length / Math.max(ents.length, 1)), roots.length - 1);
          var rootSpacing2 = Math.min(280, (viewW - 80) / Math.max(roots.length, 1));
          var rootStartX2 = (viewW - (roots.length - 1) * rootSpacing2) / 2;
          var rx2 = rootStartX2 + nearestRoot * rootSpacing2;
          svgParts.push('<line x1="' + rx2 + '" y1="' + (rootY + 32) + '" x2="' + cx + '" y2="' + (entY - 32) + '" stroke="#475569" stroke-width="2"/>');
        }
        svgParts.push('<g class="node-box">');
        svgParts.push('<rect x="' + (cx - 140) + '" y="' + (entY - 32) + '" width="280" height="76" rx="8" fill="rgba(96,165,250,0.15)" stroke="#60a5fa" stroke-width="2.5"/>');
        svgParts.push('<text x="' + cx + '" y="' + (entY - 12) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#60a5fa">ENTERPRISE CA</text>');
        svgParts.push('<text x="' + cx + '" y="' + (entY + 4) + '" text-anchor="middle" font-family="Segoe UI" font-size="11" font-weight="700" fill="#e2e8f0">' + escapeXml(ca.Name) + '</text>');
        svgParts.push('<text x="' + cx + '" y="' + (entY + 18) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + escapeXml(ca.Server || '') + '</text>');
        var tplCount = (ca.PublishedTemplates || []).length;
        svgParts.push('<text x="' + cx + '" y="' + (entY + 32) + '" text-anchor="middle" font-family="Segoe UI" font-size="9" fill="#94a3b8">' + tplCount + ' template(s) published</text>');
        svgParts.push('</g>');
      });
    }

    svgParts.push('</svg>');
    container.innerHTML = svgParts.join('');
  }

  // ==================== TAB SWITCHING ====================
  var drawn = {};
  function ensureDrawn(tabId){
    if (drawn[tabId]) return;
    drawn[tabId] = true;
    if (tabId === 'tab-forest') drawForestMap();
    else if (tabId === 'tab-trusts') drawTrustMap();
    else if (tabId === 'tab-sites') drawSiteTopology();
    else if (tabId === 'tab-replication'){ populateReplSiteSelector(); drawReplicationTopology('__ALL__'); }
    else if (tabId === 'tab-dcs') initDcInventory();
    else if (tabId === 'tab-dns') drawDnsArchitecture();
    else if (tabId === 'tab-ntp') drawNtpHierarchy();
    else if (tabId === 'tab-exchange') drawExchangeArchitecture();
    else if (tabId === 'tab-pki') drawPkiHierarchy();
  }

  function showTab(tabId, focus){
    document.querySelectorAll('.tab-content').forEach(function(t){ t.classList.remove('active'); });
    document.querySelectorAll('.sidebar nav a').forEach(function(a){ a.classList.remove('active'); });
    var tab = document.getElementById(tabId);
    if (tab) tab.classList.add('active');
    var navLink = document.querySelector('.sidebar nav a[data-tab="' + tabId + '"]' + (focus ? '[data-focus="' + focus + '"]' : ''));
    if (!navLink) navLink = document.querySelector('.sidebar nav a[data-tab="' + tabId + '"]');
    if (navLink) navLink.classList.add('active');
    ensureDrawn(tabId);
    if (focus && tabId === 'tab-forest'){
      setTimeout(function(){ selectDomain(focus); }, 60);
    }
    window.scrollTo(0, 0);
  }

  document.addEventListener('DOMContentLoaded', function(){
    document.querySelectorAll('.sidebar nav a[data-tab]').forEach(function(a){
      a.addEventListener('click', function(e){
        e.preventDefault();
        showTab(a.getAttribute('data-tab'), a.getAttribute('data-focus'));
      });
    });
    // Initial tab
    ensureDrawn('tab-overview');
  });

})();
</script>
</body>
</html>
"@

# ==============================================================================
# WRITE HTML AND FINISH
# ==============================================================================

try {
    $Html | Out-File -FilePath $OutputHtml -Encoding UTF8 -Force
    $size = [math]::Round((Get-Item $OutputHtml).Length / 1KB, 1)
    Write-Host ""
    Write-Host " +============================================================+" -ForegroundColor Green
    Write-Host " |                                                            |" -ForegroundColor Green
    Write-Host " |         ADAtlas report generated successfully              |" -ForegroundColor Green
    Write-Host " |                                                            |" -ForegroundColor Green
    Write-Host " +============================================================+" -ForegroundColor Green
    Write-Host ""
    Write-Host " [+] Output : $OutputHtml" -ForegroundColor Cyan
    Write-Host " [+] Size   : $size KB" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " Open the HTML file in any modern browser to view the interactive map." -ForegroundColor White
    Write-Host ""
    Write-Host " ADAtlas -- Map Every Corner of Your Active Directory" -ForegroundColor DarkCyan
    Write-Host " Developed by Santhosh Sivarajan, Microsoft MVP" -ForegroundColor DarkCyan
    Write-Host " https://github.com/SanthoshSivarajan/ADAtlas" -ForegroundColor DarkCyan
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host " [!] Failed to write HTML file: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
}

<#
================================================================================
 End of ADAtlas.ps1
--------------------------------------------------------------------------------
 ADAtlas -- Map Every Corner of Your Active Directory
 Version  : 1.0
 Author   : Santhosh Sivarajan, Microsoft MVP
 LinkedIn : https://www.linkedin.com/in/sivarajan/
 GitHub   : https://github.com/SanthoshSivarajan/ADAtlas
 License  : MIT
================================================================================
#>
