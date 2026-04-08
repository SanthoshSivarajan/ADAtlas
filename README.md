# ADAtlas

> **Map Every Corner of Your Active Directory**

ADAtlas is a single-file PowerShell script that produces a self-contained, interactive HTML map of an Active Directory forest. It is the visual companion to [ADCanvas](https://github.com/SanthoshSivarajan/ADCanvas) — where ADCanvas focuses on health, ADAtlas focuses purely on **what your AD looks like right now**.

No health checks. No analysis. No scoring. Just a complete picture of the current configuration.

---

## Why ADAtlas?

When you walk into a new Active Directory environment — or revisit one you haven't touched in a while — the first question is always the same: *what's actually in here?*

Existing tools either run a hundred health checks you didn't ask for, depend on cloud services, or produce a dozen Visio files you have to assemble yourself. ADAtlas does one thing: it inventories the forest and draws it as a single HTML file you can open, share, archive, or attach to a ticket.

Run it once. Open the HTML. See everything.

---

## What it shows

### Forest & Trusts
- **Forest & Domains Map** — hierarchical view of the forest tree, with domains drawn as classic AD triangles
- **Trust Map** — redesigned hierarchical layout showing forest domains on the left, external trust partners grouped on the right, with trust lines drawn as smooth curves and proper directional arrows
- **Trust Matrix** — full source-vs-target grid for all trusts, intra-forest and external

### Sites & Replication
- **Site Topology** — sites drawn as circles sized by DC count, connected by site links
- **Site Links** — full table sorted by cost
- **Replication Topology** — KCC and manual connection objects, with a per-site dropdown
- **Sites & Subnets** — combined view linking sites to their subnets and DC counts

### Domain Controllers
- **DC Inventory** — sortable, filterable, searchable table of every DC across the forest with OS, IP, site, FSMO roles, GC status, and RODC flag

### Supporting Services
- **DNS Architecture** — every AD-integrated and standalone DNS zone with replication scope, dynamic update settings, and configured forwarders
- **NTP Hierarchy** — actual time configuration on every DC, collected via `w32tm /query` with a WMI registry fallback. NT5DS-mode DCs are drawn in the AD time hierarchy they actually resolve to (forest root PDC → child domain PDCs → member DCs)
- **Exchange** — organization name, schema version (decoded), servers, DAGs, accepted domains, hybrid configuration detection — all read directly from the AD configuration partition
- **Certificate Services (PKI)** — Trusted Root CAs, Enterprise Issuing CAs, NTAuth store, and all published certificate templates
- **Authentication** — per-domain KRBTGT account state, Kerberos encryption types (with multi-path registry fallback), Default Domain Password Policy, **Kerberos Policy parsed from the Default Domain Policy GPO in SYSVOL**, Fine-Grained Password Policies, Protected Users group membership count, and Pre-Windows 2000 Compatible Access membership (with foreign-SID translation so well-known SIDs like `Authenticated Users` show up by name, not as raw SIDs)

### Hybrid Identity
- **Entra Connect detection** — automatic detection via MSOL/AAD/Sync service accounts in AD, with sync server name and tenant extraction. Entra ID is drawn as a stylized 3D blue diamond node attached to the forest root.

---

## What ADAtlas does NOT do

ADAtlas is intentionally **picture-only**. It does not:

- Run health checks or assign scores
- Test replication, DNS resolution, time drift, or trust validity
- Send data anywhere
- Require WinRM or PowerShell remoting
- Modify any AD object or registry value
- Require any external module beyond the standard `ActiveDirectory` module (and optionally `DnsServer` for the DNS view)

If you want health analysis, use **[ADCanvas](https://github.com/SanthoshSivarajan/ADCanvas)**. If you want a picture, use ADAtlas.

---

## Requirements

- **Windows PowerShell 5.1** or newer (PowerShell 7 also works)
- **Active Directory module** for Windows PowerShell (`RSAT-AD-PowerShell`)
- **DnsServer module** (optional, used for the DNS Architecture view — `RSAT-DNS-Server`)
- A user account with **read access to AD** (Domain Admin or equivalent recommended for complete data, but a regular Domain User will collect most non-sensitive data)
- For NTP and Encryption Type collection, the script tries `w32tm /query` first and falls back to **WMI (DCOM)** to read the registry. No WinRM required.

ADAtlas runs from anywhere domain-joined with line of sight to the DCs. It does not need to run on a DC.

---

## Usage

```powershell
.\ADAtlas.ps1
```

That's it. The script will:

1. Print a banner and connect to the forest using the current credentials
2. Enumerate the forest, domains, sites, subnets, site links, DCs, and trusts
3. Detect Entra Connect, Exchange, and ADCS configuration from AD
4. Query DNS (via the DnsServer module, with a PDC fallback chain)
5. Query each DC for time configuration in parallel using a runspace pool (default: 20 concurrent, configurable via `$NtpMaxConcurrent` at the top of the script)
6. Read SYSVOL for the Default Domain Policy GPO and parse Kerberos policy from each domain
7. Generate a self-contained HTML report at `ADAtlas_<timestamp>.html` in the same directory as the script

Open the HTML file in any modern browser. No web server, no CDN, no JavaScript framework — everything is inlined.

---

## Configuration

A few variables at the top of the script can be adjusted:

| Variable | Default | Purpose |
|---|---|---|
| `$NtpMaxConcurrent` | `20` | Max parallel NTP queries (runspace pool size) |
| `$NtpTimeoutSec` | `8` | Per-DC timeout for the NTP collection step |

---

## Output

A single HTML file. No external dependencies. Approximately 150-200 KB for a small forest, scaling with the number of DCs and templates. Can be opened locally, archived, attached to a ticket, or hosted as a static file.

---

## Sister project

**[ADCanvas](https://github.com/SanthoshSivarajan/ADCanvas)** — health-focused AD assessment tool. ADCanvas tells you *how healthy* your AD is. ADAtlas tells you *what's in there.*

---

## Contributing

Issues and pull requests welcome. Please test changes against a real (lab) forest before submitting — the script has to handle a lot of edge cases that only show up in actual environments.

---

## Author

**Santhosh Sivarajan** — Microsoft MVP

- LinkedIn: [https://www.linkedin.com/in/sivarajan/](https://www.linkedin.com/in/sivarajan/)
- GitHub: [https://github.com/SanthoshSivarajan/ADAtlas](https://github.com/SanthoshSivarajan/ADAtlas)

---

## License

[MIT](LICENSE) — free to use, modify, and distribute.
