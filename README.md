# VMware-ProductLocker-VMTools

PowerShell script for auditing and configuring the VMware Tools ProductLocker location across ESXi hosts, with a full VMware Tools version and status audit per VM.

Built for vSphere 7.x / 8.x / 9.x environments, including VCF deployments with vSAN where hosts have no local VMFS datastore registered in vCenter.

> Blog post: [hollebollevsan.nl](https://www.hollebollevsan.nl)

---

## What it does

### ProductLocker
- **Audit** — reads the current `ProductLockerLocation` on every connected ESXi host in a cluster via `ExtensionData.QueryProductLockerLocation()`
- **Set** — interactively pick a datastore and sub-path per cluster; sets all hosts in that cluster to the same path via `ExtensionData.UpdateProductLockerLocation()`
- **Reset** — resets every host back to its local OS volume default. Resolves the VMFSOS volume UUID per host via `esxcli storage filesystem list` and builds the full path as `/vmfs/volumes/<VMFSOS-UUID>/locker/packages/<subfolder>`. Works correctly in VCF/vSAN environments where the local volume is not registered as a vCenter datastore.
- **Path verification** — after setting a path, verifies the locker folder exists on the target datastore via the `vmstore:\` provider. Warns immediately if the folder is missing, since the host will silently fail VMTools installs/upgrades against a missing path.

### VMware Tools
- **Audit** — scans all powered-on VMs in a cluster and reports:
  - **Run Status** — `toolsOk`, `toolsNotRunning`, `toolsNotInstalled`, etc.
  - **Version Status** — current, needs upgrade, unmanaged (OSP), too old/new, etc. via `ToolsVersionStatus2`
  - **Management Mode** — vCenter/ESX managed vs. Guest Managed (open-vm-tools via OS package manager)
  - **Tools Version** — numeric version string per VM
- Colour-coded table output with a three-section summary at the end

### Distribution
- **Upload** — copy a local VMTools ISO or ZIP to a target datastore folder
- **Copy** — replicate a VMTools folder from one datastore to one or more others

### Credentials
- Prompts on first run, saves encrypted `.cred` file next to the script named `<vCenterServer>.cred`
- Username and password entered via `Read-Host` in the terminal (no GUI dialog). Saved via `Export-Clixml` / DPAPI, bound to the current Windows user — cannot be decrypted on a different machine or by a different account
- `-ResetCredentials` forces a new prompt and overwrites the saved file

---

## Requirements

| Environment | Module |
|---|---|
| vSphere 7.x / 8.x | `VMware.PowerCLI` (VMware.VimAutomation.Core) |
| vSphere 9.x / VCF | `VCF.PowerCLI` |

The script auto-detects which module is available and imports it. If neither is found it will print install instructions and exit.

```powershell
# vSphere 9 / VCF
Install-Module -Name VCF.PowerCLI -Scope CurrentUser

# vSphere 7 / 8
Install-Module -Name VMware.PowerCLI -Scope CurrentUser
```

---

## Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `-vCenterServer` | string | Yes | FQDN or IP of the vCenter Server |
| `-Cluster` | string | No | Target cluster name. All clusters if omitted. |
| `-SetProductLocker` | switch | No | Interactively set ProductLocker per cluster via datastore picker |
| `-ResetProductLocker` | switch | No | Reset all hosts to their local OS volume default path |
| `-UploadVMTools` | switch | No | Upload a local VMTools ISO or ZIP to a datastore |
| `-UploadSourcePath` | string | No | Local file path for upload (required with `-UploadVMTools`) |
| `-CopyVMTools` | switch | No | Copy a VMTools folder between datastores |
| `-CopySourceDatastore` | string | No | Source datastore name (required with `-CopyVMTools`) |
| `-CopySourcePath` | string | No | Source folder path on datastore (required with `-CopyVMTools`) |
| `-ExportReport` | switch | No | Export results to CSV in the current directory |
| `-ResetCredentials` | switch | No | Force new credential prompt, overwrite saved `.cred` file |

---

## Usage

```powershell
# First run - prompts for credentials and saves them, then shows audit menu
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local

# Force new credential prompt
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local -ResetCredentials

# Interactively set ProductLocker per cluster
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local -SetProductLocker

# Reset ProductLocker to local OS volume default on all hosts
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local -ResetProductLocker

# Upload local VMTools ISO to a datastore
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local `
    -UploadVMTools -UploadSourcePath "C:\vmtools\VMware-tools-12.4.0.iso"

# Copy VMTools folder from one datastore to others
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local `
    -CopyVMTools -CopySourceDatastore "DS-Site-A" -CopySourcePath "/locker/packages/vmtoolsRepo"

# Full run: set ProductLocker and export results to CSV
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local -SetProductLocker -ExportReport

# Target a specific cluster only
.\VMware-ProductLocker-VMTools.ps1 -vCenterServer vc01.lab.local -Cluster "Cluster-A" -SetProductLocker
```

---

## Audit menu

On every run the script shows an interactive menu before executing any audit:

```
  Select audit(s) to perform:

  [1]  ProductLocker Audit  - check ProductLocker location on all ESXi hosts
  [2]  VMware Tools Audit   - check tools version, status and management mode
  [3]  Both
  [0]  Skip audits (configuration actions only)
```

Selecting `[0]` is useful when you only want to run a configuration action (e.g. `-SetProductLocker` or `-UploadVMTools`) without running any audit.

---

## CSV export

When `-ExportReport` is specified, two CSV files are written to the current directory:

**`ProductLocker_<timestamp>.csv`**

| Field | Description |
|---|---|
| Hostname | ESXi host FQDN |
| ConnectionState | Host connection state |
| CurrentValue | ProductLocker path at time of audit/set |
| ExpectedValue | Target path (blank for audit-only runs) |
| Status | `AUDIT`, `OK`, `MISMATCH`, `SET`, `SET_PATH_MISSING`, `ERROR` |
| PathVerified | `True` / `False` / `N/A` / `N/A (local vol)` |
| Cluster | Cluster name |
| Timestamp | Run timestamp |

**`VMTools_<timestamp>.csv`**

| Field | Description |
|---|---|
| VMName | Virtual machine name |
| PowerState | VM power state |
| Host | ESXi host the VM runs on |
| Cluster | Cluster name |
| RunStatus | Raw `ToolsStatus` enum value |
| RunStatusLabel | Human-readable run status |
| VersionStatus | Raw `ToolsVersionStatus2` enum value |
| VersionStatusLabel | Human-readable version status |
| ManagedBy | `vCenter/ESX Mgd` or `Guest Managed` |
| ToolsVersion | Numeric VMware Tools version |
| GuestOS | Guest OS full name |
| Timestamp | Run timestamp |

---

## ProductLocker reset — how it works

When `-ResetProductLocker` is used, the script needs to find the correct full path for each host's local locker volume. In VCF/vSAN environments the local OS volume is not registered as a vCenter datastore, so a standard `Get-Datastore` lookup will not find it.

The script resolves this per host using ESXCLI:

```
esxcli storage filesystem list
```

This returns all mounted filesystems on the host, including the local OS volume with `Type = VMFSOS`. The UUID from that entry is used to build the full path:

```
/vmfs/volumes/<VMFSOS-UUID>/locker/packages/vmtoolsRepo   # ESXi 9+
/vmfs/volumes/<VMFSOS-UUID>/locker/packages/vmtoolsd      # ESXi 7/8
```

This is the format that `UpdateProductLockerLocation()` accepts. Relative paths such as `/locker/packages/vmtoolsRepo` are rejected by the API.

---

## Notes

- `UpdateProductLockerLocation()` takes effect immediately — no reboot or maintenance mode required
- The host does **not** validate the path at set time. A path pointing to a missing folder will silently fail when a VMTools install or upgrade is triggered. The script warns immediately if the folder cannot be found after setting
- Credential files are per-vCenter and per-Windows-user. If you run the script as a different user, a new prompt will appear
- `-WhatIf` is supported — the script uses `[CmdletBinding(SupportsShouldProcess)]` so all write operations respect `-WhatIf`

---

## Tested on

- vSphere 9.0 / VCF 9.0 with vSAN (hosts without local VMFS datastore in vCenter inventory)
- VMware PowerCLI 13.x / VCF PowerCLI 9.0

---

## Author

Paul van Dieen — [hollebollevsan.nl](https://www.hollebollevsan.nl)
