# Runbook: TraumaSoft → Samsara work order sync VM

**Purpose.** One small Azure VM whose only job is to hold the static public IP that TraumaSoft has
whitelisted for Direct Database Access, and to run `TraumaSoft_Samsara_WorkOrder_Sync.py` on a schedule.
Nothing else runs on it. TraumaSoft's third-party API has no work-order endpoint (OpenAPI v1.0.0, 94
paths), so a database read from a whitelisted IP is the only way to get work orders.

| Item | Value |
|---|---|
| Owner | *(fill in)* |
| Azure resource group / VM | *(fill in)* |
| Static IP (whitelisted by TraumaSoft) | *(fill in)* |
| VM size / OS | Standard_B1s (or B2s) / Ubuntu 24.04 |
| Inbound access | SSH (22) from the owner's IP only (NSG `<vm>-nsg`); no other inbound ports |
| Outbound | TraumaSoft MySQL (3306), api.samsara.com (443), SMTP |
| Database user | read-only TraumaSoft DDAS user *(name)* |
| Code | this repo, branch/tag *(fill in)*, at `~/wo-sync` |
| Schedule | cron, minute 17 of every hour (`crontab -l`) |
| State | `~/wo-sync/Reports/State/samsara_workorder_map.db` (SQLite mapping table) |
| Logs | `~/wo-sync/Reports/logs/` |
| Monthly cost | VM + static IP + disk, roughly a low-double-digit USD figure *(confirm in Cost Management)* |

## Build (from scratch)
1. On your laptop: `az login`, then `deploy/azure/create-vm.ps1` with the resource group, static IP
   name, VNet and subnet. It creates the VM without a public IP, locks SSH to your IP, and moves the
   static IP from the old deallocated VM's NIC to the new NIC. The old VM is never started.
2. `ssh syncadmin@<static-ip>`, copy `deploy/azure/setup-vm.sh` over, run
   `bash setup-vm.sh <repo-url> <branch>`.
3. Fill in `~/wo-sync/.env` (TraumaSoft DB host/name/user/password, Samsara token, SMTP).
4. Verify in this order:
   ```bash
   cd ~/wo-sync
   .venv/bin/python -m unittest discover -s tests          # no network needed
   .venv/bin/python TraumaSoft_Samsara_WorkOrder_Sync.py --discover          # proves DB access + shows columns
   .venv/bin/python TraumaSoft_Samsara_WorkOrder_Sync.py --dry-run --no-email
   .venv/bin/python TraumaSoft_Samsara_WorkOrder_Sync.py --limit 1
   ```
5. Set `TEST_MODE=false` and the production recipients once a real run looks right.

## Operate
- **Logs:** `tail -n 100 ~/wo-sync/Reports/logs/cron.log` and the dated logs beside it.
- **Pause:** `crontab -e` and comment out the line. **Resume:** uncomment it.
- **Update code:** `cd ~/wo-sync && git pull --ff-only && .venv/bin/python -m unittest discover -s tests`.
- **Rebuild the mapping table:** `.venv/bin/python TraumaSoft_Samsara_WorkOrder_Sync.py --rebuild-ledger`
  (reads the `[TS-WO:<id>]` tags back from Samsara).
- **Rotate secrets:** edit `.env` (mode 600); no restart needed, each run reads it.

## Decommission
1. `crontab -r` on the VM (stops the job).
2. Dissociate the static IP from the VM's NIC **before** deleting the VM so the IP resource survives:
   `az network nic ip-config update -g <rg> --nic-name <nic> -n ipconfig1 --public-ip-address null`
3. `az vm delete -g <rg> -n <vm> --yes` and delete the NIC, disk and NSG.
4. Ask TraumaSoft to remove the IP from the whitelist if it will not be reused; then delete the IP.

## Why a VM and not something else
- TraumaSoft whitelists a single static IP; the job must originate from it.
- The previous VM cannot be restarted. A static Standard-SKU public IP can be re-associated to another
  NIC without starting the old VM, so the whitelist entry carries over unchanged.
- A NAT gateway + Azure Function would also work but costs more and adds moving parts; a VPN gateway
  cannot do it at all (Azure drops internet-bound traffic through a VPN gateway).
