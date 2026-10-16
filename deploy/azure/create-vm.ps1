<#
.SYNOPSIS
  Creates the small Linux VM that holds the TraumaSoft-whitelisted static IP and runs the
  TraumaSoft -> Samsara work order sync.  Run from your laptop in PowerShell after `az login`.

.DESCRIPTION
  What it does, in order:
    1. Checks the existing static public IP (must be Static, Standard SKU, same region as the VM).
    2. Creates the VM with NO public IP, inside the VNet/subnet you choose.
    3. Creates a network security group that allows SSH (22) ONLY from your current public IP.
    4. Dissociates the static IP from the old (deallocated) VM's NIC and associates it with the new NIC.
       The old VM is never started.
    5. Tags everything so it is obvious what it is for and who owns it.

  Every step prints the command before running it.  Nothing is deleted.

.EXAMPLE
  .\create-vm.ps1 -ResourceGroup rg-ops -PublicIpName ts-static-ip -VNet vnet-ops -Subnet default `
                  -VmName ts-wo-sync -Owner "caleb@summitopsdata.com"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$ResourceGroup,
    [Parameter(Mandatory)] [string]$PublicIpName,     # the already-whitelisted static IP resource
    [Parameter(Mandatory)] [string]$VNet,
    [Parameter(Mandatory)] [string]$Subnet,
    [string]$VmName = "ts-wo-sync",
    [string]$Size = "Standard_B1s",
    [string]$Image = "Ubuntu2404",
    [string]$AdminUser = "syncadmin",
    [string]$Owner = "",
    [string]$Purpose = "TraumaSoft->Samsara work order sync (holds TraumaSoft-whitelisted static IP)",
    [switch]$SkipIpMove                                # create the VM only; move the IP later by hand
)
$ErrorActionPreference = "Stop"
function Run([string]$cmd) { Write-Host ">> $cmd" -ForegroundColor Cyan; Invoke-Expression $cmd; if ($LASTEXITCODE) { throw "command failed: $cmd" } }

Write-Host "`n== 1. Static IP check ==" -ForegroundColor Yellow
$ip = az network public-ip show -g $ResourceGroup -n $PublicIpName -o json | ConvertFrom-Json
Write-Host ("   {0}  alloc={1}  sku={2}  region={3}" -f $ip.ipAddress, $ip.publicIPAllocationMethod, $ip.sku.name, $ip.location)
if ($ip.publicIPAllocationMethod -ne "Static") { throw "Public IP is not Static; set it to Static first or you will lose the address." }
if ($ip.sku.name -ne "Standard") { Write-Warning "Public IP SKU is $($ip.sku.name). Basic SKU was retired 2025-09-30; upgrade to Standard (while dissociated) before relying on it." }
$oldIpConfigId = $ip.ipConfiguration.id
if ($oldIpConfigId) { Write-Host "   currently attached to: $oldIpConfigId" }

Write-Host "`n== 2. Create VM (no public IP yet) ==" -ForegroundColor Yellow
$tags = "purpose=`"$Purpose`" owner=`"$Owner`" app=traumasoft-samsara-sync"
Run "az vm create -g $ResourceGroup -n $VmName --image $Image --size $Size --location $($ip.location) --vnet-name $VNet --subnet $Subnet --public-ip-address `"`" --nsg `"`" --admin-username $AdminUser --generate-ssh-keys --tags $tags -o table"
$nicId = az vm nic list -g $ResourceGroup --vm-name $VmName --query "[0].id" -o tsv
$nicName = $nicId.Split("/")[-1]
Write-Host "   NIC: $nicName"

Write-Host "`n== 3. NSG: SSH only from your current public IP ==" -ForegroundColor Yellow
$myIp = (Invoke-RestMethod -Uri "https://api.ipify.org").Trim()
$nsg = "$VmName-nsg"
Run "az network nsg create -g $ResourceGroup -n $nsg --location $($ip.location) --tags $tags -o none"
Run "az network nsg rule create -g $ResourceGroup --nsg-name $nsg -n allow-ssh-from-admin --priority 100 --direction Inbound --access Allow --protocol Tcp --source-address-prefixes $myIp/32 --destination-port-ranges 22 -o none"
Run "az network nic update -g $ResourceGroup -n $nicName --network-security-group $nsg -o none"

if ($SkipIpMove) { Write-Host "`n-SkipIpMove set: VM created without the static IP. Move it later with:`n  az network nic ip-config update -g $ResourceGroup --nic-name <old-nic> -n <ipconfig> --public-ip-address null`n  az network nic ip-config update -g $ResourceGroup --nic-name $nicName -n ipconfig1 --public-ip-address $PublicIpName"; return }

Write-Host "`n== 4. Move the static IP to the new VM (old VM stays deallocated) ==" -ForegroundColor Yellow
if ($oldIpConfigId) {
    $parts = $oldIpConfigId.Split("/")
    $oldNic = $parts[[array]::IndexOf($parts, "networkInterfaces") + 1]
    $oldCfg = $parts[-1]
    Run "az network nic ip-config update -g $ResourceGroup --nic-name $oldNic -n $oldCfg --public-ip-address null -o none"
}
Run "az network nic ip-config update -g $ResourceGroup --nic-name $nicName -n ipconfig1 --public-ip-address $PublicIpName -o none"

Write-Host "`n== Done ==" -ForegroundColor Green
Write-Host "SSH:  ssh $AdminUser@$($ip.ipAddress)"
Write-Host "Next: copy deploy/azure/setup-vm.sh to the VM and run it (see RUNBOOK.md)."
