<#
.SYNOPSIS
    Blocks Entra ID Workplace Join (device registration) on Windows endpoints.

.DESCRIPTION
    Sets the BlockAADWorkplaceJoin policy registry value, which prevents Windows
    from showing the "Sign in to all apps and websites on this device?" prompt
    that auto-registers the device in whatever Entra tenant the user is signing
    in to.

    Users can still sign in to individual apps (Outlook, Teams, OneDrive, etc.)
    with their work account or any other account without registering the device.

    Idempotent: safe to run repeatedly. Exits 0 if the key is already correctly
    set or after applying it. Exits 1 only if the registry write fails.

    Designed to run as an Intune Remediation script (System context, Windows).

.EXAMPLE
    .\Block-WorkplaceJoin.ps1

    Sets HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin\BlockAADWorkplaceJoin
    to 1 (DWord) and verifies the result.

.NOTES
    Author:  Henrik Skovgaard
    Contact: henrik@cloudonly.dk

    Registry path: HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin
    Value name:    BlockAADWorkplaceJoin
    Value:         1 (DWORD)
    Run context:   SYSTEM
#>

$RegistryLocation = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin"
$keyName = "BlockAADWorkplaceJoin"

# Check if the key is already in place; if so, exit
$existingKey = Get-ItemProperty -Path $RegistryLocation -Name $keyName -ErrorAction SilentlyContinue
if ($existingKey -and $existingKey.$keyName -eq 1) {
    Write-Output "Registry key is already in place."
    Exit 0
}

# Create the registry path if it is missing
if (!(Test-Path -Path $RegistryLocation)) {
    Write-Output "Registry location is missing. Creating it now."
    New-Item -Path $RegistryLocation -Force | Out-Null
}

# Set the key value
New-ItemProperty -Path $RegistryLocation -Name $keyName -PropertyType DWord -Value 1 -Force | Out-Null

# Verify the key has been created successfully
$checkKey = Get-ItemProperty -Path $RegistryLocation -Name $keyName -ErrorAction SilentlyContinue
if ($checkKey -and $checkKey.$keyName -eq 1) {
    Write-Output "Registry key has been successfully set."
    Exit 0
} else {
    Write-Error "Failed to create registry key!"
    Exit 1
}
