<#
.SYNOPSIS
    Converts synchronized Entra ID users to cloud-only users by clearing the ImmutableId.

.DESCRIPTION
    This script clears the ImmutableId (OnPremisesImmutableId) from one or more Entra ID user accounts,
    effectively breaking the synchronization link with on-premises Active Directory. This
    converts the users from synced accounts to cloud-only accounts.

    The script uses a direct Graph API PATCH request to set the ImmutableId to null, which
    is the only method that works reliably for manually-set ImmutableId values.

    WARNING: This is a permanent operation. Once the sync chain is broken, you cannot easily
    re-establish it without potential data loss or conflicts.

.PARAMETER UserPrincipalName
    One or more UPNs of users to convert to cloud-only (e.g., user@domain.com)

.PARAMETER ObjectId
    Alternative to UPN - The ObjectId/GUID of the user in Entra ID

.PARAMETER Force
    Skips the confirmation prompt for each user

.PARAMETER WhatIf
    Shows what would happen if the script runs without actually making changes

.EXAMPLE
    .\Convert-ToCloudOnlyUser.ps1 -UserPrincipalName "john.doe@contoso.com"

.EXAMPLE
    .\Convert-ToCloudOnlyUser.ps1 -UserPrincipalName "user1@contoso.com","user2@contoso.com","user3@contoso.com" -Force

.EXAMPLE
    .\Convert-ToCloudOnlyUser.ps1 -ObjectId "12345678-1234-1234-1234-123456789012"

.EXAMPLE
    .\Convert-ToCloudOnlyUser.ps1 -UserPrincipalName "john.doe@contoso.com" -WhatIf

.NOTES
    Prerequisites:
    - Microsoft Graph PowerShell SDK must be installed
    - Appropriate permissions in Entra ID (User Administrator or Global Administrator)

    Important:
    - This script clears the ImmutableId to break the sync anchor
    - Move on-premises AD account out of sync scope BEFORE running this script
    - Uses direct Graph API PATCH to set ImmutableId to null

    Author:  Henrik Skovgaard
    Contact: henrik@cloudonly.dk
    Version: 8.0
    Date:    2026-04-15

    Version History:
    8.0 - 2026-04-15 - Added support for multiple UPNs (array parameter)
                     - Single authentication for batch processing
                     - Added -Force parameter to skip confirmation
    7.0 - 2025-01-05 - Rewritten to use direct Graph API PATCH for clearing ImmutableId
                     - Removed dummy value approach that caused issues
                     - Simplified process - no longer requires delete/restore
    6.0 - Previous version with dummy value approach
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, ParameterSetName = "ByUPN")]
    [ValidateNotNullOrEmpty()]
    [string[]]$UserPrincipalName,

    [Parameter(Mandatory = $true, ParameterSetName = "ByObjectId")]
    [ValidateNotNullOrEmpty()]
    [string]$ObjectId,

    [Parameter()]
    [switch]$Force
)

# Function to check and install Microsoft Graph module if needed
function Test-MgGraphModule {
    Write-Host "Checking for Microsoft Graph PowerShell SDK..." -ForegroundColor Yellow

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Write-Host "Microsoft Graph PowerShell SDK is not installed." -ForegroundColor Yellow
        Write-Host "Installing Microsoft Graph module..." -ForegroundColor Cyan

        try {
            # Check if running as administrator for AllUsers scope
            $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

            if ($isAdmin) {
                Write-Host "Installing for all users..." -ForegroundColor Cyan
                Install-Module Microsoft.Graph -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
            }
            else {
                Write-Host "Installing for current user..." -ForegroundColor Cyan
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            }

            Write-Host "Microsoft Graph module installed successfully." -ForegroundColor Green

            # Import the Users module
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Write-Host "Microsoft Graph Users module imported." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install Microsoft Graph module: $_"
            Write-Host "Please install manually using: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
            return $false
        }
    }
    else {
        Write-Host "Microsoft Graph module is already installed." -ForegroundColor Green
        # Import if not already imported
        if (-not (Get-Module -Name Microsoft.Graph.Users)) {
            Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
        }
    }
    return $true
}

# Function to establish Graph connection with interactive login
function Connect-ToMgGraph {
    try {
        # Disconnect any existing session to force new login
        $context = Get-MgContext
        if ($null -ne $context) {
            Write-Host "Disconnecting existing Microsoft Graph session..." -ForegroundColor Yellow
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }

        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Write-Host "Please sign in with an account that has User Administrator or Global Administrator permissions." -ForegroundColor Yellow
        Write-Host ""

        try {
            # Force interactive login with required scopes
            Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
            Write-Host ""
            Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
            $context = Get-MgContext
        }
        catch {
            Write-Error "Failed to connect to Microsoft Graph: $_"
            return $false
        }

        # Check if we have the required scopes
        $requiredScopes = @("User.ReadWrite.All", "Directory.ReadWrite.All")
        $hasRequiredScope = $false
        foreach ($scope in $context.Scopes) {
            if ($scope -in $requiredScopes) {
                $hasRequiredScope = $true
                break
            }
        }

        if (-not $hasRequiredScope) {
            Write-Warning "Current connection may not have sufficient permissions."
            Write-Warning "Required scopes: User.ReadWrite.All or Directory.ReadWrite.All"
            Write-Host "Current scopes: $($context.Scopes -join ', ')" -ForegroundColor Yellow
        }

        return $true
    }
    catch {
        Write-Error "Error checking Microsoft Graph connection: $_"
        return $false
    }
}

# Function to process a single user
function Convert-SingleUser {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$UserId,
        [bool]$IsUPN,
        [bool]$ForceMode,
        [System.Management.Automation.PSCmdlet]$Cmdlet
    )

    try {
        # Get user information
        Write-Host ""
        Write-Host "----------------------------------------" -ForegroundColor Cyan
        Write-Host "Retrieving user information for: $UserId" -ForegroundColor Yellow
        Write-Host "----------------------------------------" -ForegroundColor Cyan

        $user = Get-MgUser -UserId $UserId -Property "Id,UserPrincipalName,DisplayName,OnPremisesSyncEnabled,OnPremisesImmutableId,Mail,AccountEnabled" -ErrorAction Stop

        if ($null -eq $user) {
            Write-Host "  User not found: $UserId" -ForegroundColor Red
            return
        }

        # Display user information
        Write-Host ""
        Write-Host "User Details:" -ForegroundColor Green
        Write-Host "  Display Name         : $($user.DisplayName)"
        Write-Host "  User Principal Name  : $($user.UserPrincipalName)"
        Write-Host "  Object ID            : $($user.Id)"
        Write-Host "  Mail                 : $($user.Mail)"
        Write-Host "  Account Enabled      : $($user.AccountEnabled)"
        Write-Host "  Sync Enabled         : $($user.OnPremisesSyncEnabled)"
        Write-Host "  ImmutableId          : $($user.OnPremisesImmutableId)"
        Write-Host ""

        # Check if user is already cloud-only
        if ($user.OnPremisesSyncEnabled -eq $false -and [string]::IsNullOrEmpty($user.OnPremisesImmutableId)) {
            Write-Host "  User is already cloud-only. Skipping." -ForegroundColor Green
            return
        }

        # Check if already cloud-only but has a lingering ImmutableId value
        if ($user.OnPremisesSyncEnabled -eq $false -and -not [string]::IsNullOrEmpty($user.OnPremisesImmutableId)) {
            Write-Warning "User has OnPremisesSyncEnabled = False but still has an ImmutableId value."
            Write-Host "  ImmutableId: $($user.OnPremisesImmutableId)" -ForegroundColor Yellow
            Write-Host "  Will attempt to clear the ImmutableId to complete the conversion." -ForegroundColor Yellow
        }

        if ([string]::IsNullOrEmpty($user.OnPremisesImmutableId)) {
            Write-Warning "User does not have an ImmutableId set. May already be cloud-only."
            if (-not $ForceMode) {
                $continue = Read-Host "Continue anyway? (Y/N)"
                if ($continue -ne "Y" -and $continue -ne "y") {
                    Write-Host "  Skipped by user." -ForegroundColor Yellow
                    return
                }
            }
        }

        # Confirmation prompt
        if (-not $WhatIfPreference -and -not $ForceMode) {
            Write-Host "  WARNING: This will break the sync chain for $($user.UserPrincipalName)." -ForegroundColor Red
            $confirmation = Read-Host "  Type 'YES' to confirm"
            if ($confirmation -ne "YES") {
                Write-Host "  Skipped by user." -ForegroundColor Yellow
                return
            }
        }

        # Clear the ImmutableId using direct Graph API PATCH
        Write-Host "  Clearing ImmutableId using direct Graph API..." -ForegroundColor Yellow

        if ($Cmdlet.ShouldProcess($user.UserPrincipalName, "Clear ImmutableId and convert to cloud-only user")) {

            try {
                $body = @{
                    "onPremisesImmutableId" = $null
                } | ConvertTo-Json

                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($user.Id)" -Body $body -ContentType "application/json" -ErrorAction Stop

                Write-Host "  Direct PATCH request succeeded!" -ForegroundColor Green
            }
            catch {
                # If direct PATCH fails, try with empty string
                Write-Host "  Direct PATCH failed, trying with empty string..." -ForegroundColor Yellow

                try {
                    $params = @{
                        OnPremisesImmutableId = ""
                    }
                    Update-MgUser -UserId $user.Id -BodyParameter $params -ErrorAction Stop
                    Write-Host "  Successfully cleared with empty string method." -ForegroundColor Green
                }
                catch {
                    Write-Host "  FAILED to clear ImmutableId: $($_.Exception.Message)" -ForegroundColor Red
                    return
                }
            }

            # Verify the change
            Start-Sleep -Seconds 3
            $updatedUser = Get-MgUser -UserId $user.Id -Property "OnPremisesImmutableId,OnPremisesSyncEnabled" -ErrorAction Stop

            if ([string]::IsNullOrEmpty($updatedUser.OnPremisesImmutableId)) {
                Write-Host "  SUCCESS - $($user.UserPrincipalName) is now cloud-only." -ForegroundColor Green
            }
            else {
                Write-Host "  WARNING - ImmutableId still set: $($updatedUser.OnPremisesImmutableId)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "  Skipped (WhatIf mode)." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  ERROR processing ${UserId}: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Main script execution
try {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Convert Synced User to Cloud-Only User" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # Check prerequisites
    if (-not (Test-MgGraphModule)) {
        exit 1
    }

    if (-not (Connect-ToMgGraph)) {
        exit 1
    }

    # Build list of users to process
    if ($PSCmdlet.ParameterSetName -eq "ByUPN") {
        $usersToProcess = $UserPrincipalName
        $isUPN = $true
    }
    else {
        $usersToProcess = @($ObjectId)
        $isUPN = $false
    }

    Write-Host ""
    Write-Host "Processing $($usersToProcess.Count) user(s)..." -ForegroundColor Cyan

    foreach ($userId in $usersToProcess) {
        Convert-SingleUser -UserId $userId -IsUPN $isUPN -ForceMode $Force.IsPresent -Cmdlet $PSCmdlet
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Batch processing complete." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Error "An error occurred: $_"
    Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    # Disconnect from Graph
    if (Get-MgContext) {
        Write-Host ""
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
    }
}
