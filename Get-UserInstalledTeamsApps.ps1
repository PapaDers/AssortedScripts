# Define the Teams App ID to search for. This can be found in the Teams Admin center.
$targetAppId = ""

# Connect to Microsoft Graph with the required permissions
Connect-MgGraph -Scopes "User.Read.All, AppCatalog.Read.All"

# Fetch all users
Write-Host "Fetching all users to filter for enabled, non-external users, and excluding specific UPNs..." -ForegroundColor Cyan
$allUsers = Get-MgUser -All

# Filter for enabled, non-external users with phone numbers, and exclude specific UserPrincipalNames
$usersWithPhone = $allUsers | Where-Object {
    $_.AccountEnabled -eq $true -and
    $_.UserType -ne "Guest" -and
    ($_.BusinessPhones.Count -gt 0 -or $_.MobilePhone -ne $null) -and
    ($_.UserPrincipalName -notlike "*#EXT#*")
}

Write-Host "Total users to process: $($usersWithPhone.Count)" -ForegroundColor Green

# Initialize collections to store results
$usersWithApp = @()
$usersWithoutApp = @()

# Loop through each user and check for the app installation
$counter = 0
foreach ($user in $usersWithPhone) {
    $counter++
    Write-Host "Processing user $counter/$($usersWithPhone.Count): $($user.UserPrincipalName)" -ForegroundColor Yellow

    try {
        # Get installed apps for the user
        $installedApps = Get-MgUserTeamworkInstalledApp -UserId $user.Id -ExpandProperty "teamsApp"

        # Check if the target app is installed
        $appInstalled = $installedApps | Where-Object { $_.TeamsApp.Id -eq $targetAppId }

        if ($appInstalled) {
            Write-Host "  App found for user $($user.UserPrincipalName)" -ForegroundColor Green
            # Add the user to the 'with app' list
            $usersWithApp += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
                AppId             = $targetAppId
                AppInstalled      = $true
            }
        }
        else {
            Write-Host "  App not found for user $($user.UserPrincipalName)" -ForegroundColor Red
            # Add the user to the 'without app' list
            $usersWithoutApp += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
                AppId             = $targetAppId
                AppInstalled      = $false
            }
        }
    }
    catch {
        Write-Host "  Error checking apps for user $($user.UserPrincipalName): $_" -ForegroundColor Magenta
        # Treat as 'without app' if there's an error (optional behavior)
        $usersWithoutApp += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName       = $user.DisplayName
            AppId             = $targetAppId
            AppInstalled      = $false
        }
    }
}

# Output the results
Write-Host "Processing complete. Summary:" -ForegroundColor Cyan
Write-Host "  Users with app installed: $($usersWithApp.Count)" -ForegroundColor Green
Write-Host "  Users without app installed: $($usersWithoutApp.Count)" -ForegroundColor Red

# Export results to CSV files
$usersWithApp | Export-Csv -Path "UsersWithTargetApp.csv" -NoTypeInformation
$usersWithoutApp | Export-Csv -Path "UsersWithoutTargetApp.csv" -NoTypeInformation

Write-Host "Results exported to 'UsersWithTargetApp.csv' and 'UsersWithoutTargetApp.csv'." -ForegroundColor Cyan
