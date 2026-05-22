param(
    [switch]$NoBuild,
    [string]$Source = "https://api.nuget.org/v3/index.json",
    [string]$ApiKey
)

# Run the build script first to compile and package both legacy and modern targets, unless skipped
if (-not $NoBuild) {
    Write-Host "Running build.ps1 before publishing..."
    & .\build.ps1
}

# Find all generated NuGet packages
$packages = Get-ChildItem .\nupkgs\*.nupkg
if ($null -eq $packages -or $packages.Count -eq 0) {
    Write-Warning "No NuGet packages found in .\nupkgs\ to publish. Please run build.ps1 first."
    return
}

# Prompt securely for API Key if not passed as a parameter
if (-not $ApiKey) {
    # Read-Host -AsSecureString masks input on the terminal and prevents history capture
    $secureKey = Read-Host -Prompt "Enter NuGet API Key" -AsSecureString
    if ($null -ne $secureKey) {
        $ApiKey = [System.Net.NetworkCredential]::new("", $secureKey).Password
    }
}

if (-not $ApiKey) {
    Write-Warning "No API Key provided. Pushing without an API Key might fail if the source requires authentication."
}

# Publish each package using the modern cross-platform dotnet CLI
foreach ($pkg in $packages) {
    Write-Host "Publishing $($pkg.Name) to $Source..."
    
    # Avoid Invoke-Expression to securely pass arguments directly without shell pollution
    if ($ApiKey) {
        dotnet nuget push $pkg.FullName --source $Source --api-key $ApiKey
    } else {
        dotnet nuget push $pkg.FullName --source $Source
    }
}
