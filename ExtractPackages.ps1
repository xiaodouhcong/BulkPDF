# Extract NuGet packages script
$packages = Get-ChildItem -Path "packages" -Filter "*.nupkg"

foreach ($package in $packages) {
    $packageName = $package.Name.Replace(".nupkg", "")
    $extractPath = Join-Path "packages" $packageName
    
    Write-Host "Extracting package: $($package.Name) to $extractPath"
    
    # Create directory
    if (!(Test-Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
    }
    
    # Extract package
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($package.FullName, $extractPath)
}

Write-Host "All packages extracted successfully"