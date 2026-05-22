# Add-Type for zip extraction (used on Windows for legacy POI/IKVM)
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Determine if running on Windows
$isWin = $IsWindows
if ($null -eq $isWin) {
    $isWin = $env:OS -like "*Windows*"
}

if ($isWin) {
    Write-Host "Running on Windows. Legacy .NET Framework and modern .NET builds are supported."
    
    if (-not (Test-Path .\nuget.exe))
    {
        Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile nuget.exe
    }

    $ikvm_version = '8.1.5717.0'

    if (-not(Test-Path ikvm\))
    {
        Write-Host "Downloading legacy IKVM $ikvm_version..."
        Invoke-WebRequest -Uri "http://www.frijters.net/ikvmbin-$ikvm_version.zip" -OutFile ikvm.zip -UserAgent "NativeHost"
        [System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\ikvm.zip", "$PWD")
        Remove-Item ikvm.zip
        Get-Item ikvm* | select -Index 0 | Rename-Item -Path {$_.FullName} -NewName ikvm
    }

    if (-not(Test-Path poi\))
    {
        # Download legacy Apache POI 3.17 from the stable Apache Archive
        $poi_url = "https://archive.apache.org/dist/poi/release/bin/poi-bin-3.17-20170915.zip"
        Write-Host "Downloading Apache POI 3.17..."
        Invoke-WebRequest -Uri $poi_url -OutFile poi.zip -UserAgent "NativeHost"
        Write-Host "Extracting Apache POI 3.17..."
        [System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\poi.zip", "$PWD")
        Remove-Item poi.zip
        Get-Item poi* | select -Index 0 | Rename-Item -Path {$_.FullName} -NewName poi
    }

    Write-Host "Starting legacy .NET Framework build..."
    $poi_main = Get-ChildItem -Path poi | where {$_.Name -match '^poi-[.\d]+\.jar$'}

    $version = [Version]$poi_main.BaseName.Substring("poi-".Length)
    $version = New-Object -TypeName Version -ArgumentList ($version.Major, [Math]::Max(0, $version.Minor), [Math]::Max(0, $version.Build), [Math]::Max(0, $version.Revision))

    # Safely query latest published version
    $latestPublished = $null
    try {
        $latestPublished = (.\nuget.exe list id:poi-ikvm) | where {$_ -imatch '^poi-ikvm\s'} | foreach {[Version]($_ -replace '^poi-ikvm\s(.*)$', '$1')}
    } catch {
        Write-Warning "Could not retrieve latest published version from nuget.org. Using local POI version."
    }

    if ($null -ne $latestPublished) {
        $latestPublished = New-Object -TypeName Version -ArgumentList ($latestPublished.Major, [Math]::Max(0, $latestPublished.Minor), [Math]::Max(0, $latestPublished.Build), [Math]::Max(0, $latestPublished.Revision))
        if ($latestPublished -ge $version)
        {
            $version = New-Object -TypeName Version -ArgumentList ($latestPublished.Major, $latestPublished.Minor, $latestPublished.Build, ($latestPublished.Revision + 1))
        }
    }

    $vr = '[0-9.\-]+'
    $libs = "activation$($vr)jar|commons-[a-z]+$($vr)jar|jaxb-[a-z]+$($vr)jar|junit$($vr)jar|log4j$($vr)jar|SparseBitSet$($vr)jar"
    $new_libs = Get-ChildItem -Path .\poi\lib | select { $_.Name -imatch $libs } | where {$_ -eq $false}

    if ($new_libs -ne $null)
    {
        throw "This version of poi has new dependencies. The code below needs to be modified to determine whether to include them in the package"
    }

    $jars = @($($poi_main;(Get-ChildItem -Path .\poi\lib | where {$_ -inotmatch "activation$($vr)jar|junit$($vr)jar|jaxb-[a-z]+$($vr)jar"})) | select -ExpandProperty FullName)

    New-Item -Path .\package\lib\net40 -ItemType Directory -Force

    &.\ikvm\bin\ikvmc.exe -keyfile:key.snk -out:package\lib\net40\poi.dll -target:library ("-version:"+$version.ToString()) ("-fileversion:"+$version.ToString()) $jars

    $versionElement = Select-Xml -Path .\package\poi-ikvm.nuspec -XPath 'package/metadata/version'
    $versionElement.Node.'#text' = $version.ToString()
    $versionElement.Node.OwnerDocument.Save("$PWD\package\poi-ikvm.nuspec")

    $versionElement = Select-Xml -Path .\package\poi-ikvm.nuspec -XPath 'package/metadata/dependencies/dependency[@id=''IKVM'']'
    $versionElement.Node.SetAttribute('version', $ikvm_version)
    $versionElement.Node.OwnerDocument.Save("$PWD\package\poi-ikvm.nuspec")
    
    # Run legacy packing
    if (-not (Test-Path .\nupkgs)) {
        New-Item -Path .\nupkgs -ItemType Directory -Force
    }
    .\nuget.exe pack .\package\poi-ikvm.nuspec -OutputDirectory .\nupkgs
    Write-Host "Legacy .NET Framework build and pack completed successfully."
} else {
    Write-Host "Running on Linux. Skipping legacy .NET Framework 4.0 build."
}

# --- Modern .NET Build Section ---
Write-Host "Starting modern .NET build (poi-ikvm-net)..."

# Ensure nupkgs folder exists
if (-not (Test-Path .\nupkgs)) {
    New-Item -Path .\nupkgs -ItemType Directory -Force
}

# Run dotnet build on our modern project (which compiles POI 5.2.5 using Maven SDK and outputs NuGet package on build)
dotnet build .\package\poi-ikvm-net.csproj -c Release

# Copy the generated nupkg to our shared nupkgs directory
if (Test-Path .\package\bin\Release\*.nupkg) {
    Copy-Item -Path .\package\bin\Release\*.nupkg -Destination .\nupkgs\ -Force
    Write-Host "Modern .NET build and pack completed successfully. Package is located in .\nupkgs\"
} else {
    Write-Error "Modern .NET NuGet package was not found in the output directory."
}
