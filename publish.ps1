if (Test-Path .\nupkgs) { Remove-Item .\nupkgs -Recurse -Force }
.\nuget.exe pack .\package\poi-ikvm.nuspec -OutputDirectory .\nupkgs
.\nuget.exe push .\nupkgs\*.nupkg -Source https://www.nuget.org/api/v2/package
