# Persistent Task List: Targeting modern .NET (poi-ikvm-net)

This file tracks our progress. We perform the steps one by one, updating the status of each step and committing to git at each checkpoint.

- [x] **Step 1**: Create/Update `GOAL.md` and `TASKS.md` in the project root and commit them to git.
- [x] **Step 2**: Create the modern SDK-style project file `package/poi-ikvm-net.csproj` targeting `net9.0` with `MavenReference` for Apache POI 5.2.5.
- [x] **Step 3**: Modernize `build.ps1` to support cross-platform PowerShell 7, update the legacy POI download URL to Apache Archive, and integrate building `poi-ikvm-net.csproj`.
- [x] **Step 4**: Run the build and package processes to generate the local `.nupkg` and compiled assemblies for modern .NET.
- [x] **Step 5**: Create a sample .NET 9 project that references the built `poi-ikvm-net` package/DLL and verify that it compiles and runs correctly.
- [x] **Step 6**: Finalize task list, perform final review, and complete the transition.
