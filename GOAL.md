# Project Goal: Target Modern .NET (.NET 9+) with poi-ikvm

This project compiles Apache POI using IKVM so that it can be used natively within .NET code with identical Java namespaces (`org.apache.poi.*`). 

Our goal is to enhance this codebase to support modern .NET (targeting .NET 9 and above) as a separate package while preserving the existing .NET Framework 4.0 package target.

---

## Architectural & Design Choices

1. **Package Separation**:
   - **`poi-ikvm`**: Targets .NET Framework 4.0 (using old IKVM 8.1.5717.0). Must be compiled on Windows. Uses **Apache POI 3.17.0.1**.
   - **`poi-ikvm-net`**: Targets modern .NET (including .NET 9, 10, 11, etc., utilizing modern cross-platform IKVM 8.15.0). Uses **Apache POI 5.2.5** (the latest version compatible with modern IKVM's Java 8 bytecode constraint). Can be compiled on both Windows and Linux.
   
2. **Namespace & Version Compatibility**:
   - The legacy target will remain on Apache POI 3.17, while the modern target will use Apache POI 5.2.5.
   - The namespaces remain identical (`org.apache.poi.*`).
   - The consuming code will handle any API differences between POI 3.x and POI 5.x using `#if` directives.

3. **Maven-Based Resolution**:
   - For the modern `poi-ikvm-net` package, we use `IKVM.Maven.Sdk` and `<MavenReference>` to automatically download, resolve, and transpile Apache POI 5.2.5 and all its transitive dependencies at build time, eliminating manual dependency management in scripts.

4. **Cross-Platform PowerShell 7 Script**:
   - `build.ps1` will be refactored to be compatible with PowerShell 7 on Windows and Linux.
   - On Windows, it will support building both packages.
   - On Linux, it will skip the Windows-only .NET Framework build step and successfully build the `poi-ikvm-net` package using modern MSBuild/dotnet CLI.

---

## Persistent LLM Task & Context Workflow

To ensure seamless progress and state preservation across conversation boundaries, we use the following step-by-step git-based checkpoint workflow:

1. **Context Document & Task List**: Write the persistent context (`GOAL.md`) and task list (`TASKS.md`) and commit them to the repository.
2. **Execute Steps Sequentially**: Work on exactly one step at a time.
3. **Commit checkpoints**: At the end of every step, update `TASKS.md` (mark the step as completed) and commit all changes to git.
4. **Context Reset & Reload**: End the current LLM turn/conversation. The next turn starts with a clean context, loading `GOAL.md` and `TASKS.md` to restore full state.
