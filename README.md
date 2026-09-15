# PptCompare

PptCompare is a Windows desktop application for comparing two versions of a Microsoft PowerPoint presentation side by side. It highlights text, formatting, table, image, shape and slide-order changes without modifying either source presentation.

The application is intended for local workplace use. Presentation content is processed on the computer and is not uploaded by PptCompare.

> **Project status:** PptCompare 0.3.x is a security-hardened workplace-pilot release. Validate it against representative presentations and your organisation's deployment policies before a wider rollout.

## Features

- Independent left and right presentation selection.
- Slide matching that recognises added, removed, changed and reordered slides.
- Selectable text output with source font, colour and character styling at a configurable display size.
- Table-aware text output that preserves rows, columns and merged cells.
- Image, shape, layout and visual-change detection.
- Immediate Open XML-based preview.
- Optional high-fidelity preview rendering through desktop Microsoft PowerPoint.
- Local settings for presentation limits, output font size, preview behaviour and diagnostic logging.
- Privacy-conscious rotating diagnostic logs that exclude presentation names, paths and content.

## Requirements

### Running a published build

- Windows 11 on an x64 or Arm64 computer.
- A build published for the computer's architecture. Windows 11 on Arm can also run the x64 build through emulation.
- Desktop Microsoft PowerPoint is optional. It is used only for high-fidelity preview images; comparison and the built-in preview continue to work without it.

PptCompare supports `.pptx` and `.pptm` files. Presentation packages are opened read-only and macros are not executed.

### Building from source

- .NET 10 SDK.
- Windows with the .NET desktop development tools.
- Rider, Visual Studio or another editor capable of building .NET WPF projects.

## Build and test

Run the following commands from the repository root:

```powershell
dotnet restore PptCompare.sln --locked-mode
dotnet build PptCompare.sln --configuration Release --no-restore
dotnet test PptCompare.sln --configuration Release --no-build --no-restore
```

The repository treats compiler and analyser warnings as errors. The two committed `packages.lock.json` files should remain under source control so that package restoration is reproducible.

## Publish

Create the current self-contained x64 portable build with:

```powershell
dotnet publish PptCompare/PptCompare.csproj `
  --configuration Release `
  --profile Portable-win-x64
```

An Arm64 build can be produced independently with:

```powershell
dotnet publish PptCompare/PptCompare.csproj `
  --configuration Release `
  --runtime win-arm64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  --output PptCompare/publish/win-arm64
```

Generated `publish`, `dist`, `bin` and `obj` directories are deliberately excluded from source control. Signed executable or package files should be distributed as versioned release assets or through the organisation's managed deployment system.

## Application data

PptCompare stores per-user data under `%LocalAppData%\PptCompare`:

- `settings.json` contains validated application settings.
- `Logs` contains bounded, rotating diagnostics.

High-fidelity preview images are generated from isolated staged temporary copies and cleaned up automatically. PptCompare must never close a pre-existing PowerPoint session or alter an original presentation.

## Repository structure

| Path | Purpose |
| --- | --- |
| `PptCompare` | WPF application, view models, services, models and controls |
| `PptCompare.Tests` | MSTest unit and regression tests |
| `PptCompare/Properties/PublishProfiles` | Portable Windows publication settings |
| `qodana.yaml` | Static-analysis configuration |
| `AGENTS.md` | Detailed architecture, safety and maintenance guidance |

## Current limitations

- WPF requires Windows.
- High-fidelity rendering depends on a locally installed and activated copy of desktop PowerPoint.
- The built-in preview is intentionally less visually exact than PowerPoint's renderer.
- For decoder hardening, direct embedded-image display in the built-in preview is limited to validated PNG and JPEG data; other image formats remain comparable by content hash and are visible in the optional PowerPoint-rendered preview.
- SharePoint version history and Microsoft Graph authentication are planned but not yet implemented.

## Contributing

Before submitting a change:

1. Keep presentation parsing and PowerPoint automation out of WPF code-behind.
2. Preserve the PowerPoint ownership and shutdown-safety rules documented in `AGENTS.md`.
3. Add regression tests for comparison, parsing, cancellation or automation-lifecycle changes.
4. Complete a locked restore, Release build and full test run.

Security issues should be reported privately as described in [SECURITY.md](SECURITY.md).

## Licence

PptCompare is available under the [MIT Licence](LICENSE). Copyright © 2026 Ben Davies.
