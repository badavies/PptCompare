# PptCompare repository guide

## Product intent

PptCompare is a portable Windows desktop application for comparing two versions of a PowerPoint presentation. Its primary workplace use case is reviewing presentation changes without modifying either source file.

The application should:

- Start with no sample or placeholder presentations loaded.
- Let the user choose a left and right `.pptx` or `.pptm` file independently.
- Show slide-level text, tables, images, shapes, movement, additions and removals.
- Preserve text styling and table structure in the selectable text view, while applying one configurable output font size.
- Provide a built-in Open XML preview immediately and, when enabled and available, replace it with a high-fidelity preview rendered by Microsoft PowerPoint.
- Remain responsive while files are read, compared and rendered.
- Never modify a source presentation or lose work in an existing PowerPoint session.
- Remain ready for future Microsoft Graph/SharePoint version-history integration.

Microsoft Graph and SharePoint retrieval are not implemented yet. The current version selectors contain local-file descriptors. Keep the service boundaries and `VersionDescriptor` metadata suitable for adding remote versions later.

## Technology and deployment

- C# 14 / .NET 10, targeting `net10.0-windows`.
- WPF with an MVVM-style structure and constructor-injected services.
- `DocumentFormat.OpenXml` for local, non-destructive presentation inspection.
- Late-bound PowerPoint COM automation for optional high-fidelity PNG previews. The application does not depend on Office interop assemblies at compile time.
- MSTest for automated tests.
- Qodana plus built-in .NET analyzers. Warnings are treated as errors by the projects.
- Self-contained, single-file `win-x64` publishing for portable deployment.

The assembly version, author, product and copyright metadata in `PptCompare/PptCompare.csproj` are the source of truth. The software designer is Ben Davies and the copyright is © 2026 Ben Davies.

## Solution map

### Application project: `PptCompare`

- `App.xaml` / `App.xaml.cs`: application startup, manual dependency composition, settings load, diagnostics setup and top-level exception handling.
- `MainWindow.xaml`: WinMerge-style shell containing both selectors, slide navigator, synchronized comparison panes, preview-loading indicators and change summary.
- `MainWindow.xaml.cs`: closes disposable view models and synchronizes left/right scrolling. Keep code-behind limited to view-specific behaviour.
- `ViewModels/MainWindowViewModel.cs`: owns user commands and coordinates loading, comparison, background rendering, side swapping, settings and status updates.
- `ViewModels/SettingsWindowViewModel.cs`: validates and converts editable settings values.
- `Models/ComparisonModels.cs`: presentation, slide, rich text, table, element, diff and comparison result records.
- `Models/ApplicationSettings.cs`: persisted settings, safe defaults, upper bounds and conversion to parser limits.
- `Models/ApplicationInfo.cs`: product ownership and runtime version information used by the About window.
- `Services/PresentationServices.cs`: file picker and the core service contracts, Open XML reader, reading-order extraction, slide matcher and comparison engine.
- `Services/PowerPointPresentationRenderer.cs`: isolated STA-thread COM automation, PowerPoint warm start/retry, PNG export, cleanup and shutdown-safety policy.
- `Services/ApplicationSettingsServices.cs`: validated JSON settings persistence and the WPF settings dialog service.
- `Services/ApplicationDiagnostics.cs`: privacy-preserving rotating local logs and support summaries.
- `Services/AboutDialogService.cs`: About-window abstraction and WPF implementation.
- `Controls/SlidePreviewControl.cs`: built-in extracted-element preview and visual change overlays.
- `Controls/DiffTextBlock.cs`: selectable rich-text/table output with source formatting and inline diff styling.
- `Infrastructure/ObservableObject.cs`: property-change notification base class.
- `Infrastructure/RelayCommand.cs`: synchronous and asynchronous command implementations with centralized error callbacks.
- `Properties/PublishProfiles`: portable Windows x64 publish settings.
- `Assets`: application icon resources.

### Test project: `PptCompare.Tests`

- `ComparisonServiceTests.cs`: slide movement and text/image comparison behaviour.
- `PreviewRetryTests.cs`: background preview retry and independent loading indicators.
- `HardeningTests.cs`: diagnostics privacy, temporary-folder cleanup and PowerPoint ownership/shutdown safety.
- `PresentationSourceSecurityTests.cs`: invalid extension and malformed Open XML rejection.
- `ApplicationSettingsTests.cs`: safe defaults, validation and ownership metadata.

`PptCompare/Properties/AssemblyInfo.cs` exposes internal renderer seams to the test project. PowerPoint automation tests use fakes and must not require a live PowerPoint installation.

## Main application flow

1. `App.OnStartup` creates diagnostics, loads settings and constructs the services and `MainWindowViewModel`.
2. The main window opens empty. The user selects a file with the left or right Browse button.
3. `OpenPresentationAsync` calls `OpenXmlPresentationSourceService.LoadAsync` with the application lifetime cancellation token.
4. The Open XML reader validates the extension, path, file size, package structure and configurable resource limits before extracting content.
5. Text blocks are ordered spatially by resolved slide position (`Y`, then `X`, then document order), not merely by XML/z-order. The semantic title is used for matching but remains present at its visible position in the body output.
6. Text runs preserve font family, colour, bold, italic, underline, strike-through, subscript and superscript. Tables remain tables. Images and shapes become `SlideElement` records with bounds and hashes.
7. The built-in preview is immediately available from the extracted model. If PowerPoint rendering is enabled, high-fidelity rendering starts in the background and the affected pane shows a loading indicator.
8. The Compare command calls `TextPresentationComparisonService.CompareAsync` after both sides are loaded.
9. Slide matching first pairs identical fingerprints, which detects pure slide reordering. Remaining slides are paired by normalized title and positional proximity. Unmatched right slides are additions and unmatched left slides are removals.
10. Matched slides receive title/body diffs and element-level comparisons. The view model populates the navigator, synchronized panes, status text and changed/moved/added/removed totals.
11. Compare also retries only missing high-fidelity previews. When a new render arrives after a comparison, the view model refreshes the comparison while preserving the selected slide.
12. The swap command exchanges the complete left/right state and reverses directional diff markers.
13. Closing the main window disposes the view model, cancels outstanding work and disposes the renderer.

## PowerPoint automation safety rules

These are non-negotiable because violating them can destroy unsaved user work.

- Never kill `POWERPNT.EXE`, close presentations that PptCompare did not open, or blindly call `Application.Quit()`.
- Always render a staged temporary copy, opened read-only and without a visible window. Never automate the original file directly.
- Force-disable macros while opening the staged copy and restore the previous `AutomationSecurity` value promptly and again during cleanup if needed.
- Treat process/session detection failures as evidence of a shared user session. In uncertain states, leave PowerPoint running.
- Quit PowerPoint only when it was not running before PptCompare started it, PptCompare's staged presentation closed successfully, and the presentation collection is confirmed empty.
- Release every COM object explicitly and keep COM work on the dedicated STA thread.
- Serialize render requests with the existing semaphore; parallel PowerPoint COM automation is intentionally avoided.
- Preserve cancellation, the two-minute outer timeout, activation retry and warm-start logic.
- A rendering failure is non-fatal. Return an empty image set and retain the built-in preview.
- Preserve bulk export with per-slide export fallback.
- Keep rendered image and total-render size limits, bounded render-folder retention and stale-folder cleanup.

High-fidelity rendering may attach to an existing PowerPoint automation session or warm-start PowerPoint when none is running. Regardless of how activation occurs, PptCompare must not terminate a pre-existing or uncertain session.

## Security and privacy invariants

- Treat every presentation and settings file as untrusted input.
- Preserve all file, slide, element, table, paragraph, character, embedded-part and preview-image limits unless a reviewed requirement deliberately changes them.
- Open presentation packages read-only and do not save changes back to them.
- Do not log presentation content, filenames, full paths, usernames, document metadata or exception messages that might contain those values.
- Diagnostics should contain sanitized event names, exception types, HRESULT values and non-sensitive state only.
- Logs live under `%LocalAppData%\PptCompare\Logs`, rotate at a bounded size and retain a bounded number of archives.
- Settings live under `%LocalAppData%\PptCompare\settings.json`, are size-limited and strictly deserialized, validated and written through a temporary file before replacement.
- Keep debug logging opt-in and subject to the same privacy policy as normal logging.
- Unexpected optional-service failures must degrade safely and must not prevent the built-in comparison from working.

## Design and coding conventions

- Preserve the current service abstractions. Add SharePoint/Graph retrieval behind new or existing interfaces rather than coupling network access to WPF controls or the main view model.
- Keep business logic and file/COM operations out of window code-behind.
- Use `ObservableObject.SetProperty` for bindable properties and raise command availability when busy state changes.
- Pass `CancellationToken` through asynchronous call chains and honour it inside expensive loops.
- Avoid `async void` except framework-required UI event or command boundaries, where exceptions must be handled.
- Use ordinal or ordinal-ignore-case string comparison for identifiers, extensions and internal tokens.
- Catch specific expected exceptions inside services. Broad catches are permitted only at deliberate UI, diagnostics or optional-renderer boundaries and should include a justification and safe fallback.
- Prefer straightforward loops over LINQ where early cancellation, resource limits or reduced allocation matter.
- Use modern C# syntax when it improves clarity, but do not contort readable control flow merely to satisfy an informational style inspection.
- Do not remove apparently unused public WPF binding properties, COM-shaped test members, `VersionDescriptor.Id`, or `VersionDescriptor.ModifiedAt` without checking their framework or future-integration role.
- Preserve existing user changes in a dirty worktree and inspect `git status`/`git diff` before editing overlapping files.

The established visual identity uses `#2c4697` for the application/header blue, coral for the left presentation, and `#2ca02c`-based green styling for the right presentation and Compare action.

## Build, test and publish

Run commands from the solution directory:

```powershell
dotnet restore Solution1.sln --locked-mode
dotnet build Solution1.sln --configuration Release --no-restore
dotnet test Solution1.sln --configuration Release --no-build --no-restore
dotnet format Solution1.sln whitespace --no-restore --verify-no-changes
dotnet format Solution1.sln analyzers --no-restore --verify-no-changes --severity warn
```

Before handing over a change, at minimum complete a Release build and all tests. Add regression tests when changing slide matching, text order, formatting, comparison direction, cancellation, settings validation or PowerPoint lifecycle behaviour.

Portable single-file publishing:

```powershell
dotnet publish PptCompare/PptCompare.csproj --configuration Release --profile Portable-win-x64
```

The portable output is written to `PptCompare/publish` by that profile. PowerPoint is optional at runtime: without it, the executable must still load, compare and show built-in previews.

## Current limitations and planned extension points

- Windows is required because the UI is WPF and high-fidelity rendering uses PowerPoint COM automation.
- PowerPoint is required only for high-fidelity preview images.
- The Open XML preview is intentionally less visually exact than PowerPoint's renderer.
- Microsoft Graph authentication and SharePoint version-history retrieval remain future work.
- Local version selectors currently hold one selected local file per side; retain the collections and descriptors for remote history expansion.
