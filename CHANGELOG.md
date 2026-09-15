# Changelog

All notable changes to PptCompare are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

## 0.3.0 - 2026-09-15

### Added

- Regression coverage for external hyperlinks, horizontally and vertically merged table cells, repeated-title metrics, large word diffs, PowerPoint export numbering, temporary-source cleanup and per-user PowerPoint registration.
- A presentation-wide XML-part budget that counts each unique slide, layout, master and theme once before materialisation.

### Changed

- Replaced repeated Open XML `OuterXml` serialization and reparsing with direct typed-root conversion, while caching shared layouts, masters and themes for each load.
- Coalesced related selectable-text property changes into one WPF document rebuild.
- Preserved merged table structure and continuation text without rendering duplicate placeholder cells.
- Reduced word-diff memory pressure with rolling rows, bounded tokenization and compact chunked reconstruction data.
- Counted a spatially repeated title only once in change summaries while retaining it in the visible slide text.
- Required exact PowerPoint bulk-export slide numbering before accepting preview images.
- Isolated staged source presentations from the preview cache and hardened cleanup for read-only files, late workers and reparse-point paths.
- Added per-user PowerPoint registration discovery before machine-wide fallback.

### Security

- Bounded aggregate decompressed XML processing across presentation parts.
- Prevented staged source presentations from being retained inside cached preview folders when cleanup is delayed.

## 0.2.0 - 2026-09-13

### Added

- Repository README, security policy and MIT licence.
- Native Arm64 publication guidance.
- Comprehensive boundary, malformed-file, resource-budget, image-decoding and PowerPoint-timeout regression tests.

### Changed

- Renamed the solution from `Solution1.sln` to `PptCompare.sln`.
- Expanded source-control exclusions for generated output, IDE state, diagnostics and signing material.
- Declared and locked both `win-x64` and `win-arm64` publication targets.
- Bounded cumulative Open XML related-part processing and reused cached part hashes.
- Reworked element matching to remain globally bounded and cancellable, use the true closest candidate while the work budget permits, and avoid arbitrary matches after exhaustion.
- Hardened preview image validation with a PNG/JPEG allow-list, source-dimension limits and a cumulative retained decoded-pixel budget.
- Prevented a timed-out PowerPoint COM worker from overlapping a later render request.
- Strengthened presentation document-type and structural validation.

## 0.1.0 - 2026-09-12

### Added

- Side-by-side local `.pptx` and `.pptm` selection and comparison.
- Slide matching for changed, added, removed and reordered slides.
- Selectable rich-text output with formatting and table preservation.
- Image, shape, layout and visual comparison.
- Immediate Open XML preview with optional high-fidelity PowerPoint rendering.
- Independent preview loading indicators and safe retry behaviour.
- Configurable parsing, resource, preview and output-font limits.
- Privacy-preserving bounded diagnostic logging with an opt-in debug mode.
- About and Settings windows.
- Regression tests for comparison, settings, malformed input, cancellation and PowerPoint lifecycle safety.
- Qodana and built-in .NET analyser configuration.
