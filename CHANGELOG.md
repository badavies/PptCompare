# Changelog

All notable changes to PptCompare are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

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
