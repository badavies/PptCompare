# Changelog

All notable changes to PptCompare are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

### Added

- Repository README, security policy and MIT licence.
- Native Arm64 publication guidance.

### Changed

- Renamed the solution from `Solution1.sln` to `PptCompare.sln`.
- Expanded source-control exclusions for generated output, IDE state, diagnostics and signing material.

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
