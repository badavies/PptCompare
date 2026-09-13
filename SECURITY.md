# Security policy

PptCompare processes presentation files and optionally automates Microsoft PowerPoint. Security reports involving malicious documents, unsafe file handling, unintended data disclosure or loss of work in an existing PowerPoint session are treated as high priority.

## Supported versions

| Version | Security updates |
| --- | --- |
| 0.2.x | Supported during the workplace pilot |
| 0.1.x and earlier | Not supported |

Only the latest published build in a supported release line should be deployed. Self-contained releases bundle their own .NET runtime and must be rebuilt and redeployed to receive runtime security updates.

## Reporting a vulnerability

Please do not open a public issue for a suspected vulnerability.

Use GitHub's private **Report a vulnerability** option on the repository's Security page when it is available. If private vulnerability reporting is not enabled, contact the repository owner privately through an agreed organisational channel before sharing technical details.

Include, where possible:

- The affected PptCompare version and Windows architecture.
- A concise description of the impact and reproduction steps.
- Whether PowerPoint rendering was enabled and the installed PowerPoint version.
- Sanitised diagnostic events, exception types and HRESULT values.
- A minimal synthetic presentation if one is required to reproduce the issue.

Do not submit customer presentations, confidential slide content, unsanitised logs, usernames, full local paths, access tokens, passwords or signing material.

Reports will be acknowledged and assessed on a best-effort basis. No fixed response-time service level is currently offered.

## Security design

PptCompare is designed to:

- Process presentations locally without telemetry or application-managed uploads.
- Treat presentations and settings files as untrusted input and enforce bounded resource limits.
- Validate the Open XML document type against its extension and bound relationship references, unique related parts and cumulative decompressed related content.
- Validate PNG/JPEG headers and source dimensions before image decoding, then cap retained decoded pixels per slide preview.
- Open presentation packages read-only and avoid saving to source files.
- Disable macros when PowerPoint opens staged temporary copies for preview rendering.
- Avoid logging presentation content, filenames, paths or document metadata.
- Keep debug logging opt-in and subject to the same privacy rules as normal logging.
- Never terminate or close a pre-existing or uncertain PowerPoint session.
- Keep PowerPoint automation serialized until a timed-out or cancelled STA worker has actually exited.
- Degrade to the built-in preview when optional PowerPoint rendering fails.

These measures reduce risk but do not replace endpoint protection, supported Microsoft Office versions, signed releases, least-privilege deployment or normal organisational security controls.

## Deployment security

- Publish from a reviewed commit with a clean working tree.
- Restore NuGet dependencies in locked mode and run all tests and static analysis.
- Sign and timestamp the final executable or MSIX package after publishing.
- Never commit certificate private keys, `.pfx` files, passwords or signing-service credentials.
- Distribute workplace builds through an approved channel such as Microsoft Intune.
- Verify the signature and release checksum before deployment.
