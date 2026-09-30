# Changelog

## Unreleased

### Changed
- Library and font tool now target .NET Framework 4.5.2 instead of 4.0, matching VisioAutomation. Consumers targeting net40 must upgrade before adopting the next release.
- SDK-style projects restore reference assemblies and centrally managed dependencies; tests use .NET Framework 4.7.2 and MSTest 4.2.2.
- NuGet packages use the Release DLL under lib/net452, with README and MIT license metadata.
- Font specimen tool accepts an output path and font name, uses shared font allocation and typed cells, and keeps shapes inside page bounds.

### Added
- Drawing.ToXml() returns an independent XML snapshot using the same serializer as Save.
- Pure XML/ownership regression suite, Debug/Release CI, and build/architecture guidance.

### Fixed
- Repeated saves and retries after failed writes no longer duplicate pages, windows, or document properties.
- Rejected collection additions no longer mutate ownership or shape identity. Name lookups follow mutable names.
- Font IDs are allocated above existing sparse IDs; duplicate IDs are rejected.
- Connector creation rejects null connectors and cross-page endpoints before adding connections.
- Templates reject incorrect roots clearly and tolerate missing optional containers.
- Integration tests reliably close Visio, isolate output files, and attribute log records to unique input paths.

The existing 1.1.3 package version is retained for local package checks only. These changes have not been published.
