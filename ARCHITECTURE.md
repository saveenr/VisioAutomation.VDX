# Architecture and maintenance

## Component boundaries

- `VisioAutomation.VDX`: BCL-only document model and VDX serialization. It has no Visio COM, PowerShell, or VisioAutomation runtime dependency.
- `VisioFontCompare`: console sample. Reuses `Drawing.AddFace`, typed cells, and the library's point conversion.
- `TestVisioAutomationVDX.Unit`: XML and ownership regressions with no Visio dependency. This is the CI gate.
- `TestVisioAutomationVDX`: real-Visio round-trip checks and the published log-parser compatibility adapter.

Keep serialization rules in the library, not in tools or tests. Introduce shared helpers when they remove concrete duplication; do not couple this independent file generator to the COM automation library.

## Tooling and portability

`VisioAutomationVDX.slnx` contains only the four projects and shared configuration files. SDK-style project files use central package versions and downloaded framework reference assemblies. The shared VisioAutomation/VDX build baseline is VS 2026, the .NET 10 SDK selected by `global.json`, and explicit C# 14. This does not change the net452/net472 runtime targets or introduce a sibling-checkout dependency.

The default template is a single named embedded XML resource, loaded lazily using framework APIs. The duplicate binary template, ResX wrappers, empty test resources, unused WinForms reference, legacy solution settings, and checked-in NuGet executable have been removed. XML fixtures no longer carry personal creator metadata or absolute Office/stencil paths.

Public types, members, shipping assembly identity versions, and COM visibility/GUID metadata are preserved. The SDK generates descriptive metadata, release-specific file versions, and informational versions with Git commits. The new unit-test assembly now has complete metadata and explicit COM invisibility. Package authorship uses a project-level contributor label; copyright/license attribution and valid hosting URLs are intentionally retained until a repository transfer supplies replacement URLs.

## Model and serialization

`Template` validates a Visio 2003 XML root, retains master/resources, and removes template pages and window state. Missing optional containers are created. Masters retain their original XML; the model records their IDs and subshape counts for generated shapes.

`Drawing.ToXml()` returns an independent snapshot. The internal writer clones the cleaned template for every call, validates document-window page references, then writes current model state. `Drawing.Save(path)` uses that same serializer and disables XML formatting to protect mixed-content text whitespace. A failed file write or mutation of a returned snapshot cannot change the template.

`NamedNodeList<T>` validates before attaching ownership and exposes a read-only item view. Names remain mutable for compatibility, so lookup scans current names rather than caching stale keys. Callers must keep names unique after renaming. Faces also require unique numeric IDs; `Drawing.AddFace` reuses an existing case-insensitive match or allocates above the largest ID. Page/shape additions and connector creation validate ownership before modifying state.

These are mutable builder objects, not thread-safe documents. Use one builder per operation; do not mutate a drawing while serializing it. This is a focused VDX generator, not a general lossless editor for arbitrary Visio XML or a VSDX writer.

## Maintenance checklist

1. Add a failing pure regression test for serialization or collection changes.
2. Build Debug and Release, then run the pure suite in each configuration.
3. For file-format changes, run the Visio integration suite and inspect XML warnings as well as failures. Existing fixtures deliberately include warnings.
4. Run the font sample and inspect package contents. Update [CHANGELOG.md](CHANGELOG.md) for consumer-visible changes.
5. Before releasing, run integration tests on the supported Visio versions/architectures and confirm no new Visio process remains.

Public API spellings, including historical `GetMasterMetData(int)`, are retained. Avoid renaming public members or tightening behavior without tests and a compatibility note.

## Handover evidence

The aligned toolchain was also verified with VS 2026 and SDK 10.0.401: Debug and Release rebuilds, all 23 pure tests in each configuration, and all 6 integration-project tests in Release passed. Portability and C# style checks passed, Release packaging succeeded, and no Visio processes remained after the completed runs. Evidence is in `TestResults/aligned-*.log` and `TestResults/aligned-*.trx`. The earlier VS 2022 results below remain historical evidence.

Local verification on 2026-09-29: VS 2022 Debug and Release builds; 23 pure tests pass in both configurations, and 6 integration-project tests (5 use real Visio) pass in Release on x64 Visio 16.0.20326.20158. The original regression baseline had 14 failures among 17 tests. Evidence is under ignored `TestResults`; see [BUILDING.md](BUILDING.md) to reproduce it.

Modernization verification also compared 829 public/protected type/member signatures and enum constants against the pre-modernization DLL with no differences. A source-only export restored from nuget.org into an empty package directory, built all four projects, and passed all 23 pure tests without the original checkout's build outputs. XML comparison confirmed that fixture changes were limited to the intended metadata removal.

This is not an exhaustive schema/API audit, a multi-version Visio certification, or a completed ownership transfer. Confirm GitHub administration, NuGet package ownership, release credentials, and the minimum-runtime/version decision with the incoming maintainer. Hosted CI execution and external publishing access need separate verification.
