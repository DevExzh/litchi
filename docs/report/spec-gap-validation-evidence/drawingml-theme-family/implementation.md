# DrawingML theme-family metadata

The shared `litchi_drawingml::theme::family` owner implements the exact
`themeFamily` fragment in MS-ODRAWXML sections 2.4 and 5.17. `Family` and
`Guid` provide bounded typed authoring; `Snapshot::edit` publishes a new
snapshot and an exact-source-checked reversible `Patch`.

The three required attributes are the applied theme name, theme GUID, and
variant GUID. Name decoding follows XML 1.0 attribute normalization; authored
tabs and line breaks use character references. GUIDs use the schema's uppercase
braced form and XML token whitespace rules. The source retains its original
prefixes, quotes, whitespace, comments, attributes, and opaque payloads.

The reader validates known `extLst`/`a:ext` structure and required `uri`
presence. Foreign direct root children and descendants inside extension
payloads remain opaque and are excluded from typed projection. All ancestry
still has finite XML, depth, node, attribute, namespace, and value bounds.
Malformed XML names, unresolved entities, DTDs, unsupported declarations, and
forbidden XML characters are refused. The supported declaration profile is
XML 1.0 with UTF-8, including an optional UTF-8 BOM; processing instructions
are outside this profile. This reader is not a general XSD validator or an
MCE processor. It requires a namespace-self-contained fragment.

Clone and no-op transactions share retained allocations. Changed transactions
replace only changed attribute value spans and validate the result before
publication. Patch application uses a shared-source fast path or complete byte
comparison, so independently reopened equal sources work and stale sources
fail. Patches are in-memory snapshots, not durable serialized operations;
composition, three-way merging, and host package attachment/removal remain
open. The reader also accepts Strict DrawingML `ext` children; the retained
independent XSD evidence exercises the Transitional profile.

## Validation

The [gate receipt](gates/receipt.json) records full crate tests, strict
all-target/all-feature Clippy, warning-denied rustdoc, formatting, and diff
checks. Source hashes are captured before and after the commands and must
match. The 19 independent feature tests include exact/over-limit cases,
malformed XML, semantic whitespace, exact native replay, opaque retention,
failure atomicity, and forward/inverse application to independently reopened
snapshots. Run `python3 -B run_gates.py` from this directory or use its full
repository-relative path.

The native fixture is an exact substring of a checked-in LibreOffice test
package. Its provenance records the archive and member hashes. The standalone
[caller example](caller-example/main.rs) produces actual authored and edited
fragments; [schema validation](caller-schema-validation.json) uses the
vendored Microsoft and ECMA schemas without network access. This establishes
fragment schema validity and readback, not native Office acceptance of changed
packages. See [performance methodology](README.md) and [measurements](report.md)
for the separately scoped allocation and timing evidence.

The audit-wide specification and performance objective remains active.
