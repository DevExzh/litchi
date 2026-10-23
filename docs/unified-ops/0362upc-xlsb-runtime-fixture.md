# 0362upc — XLSB runtime fixture correction

The diagnostic XLSB/PPTX/ODF runtime check at `d4d1820dbe` reached 769
passing tests and three fixture failures. No native budget or projection
implementation failed.

The metadata limit test now snapshots the package's actual metadata before
the refused edit. Newly authored XLSB packages publish producer application,
company and security defaults, so comparing the unchanged owner with
`Metadata::default()` was too strong while the exact one-byte retained
snapshot limit remains exercised.

The two OPC projection tests now begin from a reopened package. An in-memory
builder package has no retained `[Content_Types].xml` source, while a replay
source reopened from bytes does. Comparing those unlike provenance states
made the graph helper fail before it could validate the intended result. The
fixture is now source-backed on both sides; the tests still assert that an
unreferenced edit defers projection and that a transitively referenced edit
eagerly reparses and validates it.

Cargo was not run in this leaf per coordination. The coordinator should rerun
the focused XLSB runtime group and the composed v13 check after merging.
