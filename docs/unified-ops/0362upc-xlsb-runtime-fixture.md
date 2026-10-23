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

## v15 follow-up

The next composed run on `dfd7` passed 770 tests and exposed two remaining
provenance comparisons. Direct structural commits retain the original
content-types member as stale source provenance while a replay can retain the
same graph's effective authored publication; those source byte images are
intentionally different even though the publication is identical. The graph
helper now permits only that explicit original-versus-effective distinction,
requires source presence and equal effective publication, and additionally
compares the complete serialized archive. All untouched part and relationship
source images remain exact, including the inverse path.
