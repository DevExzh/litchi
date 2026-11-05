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
compares the complete effective archive. All untouched part and relationship
source images remain exact, including the inverse path.

## v17 follow-up

The facade runtime run at `v17` passed 771 tests and left one failure in the
transitive projection test. Its old check required one retained
`[Content_Types].xml` source image to equal the current publication whenever
the two images differed. That is not a valid provenance rule: replay and the
direct candidate can retain different earlier generations even when they
publish the same effective archive. The transitive fixture now permits source
lineage differences only for the changed package root and
`/custom/transitive-root.bin` relationship owner. It continues to require the
same semantic graph and publication, exact source presence and bytes for every
untouched relationship owner, and the complete effective archive member set
and decompressed payloads. Here “effective archive” means the sorted set of
published member names and their decoded payloads: the comparison deliberately
ignores ZIP compression, local/central framing, member order and other ZIP
container metadata. This follows the existing writer contract in
[0303 §10](0303-pptx-docx-producer-structure-gaps.md): Part bytes,
relationship XML and `[Content_Types].xml` are exact, while ZIP container bytes
may be recompressed. It is not a new exception or waiver. The broader member
comparison is scoped to these transitive replay/direct/inverse cases and their
changed relationship owners; same-owner and unreferenced comparisons retain
their exact serialized-byte checks and untouched source-image checks.

The granted focused receipt was generated outside the checkout (so no large
evidence is committed):

```text
CARGO_TARGET_DIR=(deleted-coord-artifact) \
TMPDIR=(deleted-coord-artifact) \
CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 \
cargo test --offline --locked -p litchi-xlsb --all-features --lib 'opc::tests::' \
  --message-format=json-render-diagnostics -- --test-threads=1
```

Result: 13 passed, 0 failed, 759 filtered out, finished in 0.12s. Receipt:
`opc-focused-v19.log`, 159901 bytes,
SHA-256 `de5143c536f00eea1eb3b8babb4053dadb6f2e501c3473a5c6b71d1e18523743`,
stored under `(deleted-coord-artifact)`.
