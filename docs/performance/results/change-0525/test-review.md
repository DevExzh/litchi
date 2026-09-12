# XLSX row reuse test review

The applied test coverage exercises the proposed changed-row reconstruction
reuse against an independent raw worksheet parse. The semantic oracle compares
cell addresses and values plus styles, rows, columns, defaults, extents, merge
ranges, shared strings/rich text, formulas and shared-formula metadata, cell
metadata, and value metadata.

The dedicated `row_reuse_tests` module in
[`snapshot.rs`](../../../crates/litchi-xlsx/src/cell_values/snapshot.rs) covers:

- exact candidate provenance-writer bytes versus ordinary `rewrite` bytes;
- direct `Store::merge_omitted_cells` success for two disjoint omissions on one
  row around a changed cell and a following row whose first column is lower;
- numeric, boolean, escaped inline text, date, error, plain and cached formula
  replacements, including stale-dimension expansion and explicit `r` markers;
- omitted implicit cells, implicit row numbering, self-closing untouched rows,
  namespace/comment sentinels, and exact omitted-byte/address provenance;
- changed-cell readback from mutated emitted bytes and rejection of a proof
  bound to a different source snapshot;
- clear, remove, and insert within an existing row, including refusal when an
  implicit follower would be shifted;
- new-row and shared-formula fallback behavior, malformed output and
  MCE/unknown-attribute refusal boundaries; and
- no-op source editing with semantic-store sharing.

The public source integration file adds
`removing_a_cell_with_an_implicit_follower_refuses_shifted_readback`, which
checks the same refusal through the source-backed editor API.

The eligible cases assert nonempty omitted spans and compare their raw bytes to
the unchanged source records. The mutation and source-binding cases therefore
cannot pass by replaying staged values or by falling back silently to a store
from another snapshot. The direct merge test requires `Some`, so a complete
parser fallback cannot conceal an omission-range ordering regression.

## Fixture construction and preflight diagnosis

`scalar_fixture()` is intentionally pretty-printed: its whitespace, comments,
prefix aliases, and nested/default namespace scopes are part of the raw-source
coverage. Tests that call `Snapshot::load` place those bytes directly in the
private in-memory `OpcPackage`, so they do not pass through publication's
compact-XML gate. The `commit_with` and no-op tests use the source-backed editor
path. Its fixture helper now first publishes a compact placeholder worksheet,
then reads that archive and rewrites only `xl/worksheets/sheet1.xml` with the
exact fixture bytes through `ArchiveReader` and `StreamingArchiveWriter`. This
retains the intended formatted source while still exercising the normal package
ingress and generated-output compactness checks.

Preflight-2 recorded 978 passes and two setup failures. Both failures occurred
before the row-reuse assertions in `source_backed_editor`: `PackageWriter` was
given the indented `scalar_fixture()` and returned
`XmlPublication { part: "/xl/worksheets/sheet1.xml", source:
NotCompact(FormattingWhitespace, offset: 112) }`. The placeholder/archive
replacement above addresses that fixture-only failure without weakening a
production check. The latest preflight-3 receipt has exit code 0, with no failed
tests in its output and 1,288 passing test-result counts across the XLSX test
invocations.

## Exact inputs

All hashes below were captured from the current shared worktree. The patch is
generated against the clean base revision shown here; the two Rust file hashes
are the currently applied, rustfmt-formatted files used by the parent lane.

| Item | SHA-256 | Git blob hash | Size |
| --- | --- | --- | ---: |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `e974e6e94d4af3c4797d883a60410726b61081ca99adea46f786099485b6b4d6` | `323b4da587d402fb863ebac8e8086c40d649de58` | 105977 bytes |
| `crates/litchi-xlsx/tests/source_backed_cell_values.rs` | `d8270d5682af4764be54428c520b979a2cec035d881862654bbf5b21425fd166` | `24aa345a07fc6d46a597c2973c9b47d0a4e71e7c` | 103434 bytes |
| `docs/performance/results/change-0525/row-reuse-tests.patch` | `afaffd3571a018e8611ba2851b9f795ecef6788970106023138c794217b1c56f` | `6be9d72d3d81f513951ce910fb9926ac29e5cfef` | 26989 bytes / 618 lines |

The patch base is `f08daf3976714dffebe37e40bd895d87266aaf8d`. Its unified-diff
application was checked against clean temporary copies of that base. The
shared worktree already contains the applied test hunks, so an application
check against the current worktree would correctly report them as present.

For provenance, the candidate source patch artifacts used by the parent lane
have these SHA-256 hashes:

- `row-reuse-candidate.patch`: `eb3b78b043b7ab30e62dfe48dda6f107328331ad80c335c0d46d963e96763c38`
- `row-reuse-layout-followup.patch`: `7e0bbf6e830952d347e6d3d3a453e4ec65eb4348910cabdd47bf5201e54ee328`
- `row-reuse-package-simplification.patch`: `36520cbacead4c238954c186a5f5e6c2961d4137f4bf3c3309d764b425e11bff`

No tests, builds, or captures were run for this review update. Static review
found no additional correctness blocker after the candidate `Rect` omission
ordering fix; the direct `Some` assertion above guards that regression.
