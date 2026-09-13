# 0546 integrated candidate source review

This is a read-only review of the frozen 0546 candidate source bundle. The
candidate patch is `candidate-sources/candidate.patch`, SHA-256
`bd9a0e25a5ce1a7dde78d568e4c1c998dd4c7103318269da31f201a125a9d539`, based on
revision `6b866732489413366b40c94da7aa77bd37e719c3`. The source note is
`candidate-sources/source-note.md`, SHA-256
`219c60c20e4448fa183d1091e6eb2497a65b85c378888ac2aad194523994eb76`.
Preparation and this review ran no Rust build, test, or runtime command.

## Static disposition

No new source-level freeze blocker was found. The candidate is conditionally
approved to proceed through the fresh integrated campaign after the terminal
baseline stage. This review does not admit runtime retention. The integrated
candidate still has to pass every original workflow, allocation, refusal,
resource, preservation, publication, native, quality, and conditional gate,
plus the new valid sparse cap controls.

The candidate diff contains exactly the five planned files under the allowed
XLSX roots. It carries the reviewed 0544 source path and changes the marker
term in `raw/worksheet/mod.rs`; the direct oracle and exact-cap tests are
strengthened in `shared_traversal_tests.rs`. No Cargo metadata, dependency,
public API, unsafe code, event limit, source limit, or validation policy is
changed.

## Adaptive bound

The production helper retains the exact 0544 bound:

```text
1 + initial_nonmarker_text
  + count('<' or '&')
  + count('>' or ';' followed by a nonmarker)
```

Only the first marker term changes. It probes at most 16 marker hits with
`memchr2`. If fewer than 16 exist, the failed search proves the marker scan is
complete. Otherwise the suffix after the sixteenth hit is partitioned into
64 KiB chunks, and disjoint single-byte `memchr_iter` counts for `<` and `&`
give the exact remaining marker count. Checked addition still declines on a
cap crossing or arithmetic overflow. The existing full `memchr2` successor
scan for `>` and `;` is unchanged, so the mathematical bound and its
one-way conservative admission proof are preserved.

The suffix loop borrows the source and retains only scalar locals. It can read
up to one 64 KiB chunk before noticing a cap crossing, which is bounded and
safe but remains part of the resource/performance review. The 8 MiB source and
MCE input/output limits, 131,072 provisional event cap, ordinary 1,000,000
event cap, parser depth/record/scalar/formula limits, and checked collection
reservations remain in force.

## Reader, fallback, and preservation boundaries

The direct `NsReader` oracle now pins the relevant configuration explicitly:
`allow_dangling_amp = false`, `allow_unmatched_ends = false`,
`check_comments = false`, `check_end_names = true`,
`expand_empty_elements = false`, `trim_markup_names_in_closing_tags = true`,
and `trim_text(false)`. It compares the production predicate with an
independent lexical bound and checks every direct reader prefix. Its cases
cover malformed tails and references, BOM/declaration/PI, comments/CDATA/
doctype delimiters, text and successor pairs, final bytes, exact cap and
cap-plus-one streams, and marker/successor chunk boundaries. The deliberate
lexical false-positive case still exercises authoritative fallback and exact
source/no-op preservation.

`source_stream_eligible` runs before the provisional reader and retains the
existing UTF-8, MCE, source-size, MCE-marker, and x14ac-marker fences. A bound
failure returns `ProvisionalFailed`; reader, observer, or parser-transition
failure also drops provisional state and repeats authoritative worksheet
validation followed by the raw parser. The observer runs before parser
transition, so a first validation error cannot be replaced by a speculative
parser error. Completion requires the closed worksheet root and EOF; owned
parser materialization occurs in `finish_parse`, and `validator.finish()` is
checked before the result crosses the raw facade. Borrowed event data does not
escape the higher-ranked observer callback.

The moved `complete_source_parse` path preserves the historical x14ac retry:
eligible successful parses skip a redundant marker-free extension scan, while
rejected parses retain the extension capture/error precedence, and ineligible
sources keep the established two-pass route. The private
`SourceParseAttempt` lint expectation remains local to its deliberate owned
post-EOF result.

## Retained sparse cap enabler

The retained enabler diff is only
`crates/litchi-xlsx/examples/perf_cap_boundary.rs`; its baseline source patch
is SHA-256
`7cdccdd8434ea72e2169cceedda0e584c75a0c46fd9e54d6c443476b497160c3`, and the
enabler source is bound in both stage manifests as
`7a5030ef7b867338fd69d0ba5ed7d9f3222901e3c2121437e9c9b39efc6cec9f`.
Sizes 1 and 2 append exactly a one-MiB ASCII XML comment before the worksheet
close. The comment is valid, marker-free, adds one XML event, keeps the source
below the shared byte fence, and leaves the numeric cell snapshot unchanged.
The checked event formula is `5*N*N + 2*N + 6 + I(N <= 2)` including EOF.

Size 1 stays below the 16-marker probe; size 2 crosses it and exercises the
64 KiB bulk suffix count. The example times only `edit_sheets`; setup,
inspection, empty commit, source-byte checks, and exact no-op publication are
outside the clock. Each size/repeat must keep planning p50 and mean within
1.05x its matched baseline. The enabler is identical in both builds and is a
diagnostic harness input, not a production API.

## Diagnostic screen and integrated gate

The isolated screen passes its stated rule: all ten numeric/cap fixture rows
clear the 20% p50/mean reduction threshold and the remaining controls stay
within 100,000 ns. The clustered-marker control remains visible as a real
risk: its p50 rises about 174.0%/174.6% across repeats (roughly 28.2 us over
the approximately 16.2 us baseline), with corresponding mean increases of
about 171.5%/173.1%. This is scanner-only diagnostic attribution, not an
end-to-end XLSX speed claim, and it cannot be hidden by the absolute control
budget.

The integrated cap plan adds valid sparse sizes 1 and 2 to 160, 164, and 256,
with both p50 and mean required within 5% for every repeat. The full original
0544 gates remain mandatory; conditional profile, hardware, and eager lanes
can only follow a passing native/allocation/refusal/cap pilot. Any sparse,
semantic, error-order, preservation, or performance failure restores the
baseline.

## Source provenance

The source note binds the candidate files to the sealed 0544 source and the
current baseline. The candidate source hashes are:

| Candidate file | SHA-256 |
| --- | --- |
| `cell_values/snapshot.rs` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `cell_values/validation.rs` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `cell_values/shared_traversal_tests.rs` | `5c56a86bd67d014c28823c9c315d6abbed8c614faf6f458795e3015ecfde3b95` |
| `raw/worksheet/codec.rs` | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` |
| `raw/worksheet/mod.rs` | `8aedf9ab085f5dc627e76d8e13841f211a1a53d4dd02f67cf389a6ffe70cd6e1` |

The integrated plan and sparse cap plan are bound by SHA-256
`6835a4122a16bf6ec8a5f3513e27be50749d3a8b3a471261976de6f3c4302574` and
`c41066c1e0081c2aaf6385da4bf83b505de008f87f3fe4560a36be603c1ef421`;
the frozen input bundle is
`4bcdbb7c6f5d7390b847b4170c1f5a3771eb58dbfa4ddef644418b23abf72527`.
The isolated diagnostic design review is bound by
`6e69f2d6c6a9b29d74d439bec3bca18e047e174c3dad9e98747ba8af364b75e7`.

