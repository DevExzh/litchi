# 0218 — Native PPT shape edit runs

Consecutive direct top-level PPT shape text/geometry assignments now share a
native primary transaction. The coordinator closes a run before a repeated
target, selector, read, structural operation or resource boundary. Repeated
shape geometry is resolved again against the preceding committed candidate.
Original handles remain immutable; a failed atomic transaction drops all
candidate revisions. The native package owns parsing, preservation and final
publication.

The facade also streams shape candidates during ordinary path lookup. Ordinal
selection stops at its match without allocating a vector for unselected
siblings; identity lookup still counts every matching candidate to retain
ambiguity detection.

## Verification and timing

The existing all-feature facade suite passes 1,594 tests/doctests. Nine new
public transaction tests additionally pass for eager and positional owners,
covering 1,000 distinct setters in one commit, native 4,096-operation rollover,
mixed text/geometry with exact patch replay/inverse, repeated-target replanning,
query/read/raw/structural barriers,
per-item no-ops, atomic rollback, best-effort continuation and explicit limits.
The native canonical-edit suite passes 12 tests. Minimal-feature reads, writes,
structure, capabilities and allocation checks pass; full-feature Clippy,
warning-denied rustdoc, crate boundaries and the generated OfficeCLI parity
check pass. The parity tooling's 15 Python tests pass.

The allocation regression compares facade overhead against the same native
shape read and confirms that selecting the first shape does not add allocation
per unselected sibling in a 1,000-shape slide.

The final PPT-only coalescer/allocation checks pass 15 tests; the added mixed
history assertion passes separately. The harness has three focused tests for
route naming/pairing, native chunk boundaries, exact qualification and explicit
snapshot policy. Harness tests and warning-denied Clippy pass. Both repository
lockfiles remain unchanged. Facade validation uses a copied external lock with
cached cc 1.6.0 instead of unavailable cc 1.7.0; the harness uses its existing
lockfile contents copied outside the repository.

## Named paired measurements

`tools/perf-baseline` now registers twenty opt-in routes under new names:
`ppt_text_batch_{eager,positional}_{native,unified}_{1,10,100,1000,10000}_v1`.
Existing frozen route names and default selections are unchanged. Each pair
uses one deterministic source and qualifies exact output equality and source
preservation before timing, then checks every timed output against that oracle.
The harness refuses unexpected native/unified commit counts. The unified leg
uses one atomic facade edit even when native capacity requires three commits.

Five samples follow one warmup per leg, native then unified. Fixture creation,
opening, command-vector preparation, serialization and verification are outside
the timer. Unified string-path parsing and patch destruction are included. The
1,000-operation route uses the default 256-MiB retention profile; the
10,000-operation route explicitly selects a finite 1-GiB profile. Release LTO
is disabled. The worktree's uncommitted changes during final timing are docs
only. The earlier exploratory example was stopped after overlapping agent/build
work and is not used as performance evidence.

Runtime source: `0879bd5ea2` (core integration `1137b3f3d9`, harness
`3f61d48c09` / `cdaa59177a`). [Raw samples, exact output hashes and build
identity](results/0218-ppt-batch.json) cover all twenty routes. All pairs pass
exact output/source checks and expected commit counts: one through 1,000,
three at 10,000. Times below are medians; ratios divide the unified median by
the native median.

| Operations | Eager native ms | Eager unified ms | Ratio | Positional native ms | Positional unified ms | Ratio |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 0.111 | 0.116 | 1.040 | 0.116 | 0.125 | 1.072 |
| 10 | 0.169 | 0.208 | 1.233 | 0.172 | 0.232 | 1.345 |
| 100 | 2.320 | 2.704 | 1.165 | 2.345 | 2.883 | 1.229 |
| 1,000 | 193.593 | 197.383 | 1.020 | 192.041 | 196.520 | 1.023 |
| 10,000 | 24187.691 | 23992.052 | 0.992 | 23611.467 | 24325.142 | 1.030 |

The default-profile 1,000-operation closure workload meets the 2× relative
budget: overhead is 2.0% eager and 2.3% positional. The one-operation routes
are below the 10% relative review trigger. Ten- and hundred-operation batches
show 16.5–34.5% relative overhead; these do not establish a universal 10% write
claim. Blocked leg ordering can retain drift, and these are descriptive local
measurements rather than broad-corpus acceptance.

**Linear scaling fails.** Increasing 1,000 operations to 10,000 increases native
time by 124.94× eager / 122.95× positional and unified time by 121.55× / 123.78×.
The near-native 10k ratios do not close this requirement. Native publication and
drawing rewrite work below remain required follow-ups.

These routes measure the closure API with string paths; command-batch JSON
conversion, cold opens, geometry batches and a broader corpus are not qualified
by this matrix.

## Native scaling follow-up

Coalescing the facade commands does not yet eliminate native repeated work.
`slide_order::Transaction::set_shape_texts` invokes
`text_edit::replace_shape_texts_batch` once per 256-item chunk: four calls at
1,000 operations and forty at 10,000, across three primary edits at 10,000.
Each call opens an editor, parses, publishes append-only history, validates
unrelated streams and reopens for semantic readback. `replace_shape_texts_batch`
also calls `rewrite_slide` per replacement, recursively rebuilding the whole
slide drawing each time. Same-slide workloads therefore retain a quadratic
native drawing traversal. `package::Edit::stage` additionally clones staged
transaction state for each chunk.

The next native performance step is a grouped, single-traversal drawing rewrite
plus aggregate publication for each 4,096-operation profile. Preserve duplicate
ID and exact-before checks, dependency closure, cumulative resource admission,
source fences, unrelated-stream validation, complete readback and failure
atomicity. Keep the public 256-item logical chunk bound rather than merely
raising it. Qualification should cover 256, 1,000, 4,096 and 10,000 replacements,
count actual publication passes and compare exact facade/native outputs.

## Remaining native insertion work

The native transfer plan copies a complete portable donor slide and requires
matching master closure. It is not a blank-slide constructor. Rebuilding an
opened deck with `Builder` would replace opaque streams, masters, styles and
identities, so it is not a preservation-safe implementation of unified `add`.

The next native step is a typed blank-slide specification and primary-edit
insertion operation: select an existing master, allocate persist/slide/drawing
IDs, construct an empty native drawing, insert the slide group and update
Dg/Dgg counts while retaining unrelated package bytes. Existing entry points
are `slide_order`, `document_structure::Transaction::insert_slide_group`, and
writer Escher/slide-container codecs. Native patch replay/inverse must cover
the new insertion before the facade exposes it.

Top-level textbox insertion and shape removal additionally need native drawing
closure handling for shape IDs, Dg/Dgg counts, pictures, groups, links and
animations. Required regressions include custom masters, notes, noncolliding
IDs, exact publication/reopen/inverse, and security/source/limit refusals. These
engines, broader CRUD, all-format coverage and Python/CLI acceptance remain
part of the full goal.

The base remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`, already an ancestor.
