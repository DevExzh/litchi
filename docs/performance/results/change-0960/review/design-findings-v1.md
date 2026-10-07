# 0960 worksheet cell-capacity design review

Reviewed 2026-10-07 at repository `a6223afa0899d5800d21ed9dc6c49304fbb2e8ca`.
This is a read-only source/design review. No checkout source, test, build, or
workload was changed or run for this review. The checkout already had these
unrelated paths before review: `tools/check_iwork_leaf_public_api.py` modified,
`crates/litchi/examples/bench_word_parse.rs` untracked, and
`docs/UNIFIED_OPS_API_DESIGN.md` untracked.

## Bound sources

The relevant current production-file hashes are:

| file | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `94b5c115d0e44b51e66bc934a874eb51283567759ccf649a69422f3e096ea2bb` |
| `crates/litchi-xlsx/src/raw/worksheet/model.rs` | `3ad94f003d4ccaba668bab7c2fb531e6887a9594b3d6311edaae07a7f03d37d3` |
| `crates/litchi-xlsx/src/raw/worksheet/lane.rs` | `5950735336845ff546bc6cfd2ce914617709f6c5732f71940ed0dafde4879b03` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `3a1bf6412699738556dd298610580334615ae3b995dd192fddbacad94a3fc6e1` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/facts.rs` | `eae8c8be1d5af80fa097ecbd60a79d79c94aace5d79e8e00c0f3fab2b327a148` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `e153634e4aec523871e87749bb99a7c2120dfa139743891b3816b0a7b1e72215` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `183ab8ebb334a3d15e5d0312350915a756467fb578826833da17ac0565d6d53d` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `50a4ef8fd9f2e810d2b1315d8a389086134a28901a2b86a98cc8abeb84308726` |
| `tools/perf-baseline/src/lib.rs` | `b2ae866c0b10062c57170bd0ef35b5cfd6a46932388718740298134d8272c1be` |

The applicable normative files were read/revalidated against these hashes:

- `docs/GOAL.md` — `bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`
- `docs/adr/README.md` — `2af0e5814378c2765c12213e495d02c2895a27f142bc618ca1f3037824bbe028`
- ADR0001 — `813bf4faf62cedac56fbfa3bf9b7600121b42a3ed5d4cec0028f8c8f2ac2798e`
- ADR0002 — `73f4372ef6948f5a36c83fbf27a96f0b1f680225f331f201c470e04913417719`
- ADR0003 — `9200d3546d91f1604d44a9ade7ad0a6bdea50b6442d12a0517b09db5e8e802ef`
- ADR0005 — `8279fd2dc0b074aa104bf0c83df6187d12e18756cd8c8d7f25750b2a997cb0aa`
- ADR0006 — `ae21189eda0acc9524a7c5a57e87bb3566405ede5c015c87105e01337666f44b`
- ADR0007 — `8808b79dcb35eaeaa6be778b5a9593101dd302f866ef20ff4f8fdf8df5d8b100`
- ADR0008 — `5c72d79a7ba78a044471b40ba22ab3f32e4f8f878614ab05b416e60b0163a2`
- ADR0018 — `7007d34208fdd2e2072f2126b71a79a369e9ef753cb6905c2677a211e1675f7b`
- ADR0024 — `cc134b2a4eb8797a3769cf815f1f09618fd26e215bd7305de1c9890093dbb1b6`
- ADR0030 — `f0a2a679aeb4578f6ca3c4a762e546a820493d27dd205a462ffe2ba89c9f1561`
- ADR0031 — `ce7e16baa101cfdd8abfc6d606dc0a56fad6d86cc79eba997d338f919a063eb3`
- ADR0032 — `33e68f71a2845132ff53aec30422300194b1bc40caa6dd9503848278e342ed70`

The relevant constraints are bounded/fallible resource handling, unchanged
validation and refusal precedence, source freshness and semantic equivalence,
and evidence before a performance claim. The proposed reservation is only an
opportunistic capacity hint; the authoritative per-cell reservation remains
fallible and typed.

## Current allocation path

The measured owner is `Parser::finish_cell` in
`crates/litchi-xlsx/src/raw/worksheet/codec.rs:1523-1602`. It performs all
cell checks and decoding, then calls `self.cells.try_reserve(1)` at
`1586-1588` before pushing `RawCell`. `Parser` owns this `Vec<RawCell>` in
`raw/worksheet/model.rs:168-242`; `Parser::new` starts it empty at
`codec.rs:293-318`. `finish_row` has the analogous row reservation.

The ordinary `Parser::parse` path (`codec.rs:343-394`) can use the existing
strict worksheet lane. `drive` identifies the root `sheetData` start at
`codec.rs:401-441`; `lane_body` calls `lane::recognize` and then reserves the
exact admitted lane-cell count at `codec.rs:451-481`. That path already avoids
ordinary `finish_cell` growth for its admitted body.

The source-backed shared path is different. `raw/worksheet/mod.rs:394-409`
first applies `shared_event_bound_within_cap` (`327-380`), then
`codec::parse_source_with_observer` (`codec.rs:1633-1715`) constructs a fresh
ordinary `Parser`, invokes the validator/facts observer, and calls
`parser.transition` for every event. It does not enter `Parser::lane_body`, so
strict worksheet bodies still reach `finish_cell` and grow `Parser.cells` one
record at a time. `cell_values/validation.rs:266-325` makes the validator the
first observer, then the facts builder, and falls back to the authoritative
parse on provisional failure. `cell_values/snapshot.rs:467-535` is the source
snapshot caller.

`FactsBuilder` has an exact `cells` vector, but it is populated only while the
same traversal is running (`edit/codec/snapshot/facts.rs:76-168, 410-481`) and
is published only at `finish` (`483-540`). `SourceFacts.cells.len()` in
`edit/codec/snapshot/model.rs:246-284` is therefore too late to reserve the
parser's vector. Sharing that vector would be a larger ownership and facts
publication change, not a capacity hint.

## What count is authoritative early enough

- `lane::Summary` (`lane.rs:144-161`) has exact `rows`, `cells`, and `events`
  for the strict body. `lane::recognize` (`295-303`) obtains it before replay;
  `walk` (`305-341`) counts `<row>` and `<c>` only after each tag passes the
  strict grammar. A decline leaves the normal parser path intact.
- `SourceFacts.cells.len()` is exact but post-traversal, so it cannot prevent
  the growth that is being investigated.
- The `<dimension>` range is producer metadata, not a cell count. Sparse,
  noncompact, and malformed worksheets make it unsafe as a reservation amount.
- `tools/perf-baseline/src/lib.rs:22660-22710` knows benchmark-generated
  inventories, and `build_xlsx_cell_crud_corpus` records them at
  `22879-22993`. Those counts are external fixture metadata, unavailable to a
  production parser and unsuitable as a production hint. The current fixture
  families contain 9,216 cells for the four 48x48 sheets and 17,792 cells for
  the dense/sparse shape, but that is corpus information only.

The 0959 evidence identified `finish_cell` as the nearest owner for
102,826,752 requested bytes. This is an allocator-request total from the
profile, not retained bytes or a non-overlapping phase total. It justifies a
bounded experiment; it does not establish that all of that traffic is
avoidable by one reserve.

## Smallest candidate

The smallest coherent candidate is a private, best-effort strict-lane count
hint in `codec::parse_source_with_observer`, scoped to the shared
source-backed traversal:

1. When the borrowed reader delivers a `Start` event whose local name is
   `sheetData`, capture an `lane::Entry` with the same source-offset checks as
   `Parser::drive`/`Entry::locate` before resolving and moving the event.
2. Let the existing observer and `parser.transition` run first. This keeps
   validator admission, namespace resolution, parser root/`sheetData` state,
   and their existing refusal order authoritative.
3. Only when the transition has put `Context::SheetData` on the stack and the
   resolver confirms SpreadsheetML children, call the existing pure
   `lane::recognize` over the body. Require a checked summary and
   `summary.events <= MAX_SHARED_PROVISIONAL_EVENTS`.
4. If recognition succeeds, reserve only the missing number of
   `Parser.cells` slots with `try_reserve_exact`, e.g. the checked/saturating
   difference between `summary.cells` and the current vector length. Ignore a
   failed *hint* reservation as “no hint”; the existing `finish_cell` reserve
   remains the typed, fallible allocation boundary. Do not return a new error
   from this optional optimization.
5. Continue the existing reader/observer/parser traversal unchanged. The
   summary supplies capacity only; it never supplies cells, values, spans,
   validation results, tolerance, or output.

This should be implemented behind no public API and should not add retained
facts or a cache. The same recognizer must decline prefixed or namespace-changing
body tags, entities, formulas, inline strings, comments/PI/CDATA, irregular
markup, and truncated/malformed bodies. A decline uses the existing per-cell
reservations and the existing fallback behavior.

There is no safe way to count cells by merely counting `<c` markers in the
existing `shared_event_bound_within_cap`: that lexical scan deliberately
over-counts markup and cannot distinguish comments, CDATA, attributes, or
namespaced cells. Turning that into an exact production reservation would risk
large over-reservation and semantic divergence. Reusing `SourceFacts` would be
too late. A second strict walk is therefore the only currently identifiable
exact pre-count, with its cost measured rather than assumed away.

## Error, limit, and resource boundaries

The hint must not run before the initial `sheetData` start has passed the
ordinary transition. It must not bypass the validator, MCE equivalence
observer, event counter, source byte cap, or final `complete_source_parse` /
authoritative fallback. A hint failure must not turn a later malformed cell,
limit refusal, cancellation, or reader failure into a new earlier allocation
error. A successful optional allocation can occur before a later cell error,
so the candidate should document that it changes only resource timing, not
returned errors; if a stricter allocation-order contract is required, defer the
hint until after the first fully validated cell, accepting a small extra parser
state. Do not use an infallible reserve, an unchecked count multiplication, or a
public count parameter.

The strict recognizer is bounded by the already-admitted source size and the
shared event cap, but its `Summary` is not itself the authoritative parser. The
candidate must retain the ordinary parser's event/limit checks and use checked
count arithmetic. `try_reserve_exact` failure must fall back to ordinary growth;
there must be no silent semantic fallback from a required allocation. This
separation is why the hint reservation is non-authoritative while
`finish_cell` remains unchanged.

## Required trial and adoption gates

No production change is justified by the 0959 owner ranking alone. The
isolated trial should compare hint off/on on the same source bytes and build
settings, with route admission reported explicitly:

- strict valid medium and dense/sparse source bodies;
- strict bodies with sparse rows and declared-dimension mismatch;
- non-strict formula, inline-string, entity, comment/PI/CDATA, prefixed or
  namespace-changing, extension/MCE, and malformed bodies;
- provisional event/byte-cap, observer refusal, reader failure, and
  authoritative-fallback cases;
- source-backed freshness, cancellation, shared-string/style/scalar checks,
  tolerance report, semantic store equality, output bytes, reopen, and inverse
  checks.

Record recognizer time/bytes separately from parser time, `finish_cell`
allocation calls/requested bytes, total latency, and peak/requested memory.
The output and typed error text/order must match the hint-off control. Include a
route census; do not infer that the benchmark's generated cell inventory
actually reaches this strict lane. Only after the strict-route hit rate and
net allocation/latency result are measured should adoption be considered.

## Decision

Proceed, if the next isolated candidate is wanted, with the private strict-lane
hint above. It is materially narrower and more source-grounded than a dimension
hint, benchmark metadata plumbing, or a FactsBuilder ownership rewrite. Keep
production unchanged until the route census and semantic/error differential
trial demonstrate a net benefit. This review itself made no source edits,
builds, tests, or captures.
