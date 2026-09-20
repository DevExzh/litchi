# Independent review: 0721 DOCX structural scan fusion

This review covers the staged candidate under `candidate/` and the corrected
implementation during candidate qualification, based on the unchanged
`cae79bd77277921b60d5bc86fb5245da30397565` source snapshot. The individual
qualification-source hashes and quality/performance receipts are recorded below. It
makes only the bounded paired-pilot decision described below; it makes no
general hardware, RSS, cold-cache, throughput, or scaling claim.

## Decision

The fusion boundary is semantically viable as a correctness mechanism, but the
candidate is rejected for retention in this performance pilot. The candidate keeps the alt parser as the first state machine on every event,
keeps the range state independent, and defers only range-side failures. The
final writer body-capture reader is unchanged. No source-level MCE union,
unsafe code, or new unbounded source storage was introduced.

The corrected candidate source passes the DOCX formatting, all-target test, Clippy,
doctest, and rustdoc receipts currently present, including the frozen fusion
matrix. The exact public oracle and trace receipts now confirm the candidate's
observable result and one-reader mechanism, and the release harness passes its
available tests. The terminal paired analysis has deterministic output parity,
but its acceptance gates fail; the disposition restores the baseline source.

## What was checked

`AltScanState::observe` is a mechanical extraction of the former `alt::scan`
event handling. It retains the alt depth, pending anchor, opaque subtree and
properties state, `MAX_XML_DEPTH`, `MAX_CHUNKS`, relationship and `matchSrc`
validation, and the original error strings. `scan` still constructs the
sorted BTreeMap offset vector and performs its one `active` call exactly as
before.

`WordElementRangeScanner::classify` and `finish` retain the former range
scanner's independent total depth, capture depth, fragment-prefix namespace
heuristic, node count, target suppression, range conversion and callback
ordering. The ordinary `scan_word_element_ranges` wrapper now feeds this state
from `read_resolved_event`; it releases the borrowed event and namespace before
converting the event-end offset, then finishes the state, matching the old
scanner's ordering and avoiding a borrow of `NsReader` across that conversion.

`scan_with_block_ranges` performs the raw XML preflight before the shared walk,
reads one event per loop, and runs the alt state before the range state. A
range conversion, node, capture, range-conversion, or reservation failure is
remembered while alt parsing continues. At EOF it calls the first `active`
operation with the BTreeMap's anchor offsets, filters chunks, returns that
MCE error before a deferred range error, then allocates the existing exact
block-offset vector and calls `active` a second time. The block ranges are
still source-ordered and are filtered with the same BTreeSet behavior. The
inputs are therefore still two independent vectors; they are neither unioned
nor deduplicated.

The fused reader uses the same owned event and cloned resolver sequence as the
existing alt scanner: decoder before `read_event`, `into_owned`, resolver
clone, then `resolve_event`. This is equivalent to the range wrapper's
`read_resolved_event` for the current quick-xml implementation and leaves all
namespace state in the one reader. BOM removal and the final body reader stay
at their existing writer boundary, so offsets remain BOM-adjusted.

The existing resource limits and labels remain visible: 32 MiB raw XML,
256-level alt depth, 128-level range depth, one million range nodes/semantic
values, 4,096 anchors, the two MCE offset and marked-byte limits, and the
`active document block ranges` and `active document block offsets` allocation
labels. Removing `active_block_ranges` from `document_part` and making its
reservation helper `pub(crate)` does not change the helper's limits or labels.

## Oracle and trace mechanism evidence

The baseline and candidate oracle results both exit 0 and point at the same
public report: 19 cases plus the two retained package corpora, with report
SHA-256
`bea0e3f26e446d4de32cb4aa1b2398adf82da76045fb70de49957ba4a9ac24f0`.
The baseline and candidate report bytes are identical. The oracle includes the
MCE malformed and unknown-MustUnderstand refusals, active-offset count
overflow, range-depth and empty-depth boundaries, range-before-alt error
ordering, unbound fragments, malformed tails, BOM input, and the two package
corpora. The result files bind the comparison to baseline source
`9f6de4ed3f51882f1cf53bdf0fe238df8daec38bee8b1ed47d88d3a87b62d4d7` and
candidate source
`832065f45d0aca61f2cd683336151658eac00bffc47dd3e77b74ed86b69ceffa`.

The trace analysis is a pass over 21 documents. Its receipt SHA-256 is
`6540b43c30b6728faa7ddd5d8f1ea55a40368c72d3494ed20cfaf985e8e20538`; it
binds baseline trace
`74659ad0b9bb76efa4ac9f5b9e0e29ef2d3bf159b6fef6490099d640547edb8c`,
candidate trace
`248210e658b6c027c76b4505c2d427eccc156529d296c9c5dfe275c86b6bf75e`, and
the same public report hash above. Candidate range-reader documents are an
empty set. Candidate reads total 2,525, all from the alt-side reader, while
the baseline totals 4,401: 2,525 alt reads plus 1,876 range reads. On the two
representative successful corpora, `generated-medium` is 1,410 candidate
reads versus 2,820 baseline reads, and `numbered-list` is 178 versus 356.
These are instrumented event-count observations, not timings or a speed result.
The baseline range hook runs after initial admission checks and excludes a
rejecting event on a range refusal. Its all-case total is therefore not a
complete count of attempted reader calls; the representative successful
complete-pass counts remain comparable.

The candidate observer sees 2,523 events against 2,525 alt reads. The two
shortfalls are the expected early-alt-refusal cases
`empty-at-depth-256` and `range-depth-before-bad-alt`; each has 257 alt reads
and 256 fused observer events, so the event that refuses in the alt state is
not handed to the range state. There is no separate candidate range reader.
MCE call accounting is exact across the matrix: 17 cases make two calls, one
range-error case makes one call, and three earlier-refusal cases make zero
calls. This supports the intended alt-first and deferred-range error ordering
without treating the trace as a performance measurement.

## Limits exercised and limits retained

The 0721 fusion matrix directly checks the raw 32 MiB source limit with
`MAX_XML_BYTES + 1`, the alt depth boundary at 256 and 257, the corrected
empty-event depth cases at 255 and 256, the range depth boundary at 128 and
one level beyond, the 4,096-anchor boundary, the one-million range-node
boundary, and the one-million active-offset count boundary. It also checks
typed MCE refusals, malformed MCE input, and MustUnderstand handling in the
fused error-order cases.

The production MCE configuration retains a 128 MiB `MAX_MARKED_XML_BYTES`
cap, 128 MiB processing input and output caps, 256 processing depth, 4,096
namespace bindings, 4,096 directive tokens, and 1,024 choices per alternate.
The fusion matrix does not directly reach the 128 MiB marked-byte boundary:
its direct active-offset limit fixture uses a source without an MCE namespace
and therefore takes the common fast path.
The common MCE unit test does exercise a deliberately lowered marked-byte
limit, which is useful lower-level coverage but is not a 0721 production-cap
boundary result. The fusion matrix likewise does not independently reach the
one-million semantic-value reservation: its million-node fixture refuses at
the range-node check before a callback can reserve the next value. The
semantic reservation helper, MCE processing caps, and allocation labels remain
present and unchanged.

The one-million-node, one-million active-offset, and 32 MiB tests therefore
demonstrate their named limits and refusal ordering; they do not imply that
every retained byte, semantic-value, processing, or host-allocation limit has
been driven to its production maximum by this matrix.

## Resource-schedule limitation

The candidate retains the growing range vector while the alt walk completes
and while the first MCE operation runs. The baseline had already dropped its
range-walk locals before that first MCE operation. This changes the timing of
ordinary allocations and any host allocator exhaustion. It is not evidence of
a semantic limit change, and it cannot establish byte-for-byte equivalence of
global OOM scheduling. The acceptance gate should therefore require identical
typed semantic/resource refusals and preserved labels, while recording this
host-allocation limitation explicitly; it must not invent an approval gate for
an unobservable global allocator schedule.

The new import from `alt::codec` to `parts::document_part::reserve_document_value`
creates an intra-crate module dependency cycle. It is safe at the Rust item
level and does not alter the data contract, but a neutral shared helper would
be cleaner if the project’s topology checks reject the cycle.

## Follow-up before pilot sign-off

The differential file now targets the package adapter name
`scan_alt_and_block_ranges`, and the live package supplies that adapter. The
empty-depth fixture was corrected to exercise both depth 255 and 256 with two
empty siblings plus a start/end probe. The frozen alt oracle and fused helper
agree at that boundary; the all-target test receipt records the passing case.

The frozen matrix covers both MCE input vectors and call order,
inactive/active branches, nested paragraph/table anchors, fragment and foreign
namespaces, malformed tails, limit boundaries, empty-depth behavior, and typed
relationship errors. The oracle and trace receipts above cover the BOM and
source-preserving facade path and the external differential comparison. The
terminal performance analysis below rejects candidate retention; none of these
mechanism receipts should be read as a speed claim.

The extracted `classify`/`finish` calls add two state-method boundaries to the
ordinary range scanner. Release optimization may inline them, but that is an
implementation detail to measure in the read-consumer controls; adding an
`inline` attribute without a measured regression would be speculative. The
final `-D warnings` Clippy receipt reports no borrow or lint diagnostics; any
release inlining decision remains a measurement question rather than a source
claim.

The archived first DOCX test build reported mechanical blockers in the active
result naming, the removed `BTreeSet` import, the range callback conversions,
and the EOF control. The live source fixes those issues. A later Clippy run
also identified three unfulfilled test `expect` attributes; the live fixture
removed them before the final quality run. These were source-integration
findings rather than semantic changes.

The archived receipts preserve the failed controls rather than hiding them:
`initial-quality-build/quality-docx.json` records tests exit 101 for the local
active-result shadow, callback type mismatches, and EOF control;
`initial-test-helpers/quality-docx.json` records the dead `span`/`range`
helpers; `initial-empty-depth-expectation/quality-docx.json` records 998
passing tests and one stale depth-256 expectation; and
`initial-clippy-expectations/quality-docx.json` records the three unfulfilled
test expectations. Their test-log hashes are respectively
`af8a107a3390f379353b97023b51fef65a7867a419ee7645a8eacef9200e5b5b`,
`cc693482fb5a89d0405e2c2ec6fcd5ade954eddd73e2dc7aecc9292ed19f7ab3`,
`370955ffbd00531edaf24ae6e7b7adda18554cb17fc1d39e0d7da6ba3e7f71e80`,
and `ae622bd5ee3933c2cca31f3345922fdc83faf18c9c1b62bc390f51934b836e30`.
The corrected fixture and final quality receipts supersede those failures.

The first performance-control validator invocation is also retained under
`initial-control-command-validation/`: its `capture.json` has the native
baseline command at exit 0 followed by `read-controls.py capture baseline-A1`
at exit 1. The amended capture freeze records the corrected real-file control
command and resumption of the retained baseline samples. This is provenance
for the performance protocol, not a correctness result; the terminal receipts
below supersede the partial capture state.

## Current receipts and source identity

The final DOCX quality receipts all report exit code 0:

| Check | Receipt SHA-256 |
| --- | --- |
| `cargo fmt -p litchi-docx -- --check` | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` (`quality-docx-fmt.log`) |
| `cargo test -p litchi-docx --all-features --all-targets` | `dc3a526a1eb10af0375131765b30ca8d3ce49a6b4b2408a23b44b84785470c5e` (`quality-docx-tests.log`) |
| `cargo clippy -p litchi-docx --all-features --all-targets -- -D warnings` | `a4da70788e17972bd96086e021c92e6e02be6ad0eb8aa7179f2a81f52c57dbc4` (`quality-docx-clippy.log`) |
| `cargo test -p litchi-docx --all-features --doc` | `24c2258c43df5e19ff8921f8bcac4f396052ef3df8c60398aa70613febdab226` (`quality-docx-doctests.log`) |
| `cargo doc -p litchi-docx --all-features --no-deps` | `13c7f4b48aa650452d52b0e56d049094873c37398b2511f3a7aa9f481d847989` (`quality-docx-rustdoc.log`) |

The release performance harness receipt
`quality-harness.json` is SHA-256
`0044187aa7f55ba19b36bf1623d6eb444de2fcee8fe7d3114b579428ec23c301` and
its test log is
`0503b4bf822e34d4d781c4b07dc0a848a47129caa6e6774cedbc12a878157a22`.
It reports 535 passed, 0 failed, and 1 ignored; the ignored test is the
explicitly opt-in real-producer security corpus, so this receipt does not
claim that corpus ran.

The release build receipts also exit 0 in both native and allocator lanes.
`build-baseline.json` is SHA-256
`f4a0040cc6b13fe0cbf1ce5c1fe327c06f7eb6166ea1eab8a8f8bef360276108`, with
native binary
`ae35757954671eb3b3360c4c03ee90d778f112d07b6f6aa9a21ebec9cee5bbcc` and
allocator binary
`42181df7da45d8ed88a6e7f015592300e008be6bd8a36a398ef3b59eaee164c9`.
`build-candidate.json` is SHA-256
`9932b18a0c29e0cb4670e11a02906720ddc3ad3d5e7955c5a4ac1a269f7d99c1`, with
native binary
`867a39fae840dc9cf308db595256cb2e9b023222f7275770a36a70fcbdd7cf07` and
allocator binary
`db68cf04c54b12d9370acd1b134f9eaccf0b664b124b1aad6bd9e9e99b1dd4de`.
These hashes establish artifact identity only; they are not performance
results.

The quality source manifest is SHA-256
`832065f45d0aca61f2cd683336151658eac00bffc47dd3e77b74ed86b69ceffa`.
The six candidate snapshots matched the live source during qualification at
these paths and hashes; the current checkout is the restored baseline:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-docx/src/alt/codec.rs` | `edfed929ef257d7c5aa669b255429b6e8b84491eff139e7e3fb5a17868cd802e` |
| `crates/litchi-docx/src/alt/mod.rs` | `36356c6318b37f03a19fcf0e49fc6a3352b7467447f7c26a0b927fd165e99ea0` |
| `crates/litchi-docx/src/namespace.rs` | `e3521de1356e6734cebb88f8733219386623aabbeacf89bf16f485bc6dd18f4a` |
| `crates/litchi-docx/src/parts/document_part.rs` | `5d4475096253946a29ce858ba3273aa760e53387d72543494caba3ed3d7e7e21` |
| `crates/litchi-docx/src/writer/doc/package.rs` | `43bd672511bd83406321dc99cb722d8828b170e53cf04a06fbcda67c553d72f7` |
| `crates/litchi-docx/src/writer/doc/fusion_tests_0721.rs` | `cd05b403e294fe6359c65725faf7639db91b1e4e324064ea298f02a0bfb13cdd` |

The oracle, trace, release builds and harness qualification are recorded above.
The host allocator OOM schedule remains a documented limitation; semantic
resource limits, refusal types and labels remain required.

## Coordinator terminal finalization

The completed analyses reject the candidate: 1 of 64 primary hard gates and
all 16 read-control hard gates fail. Generated paragraph-list p50 regresses
6.72–8.60%; pinned-media eager paragraph-count p50 regresses 5.59–6.05%.
The first generated edit mean improves only 0.68%, below the 3% requirement.
Its candidate mean repeat spread is 15.21%; all samples remain included.
Request-count and requested-byte gates pass, but NumberedList edit peak above
region start increases from 21,310 to 23,790 bytes in both allocator pairs.
All other aligned peak-above-start and net-live vectors are unchanged.

`source-final.json` exactly matches the baseline manifest and current checkout;
`disposition.json` records `retained: false`. All six repository evidence gates,
nine corruption checks and post-cleanup analyzer replays pass. The candidate
is preserved only as evidence. A future bounded experiment may isolate fusion
inside the writer while retaining the original read-only scanner, but this
packet establishes neither an inlining cause nor a speed promise for it.
