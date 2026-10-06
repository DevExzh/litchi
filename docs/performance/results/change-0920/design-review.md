# 0920 — owned CFB output reservation design review

Historical prototype review from 2026-10-05. Each custody table and use of
“current” below refers to its own review boundary, not the final packet.
Later tooling changes supersede those runner/control hashes. The current
source conclusion is in [source-review-final.md](source-review-final.md);
final capture readiness must bind the final tooling and controls separately.

This is a read-only source review of the isolated candidate at
`/tmp/litchi-performance-0920/source`, compared with the production checkout
at `/home/zhuhe/code/litchi`. Relevant source hashes were first captured at
`2026-10-05T19:55:35Z`; a custody recheck at `19:57:20Z` found the source
hashes unchanged but Cargo had regenerated the candidate lockfile.
Production `HEAD` was `bf88f1da5f15ebf89b32709c3537b19519d89f2d`. No source,
lockfile, benchmark, capture, or production artifact was changed by this
review. No tests were run by this review, and no performance or adoption
conclusion is made.

## Reviewed files and hashes

The relevant candidate diffs are only the intended CFB core/layout and common
codec changes, plus the new owned-output integration test. The editor,
snapshot, and three reviewed manifests are byte-identical to production. The
candidate lockfile is an isolated, untracked workspace artifact whose contents
changed when the focused source closure was expanded.

| Candidate file | Bytes | Candidate SHA-256 | Production SHA-256 |
| --- | ---: | --- | --- |
| `crates/litchi-cfb/src/writer/core.rs` | 122,703 | `eac569784deb8db0e774355cac5f1001a5f3f19b96986b0ce131b179341db9e9` | `a5ead3dded13baa8bdaa350742bfe79185e39cc2f16f86e433701e939fb2b19c` |
| `crates/litchi-cfb/src/writer/layout.rs` | 96,624 | `c8d33cfab4a4c38d05395dfe665553eea88f364e4a7a5486c6b8e64730037233` | `480b30fb1c6c018c3182fc5e0ab1f53bb55f61169396e2f2d8bce257c7d1460c` |
| `crates/litchi-ole-common/src/object/codec.rs` | 52,878 | `24db2bec57e06627f31c2c68ea9f414b01fa95dfb3c8bbd6673b7d9045c47bd9` | `99dc73a358cbe5a76c6fc242e42af6524ba6d0954c2ae508cf55f84cb6a5cafd` |
| `crates/litchi-ole-common/src/object/editor.rs` |  — | `ee6079567bf85a0ff68a335fb43999d780fab335a09a9cec8b845c16b7b5b319` | same |
| `crates/litchi-ole-common/src/object/snapshot.rs` |  — | `a4456bee02288d57dfcd26c8ec9be84ccf981aacf9faf5f30b9c3e846907a847` | same |
| `crates/litchi-cfb/tests/owned_output.rs` | 2,240 | `43710312704e19ebf595c9bc77ac3cffc0585f4ccacfccb7a917d2c83d0f232c` | candidate-only |
| `Cargo.lock` (expanded sparse workspace) | 53,596 | `24b4d150878bdfa8a6ba19f214475ebc63191096e19bd85da7085aa3f3273109` | `cc7d3578cf7570a96351db585288209dd9510fa4872c7b3514eb896d6b4500dd` |

The candidate `Cargo.toml`, `crates/litchi-cfb/Cargo.toml`, and
`crates/litchi-ole-common/Cargo.toml` hashes are respectively
`94a832011482c7e0dd2a706847ff66ab140315e8b78541d31746ec43fff460e7`,
`19aaa8b5521eb417ee78d189a5af9a14b698c15925c7f3fc060bd3cf00f62573` and
`059020ac93ea3cb203007efb026ab89ddba6b317bf9664f537948e38557bdb78`.
All three manifest hashes also match production. The CFB manifest hash is
recorded exactly by the command below as
`19aaa8b5521eb417ee78d189a5af9a14b698c15925c7f3fc060bd3cf00f62573`.

The hash command used at the boundary was:

```sh
sha256sum crates/litchi-cfb/src/writer/core.rs \
  crates/litchi-cfb/src/writer/layout.rs \
  crates/litchi-ole-common/src/object/codec.rs \
  crates/litchi-ole-common/src/object/editor.rs \
  crates/litchi-ole-common/src/object/snapshot.rs \
  crates/litchi-cfb/tests/owned_output.rs Cargo.lock Cargo.toml \
  crates/litchi-cfb/Cargo.toml crates/litchi-ole-common/Cargo.toml
```

The production comparison used the same command with each path rooted at
`/home/zhuhe/code/litchi`. A complete source diff was checked with
`diff -u /home/zhuhe/code/litchi/<file> /tmp/litchi-performance-0920/source/<file>`.
The custody recheck repeated the candidate `sha256sum` and `stat` commands;
only `Cargo.lock` changed between the two observations.

## Findings

### CFB ownership, ordering, and the accepted-plan invariant

`OleWriter::write_to_vec` delegates to the private generic helper in
`crates/litchi-cfb/src/writer/core.rs:1649-1716`. For an accepted reuse plan,
the order is:

1. validate stream sizes and build the plan;
2. validate the composed plan through the ordinary CFB reader and stream
   comparer, without writing to the caller's sink;
3. compute the checked `ReusePlan::output_len`;
4. convert to `usize` and call `Vec::try_reserve_exact`, mapping failure to
   `OleError::Allocation { resource: "CFB owned output", source }`;
5. seek to zero, emit, and flush;
6. verify the owned sink's final position equals the checked length; and
7. publish `last_sector_layout` only after that verification succeeds.

The candidate revision moved the length check into the helper's post-emit
`verify` callback. This closes the earlier publication-order problem: a
defensive length mismatch now returns `InvalidData` without publishing a
successful layout report. Allocation or seek/emit/flush errors likewise leave
the previous successful report in place.

`ReusePlan::output_len` became `pub(super)` only so the CFB core can call it;
its arithmetic is unchanged. It returns `header_sector.len() +
sectors.len() * sector_size` with checked conversions. `ReusePlan::emit` writes
the header once and exactly one `sector_size` for every planned sector: free
runs write zeroes, image runs are sector-sized, and stream runs write payload
plus zero padding. Thus the accepted-plan output length is proven by the
current code, and the post-emit check is a useful failure guard. The reserved
capacity is at least that length, so the owned `Cursor<Vec<u8>>` does not need
to grow during this emission. `try_reserve_exact` does not promise allocator
capacity equality; the invariant is no growth after reservation, not an exact
capacity value.

The `usize` conversion failure is deliberately `InvalidData` (the checked
length is not representable on the host), while allocator failure retains the
typed allocation source. This distinction is appropriate and should remain
covered if the API is changed. The typed output reservation is currently only
on the accepted-reuse path. A from-scratch fallback still writes into the
existing `Cursor<Vec<u8>>`, whose `Vec` growth is infallible; the internal
builder reservations remain typed, but a fallback output-allocation failure
is not converted to `OleError::Allocation`. That is consistent with the
stated reuse-only reservation scope, but the API documentation should keep
that scope explicit (or the fallback needs a fallible owned sink before any
claim of typed failure for every owned render).

### Fallback and sink behavior

`Outcome::Declined` and the explicitly classified plan-validation errors still
enter the existing from-scratch writer. The reservation callback is not called
on that path, so `write_to_vec` intentionally uses a growing `Vec` for
fallbacks. This is reuse-only preallocation. The fallback writes only after its
builders and checked layout work complete, flushes, and then publishes a
declined `last_sector_layout` report. `write_to` supplies no-op reservation and
verification callbacks, preserving the generic seekable-sink path.

No partial candidate from a rejected reuse plan reaches the sink: plan
validation precedes the first seek or write, and a decline starts the
from-scratch path. A failure during the existing from-scratch emission can
still leave a partial generic sink, as before; the owned vector is dropped on
error. This is compatible with the common editor's failure-atomic publication.

### Common codec integration and required validation

`Package::render_with_layout` in
`crates/litchi-ole-common/src/object/codec.rs:639-720` replaces its local
`Cursor<Vec<u8>>` plus generic `write_to` call with `writer.write_to_vec()`.
There are no candidate diffs in `editor.rs` or `snapshot.rs`; their existing
routes still choose copy-through for eligible equal-length edits and otherwise
call `render_with_layout`.

The editor's existing `commit_candidate_with_rendered` sequence remains:
`Package::check`, render, `OleFile::open` on the rendered bytes,
`codec::open`, `Package::capture`, stream-allocation reuse, and target
discovery before candidate publication. The replacement therefore moves the
owned-vector allocation into CFB, but does not remove the common CFB reopen or
semantic validation. `Package::render` for selected objects also receives the
new owned path, with its existing object-size check unchanged.

### Test coverage gap

`crates/litchi-cfb/tests/owned_output.rs` now exercises byte equality for a
no-source from-scratch render, an adopted source whose owned render asserts
`reused_source_layout()`, and an adopted source with a new stream. The latter
does force a directory-shape decline in the planner, but it checks only byte
equality; it does not assert `fallback() ==
Some(SectorLayoutFallback::DirectoryShapeChanged)`. The test would therefore
still pass if the case unexpectedly took a different fallback route. Add that
report assertion, and retain a check that the accepted route's report is not
published when post-emit verification fails if a controllable failure hook is
introduced. No typed allocation-failure or capacity-growth test is present.

The available `/tmp/litchi-performance-0920/cargo-check.log` is a sparse
workspace receipt timestamped `19:48:08Z`; the final candidate `core.rs` and
owned-output test were modified at `19:53:10Z` and `19:54:19Z`. It is therefore
stale for this review. A later external receipt,
`/tmp/litchi-performance-0920/receipts/focused-11.*`, records the command
`cargo test --offline -p litchi-cfb --features write --test owned_output`, exit
0, and stdout listing three passing tests. I did not run or replay that
command, so I treat it as a receipt claim rather than independent validation;
it used the sparse lock that Cargo regenerated and does not cover the common
codec or the standalone parity build.

### Dependency and build parity blocker

The initial archived candidate source contained only `litchi-cfb`,
`litchi-core`, `litchi-codepage`, and `litchi-ole-common`; its generated lockfile
was version 4, 102 packages, and 23,438 bytes. The source was later expanded
with an unchanged focused probe dependency closure (the current root copy has
11 crate directories), producing the final sparse lockfile shown in the table:
version 4, 229 packages, and 53,596 bytes. Production has the same root
manifests but the full workspace lockfile is version 3, 584 packages, and
138,465 bytes. The provenance receipt records both sparse lock hashes and the
unchanged standalone probe lock (`abd9f47a5aaa973a760cd61392de79c2b6143026c300dc1329831652aa54c06c`,
33,633 bytes). The final sparse lock is retained outside the repository at
`/tmp/litchi-performance-0920/sparse-quality/Cargo.lock`.

The sparse root lock still does not establish full-workspace dependency or
feature parity. For example, its `serde_json` entry omits `indexmap` while the
full lockfile includes it, and the full workspace carries both `quick-xml`
0.38.4 and 0.41.0 while the sparse graph carries only 0.41.0.

The candidate `Cargo.toml` has the same workspace member glob as production,
but the missing source directories change the resolved graph. The sparse check
also warns that the `litchi-py` release profile package is absent. Do not use
this root lockfile or receipt as production build parity. Root's proposed
parity route is the standalone 0917 probe workspace: build baseline and
candidate from the same ten-crate probe closure, unchanged probe manifest and
lock, source-bound `bb19ca74bd52` baseline, explicit profile settings, and
package-level versions, checksums, features, compiler, and allocator receipts.
That route is appropriate for the DOC comparison; the sparse root check is
not a substitute. A read-only metadata comparison now verifies the candidate
standalone probe graph: the candidate probe manifest hash is
`0e6f34cc57d56b2725034d4ed067e06a9a28d571f59348d0412c3c6966ea157d`, the
probe lock hash is `abd9f47a5aaa973a760cd61392de79c2b6143026c300dc1329831652aa54c06c`,
and both match the baseline receipt. Running

```sh
cargo metadata --offline --locked --manifest-path \
  /tmp/litchi-performance-0920/source/docs/performance/results/change-0917/probe/Cargo.toml \
  --format-version 1 > /tmp/litchi-performance-0920/candidate-probe-metadata.json
```

and normalizing package identity, dependency kinds, and resolved features
produced 144 nodes with an exact normalized-JSON match to
`/tmp/litchi-performance-0920-baseline/normalized-probe-graph.json`.
Both normalized files are 131,042 bytes with SHA-256
`437357b509d1b8f377043c4de6510302f10dfa1e6a5fc38a3662d7d54df0c6c9`.
This verifies the standalone graph and features, while the identical probe
lock binds its registry checksums. The candidate performance build still needs
its own receipt proving the same release profile/compiler/allocator and final
candidate source hash; no candidate performance binary is implied here.

Compare the candidate normalized standalone probe graph against root's retained
baseline dependency and normalized-graph receipts, ignoring only path roots,
before accepting equivalence. Shared-owner quality for XLS/PPT remains a
separate required check before broader common-CFB adoption. The sparse root
lockfile should not be committed as a production dependency change.

## Questions for the candidate author

1. Please confirm that the final post-emit `verify` callback is intended to
   cover only the owned `Cursor<Vec<u8>>` path and that fallback growth and its
   infallible `Vec` allocation remain intentionally outside the typed
   reservation/invariant guarantee.
2. Please add or explain the missing fallback-report assertion for the
   directory-shape test, and provide a receipt generated after the final
   source and test hashes above. The focused receipts include failed attempts
   caused by the incomplete source tree, network access, and a stale lock; any
   successful receipt must bind its exact command and lock hash.
3. The normalized standalone graph and probe lock now match the baseline.
   Please provide the candidate build receipt binding the same ten-crate probe
   to the `bb19ca74bd52` source, release profile, compiler, allocator, and
   final candidate source hash. Confirm the separate XLS/PPT shared-owner
   quality checks before claiming common-CFB readiness.

Static review result: the accepted-plan ordering and typed reservation design
are internally coherent after the post-emit publication fix. The fallback test
coverage and complete-workspace build parity remain open review items. This
document does not authorize a capture, production merge, or performance claim.

## Addendum — final candidate custody and comparison audit

The preceding hashes and test-gap note are the `19:55:35Z` review boundary.
The candidate was edited after that boundary. At the final source recheck, the
four changed implementation/test files were:

| Candidate file | Bytes | SHA-256 |
| --- | ---: | --- |
| `crates/litchi-cfb/src/writer/core.rs` | 124,854 | `4c4b3c38533694a92e9a2d2dc4b119ef427f6ed5633f110e2d30b3fb74b4d368` |
| `crates/litchi-cfb/src/writer/layout.rs` | 96,624 | `c8d33cfab4a4c38d05395dfe665553eea88f364e4a7a5486c6b8e64730037233` |
| `crates/litchi-ole-common/src/object/codec.rs` | 52,878 | `24db2bec57e06627f31c2c68ea9f414b01fa95dfb3c8bbd6673b7d9045c47bd9` |
| `crates/litchi-cfb/tests/owned_output.rs` | 2,516 | `e1c2b74059d7efd9d9854fb2b2b987a758dc9c356a8f73f4429a4880ef024127` |

`editor.rs` and `snapshot.rs` remain byte-identical to production at
`ee6079567bf85a0ff68a335fb43999d780fab335a09a9cec8b845c16b7b5b319` and
`a4456bee02288d57dfcd26c8ec9be84ccf981aacf9faf5f30b9c3e846907a847`.
The final-quality source-hash receipt is the external 597-byte file
`/tmp/litchi-performance-0920/receipts/final-quality-source-hashes.stdout`,
SHA-256 `c7196c028ed5093c62868cffe32fb557beab9548fee88c019da1435987be7f9a`.
The final `owned_output.rs` now explicitly asserts
`reused_source_layout()` on the accepted plan and both
`!reused_source_layout()` and `Some(DirectoryShapeChanged)` on the declined
plan. This closes the earlier coverage gap at the current hash; the historical
gap remains recorded above. `core.rs` also contains the typed reservation
failure/pre-emission test and the post-emit verification/no-report test. I did
not rerun the Rust tests. The external final-29 through final-36 receipts have
exit files containing zero and their stdout reports the 15 owned-library, 3
owned-output, 5 emission-order, 11 layout-smoke, and 6 metadata tests plus the
common check and rustfmt check; these remain receipt claims rather than tests
run by this review.

### 0920 runner, protocol, and negative-control audit

The runner and protocol were rehashed after the final comparison edits:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `comparison.py` | 73,174 | `cd6315c56fa38a5cf08fbcbef7704dc1b3df76509896828a872076a56efd97c7` |
| `comparison-replay.py` | 2,871 | `01c7aa044d9893b4d74a9920713d62dd8376ee7ad765eee52d3519dbbcd2eced` |
| `comparison-protocol.json` | 9,256 | `2a289c4f9ca3d7c7fee9887575240fbae09a7755da0a73a960ceb43303363d4` |

The replay wrapper now names and requires the `change-0920` hierarchy; the
earlier stale `change-0918` path was corrected before this hash. Static
validation used:

```sh
python3 -m py_compile \
  docs/performance/results/change-0920/comparison.py \
  docs/performance/results/change-0920/comparison-replay.py
```

An import-only check of the current runner accepted the prospective protocol,
generated 192 native and 96 allocator schedule rows, and found 288 unique
labels. No child was launched. The runner's source-map check now requires the
candidate receipt and baseline receipt to enumerate the same files and limits
the changed set to `core.rs`, `layout.rs`, `codec.rs`, and `owned_output.rs`.
The freeze path also binds the explicit patch hash, both binary hashes, the
0917 probe tree, fixtures, validator, runner, and replay script. Child writes
are exclusive; failed children retain their attempt and block reduction; every
report is passed through the unchanged 0917 validator before pairing. The
capacity comparator removes only named allocator metrics and named capacity
paths, checks route/stage shape and `None` transitions, and requires candidate
capacities to be no larger than baseline.

The following no-capture comparator audit used the retained external baseline
allocator report
`/tmp/litchi-performance-0920-baseline/qualification/allocator-docnohf-e32-default/stdout`
as both sides of a temporary pair. It is a source-level check, not a
performance run. These mutations were rejected as expected: output hash,
source hash, a correctness/refusal flag, retention ceiling, a retained-stage
transition, and a candidate capacity larger than baseline. A valid smaller
retained-stage capacity was accepted. A sample/top capacity mismatch was
rejected by `validate_probe_report` before comparison. This confirms the
negative-control boundaries and the positive comparator control for the current
runner; it does not qualify either binary.

Two capture-readiness conditions remain explicit. The protocol stays
`status: prospective` with `candidate_source_head: null`, but `run` does not
itself reject a prospective protocol after `freeze`; `freeze` also does not
accept reviewer disposition or qualification receipts as arguments. Root
authorization and the final review record must therefore remain a hard gate,
or the runner needs an explicit closed/authorized state check before any
child is launched. In addition, the runner binds source/patch and binary
hashes but does not inspect compiler, profile, flags, allocator, or normalized
dependency fields in the build receipt. The external final build/parity
receipt must establish those fields before capture. No capture was launched by
this review, and this addendum does not authorize one.

## Final controls and packet-size recheck

The comparison artifacts were edited again after the preceding runner table.
The current custody is:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `comparison.py` | 74,047 | `4641c153c14d1e9909749e3da499c75946b1a99c588064602d69e054f059cd6b` |
| `comparison-controls.py` | 9,019 | `db593a8d65b47b282ca7a05832ac800f094e1c94d2f8b4fd48934604480350ac` |
| `comparison-replay.py` | 7,256 | `d6102ed795286dab59c149b25c3073e7b2d073f6c22f3d4427d5dedd67df56b6` |
| `comparison-protocol.json` | 9,500 | `2bb0ba457db283ba0d3a6673aa5ac34cbcbf5798e41605b36a8dcf6b213d69a7` |
| `capacity-controls.json` | 2,539 | `d11fdf0cb91d6c9df0673e35c26a708ac1ef6f71eb6280fdb3694efa964da7db` |

The control receipt is hash-bound to the current runner, control script,
protocol, and 0917 validator. Its `performance_children_run` is `0`; its
positive capacity-only pair is accepted; all seven listed negative controls
are rejected, including direct allocator `retained_bytes` mismatch; and the
planned 192/96 compact capacity serialization is 53,713 bytes. The external
regeneration command was:

```sh
PYTHONPYCACHEPREFIX=/tmp/litchi-performance-0920/pycache-review \
python3 docs/performance/results/change-0920/comparison-controls.py \
  --report /tmp/litchi-performance-0920-baseline/qualification/allocator-docfloat-e32-default/stdout \
  --output /tmp/litchi-performance-0920/capacity-controls-final-review.json
```

That external receipt is 2,539 bytes with SHA-256
`d11fdf0cb91d6c9df0673e35c26a708ac1ef6f71eb6280fdb3694efa964da7db`. The
source-level check also compiled the current runner, controls, and replay
scripts with bytecode directed outside the checkout, then imported the runner
and confirmed 192 native, 96 allocator, and 288 unique schedule labels. No
performance child was launched.

For the current `semantic_signature`, the removal whitelist is explicit:
top-level and sample `elapsed_ns`; allocator metric keys
`allocated_bytes`, `deallocated_bytes`, `allocation_calls`, `peak_live_bytes`,
and `retained_bytes` from the top/sample allocation objects; top-level
`output_capacity`; route `output_capacity`,
`final_retained_before_commit`, and stage `before`, `after`,
`after_release`; and sample `output_capacity` and those same stage fields.
Unknown allocation keys remain in the semantic signature. Capacity shape,
`None` transitions, and every admitted numeric field are then checked by the
capacity comparator with a candidate-no-growth rule. This is the current
whitelist reviewed against the runner hash above.

The protocol now lists the control receipt and sub-MiB population summary as
freeze requirements, but `comparison.py freeze` still has no control-receipt
argument or verification path. Before any freeze, root must verify and retain
the hash-bound `capacity-controls.json`, or the runner must be extended to
enforce that requirement. The protocol remains prospective and the current
runner does not turn that status into a runtime capture refusal after freeze;
root authorization remains required. The candidate build/parity receipts and
the separate shared-owner XLS/PPT quality review also remain outside this
read-only source audit. No capture, production merge, or adoption is
authorized by this document.

The broader external CFB receipt is also now available for custody review:
`/tmp/litchi-performance-0920/receipts/final-full-cfb.command` is 193 bytes,
SHA-256 `dabce75493ffd2d29c0ad572844d4c5bf2b75512ac58edee89fa20ab73d62992`,
and records the offline `cargo test --offline -j1 -p litchi-cfb --features
write` command with the final-quality target and profile overrides. Its exit
file is 2 bytes, SHA-256
`9a271f2a916b0b6ee6cecb2426f0b3206ef074578be55d9bc94f6f3fe3ab86aa`, and
contains zero. The 47,326-byte stdout (SHA-256
`48fbcbd7b7d035f9d72a30548cee30056cea2394275a17cd36c246e5bf6a5b4a`) reports
540 unit/integration passes and 14 doctest passes with two ignored doctests;
the 1,290-byte stderr is retained externally (SHA-256
`891863d6aa04923b5648e224d9bbf3278cbff5f5c4e2e5162f7eb84bf3453a20`). These
are inspected receipt claims, not tests rerun by this review.
