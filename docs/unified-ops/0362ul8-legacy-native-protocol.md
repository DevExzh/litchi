# 0362ul8a — legacy native qualification protocol repair

This source-only slice closes the missing driver/manifest boundary for the
RTF refusal/replay lane and the OLE/XLS-properties lanes.  It does not contain
a Cargo build, a release executable, a native run, an allocator capture, a
timing result, or an acceptance claim.  Candidate source identities remain
parameterized until the final integrated freeze.

The corrected catalog is
`.coord/final-acceptance-2026-10-24/native-qualification-lanes-76c21c49cf.json`
(audited staging `76c21c49cfc7fedb995fb1b2c11351098eddef7b`).  Its path policy
is literal external paths from the cited records; new execution output must
use a fresh external directory.  The catalog records no tracked RTF or
OLE/XLS-properties manifest.  The manifests and validator added here reuse
the existing executable drivers and the guarded process/identity runner:

| lane | reusable executable and current source SHA-256 | source fence | protocol |
| --- | --- | --- | --- |
| RTF | `crates/litchi-rtf/examples/refusal_contract_bench.rs`, `70077fcec439fda2fdf9e9635f5db2f1f7c86ed31f3f972a95c5e5e3a607a585` (reviewed contract adapter) | original `86d708878fa537eab42fd4bd5ffbc0f79c606191`; candidate `$RTF_CANDIDATE_REVISION` | `legacy-rtf.json`, five guarded ABBA blocks |
| OLE properties | raw fdb139 probe SHA-256 `3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f`; baseline public-owner/legacy-property adapter SHA-256 `6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed`; candidate split plus CFB-limit adapter SHA-256 `e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782` | original `604822b44107999aaddb0bffb96b705e9464f546`; candidate `$OLE_CANDIDATE_REVISION` | `legacy-ole-properties.json`, `read` and `edit` |
| XLS properties/lazy entry | `crates/litchi-ops/examples/ole_metadata_bench.rs`, before `b840acb620c09644ca4b89cbe42927485da9d457f3bf93b578eff90fa1008173`, after `53b26136e915fa0bf88518fd1c05e494d55ce5e8f91012be04ac7adf476c1f10` | derived pre-lazy parent `509bf28e7531ad75f278eaec3a732714a243915a` (parent of lazy feature `4f08cdc5e8751243d50077ce1a2a2ea24a0d9c51`); candidate `$XLS_CANDIDATE_REVISION` | `legacy-xls-properties.json`, existing metadata benchmark |

The XLS source parent is a provenance fence, not a historical performance
baseline: the authoritative catalog says that no XLS native baseline or probe
exists.  The current `9fe06599d6be568218446b491312e1ce18a1e2e8` source and
focused tests remain source-only evidence.  The old RTF and OLE binaries are
also diagnostic only; they are never relabeled as outputs of this protocol.

## External inputs and source preparation

The following catalog paths are read-only historical receipts and must not be
used as fresh output destinations:

- `/home/zhuhe/code/litchi-worktrees/.coord/rtf-refusals/`;
- `/home/zhuhe/code/litchi-worktrees/.coord/rtf-refusals/native-qual-20261022/block-01-01-baseline.json` (rejected clean-window receipt);
- `/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-final-20261021/protected/base-alloc/ole_w5_probe`;
- `/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-retag-20261022/protected/candidate-base-package-edit-cf3c9a4eb6c8b1d661ee57cef238de78ed67bb29/ole_w5_probe`;
- `/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-retag-20261022/counting-base-package-edit-v2/MANIFEST.txt`;
- `/home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh`.

Prepare clean detached source worktrees outside this checkout.  Preserve the
original source HEADs exactly; qualification overlays are limited to the
reusable probe, the selected lock, and the standalone feature declarations
needed to compile that probe.  Do not copy implementation files from a later
candidate into an original source tree.

* RTF: materialize `86d708878fa537eab42fd4bd5ffbc0f79c606191`, copy the
  reviewed refusal probe from the selected candidate into
  `crates/litchi-rtf/examples/refusal_contract_bench.rs`, and provide one
  reviewed root `Cargo.lock` in each endpoint.  The copied probe is the only
  baseline probe overlay; its hash must be recorded again after the final
  candidate freeze.
* OLE: materialize `604822b44107999aaddb0bffb96b705e9464f546`, copy the
  reviewed fdb139197d08677d6fa0e14d536e798e380e4cb1 `ole_w5_probe.rs`, and
  retain its raw SHA-256
  `3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f` as
  the source pin.  For the baseline, apply only the recorded feature-guard
  split (`current-owner-api` plus `legacy-property-api`), which selects the
  public `Package::open_with_options`/`Package::edit` path while retaining the
  legacy property `with_limits` calls; the materialized baseline SHA-256 is
  `6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed`.
  For the candidate, use the same split with `current-owner-api` plus
  `current-property-api`, then apply only
  `OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default())?`.
  Its SHA-256 must be
  `e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782`;
  no other probe/workload change or substitution of the newer `ad4c52c` source
  is permitted.  Add only the standalone feature declarations
  (`current-owner-api`, `current-property-api`, `legacy-owner-api`, and
  `legacy-property-api`) to `tools/perf-baseline/Cargo.toml`.  Build the
  baseline with `--no-default-features --features current-owner-api,legacy-property-api`;
  build the candidate with `--no-default-features --features current-owner-api,current-property-api`.
  The historical final-20261021 OLE lock has SHA-256
  `bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a` and is
  provenance only; the final lane must bind the lock actually selected for
  the frozen candidate.
* XLS: materialize the derived parent
  `509bf28e7531ad75f278eaec3a732714a243915a`, retain its tracked
  `ole_metadata_bench.rs` (source SHA-256
  `b840acb620c09644ca4b89cbe42927485da9d457f3bf93b578eff90fa1008173`, as recorded
  by the source-parent checkout), and provide the reviewed root
  `Cargo.lock`.  The three tracked fixtures used by the existing benchmark
  are part of the identity: `test-data/ole/doc/FancyFoot.doc`,
  `test-data/ole/xls/Simple.xls`, and `test-data/ole/ppt/SampleShow.ppt`.
  The source-parent `tools/perf-baseline/Cargo.lock` has file SHA-256
  `4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184`, but
  the root workspace build still records the actual root lock supplied for
  the frozen paired endpoints.

Every endpoint must be a full, clean materialized HEAD tree apart from these
explicit preparation overlays.  `prepare_legacy_native_inputs.py` emits a
source-only `litchi.legacy-native-preparation.v1` receipt: it binds the
manifest and preparer hashes, baseline/candidate HEADs, exact status/diff
digests, reviewed dirty-path allow-lists, lock/probe/fixture paths and file
hashes.  The runner and validator require that receipt and compare it again
after each build/run event; a fabricated event, arbitrary hash path, or text
file marked executable is rejected.  `$RTF_CANDIDATE_REVISION`, `$OLE_CANDIDATE_REVISION`, and
`$XLS_CANDIDATE_REVISION` are required full 40-character SHAs at execution;
no historical diagnostic tip is a substitute.

## Guarded workload

All three manifests use `tools/native_qualification_runner.py`.  It injects
`--offline --locked` before Cargo's separator, keeps build targets and logs
outside the source worktrees, records every binary SHA-256, rejects stale
handoffs and source mutation, and checks the process guard after every child.
The new validator additionally checks the exact source/lock/probe/fixture and
binary records, the ABBA event order, route output shape, quiet-host status,
target size and memory caps.  It emits `performance_claim: false`.

The common execution envelope is:

```sh
export PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH
unset LITCHI_CAPPED LITCHI_BLESS
export CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0
export LITCHI_MEM_MAX=20G
```

Use a new external output, target and temporary directory for every attempt;
never reuse the historical catalog paths.  The runner's fixed limits are a
19,200,000,000-byte target preflight ceiling, a 20,000,000,000-byte target
hard ceiling, and a 20,000,000,000-byte child address-space ceiling.  Pass a
pinned CPU to the ABBA phase and stop if the process guard observes another
Litchi Cargo/compiler/profiler workload.

### RTF

The reviewed probe emits the fixed `rtf-refusal-contract-v1` workload.  It
internally asserts all eight route dispositions and the validator independently
requires their exact output/error/fidelity fields:

`open_1k`, `noop_1k`, `text_commit_1k`, `text_commit_save_1k`, `replay_1k`,
`tail_commit_1k`, `unsupported_content`, and `protected_commit`.

The manifest runs the same executable in five separate
`before,after,after,before` blocks (20 report streams total).  Each executable
run retains the probe's fixed three warmups and 31 samples and reports one
median per route.  The validator requires exactly those eight JSON rows in
every stream, positive finite medians, all block/leg identities, and the
runner's quiet/cap evidence.  `unsupported_content` must be
`RefusalKind::unsupported_content`, operation `edit`, fixed reason `underline
edits require non-empty text within one paragraph`, and `unchanged` fidelity.
`protected_commit` must be `RefusalKind::protected`, operation `edit`, fixed
reason `document protection refuses changed publication`, and `unchanged`
fidelity.  Successful routes must report `exact` fidelity.  All refusal text
remains content-free.

### OLE properties

The manifest reuses `ole_w5_probe` for exactly `read` and `edit`, with
`--warmup 20 --samples 500 --batch 1`, in one A-B-B-A block.  The reviewed
fdb139 probe owns the six-property generated corpus, link-base/hyperlink read
checks, edit publication and exact section-output check.  The validator binds
raw probe SHA `3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f`,
baseline adapter SHA `6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed`,
candidate adapter SHA `e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782`,
source fence `fdb139197d08677d6fa0e14d536e798e380e4cb1`, the exact feature-split
and one-line CFB recipes, fixture ID `ole-w5-generated-corpus-v1`, exact route
argv, and output size `3` for both `read` and `edit`.  It requires the exact
CSV header, 500 ordered rows and all eight streams (two routes × four legs);
nonzero child/error or fidelity failures remain fatal.

### XLS properties/lazy entry

The manifest uses the existing `ole_metadata_bench` rather than the unrelated
cell-history routes.  It builds `litchi-ops` with `--no-default-features
--features doc,xls,ppt --example ole_metadata_bench` and runs one A-B-B-A
block.  The driver constructs the tracked DOC/XLS/PPT fixtures and performs
the native `Package::edit_metadata` property write beside the unified
metadata write; its byte-equivalence assertion remains inside the timed
driver.  The validator binds fixture ID `ole-metadata-fixtures-v1`, exact
fixture hashes and sizes (DOC 27136, XLS 13824, PPT 125440), and requires
ordered `doc`, `xls`, and `ppt` JSON rows with exactly 50 repeats and those
expected input byte counts.  It also binds all source/fixture identities.
This lane is explicitly a source-parent comparison, not a claim that a
historical XLS native latency baseline exists.

## Allocation and fidelity boundary

Normal timing binaries are not allocator-instrumented.  Allocation evidence
must be run as a separate fresh phase with the reviewed counting/heaptrack
driver and must never be mixed into the elapsed-time rows.  For RTF this is a
fresh five-block heaptrack path under the RTF external lane; the existing
`rtf-refusals` captures are rejected diagnostic receipts.  For OLE the
historical counting manifest and protected binaries under the catalog paths
are attribution-only; prepare fresh protected binaries and preserve the
read/edit plus XLS residual checks before any claim.  XLS lazy-entry has no
paired native allocation baseline in the authoritative record, so its
allocation phase remains a coordinator prerequisite.  The validator records
`separate_phase_required: true` and `elapsed_claimable: false` for this
boundary.

No source-only focused-test receipt, old binary hash, allocation count, or
historical heap/callgrind result is promoted to native acceptance.  The final
freeze must bind the source SHA, lock SHA, probe/fixture hashes, binary hashes,
fresh output paths, CPU, and quiet/cap receipts together.

## 0362ul8a return coverage

This return preserves every earlier repair: the reusable RTF/OLE/XLS manifests
and workload counts; source-preparation receipts with allow-listed overlays;
manifest/runner/journal/argv/cwd/cap/target/binary provenance; exact RTF
refusal kind, fixed reason and fidelity; strict OLE output bytes and XLS
fixture IDs; and raw-path symlink/lock binding rejection.  It additionally
separates the unchanged fdb139 OLE source
(`3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f`) from
the baseline public-owner/legacy-property feature adapter
(`6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed`) and
the candidate's single consumer-only `OleFileLimits::default()` addition
(`e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782`).  No
production overlay, workload change, or newer `ad4c52c` probe is accepted.

## Build-free validation performed here

No qualification Cargo build/release, native, allocator, heaptrack, or
profiler workload was run in this preparation lane.  The following checks
exercised only manifest,
validator, identity, route-shape and cap logic:

```text
python3 -m unittest tools.test_legacy_native_qualification tools.test_native_qualification_runner
Ran 35 tests
OK

python3 -m unittest tools.test_no_3rdparty_references
Ran 3 tests
OK
```

The Python modules compiled successfully with `python3 -m py_compile`; all
three manifests parsed as JSON and loaded through the strict runner with
parameterized source/candidate paths.  No external output or executable was
created by this commit.  The shared `python3 tools/check_feature_matrix.py`
helper was also attempted as a repository diagnostic; it launched 31
`cargo check --locked` commands, all exit 101 before compilation because this
source-only worktree has no root `Cargo.lock` and the helper omits `--offline`.
Those failures are not qualification evidence and no retry was made.
