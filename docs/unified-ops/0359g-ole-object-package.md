# 0359g — OLE object package owner and consumer migration

- Date: 2026-10-21.
- Base: `45a5f9e9d1` plus the 0359f raw-bounds commits.
- Status: implementation and consumer migration are complete on `uo2/w5-ole`;
  the corrected candidate still requires a fresh coordinated performance
  measurement before acceptance.

The shared object owner now follows the common package contract. `Package` is
an immutable read owner; `Package::open(bytes)` uses the finite default policy
and `Package::open_with_options(bytes, OpenOptions)` accepts caller-selected
targets and bounds. `Package::edit()` and `edit_with_options(EditOptions)`
start an isolated `Edit`, and `Edit::commit()` returns the shared
`Commit<Package, Patch, Diagnostics>` triad. `Package::apply(&Patch)` and
`Patch::apply(&Package)` use the recorded source content and return the same
triad. Dropping an edit discards its staged state.

The owner retains admitted source bytes and captured stream allocations in
`Arc` values. Read projections and output methods use immutable receivers;
`write_to(&self, ...)` streams an unchanged source directly. Changed
candidates are validated internally while staging and are published only by
`commit()`; no staged-output method exposes those bytes. `source_shared()`
exposes the retained source allocation to deferred format owners, and
`OpenOptions`/`EditOptions` provide the policy getters with the required
`_value` names. DOC, XLS, PPT tests, and the OLE performance harness were
migrated to package ingress and edit/commit publication in the same slice.

## Publication handoff correction

The first owner migration exposed `put_stream_shared_with_rendered` and
`put_streams_shared_with_rendered` as staged-output bypasses. The correction
removes both public methods and migrates the DOC tracked-revision path plus
the XLS visibility, comments, cells, user-names, and worksheet-feature paths
to `Edit::commit().package()`/`Commit::into_package()`. A common `Edit` now
retains the already validated candidate render internally and reuses it when
`commit()` publishes the package, so these consumers do not serialize the
same candidate a second time. A changed sector-layout policy bypasses that
retained render and performs its required validation under the selected
policy. Failed clone-first mutations leave the prior retained render intact.

The exact `Patch` debug regression asserts structural lengths and the no-op
flag rather than merely checking that authored payload strings are absent.
The common error is non-exhaustive; CFB diagnostics remain structural and the
broader redaction audit remains the single deferred memo item.

The migration removes the former public object-editor open constructors and
staged-snapshot publication path. The raw OLE directory, OLEDS, link, and
protection bounds and typed errors from 0359f remain unchanged. Exact no-op
bytes, source policy fencing, atomic edits, reversible replay, and opaque
stream fidelity are covered by the existing object and format regression
suites; their post-migration run and ABBA receipt will be added before this
slice is accepted.

## Driver preflight

The paired standalone driver is `tools/perf-baseline/examples/ole_w5_probe.rs`
at probe source `fdb139197d`. It uses the same 64-sibling OLE corpus and six
routes for every revision: `open`, `begin_edit`, `noop_commit`,
`replace_stream`, `read`, and `edit`. The owner pair is source
`0ce4bb8b61` → `fdb139197d`; the typed-overlay pair uses the owner-migrated
source `604822b441` → `fdb139197d`. The baseline and candidate binaries were
built in distinct absolute target directories with the capped W5 environment:

| role | build source | binary SHA-256 |
| --- | --- | --- |
| owner baseline (`g`) | `0ce4bb8b61` | `7842e775c2076ffd89de2b6aa2b6c99f4ae6e2a5b4c431caaeb94552fbfdad22` |
| owner/property baseline (`h`) | `604822b441` | `f5520c923c7ec1d763c96bd58e5ad837c08dcc31ee231727fa9ad17535e82e71` |
| corrected candidate | `fdb139197d` | `44b3e0427841cbb0ba3fed146f7da67926f096f05cf266ea9f40ec71f9498700` |

The legacy binary used `--no-default-features --features
legacy-owner-api,legacy-property-api`; the `h` binary used the current owner
and legacy property API, and the candidate used the current API. One-sample
preflights passed for every route on all three binaries. Owner routes emitted
14,848 bytes and verified the original/replacement target stream; property
routes emitted six counted properties. The preflight output is retained
outside the checkout under `/tmp/w5-ole-preflight-20261021T164402Z`.

The acceptance run remains pending a quiet native grant. For each route, run
the four legs serially in `A1, B1, B2, A2` order with identical warmup/sample
arguments, using the baseline and candidate paths above. The owner routes use
`g`/candidate; the property routes use `h`/candidate:

```sh
for route in open begin_edit noop_commit replace_stream read edit; do
  "$A" --route "$route" --warmup 20 --samples 500 >"$OUT/a1-$route.csv"
  "$B" --route "$route" --warmup 20 --samples 500 >"$OUT/b1-$route.csv"
  "$B" --route "$route" --warmup 20 --samples 500 >"$OUT/b2-$route.csv"
  "$A" --route "$route" --warmup 20 --samples 500 >"$OUT/a2-$route.csv"
done
```

The four paths are expected to be kept distinct per route. Set `A` to the
appropriate baseline binary and `B` to the candidate binary; do not mix the
`g` and `h` baselines within a route. No native performance result is claimed
until this paired run is completed under quiescent host conditions.

The binaries and one-sample preflight listed above predate the publication
handoff correction and are diagnostic only. A corrected before/after pair must
be rebuilt from the exact post-correction source before any performance claim;
the route set must additionally cover ordinary `put_stream` + `commit` and the
former XLS/DOC rendered-handoff consumers through their new commit paths.
