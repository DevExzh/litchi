# 0359g — OLE object package owner and consumer migration

- Date: 2026-10-21.
- Base: `45a5f9e9d1` plus the 0359f raw-bounds commits.
- Status: implementation and consumer migration are in progress on
  `uo2/w5-ole`; validation and coordinated performance measurement remain
  pending.

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
`write_to(&self, ...)` streams an unchanged source directly, while changed
output is rendered only at publication. `source_shared()` exposes the retained
source allocation to deferred format owners, and `OpenOptions`/`EditOptions`
provide the policy getters with the required `_value` names. DOC, XLS, PPT
tests, and the OLE performance harness were migrated to package ingress and
edit/commit publication in the same slice.

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
