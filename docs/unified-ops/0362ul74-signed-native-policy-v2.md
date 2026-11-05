# 0362ul74 — signed-native qualification policy v2

Date: 2026-10-25.  This change records the next source boundary for the
signed-native qualification.  It is source-only: it does not run Cargo,
rustdoc, a native process, or a timing/allocation workload.  The historical
0362ul24 receipts remain v1 evidence and are not relabelled.

## Current disposition

The final candidate source commit was not frozen at this checkpoint.  The
checked-in policy therefore has status `pending_final_source_freeze`, and its
candidate binding is deliberately `None` in
`tools/native_qualification/prepare_signed_native_inputs_v2.py`.  The
`after` preparation/build/execute path fails before creating output or
reading a candidate source tree; the independently pinned `before` f2
assembly remains available for its compatibility preparation.  There is no
command-line digest, commit, archive, or JSON override that can turn this
incomplete candidate policy into qualification evidence.

The three known candidate lock identities are recorded as role-specific
inputs, not as source commit identities:

| role | candidate lock SHA-256 | bytes |
| --- | --- | ---: |
| workspace `Cargo.lock` | `02fda2543226f28c2ded9bbf7e18618f8b0aab1c4d02c7b2c4fdd73511e4ff7a` | 139,207 |
| perf `tools/perf-baseline/Cargo.lock` | `de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea` | 56,941 |
| native `tools/native-resave/Cargo.lock` | `989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae` | 56,689 |

The f2 baseline role locks are also pinned: workspace
`Cargo.lock` is `1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c`
(139,441 bytes), perf `tools/perf-baseline/Cargo.lock` is
`57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f`
(57,028 bytes), and native `tools/native-resave/Cargo.lock` is
`6aaa0e3800a4db72afbe8783aab382accdb4a6355d06bb97fb1c12dd78839408`
(56,929 bytes).

The root lock is currently absent from the sparse candidate leaf.  A later
freeze must materialise it and prove all three role locks, source manifests,
and the complete lock graph.  The `allsorts`, `flate2`, and `zlib-rs`
registry-to-path transitions are explicit in the manifest; no other vendor
root is admitted.

## Binding rules

Before preparation can be enabled, the same reviewed change must bind all of
the following:

* a full 40-character candidate commit, its Git tree id, a canonical tracked
  tree digest, and an external source archive with recomputed byte count and
  SHA-256;
* the three source manifests, role locks, and their dependency/lock-graph
  identities;
* every tracked regular blob below `vendor/allsorts`, `vendor/flate2`, and
  `vendor/zlib-rs`, including mode, Git blob id, byte count, and SHA-256;
* the reviewed toolchain/configuration, selected `litchi-ops` feature set,
  probe/driver/validator/memory/final-collector inputs, and the signed fixture
  payloads;
* the exact route and format lists from the current signature policy tests;
* the owner’s signed-batch choice, which is `immediate_refusal` with no
  iterator reads or publication attempt.

The verifier enumerates the candidate Git tree and reads every bound vendor
file.  A sparse checkout, missing file, symlink, modified blob, extra vendor
root, changed fixture payload, or changed probe input fails closed.  The
archive is checked as an external artifact and no generated archive or binary
is committed.

The fixture manifest intentionally records signature *presence* only.  Its
`cryptographic_verification` field remains `not_performed`; synthetic OPC/CFB
signature graphs are not presented as producer certificates.

## Route semantics retained

The policy carries all seven formats (DOCX, PPTX, XLSX, XLSB, DOC, PPT, XLS)
and all sixteen leaf/facade routes: exact no-op commit/write/replay,
protected refusal, explicit unsign-and-write, unsigned negative,
source-freshness, and budget.  The route contract also records the existing
cancellation rule: a cancelled intermediate edit remains signed and does not
publish.  The owner-confirmed signed-batch behavior is immediate refusal
before the iterator is touched; this is not a timed or native acceptance
claim.

The expected typed dispositions are unchanged: `ProtectedDocument` /
`document_protected`, `PatchConflict` / `invalid_input`, `SourceChanged` /
`io_error`, and `LimitExceeded` / `limit_exceeded`.  Successful no-op/replay
routes require unchanged diagnostics, a no-op/content-bound patch, signed
state, and exact source/output bytes.  Unsign-and-write requires absent
signature parts, a real semantic change, bounded streaming, and a signed clone
that still refuses edits.  These are input contracts for a future run, not
timed or native acceptance claims.

## Source-only checks

Run the focused tests from the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_signed_native_policy_v2
```

The suite checks the pending fail-closed boundary, refusal to accept an
arbitrary candidate digest, rejection of old v1 receipts, fixture payload and
signature-presence policy, route/cancellation coverage, and sparse/modified/
symlinked vendor trees.  It passes ten tests without invoking Cargo or a
native executable.

After a separately reviewed source-freeze change supplies the complete
binding, the owner must rerun this suite, then run `prepare` against two
distinct immutable source roots.  Only after the preparation receipt, signed
correctness routes, bounded native timing, and external per-fixture memory
evidence have independently passed may a final collector qualify the run.
