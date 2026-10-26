# 0362ul165l — Final-source admission and test returns

The remaining-29 diagnostic continues on Q472a. ODS Clippy, the full ODS
suite and workspace Clippy pass. The leaf suite reaches DOCX and reports
1,525 passing library tests and two failures; its later leaf chunks are
unrun. This is a code/test failure, with unchanged source and locks and no
resource failure. The driver preserves it and proceeds to later checks.
Neither this diagnostic nor the earlier failed C33 run qualifies the final
all-65 gate.

## DOCX metadata assertions

The two failures are stale assertions. `Package::new()` deliberately contains
an empty core-properties part, so its metadata is `Some(Metadata::default())`.
The app/custom fixture preserves that part. The other assertion incorrectly
expects a core-properties entry in the exposed `metadata::Patch::parts()`;
that projection covers app/custom changes, while the core transition is
represented by `Patch::base()` and `Patch::target()`.

The isolated `uo2/docx-metadata-repair` leaf corrects these assertions in
`dbebbf9b81`, then strengthens the existing regression in
`ec243ebb44fc5df2cb939f780ac81135a322b231`. The test reopens the serialized
empty core part, replays onto an independently reopened cleared package,
and retains inverse checks for absence and exact original bytes. The
app/custom test compares source metadata with its captured initial value.
Production code is unchanged. Formatting and diff checks pass; focused Rust
validation remains pending while the diagnostic owns the Cargo slot. The
leaf is not merged at this checkpoint.

## Actual final-C admission

A cross-protocol audit confirms that the final source cannot be identified
by merely relabeling C33 or Q472a. Those sources have the old 56,941-byte
performance lock; merged tooling uses the 57,961-byte `89ee807d…` lock.
Signed and DrawingML generation check the candidate's actual lock and also
pin its commit/tree. Their fresh successors must bind the selected final C
and the matching P policies together.

ODF raw preparation has an additional executable mismatch: it requires an
absent candidate probe and the old pre-overlay lock. Trunk now has the
reviewed probe, manifest and repaired lock committed. The required successor
must verify these files in a clean final C without writing candidate files.
Its external identity belongs to P, avoiding a commit that contains its own
hash. The historical 70d baseline still receives its reviewed consumer
overlay and separately resolved lock. The 3401 adapter supplies consumer-file
provenance; its entire production tree must not be equated with a later C.

The new `uo2/odf-final-source-admission` leaf commits that migration as
`6f6e752e857519deed18658741f79f8130ffa14c`, changing the preparer, paired
runner, manifest and tests together. Twelve source tests and AST/JSON/diff
checks pass. Independent review returns the leaf for a missing actual P Git
trust anchor, ignored adapter-file admission, external identity drift after
initialization, incomplete row/timing validation, assume-unchanged index flags
and a report-read symlink gap. These require a successor before merge.
Actual final-C selection,
P installation and rebinding of all 28 qualification rows remain required;
the existing C33 packets are preserved.

## Protocol composition phases

The v5 preview carries all 22 files changed by the ODF leaf: 17 are byte
identical and five incorporate the reviewed DrawingML/signed migrations.
Its eleven alleged conflicts are actually unchanged C33 input blobs. Root
records that comparison separately. The reported 46 tests ran on the active
ODF leaf, not the composed preview.

Independent review reproduces a real phase mismatch: v5 includes the signed
P2 bound-policy test while its production binding is still pending. V6
separates P1 and P2, but its new installation guard expects a binding schema
the real generator never emits. V7 consumes the real producer shape and
installs all three consumer files, but is returned for unpinned patch bytes,
incomplete archive/source linkage and install-path symlinks. These are not
accepted exceptions.

V8 preserves the latest 3401 signed-policy test in P1, regenerates P2 against
that exact base, pins patch/base/target identities, uses the reviewed
finalizer's provenance checks, validates install paths before mutation and
rolls back a partial write. Its twelve controls pass, but independent review
returns incomplete protocol-pin and source-assembly validation: changing
those fields and refreshing their dependent hashes still passes admission.
V9 must compare the full closure with the actual C/P sources before writing.
Actual ordering remains: current archive/binding and reviewed
consumer installation, clean P1 commit, signed generation against P1, then
signed code binding, manifest and test together in P2. No binding, archive
or integrated P1/P2 execution is claimed by the static previews.

## W6 preparation and historical ODF compilation

The W6 draft's real public preparation exposes four errors before succeeding:
a Bash variable expansion, an unwanted `.orig` file, workload-flag parsing,
and stripping a leading porcelain-status space. All four failed logs remain
preserved. The successful TDnt83 preparation binds the original f755/e4a
endpoints, exact post-overlay probes and locks, the two original W6 routes
and four ODS routes, and passes through the bounded parent/producer handoff.
No Cargo or native child runs.

The published receipt mistakenly attributes the first failed attempt's log
hash to the successful log. An append-only addendum records the successful
log's actual 1,346-byte identity; the original receipt and seal are unchanged.
Independent review verifies this correction and the endpoint, workload and
handoff evidence, then returns v5 for remaining path and invocation defects:
some public paths are resolved before symlink checks, and direct child/worker
entry accepts a nonexistent parent receipt path. V6 repair must bind an
actual parent invocation and retain final parent-receipt validation, without
requiring a completed receipt before its child can run. The sealed v6
successor adds lexical path checks and a live parent context; fourteen
controls pass and independent review is pending. W6 remains open.

The original ODF raw compiler source is materialized at exact ancestor
`70d64634c9d1d4328447d8724d3b076114a91fd0` with the reviewed 3401 manifest
and probe. Its historical 57,158-byte lock is preserved, but lacks seven ODF
path packages. Root replaces the proposed `generate-lockfile` command with
offline `cargo metadata` for the later intentional lock edit, followed by
comparison of all 204 existing registry entries and review of local changes.
This preserves resolved versions. Lock resolution and both original release
compiler variants await the Cargo slot; no baseline compilation is claimed.

## Cleanup and scope

Root removes 8,345,409,478 apparent bytes from the terminal failed-C33 and
passed-Q-Clippy build caches after checking service termination and process
references. Their source, logs, receipts and artifact quarantine remain.
The active diagnostic and ODF compiler caches are untouched.

The 77-head/65-worktree checkpoint includes the new DOCX leaf and required
70d compiler input; all 70 excluded heads are unchanged. The subsequent ODF
admission leaf is also active and awaits review. Existing unified control
branches remain merged. Root preserves the four failed W6 clones' HEADs,
trees, consumer patches and small probe/lock inputs, then starts retirement
of only their obsolete `prep/` directories. Failed logs/state and the
successful TDnt83 endpoints remain. No registered worktree is removed by
that source-copy cleanup.

Final gate, census, corpus, fidelity and performance qualification remain
open. C/P stay frozen until the repaired final source is selected. Memo 22
is the only deferral. Nothing is pushed.

## Evidence

Paths below are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `rest29-docx-metadata-failure-20261026/receipt.json` | 2,394 | `8cbe9ecc3e98a77702e51c2f91aa4f3964c374025974df825c0890f8004c04ca` |
| `protocol-composition-v5-root-clarification-20261026/comparison.json` | 10,604 | `1d02f674b8b9ea433193f9abbefeb8801457bc3bd4425b46bea6fab71a5b792f` |
| `protocol-consumer-composition-C2a-20261026-v7-independent-review-fresh-20261026/review-receipt.json` | 9,970 | `42b2f33a1603d38e75041405c9c62ef7907aca7195765965a9db7f1ff070d67a` |
| `w6-v5-real-prepare-receipt-correction-addendum-20261026.json` | 2,469 | `d9cffcb97999095b08f28c02ee36c149dea2c758eada870e9640afeedb13c97a` |
| `w6-odt-ods-C2a-prep-20261026-v5-independent-review-fresh-20261026/review-receipt.json` | 11,054 | `d3b6c7aff86b81b092fcfacc3a6d331b189c08aa09dd312338b1bc735e8e89f1` |
| `odf-raw-70d-lock-plan-20261026/plan.json` | 1,638 | `5591d476e78489b2bbfd49f3d25149e81fba4b84b1695cfa229ef28b960b3547` |
| `retired-gate-cache-cleanup-20261026/receipt.json` | 1,669 | `4ed47d86480321a8e79a7386610503d8244b081d036e5ff2aee8853355d96de0` |
| `branch-worktree-165l-20261026/receipt.json` | 1,948 | `1b00f857d964ff897a886a8dcb238c05d5436797f40cc6119e94409bc978010d` |
| `w6-failed-source-retirement-20261026/preservation.json` | 23,472 | `71384610a411de59ab9e0e23cf5f6d905cd1b65fb77244b867ca41a082f2f0c1` |

The live diagnostic is `litchi-q472a-rest29-v3.service`; its original plan
and launch remain those recorded in 165k. Source-copy retirement runs as
`litchi-retire-w6-failed-sources-v1.service`, with driver SHA-256
`5c5b52e3fa46c71a0d5761465c987a76003b9ccf522f34baa0d704f57e935230`.
