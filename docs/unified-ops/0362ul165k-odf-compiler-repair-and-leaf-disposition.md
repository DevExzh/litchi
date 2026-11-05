# 0362ul165k — ODF compiler repair and leaf disposition

The ODF tooling leaf at `340191f1effbd9e47d68f7fcabbe21a526279359`
passes ordinary and allocator release compilation. Its 82 focused Python
tests also pass after materializing the tracked signed-fixture manifest.
Independent source review passes and root merges the leaf without conflicts.
This validates tooling;
historical compilation and actual paired qualification remain required.

## ODF compiler and consumer repair

The first compiler repair, `183bde41d9`, fixes the event-buffer borrow and
uses a macro to dispatch each closure to the existing generic timing
function. It adds no boxed or dynamic dispatch and preserves the route
boundaries. It also completes the candidate consumer-lock overlay: the
production endpoint starts with the pinned 56,941-byte lock, and receives
only the adapter's manifest, probe and 57,961-byte lock. Preparation and
the consumer verify the before/after identities and unchanged production
source closure.

The compiler then exposes another API mismatch. Changing the ZIP import
alone (`6cac74084f`) does not make the old ZIP calls compatible.
`d5eff369fa` uses the declared soapberry dependency's preservation index
and archive reader. The outside-clock audit retains archive/member comments,
member order, raw names, payload hashes/sizes, compression, compressed size,
CRC, version-needed, flags and external attributes; it also records raw
comment hashes and declared uncompressed sizes. Both release variants pass.

Root checks these calls against historical baseline `70d64634c9`. Its two
ZIP constructors take fewer arguments and internally use the same default
limits. `340191f1ef` adds the corresponding compile-time API branches and
updates the manifest's probe hash in the same commit. Fresh locked, offline
release checks pass for both current variants. The dependency cache is
reused; command receipts, logs, source identities and post-exit checks are
new. Neither check executes a native benchmark.

The broad Python run initially passes 80 of 82 tests. Both errors identify
the missing tracked `docs/unified-ops/results/0362ul24/fixtures.json` in the
sparse leaf. Materializing that directory makes all 17 signed-policy tests
pass, followed by all 82 tests across the eight affected modules. All 82 also
pass on the merged trunk tree (10.292 seconds). No fixture or assertion is
changed. `git diff --check` passes.

## Remaining-29 diagnostic

Independent review returns v2 because OfficeCLI row 37 could preserve a
symlink target. V3 rejects lexical symlink components on both paths, requires
a regular source, and verifies stable source identity, hashes, destination
type, byte count and mode around the approved preservation helper. All
17 controls pass. Independent review confirms the defect is closed and all
29 selected commands remain equivalent to the original plan after the
recorded path substitutions. The suggested additional runner-composition
test is explicitly optional coverage, not another defect.

V3's driver and controls are committed under `results/0362ul165k/` with
bytes identical to the reviewed external packet. The root dry-check passes,
then root starts `litchi-q472a-rest29-v3.service` under a 20 GiB memory cap.
This diagnostic covers only the 29 previously
unrun checks; it imports no passed rows and cannot close the final all-65
gate. The failed C33 run and v2 review remain preserved.

## Branch and worktree disposition

Root remains checked out on `feat/unified-ops-api`. The refreshed census has
76 heads and 63 worktrees. All 70 excluded main/office/independent-office
heads match the preceding audit. The four unified control branches are
already ancestors of trunk: coordinator integration, protocol installation,
approved gate runner and DrawingML test repair. Their worktrees remain
active validation inputs. The ODF tooling leaf is the sole unmerged unified
branch at the audit checkpoint; its subsequent merge brings every retained
unified leaf branch into trunk. The ODF worktree remains in use as a sealed
adapter/consumer input until protocol installation has replaced that role.

The temporary ODF lock-overlay design branch was merged into the ODF leaf,
then its clean worktree and branch were removed. Cumulative cleanup is
278 worktrees and 440 refs. Historical detached sources remain necessary
for the original comparisons. No additional unused unified leaf was found.

Flat-parts v4 is returned for incomplete tool identity, missing report-row
hash consumption and lexical manifest-symlink admission. V5 passes independent
source review, including four public tests, all 36 report rows and their
order/hash/status checks, historical API/feature roles and the approved
Cargo/rustc identities. Actual admission must pin the external `SHA256SUMS`
digest and verify native child/binary provenance; the source tests do not
establish those execution facts. W6 v5 repair continues. C/P remain frozen at `33a0a63d35`;
all 28 actual qualification rows and final gate/census/corpus work remain
open. Memo 22 remains the only deferral. Nothing is pushed.

## Evidence

All paths below are beneath `/home/zhuhe/code/litchi-worktrees/.coord/`.
Logs, binaries and build caches remain outside the repository.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `odf-raw-example-compile-3401-20261026/root-verification.json` | 1,601 | `bef9e10fd09abf54b7893acf8c4a1c23a39f22f845eea4c491b0ce36c0d7b349` |
| `final-gate-q472a-rest29-diagnostic-20261026-v3-packet/plan.json` | 87,551 | `84acd594c69cb46380c98bce1da3667d583db7b9ddaf700507cd2558eb02bc28` |
| `branch-worktree-165k-20261026/observation.json` | 17,845 | `ce616cf78b7eda82975fa0c5fdc1b647e6c84a69bf9012b7fa50453a4f6c578a` |
| `odf-lock-overlay-design-retirement-20261026/receipt.json` | 724 | `efc81c2f2ee25dc5dc8ba33e2bbc4f9cbf661d88b5df95c319235a1f80adeb1b` |

The current compiler commands are recorded by
`odf-raw-example-compile-prep-20261026-v6/run.py` (SHA-256
`83e8de77bab66b18cfbcded912e0c4d6ec3f05cc75632e71385dd096bb2586a9`).
The 82-test command and log are in `odf-raw-3401-source-tests-20261026/`;
the 293-byte log hashes to
`9bef40512e4fcd5e92ced339f53867480753f365a44c855773e7b609a3e51e06`.
The v3 independent receipt is in
`final-gate-q472a-rest29-diagnostic-20261026-v3-independent-review-20261026/`
and hashes to `c406a15aeb11513ba3bd13057cdeda3dc93e5e2ae6df7638da3ef21817daa2bc`.
The 8,000-byte flat-parts v5 independent receipt is in
`odf-flat-parts-C2a-prep-20261026-v5-independent-review-fresh-20261026/`
and hashes to `e4d06ce9e5c2b20eaa656a519eb30bc340e6c027e739e8d1fbcf8eff0cdee901`.
Its packet's 1,643-byte `SHA256SUMS` hashes to
`2391ee95038fc5411d0bf70084763c919db852b7daca75abf05d08f7368992c3`.
The diagnostic launch is recorded in
`final-gate-q472a-rest29-diagnostic-20261026-v3-launch/launch.json`, SHA-256
`6e8324f87b572aebf3acbe9f8b8e4797f7c4ef73c9ad690a86fe4893950d7ff7`.
The 8,167-byte final ODF independent receipt is in
`odf-raw-native-final-3401-independent-review-20261026/`, SHA-256
`638d5e9203c52fe67c99cfeb1cca2b6145a162bd01b340e3be843518d2d4b71e`.
