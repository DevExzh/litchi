# 0362ul165ub — Golden checks and bound protocol recipe

Current-7c diagnostic rows 42–45 pass: golden checks, golden-file cleanliness,
the existing performance harness compiler check, and native-resave compiler
check. Root verifies their logs, Cargo target traces, clean source, four
unchanged locks, resource guards and post-exit process samples. Target usage
after row 45 is approximately 11.74 GB. The API glossary audit follows.

These compiler checks use the existing 7c consumers. They do not compile or
execute the new DOCX metadata probe, whose separate release builds remain
queued. The original Python and missing-fuzz-lock failures remain preserved.

## Explicit recipe handoff

Protocol v15 adds source rebinding, but independent review returns hidden
working-file changes, symlink scope, missing producer-to-P2 wiring, stale
internal references and unqualified historical metadata. V16 repairs those
issues but leaves the recipe argument optional, allowing silent fallback to
the frozen C33 generator and finalizer.

V17 removes that fallback. The P2 CLI, production admission functions, and
generator/finalizer loaders require an explicit bound recipe. Current
callers and tests migrate together. The guard verifies the rebound recipe,
source identity and generation artifacts before installation. Frozen input
recipes and earlier packets remain preserved.

Root reviews the narrow v17 delta against the independent v16 review. Direct
signature inspection, explicit-None refusals and a real CLI omission check
confirm that absent recipe input fails before loading a frozen recipe or
creating output. AST checks pass. The author reports 20 rebinder, 24
composition, 21 phase and eight handoff controls passing. Root does not
relabel those as independently rerun tests.

This closes source-preparation review. The concrete output overlays remain
the reviewed virtual merge preview. Actual final C is still unselected;
clean P1, real producer generation, P2 installation and finalization must
still run against the eventually selected source. No archive, native timing,
allocator measurement or final acceptance is claimed.

## Fuzz preparation return

The v2 fuzz handoff repairs the caller-dependent limits, source/tool checks
and output protection. Its sparse source still omits Cargo targets needed
for metadata: nine ops fuzz bins, 58 facade fuzz bins, implicit library
targets and tracked build scripts. Those exact 7c blobs must be materialized
and included in the verified source inventory. Source-only preflight is
insufficient when the package targets are absent. A v3 repair is required
before any Cargo execution.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `remaining20-goldens-consumers-root-verification-20261026/receipt.json` | `82e42437458910b313880073ec865fd2c82c48d5857873c7b721f1f408ad9122` |
| `protocol-consumer-composition-C2a-20261026-v16-independent-review-20261026-fresh/review-receipt.json` | `a3066b08ffa7102ee477048134283ded8f81243cdf59667ac5868982e23def8e` |
| `protocol-consumer-composition-C2a-20261026-v17/receipt.json` | `1b64e3b6d120cb38ff119f5f14d62aec13f70ae72c2cf95b50452d84c759c2ab` |
| `protocol-consumer-composition-C2a-20261026-v17-root-review-20261026/receipt.json` | `122f000789209e0155e986c0fd43a376e157ad9844b336f275d270a2d372967f` |
| `fuzz-lock-handoff-7c-20261026-v2-independent-review-20261026-fresh/review-receipt.json` | `b86788ea5faec5589b4626184ce224ca88a192324bd9c0f618afeccc09566411` |

All 29 native groups and the fresh final all-65 gate remain required.
Nothing is pushed.
