# 0362ul92 — Census preparation, compiler returns and cleanup 64

Date: 2026-10-25. Engineering and final acceptance remain active. Only
redaction audit 22 is deferred; its rules still apply. Nothing is pushed.

## Current census preparation

Leaf `2208959c1278ac9d7016ef10409acc0c19fd52f6` merges into integration at
`cd29aa8bdaa4f416f86195db737e4e3651b61d0e`. The new deterministic preparation
tool derives 69 glossary profiles and 15 package-only profiles, separately
from the 68 direct packages required by the all-features inventory. The
matrix deliberately has state `incomplete` and empty evidence arrays; it
does not replace or weaken the historical raw-coverage validator.

Independent review returns two source-binding bugs before acceptance.
The repair rejects relevant dirty input bytes and binds the legacy policy
manifest to the selected commit. Real synthetic reproductions reject both
cases. Final execution guidance now requires the approved F5 outer runner,
actual ownership/isolation receipts and 18/19/20 GB target guards with 12 GB
free space; the inventory payload alone is not an acceptance boundary.

All 107 selected glossary, inventory, census and fixture-reference tool tests
pass on integration, and the matrix check passes. The earlier broad test
attempt in the intentionally sparse author checkout reports three missing
module imports; it is not a product-code failure or a passing full suite.
The hydrated integration run supplies the complete 107-test result.

## Compiler and source-review returns

The first actual combined PPTX capture on `2915ef1762` passes formatting over
all 384 changed Rust paths, then stops on three OPC `unused_mut` errors.
Private repair `00a42a7d4c9bfead038e2e40c57a0c14291e5ecb` removes only those
qualifiers. Record 0362ul89f binds the failed capture. A fresh two-row plan
is prepared; no combined test, Clippy or performance acceptance is claimed.

Web/Ribbon capture `beac39a4c5` passes formatting and all four compiler rows
(common, four leaf formats, unified operations and facade). Common Clippy
then reports two derivable default implementations. Repair
`c877dee369346ba725f6d052562ef8e509cbb6c1` derives those defaults without
changing the selected limits. Its fresh 33-row validation is in progress.

Properties capture `029446d5d9` passes formatting, then the common compiler
reports two redundant test-helper result qualifications. Repair
`ac7b91e8d8aaf0b2d856cc036bc499e4750b685b` removes them; its fresh 37-row
plan is prepared. Earlier error-shape repairs and failed runs remain recorded.

VBA `b62c2e53590023df927ee520e8c769804497e5dd` passes independent source
review after fixing generated root relationship metadata accounting.
Regressions cover exact/one-under metadata ceilings for absent root members,
equal payloads and reverted relationship edits. The shared bounded comparison
and single admission scan remain intact. An inert prelaunch reveals two
mistyped fixture hashes; a unique corrected plan recomputes all hashes from
tracked files. No VBA compiler or runtime result is claimed yet.

The coordinator independently verifies actual completed rows against their
commands, source/lock identities, logs, target traces, resource guards and
terminal process ownership. Additional checked failures are retained outside
Git under `/home/zhuhe/code/litchi-worktrees/.coord/`:

| Capture directory | Result bytes | Result SHA-256 |
| --- | ---: | --- |
| `0362ul87b-f5-retry-beac39a4c5` | 25,496 | `8319cf286c414d1b54cb86315f3ecadb934af39cb5c97697e2d7e6272c0946d2` |
| `0362ul87a-f5-029446d5d9` | 10,000 | `841520bf2db05a6f58f6479629dd6f1f3263e6aa8adbf5f78703214c80a0754f` |

Each directory retains its exact `plan.json`, source-bound `run_f5.py`, row
logs and traces. The regeneration entry is that directory's driver with
`--go`, after a fresh reviewed source/target freeze; existing receipts must
not be overwritten. All isolated captures continue to run serially.

## Cleanup 64

The merged, quiescent image-conversion leaf
`uo2-imgconv-raw-record-limits` is removed without force. Branch
`uo2/imgconv-raw-record-limits` remains at
`aaeaa79a0bc0689e87069c78c4c87726c7bc8523`, an integration ancestor.
Four ignored files (294,862 logical bytes) are moved to same-device quarantine.
Dry-run and actual sparse restoration pass; final inode/device/mode/size/hash
checks preserve every payload. Both external validation roots remain intact.

The coordinator verifies the receipts, absent path/registry entry, retained
ref, held heads and unchanged protected constructor probe. Holds are released.
The worktree occupied 148,369,408 allocated bytes; final quarantine occupies
344,064 bytes, giving **148,025,344 allocated bytes reclaimed**. No build
target, validation evidence or native input is removed.

Artifacts reside at
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave64-imgconv/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 2,505 | `9b3ae670c894d4ac0884146c2f03decb564e59244d156c7b22b5cc559fc16d80` |
| `restore.py` | 4,544 | `9d65c99b923ef672103666e517d4291b720024acffa9a729eb4546f901fc56b7` |
| `restore-check.json` | 844 | `d7f6ce2d9436222be1d1763b6f07e6e55794ee75bddd09802cc9b8616d8cc174` |
| `receipt.json` | 2,221 | `35fe7c034d6223ee7d428c91ed45b323a5b94ab54b796c28ef8560227e8f1505` |

Final full formatting, complete gate including isolated workspace docs,
fresh raw/all-features inventories and every non-draft disposition, actual
corpus/native and quiet ABBA qualification remain required. These preparation,
source-review and cleanup results do not close those acceptance requirements.
