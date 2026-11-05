# 0362ul130 — diagnostic suffix and cleanup98

The remaining43 diagnostic packet on frozen
`a8d32513b8e427c6d234ddd17fb836b0e8d2477b` terminates failed after 27 rows:
24 pass and three fail. Root verifies all row commands, source/three-lock
fences, physical log/trace identities and process/target guards before
releasing the heavy slot. The failure-admission journal also matches the
actual rows: indices 18 and 20 are direct Cargo-test exit 101 and continue;
index 26 is a helper exit 1 and stops. Aggregate acceptance remains false.

All-feature PPTX chunks 32–53, all six examples and doctests execute. Two
chunks fail: primary-text-box output/memory admission (128), and streaming
configuration error classification (129). The minimal-feature helper passes
chunks 0–30, then reproduces the known readonly physical-event failure (126)
in chunk 31. Its later chunks, examples and doctests remain unrun. Sixteen
outer suffix rows are prepared separately; their execution is not claimed
by this record.

External packet, relative to the worktree parent's `(deleted-coord-artifact)`:
`ooxml-pptx-joint-a8d32513b8-f5-runtime-remaining43-v2`.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 114,079 | `5d664ddc3f3b06cf7350ee6cd6eca424bade2cdc248be452010a4683e10d6d52` |
| `result.json` | 124,193 | `e2de338cf4cff99debd189f4f4c292c283b375f833f2f341634a007f71316611` |
| `pptx-minimal-tests/0031.log` | 148,446 | `919f39bb5e86587730726c3bb247e3ec28023f7e8b4c49bdbfb42863371ca546` |

Use a fresh source-bound packet, target and temporary directories to
regenerate via `python3 -B run_f5.py --go`. The captured failed result is
immutable and does not qualify any repair.

129 passes root and independent source review at
`77ebd7e3b65c5edf7d851f3cf87ad66d6a0eed63`: the test now requires exact
`InvalidLimit(Codec(OutputBytes))`, value 512 MiB + 1 and maximum 512 MiB.
Configuration validation precedes sink construction; the unchanged-sink
assertion and neighboring runtime `Limit` coverage remain intact.

128 identifies a production dimension bug: transient memory was bounded by
serialized output bytes. Root returns the initial one-helper repair for the
same defect in five adjacent helpers. 128a corrects all six without adding a
parse or allocation pass, but its test incorrectly constructs private
`ChunkAdmission` fields and overstates hierarchy coverage. It is returned
as 128b before composition. Record0152's output-first guarantee concerns
text-box authoring; the other five paths preserve their existing native
exact-output order, including shortening and no-op behavior. There is no
new uniform precedence requirement or accepted exception.

## Census preparation

Raw census v4 is returned before execution because real inventory receipts
use `identity.package`, whereas `packages.json` uses `name`; its fake fixture
had concealed the mismatch. V5 uses the actual producer for its fixture,
normalizes those two fields explicitly and checks exact package/library/
manifest identities. It also corrects plan indices to probes 0/1, inventory
2, raw profiles 3–86 and validator 87. Root and independent review pass all
nine build-free tests. The inherited commit-object, cleared-toolchain and
immediate per-profile artifact archival checks remain intact.

External adapter:
`(deleted-coord-artifact)/run_final_f5_v5.py`, SHA-256
`0721536141117f300c51ba876f8af8b90159c624b7c187ac37a207b0de59def1`.
Plan SHA-256:
`8ec3fb78b9147155295e23c11a77040c346608e97d47dd6d1850897f65c30037`.
This is reviewed preparation only. Final composed source, the S→C matrix
commit and fresh binding are required before the 84-profile/68-package run.

## Cleanup98

Root removes the obsolete detached DrawingML checkout
`uo2-drawingml-diagram-correctness-fix` at
`b6a2174db0b827304a1e3a114a534257f71be965`. Its dirty state consists solely
of 460,785 tracked deletions. The full exact deletion manifest is preserved
externally; all 460,792 Git blobs remain reachable, and the seven remaining
files match their recorded blob, size, mode and digest. There is no staged,
untracked or ignored payload. The tip and superseding signed-path coverage
are ancestors of retained integration. No ref is deleted.

The first removal proposal is returned for using non-force removal on the
dirty checkout. Its replacement guard is returned for silently ignoring FD
permission failures. Reviewed v2 classifies CWD, FD-list and FD-readlink
errors and matches the explicit same-user service exceptions by PID, UID
and command. Root reruns all fourteen checks into a fresh receipt, finds no
visible holder or unclassified error, then performs the one bounded force
removal and verifies the absent path/registration, retained Git history,
unchanged heads, manifest and protected probe. Inaccessible other-user
system/container processes and the named services remain disclosed; no
complete visibility or zero-descriptor proof is claimed.

Evidence under `(deleted-coord-artifact)`:

| Artifact | SHA-256 |
|---|---|
| `cleanup98-drawingml-deletion-manifest.tsv` (57,980,916 bytes) | `0a296cbebad5caa3df8f46e0528d4e0b4f0af4690a20c25495ea05bdfcd828be` |
| `quarantine/wave98-root-removal/preflight.json` | `de71111f0884d0358aaa8c86f3613303977f203c42bedbf72919beb637f6e5d7` |
| `quarantine/wave98-root-removal/removal-receipt.json` | `4dcb3634a2965c8b3a95777ed0144aef6f555603a04b8f0d9389da027c39f90a` |
| `quarantine/wave98-root-removal/postflight.json` | `e6860ca15afd80fa0e21f9d4deb16459a11c2435f7be5937fbc2c8b0eeeb4e8c` |
| `cleanup98-root-allocation.json` | `c4c9dd5652c0635a1393b86973cde4db037072f5ca999e2e761082a47efe52d8` |

The manifest can be regenerated from the recorded base and deletion overlay
with `git diff --no-ext-diff --name-status --no-renames`; its exact bytes are
retained externally after removal. This batch costs space: the checkout had
491,520 allocated bytes, while the enumerated evidence uses 58,355,712.
The ledger increment is therefore **−57,864,192 bytes**, bringing the total
to 753,286,500,352 bytes under the existing exclusions. This is not a
filesystem free-space measurement. Totals are 257 removed worktrees and
422 deleted refs, with prior evidence limitations preserved.

Root remains on `feat/unified-ops-api`. Repaired-source qualification,
composition, final gates, census, fidelity and performance remain open.
Nothing is pushed.
