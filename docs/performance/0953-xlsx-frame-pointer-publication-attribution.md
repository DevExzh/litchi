# 0953 — frame-pointer qualification of XLSX publication

All four frame-pointer captures pass the frozen owner-quality and semantic
gates. Longest-match search accounts for 60.00–65.88% of admitted publication
period weights across the two shapes and repeats. This is attribution in the
profiling build, not a speedup or an ordinary-native timing claim.

0952 rejected its DWARF profile after no publication-owner frames survived
reconstruction. This batch builds a separate profiling executable from the
same frozen 0950 candidate with `-C force-frame-pointers=yes`. It does not
change production source or replace the ordinary native baseline.

## Source and build

The starting checkout is `d54e88bf89`; all 33 changed paths since 0950 adoption
are performance documentation or bounded evidence. The source build remains
the frozen 0950 candidate based on `b6c5c80ffbcb60736d6c4c2f3b590838640008a9`
plus its twelve-file adopted delta, with the previously declared ignored root
lock exception. The standalone tracked lock controls the build. All 10,790
files (260,102,337 bytes) matched their manifest before and after compilation.

The build uses Rust 1.95.0, release opt-level 3, fat LTO, one codegen unit,
debug level 1, no incremental compilation, and two Cargo jobs. It retains the
0950 environment apart from separate target/temp paths and the frame-pointer
flag. The build returned zero without timeout; its process group was reaped
and clear before binary inspection. Actual `rustc -vV` is retained.

The 264,707,944-byte profiling binary has SHA-256
`0df5a52fc7b716cc0e7810b798987501f0053afcd85fcbe5a04cc00ec57a2274`
and GNU build ID `81b9bad2e48766f5e60bd2e853c1f3d56fd4519e`. All four
publication-owner instances and the measured harness caller establish an RBP
frame chain. The plan uses their new ELF ranges, not the old executable's
addresses. Frame-pointer sampling still requires exact mapping, build-ID,
PID, time, executable segment, owner, and immediate-parent checks.

## Compression-task census

Source review proves five changed Deflate members: workbook plus four
worksheets. Twelve other members are copied, including eight 512 KiB media
members. The six materialized source payloads are not six compression tasks.
The reports retain original compressed-range read overlap, not replacement
logical task lengths. Neither whole-package bytes nor those read counters
prove a useful parallel wave. Exact total/largest/remainder replacement bytes
remain a required follow-up before evaluating more than one worker. The
existing managed benchmark profile also limits workers and in-flight tasks
to one.

## Claim boundary

This profiling executable changes generated code and is for attribution only.
Its timers, allocations, and RSS cannot support a before/after optimization
claim against the ordinary 0950 executable. iWork is outside this batch.


## Capture and attribution

The frozen schedule is medium, dense-sparse, dense-sparse, medium, with 200
semantic iterations and zero warmups each. Sampling uses CPU 12, `cycles:u`,
997 Hz, frame-pointer callchains, fixed C/UTC environment, no inheritance, and
a 900-second process-group deadline. The first child required 100 fully
resolved owner samples, 100,000 owner period weights, zero lost records, and
all semantic checks before later children were released. All four passed.

| Child | Shape | Perf events | Qualified owner samples | Owner period weights |
| --- | --- | ---: | ---: | ---: |
| 0 | medium | 8,884 | 1,157 | 5,076,338,812 |
| 1 | dense-sparse | 13,649 | 2,335 | 10,275,952,066 |
| 2 | dense-sparse | 13,661 | 2,342 | 10,306,403,757 |
| 3 | medium | 8,883 | 1,150 | 5,046,933,766 |

All 800 semantic iterations pass and all captures report zero lost records.
A sampled event qualifies only when its physical owner and immediate caller
map to the exact new executable ranges and the callchain is resolved. The
unqualified population includes setup, planning, lifecycle/readback, and
missing/ambiguous membership; it is not a publication denominator. Owner
scope excludes returned-snapshot destruction, which the harness publication
timer includes. Period weights are sampled event weights, not exact cycle
counts or elapsed time.

After merging only terminal `+0x...` instruction offsets, longest-match search
is the largest exclusive leaf in every capture. Its shares are 60.00% and
60.45% for medium, and 65.40% and 65.88% for dense-sparse. The next individual
leaf, `audit_source_part_replacement_xml`, contributes 10.88–12.69%, with
`deflate_medium` at 8.26–9.54%. Exclusive leaves conserve the entire admitted
owner period. Nested stack rows overlap and must not be added to those leaves.
The two repeats establish observed consistency, not a confidence interval.

This makes compression the next measured opportunity in this profiling build.
The next bounded experiment should census exact replacement task lengths and
admitted scheduler width before testing the existing explicit managed write
wave. It must preserve unchanged-member copying, source freshness, full
original XML auditing, exact no-ops, reversible patches, deterministic errors,
and resource/cancellation limits. Ordinary native and allocation measurements
remain required before any adoption, including both one-edit controls and the
0950 tail-regression review flag.

## Preflight corrections and retained evidence

The corrected decoder was first exercised offline on the existing 0952 raw
capture. Samples omit `--debug-mode`; the mapping/task export uses it to retain
mapping records without sample printing. Both exports bind the same raw file
and have terminal receipts before parsing.

Review tightened exact schedule, argv, environment, artifact descriptors,
thresholds, and first-child admission bindings before capture. During final
preflight, a helper received a one-character `samples` to `sample` path-guard
correction after its hash had been frozen. The launch refused the mismatched
bytes before any workload or capture directory existed. Both exact helper
versions and the failed prelaunch receipt remain. Root then bound an immutable
v6 copy and a new plan/GO receipt; all four workload captures used those bytes.
There was no repeated or replacement workload capture.

Final plan is `plan-0953-v3.json`, SHA-256
`8e295dc6a23059400a5855e3bd15d9aede22c9ff54d02a2261671e5b669bc3a5`.
The runner is `capture_0953_v6.py`, SHA-256
`21dc19f3ad3b4b970b2dd72d97754cc4b02cba61a53e716674345f5dd6ab2def`;
the reader is `analyze_0953_v3.py`, SHA-256
`063eef6c4d29f01eb512100ed0d755e13ea1724eef4ccf6dbcae4fe96c165ac5`.
Older draft check outputs are historical diagnostics; replay authority is the
final frozen helper set and final admitted results.


## Archive, replay, and cleanup

The external core archive is
`/home/zhuhe/litchi-performance-0953-archive-sealed-v1/archive-0953-v1.tar.zst`,
2,584,519 bytes, SHA-256
`f0da81fc4c33f42416f2be49325cc2e85c72245b62e0c5f00d61d70c5a46902b`.
Its [manifest](results/change-0953/archive-manifest.json) binds 154 files,
24 directories, and 115,376,952 logical file bytes. Every raw artifact of
1 MiB or more is listed by name, size, and hash in the
[extraction verification](results/change-0953/extraction-verification.json).
The profiling binary remains an explicit, separately hashed external
build artifact; no binary or large raw evidence is committed.

The complete extracted set, sizes, hashes, modes, and absence of symlinks were
verified before the final reader replays. All four outputs reproduce exactly
after replacing only the known relocated root with the historical root.
Tracing permits the exact external binary file and metadata checks of its
parent directories; no original packet helper or capture file was opened.
The [replay receipt](results/change-0953/replay-verification.json) distinguishes
these allowed dependency accesses from a claim of zero original-path accesses.
The [aggregate replay](results/change-0953/summary-replay.json) also reproduces
the full summary after root normalization. The committed
[summary](results/change-0953/summary.json) is a compact projection of that
hash-bound full summary.

Nine actual [manifest controls](results/change-0953/manifest-controls.json)
reject missing/extra files, file/directory symlinks, wrong hashes/modes,
duplicate paths/JSON keys, and path traversal. The first directory-symlink
fixture failed during setup; its corrected version exercises the verifier
itself. Earlier failed invocations remain diagnostic history, not passing
validation evidence.

To recover the raw evidence exactly, extract the named archive into a fresh
directory using `tar --extract --same-permissions --no-same-owner --file
<archive> --directory <fresh-directory>` and verify against its manifest before
parsing. The [closure](results/change-0953/archive-closure.json) records exact
build, capture, archive, and replay dependencies and commands. Repacking uses
a fresh output directory and does not promise byte-identical compressed bytes.
The committed Python files are source copies; their archived names and bound
layout govern replay, including the adapter for captured absolute paths. The
current adapter expects the fresh extraction root
`/home/zhuhe/litchi-performance-0953-replay-v1`; its `--packet` option alone
does not change that mapping constant.

Owned replay trees, generated build cache, temporary control fixtures, and
bytecode were cleaned after verification. Canonical evidence, the archive,
and the bound profiling binary remain. Production Rust source is unchanged;
validation in this batch covers the build, 800 semantic iterations, static
prologues, raw owner admission, reader/runner controls, and archival replay.
