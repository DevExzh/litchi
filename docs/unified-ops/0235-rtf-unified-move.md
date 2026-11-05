# 0235 — Unified RTF paragraph moves

RTF now exposes paragraph `move_to` through `File`, explicit edits and batches.
The facade resolves Start, End, Position, Before, After and Into placements
against the original native snapshot, then passes one final zero-based ordinal
to native `Edit::move_paragraph`. Before/after destinations account for detaching
the source paragraph. Into accepts the body only. Self and adjacent anchors
resolve to exact native no-ops, including protected sources.

The adapter bounds expanded targets, staged retained text/action bytes, operation
count and retained history; it keeps native source/protection errors, origin,
strict policies and process-local patch identity. Mixed edits resolve sequentially.
Capabilities advertise paragraph moves while retaining the native source-closure
gap. Formatted, opaque, inline-break and compressed changed moves still require
broader native closure support. Add and swap remain explicit gaps.

Qualification exposed a native boundary bug: moving an empty paragraph to the
end could drop it during canonical publication. The native fix preserves the
terminal boundary needed to represent that paragraph and validates paragraph
count as well as body text after publication. Regression coverage includes
ordinary and terminal-`\par` sources, reopen and exact process-local inverse.
The current durable move vocabulary cannot represent exact boundary spelling;
serialization of moves involving empty paragraphs or terminal body breaks returns
`PatchError::InvalidText`. Ordinary moves retain their durable replay contract.
Boundary-aware durable preconditions remain a native follow-up.

## Verification

The complete native RTF suite passes 1,476 tests/doctests; the full-feature
facade passes 1,660 and the RTF-only facade passes 525. Eight new facade tests
cover native byte parity for all placements, self/adjacent protected no-ops,
empty-paragraph order, mixed sequential edits, exact identity-bound history,
batch attribution/rollback, typed limits, protection and atomic save/origin.
Native regressions cover boundary preservation and truthful durable refusal.

Warning-denied all-target native/facade Clippy, RTF-only Clippy, rustdoc,
no-feature compilation, scoped formatting and diff checks pass. Crate boundaries
pass for 66 packages, 36 Python tooling tests pass, and the generated OfficeCLI
delta is unchanged.

## Paired timing

The `rtf_move_bench` example compares native and unified first-to-end moves on
warmed retained owners, with matching native limits and a prepared path. Three
trials on one pinned CPU use two warm-up rounds followed by nine ABBA rounds
(18 samples per route). Setup, serialization, output checks and drop are outside
timing. Every sample verifies exact bytes and reopened paragraph order; the
source remains unchanged. The one-paragraph case is an exact no-op.

| Paragraphs | Native median µs | Unified median µs | Unified/native |
|---:|---:|---:|---:|
| 1 | 0.180–0.190 | 0.440–0.440 | 2.316–2.444× |
| 10 | 6.450–6.585 | 6.720–6.750 | 1.025–1.042× |
| 100 | 44.980–45.490 | 44.906–45.110 | 0.990–1.003× |
| 1,000 | 404.297–411.782 | 404.527–408.971 | 0.993–1.001× |
| 10,000 | 4487.256–4563.391 | 4470.470–4514.206 | 0.989–0.999× |

Changed moves in these cases remain within the 10% facade-overhead comparison
reference. The single-paragraph no-op is 2.32–2.44× native and remains a fixed
cost follow-up. These shared-host samples do not establish allocation or
repeated-command scaling acceptance.

Receipts: [trial 1](results/0235-rtf-move/trial-1.json),
[trial 2](results/0235-rtf-move/trial-2.json),
[trial 3](results/0235-rtf-move/trial-3.json). They record source/tree hashes,
binary size/SHA-256, compiler, CPU, raw samples and reproduction commands.
Measured source: `920f759a4f`. Build with
`cargo build -p litchi-ops --no-default-features --features rtf --example rtf_move_bench --release`,
then run `target/release/examples/rtf_move_bench` on an available pinned CPU.
Validation used external copied locks with cached `cc 1.6.0`; repository lockfiles
are unchanged. Owned temporary worktree/branch, builds and scratch artifacts
were removed; only the source and small timing receipts are committed.

## Remaining work

This milestone does not complete the goal. RTF insertion, swap, broader element
coverage, repeated-command coalescing, full performance/allocation qualification,
remaining native glossary families, other delayed formats and Python bindings
remain governed by the original goal. No new Python binding surface is claimed.
