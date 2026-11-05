# 0362ul165h — core suite and runtime handoff

Root verifies 35 completed checks from the C2a full65 gate before its later
Clippy failure, recorded below. The
core suite completes all six chunks, with 332 passing tests reported in the
logs, including 253 library tests and 59 doctests. Every chunk log matches
its producer's size and SHA-256; command order, successful exits and resource
checks pass. The reviewed root verifier validates the new top-level row
alongside the previous 34. C/P remain clean and frozen at
`33a0a63d3553712c1276da592f7846a0c7dcabd6`. Remaining gate checks, census,
corpus and native qualification still require their own evidence.

## ODG v4 reaches the next consumer

Independent review verifies both repaired v3 producer seams in v4. ODG's
public preparation, bounded synthetic build and sealing succeed; ODP's
producer now records its required 4 MiB output cap, and sealing rejects
changed artifacts. Forty-eight source tests and Python AST checks pass.
No Cargo, native build, archive or timing run is claimed.

The ODG manifest nevertheless fails the actual runtime validator. Its fixture
is an external packet copy with the correct bytes, but the runtime requires
the fixed reviewed fixture under the harness checkout. The handoff test stops
at `builder.validate_build_receipt`, so it misses that later consumer.
Independent review reproduces the runtime refusal and confirms that the
reviewed harness fixture passes with the same 738-byte identity/provenance.

V4 is returned for a v5 correction that preserves the fixed fixture binding
and exercises runtime admission after sealing. The future integration recipe
must bind paths to the actual reviewed harness checkout: its root changes
when the tooling moves from the leaf into protocol P. An external copy cannot
silently replace a fixed reviewed input. All prior packets remain immutable.
The existing ODF leaf is retained for this repair and subsequent compilation
and integration; it is not merged on the basis of synthetic tests alone.

V5 subsequently closes this return. Independent review passes the public
ODG preparation/build/seal/runtime-admission flow with the fixed tracked
fixture, all four reviewed probes, both endpoint identities and the required
build receipt. All four normal/allocator build roles are covered. The ODP
flow also passes with the 4 MiB cap and failed-receipt/artifact-drift refusals.
Both lanes are now source-ready, bringing the count to 25 of 28. Their actual
builds, native correctness, fidelity and performance remain pending. After
integration into P, fixture/probe paths must be rebound and source preparation
rerun; the leaf's receipts must not be relabeled as integrated-P receipts.

Root rechecks all eight source hashes in the independent v5 receipt against
the exact Git tip and file bytes. One recorded test-file digest lacks its last
hex digit. A separate immutable correction supplies the full digest and Git
blob proof; the original receipt is preserved and its source-only conclusion
is unchanged.

The new six-file ODG/ODP tooling patch also applies cleanly to the previous
26-file protocol preview. Independent review returns v3's receipt because
its purported pristine input paths point to files subsequently patched in
place. Root preserves v3 and rebuilds the full ten-patch sequence in v4,
using 23 separately sealed, read-only pristine inputs from C2a. Independent
review reproduces all 32 output files and parses all 26 Python files; the
output bytes match v3. This proves composition only. Actual current/signed
bindings and integrated tests remain pending, with the same P1/P2 ordering.

## W6 v2 still lacks its producer

W6 v2 repairs its source/lock declarations and overlay expectations. Eight
policy tests, three validator tests, AST parsing and shell syntax checks pass.
Its build/run recipe nevertheless requires an external `W6_RUNNER` and SHA
placeholder for which no implementation exists in the packet. The public
prepare command refuses before work, and the validator test hand-constructs
a receipt rather than consuming an implemented producer. Independent review
returns v2. V3 must implement and bind the complete guarded producer and test
its actual output through the consuming validator. This work is assigned;
it is not delegated to a future unspecified coordinator action.

V3 supplies a concrete producer but is also returned. Its source-backed
manifest selects an empty scenario list and therefore skips all four ODS
routes. The public phases check the SafeGate hash without invoking it, and
some child timeout/output limits are unused or checked only after execution.
Seven tests pass but do not exercise the composed public flow through real
producer receipts. V4 must repair those executable paths and test the full
prepare/build/seal/run/validate handoff before another review.

Flat-parts v2 is returned for reproducible source-parent symlink admission,
mutable build-profile admission, missing allocation-role/counter checks,
unenforced cap maxima and a mismatch between the build command's output
and its claimed binary. Additional source review finds an outdated candidate
API call, incomplete transitive lock binding, unbound manifest/oracle data,
an allocator enabled in the purported ordinary timing binary, detached timing
receipts and unused memory-wrapper enforcement. V3 must address these paths;
passing source-only tests is not evidence of compilation or native execution.

## Terminal gate result and isolated repair

The full65 run terminates with 35 checks passed, one failed and 29 unrun.
`clippy-leaves` reports seven test-only DrawingML errors: one redundant
`matches!(..., Ok(_))` and six needless borrows. Cargo exits 101; no resource
failure is recorded. The target peaks at 7,717,645,942 bytes, post-exit samples
show no owned processes or builders, and C/P plus all three lock hashes remain
unchanged. The terminal run and its evidence are preserved; it is not restarted
or relabeled as a successful run.

An isolated leaf commits the seven fixes as
`472a89a3fae1cbe81d177b6c2f139e03d788bf43`. Only the two test files change;
the production prefix of `diagram/model.rs` is byte-identical. Rustfmt and
diff checks pass. Root starts the original leaf Clippy command with
`--offline --locked --keep-going` on that clean leaf, in a fresh external
target under the approved runner and a 20 GiB service cap. This is a repair
diagnostic, not a successor full65 result. C/P remain frozen while the
diagnostic and source-only harness repairs proceed.

## Branch retention

Root remains on `feat/unified-ops-api`. The refreshed census has 76 heads
and 63 worktrees; all seventy excluded refs are unchanged. The new test-fix
leaf and existing ODF tooling leaf await their required validation and merge.
The four merged unified controls and original baselines still have active
consumers. No unused unified-derived leaf is found at this checkpoint, and
cumulative removal remains 277 worktrees and 439 refs.

## Evidence

Paths are relative to `(deleted-coord-artifact)`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `final-gate-C2a-root-verifier-20261026-v2/partial-3-core-complete/verification.json` | 26,604 | `4e8e1e6ac8a26e3221c00cb9213cabdb7eafab8cf1786a18f7cab346607a496d` |
| `final-gate-C2a-root-verifier-20261026-v2/core-6-chunks-verified.json` | 5,295 | `7fb9bda9cb19cd929c1b27bfdf238336a8369ae28a7e80f40f344a92215c188b` |
| `final-C2a-three-lanes-source-only-20261026-v4-independent-review-20261026/review-receipt.json` | 6,546 | `826bcf03419c47e4faba180bf2f6b1aeaf2a8cded5a70c031260b8d5ee6f9ceb` |
| `final-C2a-three-lanes-source-only-20261026-v5-independent-review-20261026/review-receipt.json` | 6,145 | `2d0d53b602299caccc1251a7707160c01d72c7386c3aa122432d0b50bd494538` |
| `final-C2a-three-lanes-source-only-20261026-v5a-root-correction-20261026/correction-receipt.json` | 3,845 | `086fa86352710766077374e6c7ecf058b5004e2308a81fce8cdf049fd8994551` |
| `protocol-consumer-composition-C2a-20261026-v3/receipt.json` | 13,193 | `dd07bde642e01bbe918198c6fd5095a23571fc893c32bd207849602d69e8f5db` |
| `w6-odt-ods-C2a-prep-20261026-v2-independent-review-20261026/review-receipt.json` | 8,290 | `f1079d1a32e3670b0d9dea82d5ff8e19a72901a6718413225aaf3353f95c8671` |
| `protocol-consumer-composition-C2a-20261026-v3-independent-review-20261026/review-receipt.json` | 7,006 | `d1d784411e5afae372e5ba8833e8817eafe1ff59220cd07de531454f975760d0` |
| `protocol-consumer-composition-C2a-20261026-v4/receipt.json` | 26,702 | `4c51bc167a573fb9a83d2bfd96b8ecfaad7b686bd9311325f46d8ec5d815611e` |
| `protocol-consumer-composition-C2a-20261026-v4-independent-review-20261026/review-receipt.json` | 5,302 | `faabb2e2342a2a63a4b3e02d534c3ac8f723d1c98176e023f7e902c20a402ea0` |
| `w6-odt-ods-C2a-prep-20261026-v3-independent-review-20261026/review-receipt.json` | 10,608 | `60ef0d933ab9edf22a5dd6f1f1092237953806e8430f7f7c7f77a6ef38379313` |
| `odf-flat-parts-C2a-prep-20261026-v2-independent-review-20261026/review-receipt.json` | 9,454 | `688ff277e4adec57a7f0e7d7b064d401b4827400647bf9becd26a23cf4c0b71a` |
| `final-gate-C2a-terminal-clippy-failure-20261026/failure-receipt.json` | 3,920 | `b872fe1b2cd67e120f6e2709ef888878f0e7f9018ccb10fadbb57e96497e7811` |
| `final-gate-test-repair-diagnostic-20261026/run.py` | 4,861 | `9ca0cede24f58347782fc200f893914561205fa14e53902fbf87c21ad12bf353` |
| `final-gate-test-repair-diagnostic-20261026/systemd-launch.json` | 2,390 | `3cc5a3451c488717c3e64b4e83e4b8a3dbbb9b4b81e9232e5dfb78135d621ca8` |
| `branch-worktree-165h-20261026/observation.json` | 17,527 | `9fcdb40cf39bdd82ea70af424e197f9798d0bb8527f9eb243ae9f480ea07f9fe` |

The restored flat-parts and W6 preparations remain in repair. The preserved e4a
baseline at `(deleted-coord-artifact)` contains a 139,425-byte root
lock, SHA-256
`04665b965a919412e29c8772956ddf6267264fa40e3e0ef7c9967fd2eb584171`;
the W6 worker is checking whether its provenance and dependency graph can
support that lane. Finding the file does not admit it or authorize modifying
the existing ODP baseline. All 28 final qualification rows remain open.
Only memo 22 is deferred. Nothing is pushed.
