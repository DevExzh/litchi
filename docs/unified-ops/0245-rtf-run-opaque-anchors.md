# 0245 — Opaque body anchors through RTF run edits

## Contract

Run text replacement and removal now preserve opaque body nodes whose bytes lie
outside the changed lexical spans. The native editor derives selected run spans
in semantic UTF-8 byte positions in one traversal, prepares cumulative position
changes once, and uses binary search for each retained anchor. This replaces the
previous per-anchor collection/sort and includes run text deltas that the old
projection ignored. Sources without body opaque nodes skip projection setup.

An anchor at the start of a replacement remains before its new text; an anchor
after the selected run follows the replacement. When all runs of an unterminated
final paragraph are removed, the native publisher emits a paragraph delimiter to
retain the empty paragraph. That semantic byte participates in projection, so
opaque syntax after it receives the corresponding post-delimiter anchor.

Candidate validation still compares opaque node count, order, kind, exact source
bytes and projected anchors. Source-bound patch replay and inverse preserve exact
endpoints. The facade continues to delegate to the same native transaction, with
no new public API or alternate rewriting path. Capability wording now distinguishes
preserved boundary anchors from opaque syntax overlapping changed source spans.
Other native source-closure restrictions remain in place.

## Verification

The new facade regressions failed with the prior implementation, then passed
with native projection integrated. Native RTF passes **1,515 tests/doctests**;
all-feature facade passes **1,738**; RTF-only facade passes **603**. Focused cases
cover opaque nodes before, between and after edited runs, source Unicode and
Unicode replacements, expanded/shortened/empty replacements, unchanged prefix
anchors, anchors after paragraph breaks, terminal-empty paragraph preservation,
query fan-out, reopen, source immutability, and exact patch replay/inverse.

Warning-denied native/all-feature and RTF-only Clippy, warning-denied rustdoc,
no-feature checking and both native/facade RTF glossary gates pass. Scoped formatting
and diff checks pass. Tooling passes **1,166** unit tests; crate boundaries pass
for **66 packages / 268 dependency declarations / 11 existing debt items**.
The OfficeCLI capability delta check passes. Final read-only review found no
remaining concrete projection defect. Python bindings are not qualified here.

Parser smoke checks pass for **792 query inputs / 748 plans**, **909 paths /
435 accepted**, and **776 values / 627 successful coercions**. The frozen performance harness passes all **nine** RTF tests, including
stable large-fixture output hashes.

The main verification commands are:

```sh
cargo test -p litchi-rtf
cargo test -p litchi-ops --all-features
cargo test -p litchi-ops --no-default-features --features rtf
cargo clippy -p litchi-rtf -p litchi-ops --all-features --all-targets -- -D warnings
cargo clippy -p litchi-ops --no-default-features --features rtf --all-targets -- -D warnings
RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi-rtf -p litchi-ops --all-features --no-deps
cargo test --manifest-path tools/perf-baseline/Cargo.toml --lib rtf
```

## Paired measurements

The existing run-text and run-removal harnesses now accept
`LITCHI_RTF_BENCH_OPAQUE=1`. This inserts an unknown destination after each
middle run and independently checks its body-byte anchor and exact bytes on
reopen. The same 20-case matrix covers separate paragraphs and dense runs,
1–10,000 groups, and direct last-run versus all-italic query edits. Fixtures,
selectors and native updates are prepared outside timing; staging and commit
are timed. Serialization and readback checks remain outside timing.

Ranges span three trial medians; durations are microseconds. Each trial uses two
warm-up rounds and nine ABBA rounds (18 samples per route). All 120 cases pass
their exact byte, text, formatting, count, source and opaque-anchor checks.

| Edit / shape / groups / route | Native µs | Unified µs | Unified/native |
|---|---:|---:|---:|
| text / paragraphs / 1 / all | 6.88–6.94 | 9.46–9.72 | 1.37–1.41× |
| text / paragraphs / 1,000 / all | 3,503.44–3,547.91 | 3,904.85–3,992.45 | 1.10–1.14× |
| text / paragraphs / 10,000 / last | 45,276.42–45,404.66 | 45,178.92–45,575.78 | 1.00–1.00× |
| text / paragraphs / 10,000 / all | 61,643.58–62,006.22 | 67,552.56–69,061.29 | 1.09–1.11× |
| text / dense / 1 / all | 6.96–7.17 | 9.49–9.64 | 1.34–1.37× |
| text / dense / 1,000 / all | 3,187.49–3,198.28 | 3,508.33–3,618.51 | 1.10–1.13× |
| text / dense / 10,000 / last | 33,447.12–33,615.94 | 33,535.59–33,826.65 | 1.00–1.01× |
| text / dense / 10,000 / all | 47,670.23–48,150.38 | 51,385.24–51,696.07 | 1.07–1.08× |
| remove / paragraphs / 1 / all | 5.72–5.78 | 8.21–8.39 | 1.43–1.45× |
| remove / paragraphs / 1,000 / all | 2,705.18–2,730.13 | 3,125.44–3,223.35 | 1.14–1.19× |
| remove / paragraphs / 10,000 / last | 44,655.38–45,149.88 | 44,734.23–45,591.78 | 1.00–1.01× |
| remove / paragraphs / 10,000 / all | 47,684.71–47,711.81 | 53,505.87–54,556.37 | 1.12–1.14× |
| remove / dense / 1 / all | 5.95–6.05 | 8.12–8.13 | 1.34–1.37× |
| remove / dense / 1,000 / all | 2,355.41–2,408.39 | 2,674.04–2,802.27 | 1.14–1.16× |
| remove / dense / 10,000 / last | 33,284.04–33,529.97 | 33,891.15–34,179.98 | 1.01–1.03× |
| remove / dense / 10,000 / all | 34,430.64–34,840.17 | 38,188.66–39,031.04 | 1.10–1.12× |

Large direct-path cases remain within the 10% reference. Generic fan-out still
exceeds it in several cases, while tiny queries retain fixed overhead. From
1,000 to 10,000 groups, unified all-target replacement grows **17.05–17.30×**
for separate paragraphs and **14.29–14.69×** for dense runs; removal grows
**16.93–17.20×** and **13.93–14.28×**, respectively. These results do not prove
linear end-to-end scaling or full performance acceptance. Native parsing/source
mapping, generic query resolution, allocations and repeated-command coalescing
remain follow-ups. The prepared anchor projection avoids repeated per-anchor
collection and sorting, but this milestone does not isolate its timing from
the rest of publication.

Receipts:
- Text: [trial 1](results/0245-rtf-run-opaque-anchors/text-trial-1.json), [trial 2](results/0245-rtf-run-opaque-anchors/text-trial-2.json), [trial 3](results/0245-rtf-run-opaque-anchors/text-trial-3.json).
- Remove: [trial 1](results/0245-rtf-run-opaque-anchors/remove-trial-1.json), [trial 2](results/0245-rtf-run-opaque-anchors/remove-trial-2.json), [trial 3](results/0245-rtf-run-opaque-anchors/remove-trial-3.json).

Each receipt includes raw samples, source commit/tree, compiler/CPU, binary
size/SHA-256 and regeneration commands. Measured implementation: `bf31471deb`.
No executable is committed. Builds use external copied lockfiles with cached
`cc 1.6.0`; repository lockfiles remain unchanged. The isolated implementation
worktree/branch and this turn’s temporary build artifacts were removed. The
base checkout remains at `32ad64e3207655b4ba558fc2793d9a7396655cdf`; the branch
is rebased onto it.

## Remaining work

This milestone does not finish RTF CRUD, durable replay for new run operations,
repeated-command coalescing, allocation qualification, query/scaling performance,
other format coverage, the complete glossary migration or Python bindings.
The full goal remains active.
