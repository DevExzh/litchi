# 0220 — Aggregate PPT text publication

Record 0219 removed per-target drawing rewrites, but every 256-target chunk
still opened and published another append-only editor. A 10k direct text edit
performed forty publications and copied progressively larger history snapshots.

Native `Edit::set_shape_text_chunks` now admits a complete bounded plan and
publishes it once. Each inner slice retains the ordinary 256-target contract;
at most 4096 chunks (including empty chunks) and 4096 globally distinct targets
are accepted. Checked replacement-byte admission precedes target resolution.
The primary edit charges every requested target, including exact no-ops, against
its selected operation limit. Empty plans charge zero operations.

The existing native engine resolves projected positions, validates duplicates,
encodings and dependency closures, groups drawing rewrites, checks persisted
record identities, preserves unrelated streams, and reopens the entire candidate
for semantic readback before adoption. Signed exact no-ops preserve source bytes;
changed signed writes refuse. The canonical adapter stages one cloned transaction
and retains its source fences, so late failure leaves both semantic state and
operation accounting unchanged.

Direct facade text runs and text-only query fanout use the aggregate native call.
A repeated target or mixed text/geometry run retains its existing sequencing.
More than 4096 direct assignments still roll over into bounded native edits,
within one atomic facade edit. Native leaf parsing/publication remains authoritative.

## Verification

The existing all-feature native suite passes 1,411 tests/doctests with the
aggregate implementation. Eight additional integration tests pass after final
admission fixes. The full all-feature facade suite passes 1,603 tests/doctests;
40 final PPT-only write, coalescing, structural, capability and allocation tests
also pass. The three focused aggregate harness tests pass.

Native/facade and harness all-target warning-denied Clippy, warning-denied
rustdoc, the PPT glossary profile, 66-package crate-boundary validation,
OfficeCLI capability delta check and its 15 Python tests pass. Both repository
lockfiles remain unchanged. Validation uses an external root lock with cached
cc 1.6.0; the harness uses an external exact copy of its own lockfile.

The regression tests exercise the exact 256/4096 boundaries, oversized chunks
and chunk count, aggregate count, lower edit budgets, byte limits before target
or signed policy admission, duplicates across chunks, late encoding/missing
shape failure with retry, no-op accounting, signed preservation, stale source
precedence, unrelated streams/shapes, full readback and exact patch/inverse.
The 4096-target case compares semantic output with sequential old chunks and
asserts exactly one new UserEdit predecessor link. Facade tests compare 1k
coalesced output directly with the aggregate native bytes and replay both ways.

## Benchmark identity

New opt-in routes are named
`ppt_text_batch_{eager,positional}_{native,unified}_{count}_aggregate_v2` for
counts 1, 10, 100, 256, 1000, 4096 and 10000. They use the same deterministic
source corpus as records 0218–0219. Qualification checks exact native/unified
bytes, every resulting shape text, unchanged source, expected operation/commit
counts, inner-chunk counts and the actual UserEdit predecessor-chain growth.
One aggregate call creates one publication per native edit; it does not silently
collapse distinct sequential public calls.

The former `*_v1` selectors are retired from the active harness. Their frozen
JSON evidence and route-name records remain unchanged. The former native route
called `set_shape_texts` repeatedly, whose extra history bytes now differ from
the aggregate facade output above 256 targets. Renaming the active routes keeps
the old exact-byte oracle intact rather than weakening it to text equivalence.
This follows the earlier `chunk_publication_v2` precedent in record 0158.

## Paired measurements

Runtime source: `ba7755e4a0` (native API `eee84e1c40` / `44bac4ab92`).
[Raw samples, build identity and exact hashes](results/0220-ppt-aggregate-batch.json)
cover all 28 routes. All source hashes match record 0219. Output hashes also
match through 256 targets; above that boundary the new aggregate history
intentionally differs. Every native/unified pair has identical output bytes.

One warmup precedes five samples per leg, native then unified. No builds or
tests overlapped timing. Runtime sources were committed; only documentation
was uncommitted. Release LTO was disabled. Opening, fixture and command-vector
preparation, serialization and proof checks are outside timing; string-path
parsing, publication and patch destruction are included. The default 256-MiB
snapshot profile is used through 4096 operations; 10k retains the explicit
finite 1-GiB profile from prior measurements.

| Operations | Eager native ms | Eager unified ms | Ratio | Positional native ms | Positional unified ms | Ratio |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 0.116 | 0.120 | 1.038 | 0.115 | 0.135 | 1.176 |
| 10 | 0.144 | 0.190 | 1.324 | 0.146 | 0.210 | 1.439 |
| 100 | 0.541 | 0.906 | 1.676 | 0.550 | 1.069 | 1.943 |
| 256 | 1.230 | 2.104 | 1.711 | 1.227 | 2.420 | 1.972 |
| 1,000 | 4.331 | 7.577 | 1.749 | 4.409 | 9.170 | 2.080 |
| 4,096 | 17.510 | 31.924 | 1.823 | 17.521 | 37.344 | 2.131 |
| 10,000 | 122.643 | 228.629 | 1.864 | 129.047 | 236.572 | 1.833 |

Relative to record 0219 on identical source bytes and requested replacements,
1k native improves 3.31× eager / 3.11× positional; unified improves 2.37× /
2.03×. At 10k, native improves 36.20× / 32.94× and unified improves 20.01× /
19.71×. Those comparisons deliberately include the reduction in publication
history size; they are not claims about producing identical historical files.
The actual UserEdit-chain proof confirms one added publication through 4096
and three at 10k, replacing the prior four/sixteen/forty at 1k/4096/10k.

The default-profile 1k eager route meets the 2× relative budget (1.749), but
**positional 1k remains over budget (2.080)**. Its native samples span
4.403–4.418 ms and unified samples 9.045–9.256 ms; this run does not justify
rounding the result into a pass. Single-operation eager overhead is 3.8%,
while positional overhead is 17.6%, above the 10% single-write trigger.

Within one native profile, 1k→4096 growth is close to the operation-count ratio.
Across the rollover boundary, **linear scaling remains unproven and the measured
1k→10k growth still exceeds 10×**: native 28.32× eager / 29.27× positional,
unified 30.18× / 25.80×. A 10k facade edit still retains three native
publications and their exact inverse images. The aggregate fix removes
per-256-chunk publication, but does not complete the all-size linear-scaling bar.

These are local synthetic same-slide text closure edits on warm owners.
Blocked leg ordering can retain drift. Cold opens, command JSON conversion,
geometry, broader corpus and allocation qualification remain outside this
matrix. The final source checks, counted history links and complete semantic
readback qualify correctness independently of timing.

## Remaining work

The next performance work is to profile facade planning/freshness overhead on
positional 1k edits and retained publication cost across native rollover. The
native operation bounds and per-item source/error guarantees must remain
explicit; lower relative overhead cannot be obtained by weakening them.

Both owned agent worktrees/branches and this turn's temporary target, copied
locks, probe binary, logs and raw result copies were removed after reduction.
No compiled artifacts are committed.

Performance acceptance remains measurement-driven. Broad CRUD, PPT slide/shape
insertion and removal, delayed formats, Python/CLI coverage, allocation and
corpus qualification, and the full unified-operations goal remain open.
