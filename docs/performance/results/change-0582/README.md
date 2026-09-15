# change-0582 evidence packet: a deterministic differential for change 0580's ZIP strict-layout scope

Change record:
[`docs/performance/0582-zip-strict-scope-differential-fuzz.md`](../../0582-zip-strict-scope-differential-fuzz.md).
Subject under review: [`0580`](../../0580-zip-target-scoped-strict-layout.md),
whose `parse_zip` fuzz gate could not be run on this host.

Disposition: retained. `performance_claim: none`; this is a correctness and
safety gate. Nothing under `crates/` was modified.

## Contents

| Path | What it is |
| --- | --- |
| `build_corpus.py` | The corpus generator. `RNG_SEED = 0x05820580`, seeded once and drawn from in a fixed order over sorted paths, so the corpus regenerates byte for byte. Emits four families: the 7 retained ZIP seeds, 516 distinct real ZIP containers from `test-data/`, 45 hand-built archives aimed at the strict-layout proof, and 22,307 deterministic mutations. |
| `strict_scope_differential.rs` | The harness. Built as an `examples/` target of `soapberry-zip` inside each extraction, so it links the crate under test without adding a dependency to any crate. Ports the whole body of `crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs`, then adds a full sweep of every member through all four public entry points that reach `strict_layout_for`, forward and reverse, on fresh and reused readers, under two limit profiles. |
| `strict_scope_coverage.rs` | The coverage probe. Built against a **third**, instrumented copy of the after tree carrying two atomic counters. Not one of the differential builds; no verdict in the record comes from it. It answers "how many inputs reached the changed code at all". |
| `classify.py` | The comparator. Pairs the two reports on (input, profile, API, member) and sorts every difference into class A (the approved overlap narrowing), B (a narrowing whose pre-change refusal was something else), C (accept → refuse), D (different bytes), E (error-identity drift), plus panics and intra-build oracle failures. |
| `run.sh` | Both `git archive` extractions, the overlay, the corpus, both builds with their own `CARGO_TARGET_DIR`, both runs, the classification and the coverage probe — end to end. |
| `corpus-fingerprint.txt` | Family counts, the RNG seed, per-file SHA-256 for all 45 crafted archives, and one hash over the whole sorted manifest. |
| `classification-summary.json` | Every count the record cites: the class histogram, the per-API and per-family breakdowns, the class-B refusal identities, the 130 class-E identity pairs, and the panic and oracle lists (both empty). |
| `coverage-summary.txt` | The coverage probe's per-family table. |
| `crafted-verdicts.txt` | The full before/after verdict table for all 45 crafted witnesses, with the changed rows marked. This is where the approved witness, the residual-window boundary cases and the two smuggling witnesses can be read directly. |

The two full reports are 238,685,636 and 233,872,704 bytes and are **not** retained. `run.sh`
regenerates them; `corpus-fingerprint.txt` and `classification-summary.json` are
what the record cites.

## Result

2,692,431 verdicts compared over 22,875 inputs, on `93a610ded` versus the same
tree with change 0580's two files overlaid:

```
panics, either side                              0
intra-build oracle failures, either side         0
accept -> refuse                                 0
accept -> accept with different bytes            0
refuse -> accept, pre-change refusal was overlap        37,176   (class A, approved)
refuse -> accept, pre-change refusal was something else 1,094,221 (class B)
refuse -> refuse with a different error identity          351,525 (class E)

real archives (516) and seeds (7): zero divergences of any class
inputs that entered the strict-layout proof: 15,468 of 22,875 (3,817,336 entries)
```

No counterexample to the soundness of the target-scoped proof was found. The
findings are that the shipped delta is 30× the approved one-row table and spans
twelve refusal identities rather than one, and that a local-versus-central compressed
size disagreement on an untouched record now lets a member be read out of a
region a local-header-trusting reader assigns elsewhere
(`crafted/neighbour-local-csize-smuggles.zip`). The record recommends a hold for
re-approval, not a revert.

## Reproducing

```sh
docs/performance/results/change-0582/run.sh /path/to/scratch
```

Requires only the stable toolchain. No `cargo-fuzz`, no nightly. About 45 s of
runtime after the two builds; the corpus is 334,625,602 bytes and the two
reports another 472 MB, so give the work directory about 1.5 GB on top of the
three `target` directories (roughly 210 MB each).
