# 0924 — attribute owned-reader allocation requests before reprioritizing

This batch analyzes retained 0923 evidence at source revision
`679252bad65004085eb62a28b056b1a16b3f7468`. It introduces no production change,
new performance capture, or speedup claim. The preceding turn made progress
by committing 0923; the broader non-iWork goal remains active.

The 0923 reader-versus-vector requested-byte gap is a question about allocator
requests, not proof of excessive resident memory or an end-to-end bottleneck.
This investigation connects that gap to the current shared OPC input helper
and the already rejected [0779 growth experiment](0779-opc-bounded-input-growth.md).
It supersedes the open attribution recommendation at the top of HOTSPOTS.md.

## Source mechanism and semantic boundary

`SourceBackedPackage::from_reader_with_limits` calls `read_limited` and moves
its returned vector into the immutable owned source. `from_vec_with_limits`
moves an existing vector directly. The reader helper reserves up to 8 KiB
initially, then requests exact additional capacity after each successful read.
For the 2,101,427-byte fixture, 256 full chunks and a 4,275-byte final chunk
produce 257 nonzero reads. An ordinary EOF read produces call 258; this fixture
does not reach the configured input ceiling or its separate sentinel path.

The capacity-request sequence is 8,192, 16,384, …, 2,097,152, 2,101,427 bytes.
Its sum is 271,585,459 bytes, comprising the initial allocation and 256 growth
requests. This arithmetic counts full requested capacities; it does not count
physical copies, committed pages, peak live memory, or native CPU time.

The helper's current source SHA-256 is
`3dfa9ec92c0d090b0bc753ecacc63f263975a470c267b79b03ae1c5898ac3ab8`,
identical to the frozen 0923 candidate source. It also serves eager OPC and
physical package ingress, so changing it has a wider impact than this DOCX case.
Any future change must retain fallible allocation, maximum-input enforcement,
Interrupted retries, short reads, ordinary EOF, invalid read-count refusal,
late I/O errors, and the exact-limit one-byte sentinel before ZIP validation.
Sequential ingress has no execution context; managed cancellation belongs to
the separate context-bearing positional APIs.

## Existing-profile stack attribution

The final decoder verifies all 25 raw profiles against the retained 0923
independent-audit receipt. It assigns each allocation event once to its complete
leaf-to-root stack; it does not add inclusive parent totals. All eight reader
profiles have 257 events and 271,585,459 requested bytes in the `read_limited`
stack. None of the eight vector profiles has that stack.

| Whole-process requested bytes per profile | 0923 baseline | 0923 candidate |
| --- | ---: | ---: |
| Owned vector | 151,803,020 | 151,737,589 |
| Owned reader | 423,388,539 | 423,323,108 |
| Reader minus vector | 271,585,519 | 271,585,519 |
| `read_limited` contribution to difference | 271,585,459 | 271,585,459 |

Each cell repeats exactly across four profiles. The 60-byte residual is outside
`read_limited`: common stack keys net 48 bytes at a boxing frame and 12 bytes
across environment/argument frames per profile. Route-specific constructor
stacks cancel in requested-byte totals. These trace categories identify the
small remainder without assigning it to a particular timed operation. Optimized
symbol merging/inlining limits exact source-line interpretation; the trace's
`Package::from_reader` name is not proof of the literal harness call spelling.

The complete stack accounting is reproducible from the same raw files. This
analysis performs no recapture and does not infer physical-copy savings or
an achievable latency reduction from allocation requests.

## Native phase priority after 0923

Sum the raw six-phase arrays over all eight candidate processes per route
(1,600 samples each), then divide each phase sum by the total phase sum.
These are elapsed-time shares from the retained native experiment, not shares
of allocator requests, medians, CPU instructions, or a new timing result.

| Route | Open share | Query share | Publication share | Ideal speedup if all open work vanished |
| --- | ---: | ---: | ---: | ---: |
| Owned reader | 9.9652% | 34.5514% | 43.6172% | 1.1107× |
| Owned vector | 1.2168% | 35.9398% | 45.6628% | 1.0123× |
| Generic ReadAt | 1.0087% | 29.7193% | 59.5310% | 1.0102× |

The last column is the conditional Amdahl bound `1 / (1 - open_share)` with
all other work held fixed. It unrealistically removes the entire open phase,
including required catalog work; eliminating only vector growth has a smaller
potential. Cache/allocator interactions can change other phases, so this is a
prioritization model, not a performance guarantee. The separate 0779 experiment
shows why fewer cumulative allocation requests alone cannot predict latency.

Query and publication together still account for 78.17% of the owned-reader
interval and 81.60% of owned-vector time. Source-provider intersections are a
more useful next qualification target than another unproven growth policy.
Existing 0491–0494 provider infrastructure and measurements must be reused and
reconciled; those older paragraph-edit results do not measure this current
section-layout workflow. The next experiment needs current same-corpus
positional-file and controlled range-provider baselines, explicit source and
sink counters, and unchanged preservation/source-refusal checks. It must keep
simulated latency separate from physical cold storage and real networking.
No provider speedup or coverage promotion follows from this recommendation.

## Why the old growth candidate stays rejected

0779 measured a one-eighth geometric growth policy on generated and real XLSX
inputs. Generated open requested 96.18% fewer cumulative bytes, but retained
323,002 more bytes. Generated lifecycle/save whole-process RSS rose about 20%,
while lifecycle median changes were less than 1%. That candidate was rejected
and production restored. Its allocation savings cannot justify adopting it
now on a different fixture without a materially better measured tradeoff.

An exact-length hint is also not an implementation decision from this evidence.
The benchmark already owns a vector and can use `from_vec`; a new ingress API
would require a representative caller, bounded hint semantics, and separate
correctness and performance qualification. Reserving the default input ceiling
would not be a suitable substitute for knowing the input size.

## Evidence scope

The prior allocation protocol remains invalid as a qualification: one initial
profile was followed by an unapproved separate 24-process rerun. All 25 retained
profiles remain visible in 0923. Reanalyzing their existing allocation stacks
adds descriptive attribution and does not repair that protocol or turn
whole-process measurements into operation-scoped allocation measurements.
Native phase evidence comes from the independent, valid 0923 comparison.

The current production source matches the measured candidate. Its adopted
harness has only the separately recorded formatting changes from 0923.
No timing values in this record are newly captured at the 0924 revision.

## Reproduction and review

The [source review](results/change-0924/source-review.md) and
[test boundary inventory](results/change-0924/source-test-matrix.json) identify
current owners and existing controls. No Rust tests, builds, or performance
captures are rerun for this diagnostic-only batch. Independent evidence replay
and corrupted-input rejection validate the analysis tooling instead. The
[phase summary](results/change-0924/phase-share-v2.json),
[phase corruption controls](results/change-0924/phase-share-negative-controls-v1.json),
[stack validation](results/change-0924/validation-v1.json), and
[root replay](results/change-0924/root-replay.json) retain those checks.
The [independent review](results/change-0924/independent-review-receipt.json)
binds the final analysis and source/test audit.

Run the phase audit entirely from committed 0923 reports:

```sh
python3 -B docs/performance/results/change-0924/phase-share-v2.py \
  --repo . --output /tmp/0924-phase-replay.json
```

Run stack attribution against the retained external raw profiles:

```sh
python3 -B docs/performance/results/change-0924/stack-attribution-v2.py \
  --repo . --output /tmp/0924-stack-replay.json \
  --stacks-output /tmp/0924-stack-rows.json
```

Raw inputs retain their 0923 names, sizes and SHA-256 identities in the
[stack summary](results/change-0924/stack-attribution-v2.json); all 25 are
required. The separate 0923 custody records bind the source/build and capture
commands. Replaying this analysis is not permission to replace historical
profiles with new processes. Use the tool's explicit path overrides for moved
raw evidence.

The complete derived `stack-rows-v2.json` remains outside the checkout at
`/tmp/litchi-performance-0924/`: 150,263,792 bytes, SHA-256
`7db5a787e43d1b7c700c93d609d66d298f4c5cbd09884c58fca1e001c1a6340b`.
The stack command above regenerates it; the reduced committed summary is
135,378 bytes. No binaries or large evidence files are committed.

Cleanup removed obsolete draft analysis files after the final replay stopped
depending on them. Moved-packet controls and duplicate full-stack replay copies
were also removed; canonical external evidence remains retained. See the
[cleanup record](results/change-0924/cleanup.json).
