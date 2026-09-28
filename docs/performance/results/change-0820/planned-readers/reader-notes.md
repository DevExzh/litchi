# 0820 reader notes

`analyze.py` and `validate.py` are offline, fail-closed readers. They consume
retained packet evidence only; they do not invoke Cargo, a binary, an exporter,
a workload, or a child process. Root should run them only after every capture
child has terminated and the capture completion receipts are terminal.

The frozen packet is `litchi.performance.0820.plan.v1` at base
`096b810f23cc66fe01ec8c96c74f36888f954eff`. Its timed selectors are the 24
flat rows whose identity is `id = case + "__" + policy`, with the fields
`id`, `case`, `format`, `phase`, `input`, and `policy`. The rows cover three
caller-named real files, the `lifecycle` and `atomic_publish` phases, and
`default`, `full`, `file-only`, and `no-sync`. `custody.ordered_cases(plan,
lane, block)` is the single acquisition-order expansion used by both capture
receipts and this reader; it applies each block's forward/reverse group order
and policy order.

Expected evidence is 24 qualification reports/24 samples, 144 native
reports/4,320 samples, and 48 observer reports/144 samples: 216 reports and
4,488 samples. Native blocks have six process blocks, 30 retained samples, and
three warmups. Observer blocks have two process blocks, three samples, and no
warmups. The 24 qualification reports are kept separate from observer timing.
Each observer report retains 32 empty procfs controls. Controls and observer
operation counters are diagnostic evidence and are never subtracted from a
sample or pooled with native elapsed values.

Each report must retain the harness's warm and `cold-requested` filesystem
metadata, fresh-child and process-isolation flags, and the raw elapsed vector.
`cold-requested` is metadata about the requested harness state; it is not a
physical cold-cache claim. Logical source/output byte rates are descriptive
workload fields only and carry no physical-I/O or memory-bandwidth claim.

The reader checks the generic `Phase::timing_scope` string exactly as emitted
by the producer. That string describes the full phase-level atomic boundary
for every policy, including both synchronization steps. Policy attribution
comes from the exact `save_durability` field and the exact
`atomic_publication_steps` string: default omits `save_durability` and uses
the documented full route; full is explicit `Full`; file-only syncs only the
temporary file; and no-sync skips both synchronization calls. The default/full
pair is an explicit API-route control. File-only and no-sync are weaker
durability configurations whose differences remain semantic configuration
observations.

For each native selector, the reader computes nearest-rank p50, p95, and p99
inside each process block, then retains the raw harness p50 integer midpoint
and takes the median across six process blocks. Absolute and paired ratio
bootstrap intervals use 10,000 resamples, seed `820820`, and sorted ranks
250/9749. A policy's paired ratio is its p50 divided by the default p50 in
the same process block, format, and phase; the reported ratio statistic is the
median of those six ratios. Ratios are configuration attribution evidence,
not speedups, optimization measurements, sync syscall prices, historical
comparisons, or a reason to change the Full-by-default route. Lifecycle and
atomic values are not subtracted or treated as additive phases.

Artifact admission remains a separate correctness gate. It binds every flat
policy selector to the same admitted source digest, published digest, output
digest, edit outcome, and reopen/preservation evidence. The five untimed
artifact outputs per case are `default`, `full`, `file-only`, `no-sync`, and
`stream`; they do not add timing rows. The reader requires the publication and
edit-outcome vectors to have the planned count and to match the admitted hash.

Cleanup and sealing are owned by root. Derived analysis intentionally omits
the cleanup object so replay bytes remain identical before and after removal.
`validate.py --final` requires the 0820 cleanup witness and replays
`seal.json` only when it exists; `seal.py` remains the owner of seal creation.
Cleanup descriptors are checked as flat binary records sorted by path and
matched to the build descriptors. No 0819 evidence or prior broken reader is
an input to this packet.
