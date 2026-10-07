# 0954 census review

Reviewed 2026-10-07T10:04:45Z. This was read-only: no build, workload, or
repository source edit was performed.

## Reproduction

From `/home/zhuhe/code/litchi`:

```sh
python3 docs/performance/results/change-0954/census.py \
  /home/zhuhe/litchi-performance-0947/inputs/artifacts \
  --output /tmp/census-0954-recomputed.json
sha256sum /tmp/census-0954-recomputed.json \
  docs/performance/results/change-0954/census.json
cmp -s /tmp/census-0954-recomputed.json \
  docs/performance/results/change-0954/census.json
```

The recomputed JSON is byte-identical: SHA-256
`4edde4f45601f95057264b0bdafa3b219614129dbbab862af5e45bec30a6a719`.
The census script SHA-256 is
`cbd01b77b9365237f489a0d299a1ea2b4f681c73d1393339b12a49039914ebac`.

Independent ZIP inspection reproduced the changed-member sets, verified every
untouched compressed payload byte-for-byte, checked Deflate methods and decoded
lengths, and reproduced these aggregates:

| fixture | tasks | total | largest | remainder |
| --- | ---: | ---: | ---: | ---: |
| medium one-percent | 5 | 269,285 | 68,488 | 200,797 |
| medium one-edit | 2 | 63,821 | 63,294 | 527 |
| dense-sparse one-percent | 5 | 508,961 | 462,569 | 46,392 |

The retained medium and dense-sparse source/output SHA pairs match the 0953
source census identities. The 0954 artifacts contain 12 untouched members for
each five-task primary and 15 for the one-edit control.

## Policy verification

`write_limits` in `crates/litchi-opc/src/source_backed.rs:5431-5446`
uses the ZIP defaults. The current ZIP source defines the warm threshold as
8,192 bytes and the cold local-pool threshold as 65,536 bytes
(`crates/soapberry-zip/src/office/parallel_write.rs:21-44`).
`ParallelWriteSession::wave_for` selects the cold threshold when there is no
caller facility and no already-built local pool, and the warm threshold when a
facility or existing pool is present (`parallel_write.rs:243-260`).
`PreservationPlan::deflate_wave` then requires at least two Deflate tasks and
caps width by workers, max in-flight tasks, task count, and a positive
per-task floor (`preserve.rs:430-459`).

The census models `min_task_bytes=0` and sets max in-flight equal to workers,
so its advisory rows are consistent with that exact policy: medium is cold
parallel at widths 2/4/5 for workers 2/4/8; dense-sparse is serial cold but
warm-capable at 2/4/5; one-edit is serial in both states. These rows do not
model other worker/in-flight combinations, finite budget refusal, worker-state
memory, or actual execution. “Warm” means session/pool lifecycle, not CPU or
filesystem cache warmth. No performance or adoption claim follows from the
census.

No material correction is required. The claim boundary is sound provided the
reported widths remain labeled advisory calculations rather than observed
preservation-plan execution; the existing README/report already states that
boundary.
