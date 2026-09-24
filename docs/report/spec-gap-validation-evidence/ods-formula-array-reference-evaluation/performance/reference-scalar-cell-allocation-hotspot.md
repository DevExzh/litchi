# Scalar single-cell reference allocation hotspot

This note records a source trace for the reference rows in the archived
candidate-release-01 diagnostic capture. It proposes one narrowly scoped
follow-up; it does not change the evaluator or claim a performance win. The
capture is a single shared-host window, and its value lanes have no A/B
baseline.

## Observed shape

The evaluate-phase p50 rows below come from
[the archived value raw CSV](diagnostics/candidate-release-01.tar.gz), also
reported in [candidate-release-01-analysis.md](candidate-release-01-analysis.md).
`repeat` is inside the timed batch, so allocation and timing counters are per
batch. `requested` is allocator requested bytes; `retained` is the evaluator's
post-evaluation memory reservation, not a transient allocator peak. `RSS` is
the process maximum RSS.

| case | repeat | reads / distinct | p50 ns | alloc calls | requested | retained | work | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| reference-range-256 | 8 | 2,048 / 256 | 339,502 | 80 | 369,600 | 22,528 | 4,152 | 3,328 |
| reference-repeat-256 | 8 | 2,048 / 1 | 1,278,686 | 6,240 | 1,085,248 | 0 | 8,176 | 3,260 |
| reference-distinct-256 | 8 | 2,048 / 256 | 1,281,486 | 6,240 | 1,085,248 | 0 | 8,176 | 3,196 |
| reference-range-1024 | 2 | 2,048 / 1,024 | 351,251 | 20 | 362,736 | 90,112 | 4,110 | 3,684 |
| reference-repeat-1024 | 2 | 2,048 / 1 | 1,283,996 | 6,172 | 1,082,320 | 0 | 8,188 | 3,984 |
| reference-distinct-1024 | 2 | 2,048 / 1,024 | 1,290,616 | 6,172 | 1,082,320 | 0 | 8,188 | 3,992 |
| reference-range-4096 | 1 | 4,096 / 4,096 | 772,973 | 10 | 722,040 | 360,448 | 8,199 | 4,428 |
| reference-repeat-4096 | 1 | 4,096 / 1 | 2,529,512 | 12,304 | 2,163,176 | 0 | 16,382 | 5,832 |
| reference-distinct-4096 | 1 | 4,096 / 4,096 | 2,598,451 | 12,304 | 2,163,176 | 0 | 16,382 | 5,996 |

The two scalar chains have identical allocation counts and requested bytes at
each scale even though their provider distinct-key counts differ. At 4,096,
`12,304 = 3 × 4,096 + 16`; at 256 with eight repetitions,
`6,240 = 3 × 256 × 8 + 96`. These exact affine shapes align with three fresh
single-item metadata buffers per scalar reference plus batch-level evaluator
activity. The source trace below identifies those buffers. This is evidence
from allocator counters and source inspection, rather than a claim that every
counter can replace a heap profile.

The range has one retained area and one output materialization for the whole
rectangle, so its 10 calls at 4,096 are not a like-for-like per-cell control.
The scalar chains immediately consume each one-cell reference and retain no
reference metadata (`retained = 0`). Their roughly three-times larger requested
bytes at 4,096 are consistent with the temporary metadata path, while RSS is
only a process-level signal and should not be attributed to those buffers
without a profile.

## Current allocation path

1. A reference AST node enters
   [`reference_value`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value/references.rs:49).
   A local reference is sent to
   [`reference_areas`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value/references.rs:69).
   For `Address::Cell`, `endpoint_area` resolves the sheet, converts the
   endpoint to a one-cell rectangle, and validates the provider extent before
   [`RuntimeAreaSet::direct`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value/references.rs:78)
   is called.

2. `RuntimeAreaSet::direct` creates an empty set and reserves one
   `RuntimeReference` record before pushing the lexical reference
   ([`direct`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:986),
   [`ensure_record_capacity`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:1013)).
   `areas.push(area)` then goes through
   [`push_raw`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:1109),
   which reserves the set's area vector, and
   [`append_record_area`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:1029),
   which reserves the record's public-area vector. Each vector starts empty
   for a fresh one-cell set, giving three potential heap growths.

3. The evaluator keeps those buffers only long enough to project the scalar.
   [`project_scalar`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:1633)
   dispatches an area to
   [`project_area_value`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:1730).
   That method performs the current-sheet probe, charges one scalar work unit,
   selects the single cell, and calls
   [`read_reference_cell`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:4440).
   Arithmetic reaches the same projection through
   [`map_binary`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:4266)
   and `value_to_slot`.

4. The cell read and conversion already carry the important policy. The read
   helper updates the cumulative reference-cell limit, checks cancellation
   before and after the resolver call, and passes the same execution context
   to `Resolver::read_cell`. [`read_to_element`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:4743)
   preserves typed `Unsupported(CellValue)` refusal and maps text through
   `TextValue::borrowed` without copying. The outer
   [`evaluate`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value.rs:1263)
   still fences the whole operation with source-version probes.

The three growths use the shared checked
[`ensure_capacity`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation.rs:1364)
helper. It checks the relevant maximum, cancellation, storage budget, and
fallible reservation before the vector grows. The temporary reservations are
dropped when projection consumes the area, which explains the zero retained
reservation in the scalar rows.

## Smallest compatible optimization

Add a private scalar single-cell projection path before
`RuntimeAreaSet::direct` is entered. It should resolve the existing
`endpoint_area`, admit one reference cell and one reference area, perform the
same current-sheet probe and one work charge, call `read_reference_cell`, and
finish through `read_to_element` and `element_to_runtime`. This removes the
three temporary vectors while preserving the existing resolver and conversion
helpers.

The selection must be based on scalar *demand*, not only `self.mode ==
Mode::Scalar`. A direct cell under `:`/`!`/`~` must remain an `Areas` value so
the reference operator can construct a range, intersection, or ordered list;
the operator path calls
[`coerce_reference_areas`](/home/zhuhe/code/litchi-spec-gaps/crates/litchi-ods/src/codec/formula/evaluation/value/references.rs:606)
and must not receive a scalar.
Likewise, a bare reference in matrix mode must continue to expose the public
`Value::Reference` view. A small internal scalar-demand bit on visit frames,
propagated through parentheses and cleared for reference-operator operands,
is one way to select the helper. An equivalent borrowed internal
`ScalarCell` token can be used if the implementation prefers to materialize
metadata only when `coerce_reference_areas` needs it. Returning a public scalar
directly from every scalar-mode `reference_value` call would break the
reference-operator cases covered by the integration tests.

The fast path must retain these details:

* `endpoint_area` remains the authority for sheet lookup, endpoint geometry,
  and finite extent validation. A missing sheet or out-of-extent endpoint must
  still become the existing scalar reference error without a provider read.
* Call `check_reference_cells(1)` and `check_reference_areas(1)` before the
  read. The latter is currently enforced indirectly by the metadata
  reservations; removing those vectors must not make a zero-area budget pass.
  Cumulative reads still go through `read_reference_cell`.
* Preserve the current-sheet resolver probe and its execution context, the
  order of cancellation checks, and the single-area implicit-intersection
  behavior. Named endpoints must continue to read the named sheet; current
  endpoints must continue to use the caller's position.
* Keep the source-version fence in `evaluate`, scalar work accounting, and
  remaining stack/array/text storage checks. Eliminating unneeded metadata
  reservations should reduce actual temporary storage; it must not bypass
  limits for the storage that remains. Low-storage tests should record any
  intentional change in refusal threshold.
* Leave `Address::Cells`, whole-row/column references, lists, source-qualified
  references, and matrix materialization on the existing paths. Lazy `IF`,
  `IFERROR`, and `IFNA` must visit and read only the selected branch; the
  scalar helper should not be called for an unvisited branch.

## Paired proof plan

Build one candidate containing only this scoped evaluator change and compare it
with the preserved candidate-release-01 value binary and input corpus. Use
serial CPU-6 AB/BA/AB windows with the same warmups, iterations, parsed-AST
evaluate boundary, and parse-evaluate control. Retain source, harness, binary,
command, and host hashes with the raw CSV and exact compare receipt.

The primary lanes should be `reference-repeat-{1,16,256,1024,4096}` and
`reference-distinct-{1,16,256,1024,4096}`. Keep these controls in the same
run: `reference-range-{1,16,256,1024,4096}`, direct numeric/text/empty/logical
and error cells, lazy selected and unselected references, cancellation,
reference-cell and area-limit refusals, a matrix bare-reference public-view
case, and a typed `Unsupported` cell. The range and matrix controls ensure
that retaining first-class reference metadata was not accidentally shortened.

Expected evidence is falsifiable and bounded:

* Scalar direct-cell allocation calls should lose approximately three growths
  per referenced cell in the timed batch. At 4,096, the observed 12,304 call
  shape should move toward the fixed evaluator baseline plus any remaining
  scalar-stack activity; the exact number is an implementation result, not a
  promised target.
* Requested temporary bytes should fall with the removed reservations, while
  `memory_retained_used` for the scalar chains should remain zero. RSS may or
  may not move; it is not proof of allocator savings.
* Result/status/checksum parity, resolver read counts, distinct-key counts,
  source-version behavior, work charges, typed failures, cancellation, and
  lazy branch reads must remain exact. Text pointer/borrow counters must show
  no new copy.
* Range and matrix lanes should remain within the paired review threshold. If
  allocations fall but p50 does not, report an allocation/budget improvement
  only. A CPU speedup or code-level cause requires a separate profile; this
  note supplies neither.
