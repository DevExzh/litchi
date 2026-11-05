# 0359qd — DOC body-toggle final publication in place

- **Status:** source repair committed; matched ordinary Ir, output, and heap
  receipts recorded. Native qualification remains pending under host load.
- **Date:** 2026-10-22.
- **Input:** [0359qc](0359qc-doc-body-toggle-repair.md).
- **Implementation:** `ae094d7d055c59aa9980d52339000b959689d376`.

## Repair

The final body commit now calls a private in-place finish on the owned
`RevisionEditor`. Pending CHPX is flushed, then the owned WordDocument, table,
and optional Data vectors are moved into the lower object editor's shared
stream replacements. The prior final route cloned the complete revision
editor before taking those vectors. Candidate renders used by validation and
interleaved operations still clone first, so their failure-atomic behavior is
unchanged; only the consuming publication route moves buffers in place.

The existing mixed-format/table, caught-refusal, inverse, and exact no-op test
continues to exercise the final body commit. The source repair was built from
the same probe consumer and produced the protected output SHA.

## Matched ordinary qualification

The exact selected route was `LITCHI_PROBE_ONLY=doc/leaf_write`,
`LITCHI_PROBE_BATCH_ROUNDS=1`, runner argument `1`, with the stable runner
path used by 0359qb and 0359qc. The output SHA is
`6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11`, equal
to both protected legs.

The candidate route-local Callgrind Ir is `257,805,703`, or `1.004829412` of
the original `256,566,637` and `0.996268196` of the 0359qb after leg
`258,771,387`. `Package::Edit::commit` is `79,832,201`, or `1.088973577` of
the original `73,309,585` and `0.988289311` of the previous after leg
`80,778,169`. This removes the final full-editor clone cost measured in
0359qc; the remaining route delta stays open for a separate repair.

The largest changed commit callees against the original are:

- `Document::from_ole_with_options`: `59,887,476` Ir, down from `60,195,583`.
- `RevisionEditor::open_from_package`: `7,340,384`, up from `7,287,067`.
- `Edit::put_streams_shared`: `4,803,020`; the old setter's
  `commit_candidate` was `4,789,219`, so deferred publication moved this
  work to the final commit.
- `memcpy`: `1,999,420`, up from `1,358,348`; this `641,072` Ir increase is
  retained as the next concrete optimization target.
- `RevisionEditor::rewrite_chpx`: `653,486`, comparable to the original
  setter's `655,124`; `Arc::from_box_in` is `678,625`, comparable to the
  original `678,715`.

The matched Massif peak is `7,910,440` bytes (`7,863,939` useful heap and
`46,501` extra), or `0.961329249` of the original `8,228,648` and
`1.038094778` of the previous after peak `7,620,152`. The lower peak against
the original is retained as evidence; the increase against the previous
after leg is open and is not waived.

The build used toolchain `1.95.0`, one Cargo job, incremental compilation
disabled, release debuginfo disabled, and the existing 20 GB systemd cap.
Native timing was not run because unrelated host builds were active.

The compact receipt is in
[`results/0359qd/qualification.json`](results/0359qd/qualification.json).
