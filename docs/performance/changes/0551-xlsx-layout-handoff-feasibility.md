# 0551: select compact cell proofs for XLSX layout reuse

`performance_claim: none; retained-profile reanalysis and source audit`

Replaying eight sealed 0550 commit dumps shows that named XML reader and
namespace dispatch account for 29.04–34.39% of scanner instructions. Cell
setup remains substantial, especially for prefixed/typed cells. Retaining a
full `Layout` would move its tag/span construction into planning and extend
its lifetime; it would not eliminate that work.

A source-bound corpus calculation also rules out edited-row reparsing as the
selected approach: the dense/sparse 1% workload touches every dense-sheet row,
so those rows contain 16,769 of 17,792 cells. The next implementation direction
is compact per-cell source offsets with changed-cell tag materialization.

The source review inventories all `Layout` fields and scanner refusal
families. Completed source validation proves several hazards absent, but its
attribute allowlist does not establish the scanner's complete decoding and
normalization contract. Preserving latent errors on unchanged cells is a
mandatory differential test obligation.

The [evidence bundle](../results/change-0551/README.md) retains deterministic
replay, input hashes and independent review. No candidate or runtime speedup
is admitted. Production/harness source and the performance baseline are
unchanged. OLE2/OOXML remain first, ODF is deferred and iWork excluded.
