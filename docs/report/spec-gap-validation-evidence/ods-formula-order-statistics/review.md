# Independent order-statistics semantic review

Status: **PASS** for the frozen production implementation and the staged
contract. This review covers `MEDIAN`, `MODE`, `LARGE`, `SMALL`, `PERCENTILE`,
`PERCENTRANK`, `QUARTILE`, and `RANK` against the repository-local ODF 1.4
Part 4 source and `contract.md`. The normative archive SHA-256 is
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`; the
Part 4 HTML member SHA-256 is
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The frozen production inputs are unchanged. Their handoff SHA-256 values are:

| Input | SHA-256 |
| --- | --- |
| `evaluation.rs` | `4a7a3ec13a425ed44432a220a8e14c75c8781934cdce6b9632b085a5e9967c1d` |
| scalar `evaluation/order.rs` | `516cc2f144de18522ef40f708e542573a259287aff9cfa0741c04beb3f86d86a` |
| `evaluation/rounding.rs` | `82a46b6ba227a79979f284e9fbaf9838626fe2a1cad3bafdac968c77a67a6d03` |
| `evaluation/value.rs` | `190609b5036a3a64a0ac035f9956e215dbb4c20d3152257441e6ed803f7a46e3` |
| resolver `value/order.rs` | `9f66676154ed7945f02d7e4cdb321d9c5ae9d5e8c61ef3d2cf9ab991a288a88b` |
| resolver `value/matrix.rs` | `95ac99a7ded56e77966029f9cc1433aa0b70fe07f1a7f30020589dc973e3a5db` |
| semantic tests | `8e464969151e6ebf27f3bfd9c7a27932885581c4f7016e3b801914e1d9f183c2` |
| contract | `d57a13579f769d35d750c95c3480c382ce6df4888c223bb3fc7a6145bd61611d` |

The scalar kernels implement the specified tie and domain rules. MEDIAN uses
the fixed-width numeric average for its even midpoint; MODE selects the
smallest value among tied frequencies and reports the selected no-mode profile
error; LARGE and SMALL require exact positive integer ranks and retain duplicate
positions. PERCENTILE and QUARTILE use the sample rank
`1 + X * (n - 1)`, safe interpolation, exact quartile bounds, and canonical
`+0` for direct endpoint selections. PERCENTRANK uses the first occurrence of
duplicate `X` values, safe adjacent interpolation, default significance `3`,
and the shared ROUND kernel. RANK uses competition ranking, descending order
for zero and ascending order for every nonzero finite order value.

The signed-zero policy is consistent: directly selected zeros publish `+0`,
while computed midpoint or interpolation underflow retains its computed sign.
The interpolation path does not form an overflowing endpoint difference.

Matrix behavior was checked independently. Complete sequence/reference
arguments remain whole descriptors, while scalar Number/Integer parameters
lift and broadcast in matrix mode and project in scalar mode. LARGE and SMALL
retain the explicit Array `N` result with its shape through scalar publication,
parentheses, and scalar-condition IF. The §3.3.2.2.1 rule applies to the
input of an Array-returning function: MUNIT's own array-valued size expression
uses its `[0,0]` input element. Once MUNIT has produced an Array, a scalar
consumer such as `PERCENTILE(Data;MUNIT(2))` applies ordinary matrix lifting;
the focused assertions cover matrix, scalar, and projected-IF results.
MUNIT remains position-sensitive in conditional criterion evaluation and is
excluded from complete-argument demand propagation.

Formula errors retain source argument and cell order, and a retained formula
error does not hide a later typed resolver, cancellation, source, resource, or
allocation failure. Missing optional slots are errors; omitted PERCENTRANK
significance and RANK order use their defaults. The scalar profile and
resolver-backed profile agree on conversions, empty and no-mode behavior, and
error subtypes selected by the contract.

Evidence reviewed at handoff includes 17 semantic tests, 16 resource-limit
tests, 2 native tests, and 8 independent oracle tests, plus the isolated
1,421-test ODS run with zero failures or ignored tests. The refreshed seven
gates passed with the added MUNIT assertions and updated contract, including
strict clippy, rustdoc, formatting, boundaries, and source-diff checks. Root's
final verification receipt records all 4,290 performance samples as passing
with stable inputs ([verification receipt](verification-receipt.json)). The
separate threshold review retains one accepted-with-flag SUMIFS control
observation at +6.39% parse-evaluate; its evaluate phase is +0.53%, work,
reads, checksums, and allocation counts are unchanged, and the retained
bootstrap interval crosses zero ([threshold review](performance/results/threshold-review.json)).
That root performance disposition is separate from this semantic PASS and is
not a causal regression claim about the order reducers. No production or test
files were edited by this reviewer.
