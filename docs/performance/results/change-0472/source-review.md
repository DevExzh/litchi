# Elide owned tags for ordinary worksheet cells

The retained 0468 profile attributes 24.41% of inclusive commit weight to the
lossless snapshot scan, including `cell_address` at 4.93% and `wire::tag` at
3.69%. These contexts overlap. 0470 removes ordinary web traversal but leaves
this snapshot path intact; 0471 rejects a buffer-lifetime change after finding
no useful peak-memory reduction.

An ordinary unprefixed `<c r="A1">` currently retains owned element and
attribute names, an owned normalized coordinate and an attribute slice.
`write_cell` always removes the original `r` and emits its checked address.
Unedited cells retain their original byte spans and are copied verbatim.
The candidate therefore represents only exact unprefixed `c` tags with zero
attributes or one unqualified `r` by `None` in `Option<Tag>`.

`cell_address` still runs first and checks/infer coordinates. `cell_tag` still
iterates checked attributes in order, validates UTF-8 names, decodes and
normalizes every value, and rejects duplicates. A prior normalized `r` is
held locally in a Cow until any extra attribute forces the original owned
Tag representation, including attribute order and that `r`. Prefixed names
use the existing tag reader. The writer emits the same generated attributes
and uses `c` for payload/closing names only for the proven plain case.

This is an ephemeral representation change, not a skipped XML traversal,
source-borrow redesign, semantic Store shortcut or persistent cache. Unknown
attributes, namespace declarations, metadata, style and type attributes keep
owned tags. MCE processing, original spans, no-op behavior, source identity,
reversible patches, publication audit and the 4,096-cell / 1 MiB Store handoff
remain unchanged. No deliberate panic, unsafe code, public API, runtime or
container dependency is introduced.

| ADR | Obligation |
| --- | --- |
| 0001, 0003, 0006 | Preserve bytes, checked coordinates, duplicate/error order, no-ops, atomic publication and inverse semantics. |
| 0002, 0010, 0011, 0024 | Keep format-owned worksheet grammar and existing crate boundaries. |
| 0005 | Keep transient state bounded and measure allocation/peak-memory effects and every normal latency/RSS flag. |
| 0008 | Differential old/new tag tests and full XLSX plus scoped correctness gates. |

The frozen measurements cover six ordinary worksheet commit/save rows and a
payload-heavy PPT guard in 100/5 ABBA, a 201-row 15/3 default guard, and dense
one-percent 5/1 whole-process Heaptrack. No registered latency claim follows.
Measurements for arbitrary prefix-rich or attribute-rich worksheet populations
remain outside this ordinary synthetic timing matrix; focused differential
and existing preservation tests establish their exercised correctness only.

Independent review confirms the None case has no metadata attributes; the
shared-formula cm/vm guard now checks materialized tags and treats None as
attribute-free. All 16 focused codec tests pass, including six new tests and
the Option<Tag> size guard. Initial test attempts exposed malformed raw-string
delimiters and an incorrect expected duplicate-error wording; both failures
and the corrected final run are retained under `validation/`.
