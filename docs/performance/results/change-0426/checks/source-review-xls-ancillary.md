# Change 0426 XLS Formula `RgbExtra` source review

This is an independent source-only review of the Formula ancillary-data
change against baseline `f22917bbe973709e93dba9d1d4bc85a4dc0778ce`, the
accepted ADRs, the checked-in `[MS-XLS]` sections 2.5.198.3, 2.5.198.61,
2.5.198.70, 2.5.198.103, and 2.5.209, and
`xls-ancillary-design-review.md`. I did not run Cargo, a build, tests, a
fuzzer, or a CPU/allocation measurement.

The implementation has the right ownership shape. The private
`Ancillary` owner retains the source cell, original token bytes, and exact
validated suffix behind one `Arc`; ordinary formulas with an empty suffix do
not scan tokens or allocate an ancillary owner. The parser accepts
`cce <= payload.len() - 22`, validates a nonempty suffix to exact
consumption, and attaches the owner through `Metadata`, so both strict and
defect-preserving Formula paths reach the same validation.

## Findings

### Initial blocker: ordinary Name and 3-D tokens were rejected before supported extras were scanned

At the initial review point, `scan_formula_extra_kinds` returned
`UnsupportedFeature` immediately for every `PtgName` (`0x23`) and every
`PtgNameX`/3-D form (`0x39..=0x3d`) at
`crates/litchi-xls/src/formula_metadata/extra.rs:121-187`. The local RgbExtra
grammar requires revision structures for those Ptgs only when the containing
formula is in revision context. A normal CellParsedFormula can therefore have
one of these ordinary tokens and a later `PtgArray` or `PtgMemArea`, with the
suffix containing only the supported Array or Mem structure. For example, a
framed ordinary RPN stream consisting of a value-typed `PtgName` (five bytes),
a value/array-typed `PtgArray` (eight bytes), and `PtgAdd` can have one valid
`PtgExtraArray`; it is rejected at the first token even though no revision
structure is present or needed. The analogous nonrevision `PtgRef3d` plus an
Array extra is rejected for the same reason.

The scanner needs an explicit ordinary/revision context, or the ordinary
CellParsedFormula path must frame these tokens and leave them without an
extra. It can still return a typed refusal when an actual revision-owned
structure is selected by a revision-aware caller. This is a compatibility
failure for valid mixed formulas, rather than an intended rejection of an
unsupported extra. The `PtgElf` branch at lines 188-198 has a similar
conservative behavior for all ELF subtypes; at minimum it should distinguish
the subtypes that actually require `PtgExtraElf` from ordinary ELF tokens, or
document that the no-context path intentionally refuses any nonempty suffix
containing ELF.

The scanner also treated every classified `0x20` form as `PtgArray` without
checking the required type bits. The local grammar requires PtgArray type 2 or
3, so `0x20` plus an otherwise valid array suffix could be accepted as a
corresponding Array structure. That was a smaller fail-closed gap beside the
context issue; the strict cell-formula scanner now checks those bits while the
list-object adapter retains its historical permissive framing.

### Formula suffix ownership and parsing are otherwise sound

`formula_metadata/extra.rs:131-228` bounds the token scan, rejects reserved
high bits for ordinary cell formulas, rejects zero or overrun `PtgMemArea`
`cce`, and requires each memory expression end on a framed token boundary.
The extra pass reads each `PtgExtraMem` count with checked arithmetic and
validates every `Ref8U` row/column ordering and `0x00ff` column ceiling.
`PtgExtraArray` dimensions use checked multiplication and reuse the strict
SerAr validator, including finite Xnum values, Boolean/error restrictions,
UTF-16 validity, and exact value consumption. Invalid SerAr errors are mapped
back to Formula record context at lines 277-283. Empty suffixes deliberately
skip this scanner, preserving the pre-existing opaque-token acceptance.

`cce == 0` is now rejected for `PtgMemArea`, satisfying the checked-in
Formula-reference policy and avoiding a zero-length memory expression being
treated as a valid owner. The scanner remains a framing pass rather than a
complete RPN/stack validator, which is consistent with the existing no-extra
opaque contract; the nonempty suffix path nevertheless rejects truncation,
unowned tail bytes, malformed memory ranges, malformed SerAr values, and
record payloads above 8,224 bytes.

### Source-bound writer preflight is correctly ordered and atomic

At `writer/biff/cells.rs:213-306`, validation first resolves shared/array
owner substitution, then compares the actual emitted cell and token stream
to the private ancillary owner. A mismatch returns `Error::UnsafeEdit` before
the record header or any payload is written. The exact suffix is appended
after the unchanged token stream, while the 8,224-byte limit includes both
tokens and suffix and `cce` remains the token length. Existing unit coverage
exercises original-cell/token success, changed cell, changed tokens, shared
owner substitution, and array owner substitution.

The semantic handoff clones Formula metadata into `Cell`; array-owner binding
only changes the existing array/shared fields and does not discard ancillary
metadata. Cache-only edits modify the fixed eight-byte cache field in place,
so a retained suffix remains byte-identical.

### Structural and canonical-resource preservation paths fail closed

`cell_values/structural.rs:2386-2400` uses the common framing helper. A valid
nonempty suffix is validated and then returns `Error::UnsafeEdit` before token
patching; a malformed suffix returns `InvalidData`. The check runs on the
candidate payload used by the structural certification pass, so the source
snapshot is unchanged on refusal.

`cell_values/mod.rs:4123-4154` fingerprints Formula bytes through the end of
the record, including `rgcb`. `authored_formula_record_matches` at
`:4157-4205` reads the complete record using its BIFF payload length, validates
any suffix, and returns typed `UnsafeEdit` for a valid nonempty suffix before
canonical no-suffix matching. The full record comparison also checks equal
record length, so it cannot match a source suffix through a prefix comparison.
The worksheet index validates Formula framing and every nonempty suffix before
cache entries are exposed at `:5910-5958`.

A focused internal regression now exercises
`authored_formula_record_matches` against a valid ancillary Formula and asserts
the typed `UnsafeEdit` plus unchanged source bytes. The public preservation
fixtures separately cover cache/style edits and structural refusal. The helper
reads the complete BIFF payload length, validates the suffix, and compares the
complete record, so a canonical no-suffix record cannot match through a
prefix.

The list-object adapter uses the extracted scanner with its old permissive
high-bit, memory-cce, and unsupported-token settings, then retains its own
array/memory payload parser and list-specific error mapping. This keeps the
list formula policy separate from ordinary CellParsedFormula validation.

### Memory/layout observation, without a performance claim

The no-extra path performs no ancillary scan or ancillary allocation, but
adding `Option<Arc<Ancillary>>` to `Metadata` adds one pointer-sized field to
every metadata value, including ordinary formulas with no suffix. The retained
owner's combined `Vec` removes a separate token/suffix ancillary allocation,
while its source buffer still duplicates the Formula token bytes already held
by `CellRecord::Formula`; this is an ownership tradeoff, not a measured
throughput or RSS result. The baseline layout probe recorded `Metadata` 24 B,
`CellRecord` 80 B, and `Cell` 152 B; no candidate layout or performance
measurement is claimed by this review.

## ADR applicability

ADR 0003 requires immutable snapshots, source-bound identity, and atomic
publication. The private owner and pre-write writer check follow that rule;
the structural refusal is staged against a candidate before publication.
ADR 0005 requires finite resource bounds, fallible allocation, and measured
evidence before performance claims. The Formula payload ceiling, checked
arithmetic, and `try_reserve` calls satisfy the bounded path; this change has
no performance claim. ADR 0006 requires preservation by default, inert formula
handling, and typed fail-closed refusal. Exact suffix retention and the
`UnsafeEdit` paths satisfy it. ADR 0008 makes compile/test/evidence gates part
of migration; those
gates are owned by root and were not run in this review. ADR 0012 requires
checked BIFF8 references, zero-length `cce` rejection, and panic-free typed
failures; the new memory and Ref8U checks align with it. ADR 0024 is a topology
inventory; the private `formula_metadata` owner and list adapter remain within
the documented XLS ownership boundary.

## Source coverage reviewed

I inspected `xls_formula_ancillary.rs`, including the native five-tail census,
eager and source-backed access, exact no-op/cache/style preservation, and
atomic row/column refusal, plus
`xls_formula_ancillary_rejections.rs`, including mixed ordered extras,
truncations, Ref8U and memory-cce failures, strict SerAr mutations, legacy
empty-suffix opaque acceptance, and the exact 8,224/8,225 payload boundary.
These files were not executed during this review, so this report makes no test
pass claim.

## Follow-up resolution

The initial review findings are retained above for audit history. The final
source re-read resolves the two parser findings: ordinary Name/NameX/3-D
tokens are framed at their ordinary widths without requiring revision extras,
and strict cell-formula PtgArray framing accepts only value/array operand types
2 and 3. The list-object path keeps its historical permissive high-bit,
array-type, and memory-expression policy. A nonempty ELF tail remains an
explicit typed unsupported scope item; ELF subtype grammar is outside change
0426. The canonical-resource refusal regression is now present in
`cell_values/tests.rs`.

Review result: **no remaining production blocker in the reviewed 0426 scope**.
This is a source-only result; no Cargo command, build, test, fuzzer, or
performance measurement was run by this reviewer.
