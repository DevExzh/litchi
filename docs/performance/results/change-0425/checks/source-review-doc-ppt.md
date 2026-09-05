# 0425 DOC/PPT source review

This is a source-only review of the current worktree changes under
`crates/litchi-doc` and `crates/litchi-ppt`, compared with baseline
`340cc91ae2bdec338dfe7682b4b5d8c219a2d288`. The review covers the constant
width chunk migration and the requested iterator, remainder, error, and
UTF-16 behavior. It does not make a build, test, lint, profiler, or workload
claim.

## Findings

No correctness blocker was found in the visible DOC/PPT migration.

`as_chunks::<N>().0.iter()` and `as_chunks_mut::<N>().0.iter_mut()` enumerate
the same complete fixed-width groups, in the same order, as the corresponding
`chunks_exact(N)` and `chunks_exact_mut(N)` calls. Every migrated width is a
positive compile-time literal or format constant; no runtime-width iterator
was changed. The paragraph scanner keeps its checked byte count and reads the
same UTF-16 units in order (`crates/litchi-doc/src/body_text/source.rs:1703-1724`).
The fixed 336-byte encryption buffer remains sixteen 21-byte blocks
(`crates/litchi-doc/src/encryption/codec.rs:528-533`).

The one parser that intentionally observes a tail keeps that behavior. PPT's
persist directory parser now retains the tuple remainder and still reports an
unaligned directory after parsing all complete records; missing offsets and
zero entry counts still return at their original points
(`crates/litchi-ppt/src/persist/ptr_holder.rs:58-101`). The other loops either
operate on a slice whose size was already checked, such as DOC table border
records (`crates/litchi-doc/src/parts/tap_parser/codec/cells.rs:287-301`) and
PPT RecolorInfo entries (`crates/litchi-ppt/src/recolor.rs:280-305`), or retain
the old deliberate incomplete-tail behavior. In particular, DOC Unicode text
still truncates an odd byte before iteration
(`crates/litchi-doc/src/parts/text.rs:406-417`), while PPT's lossy text helper
still consumes only complete UTF-16 units and stops at the first zero
(`crates/litchi-ppt/src/text/extractor.rs:32-44`).

The UTF-16 error and boundary contracts remain in place. DOC paths that use
strict decoding retain their prior even-length checks and `String::from_utf16`
or `decode_utf16` error mapping. PPT animation string variants retain their
odd/even payload checks. PPT embedded-font EOT names still reject odd lengths,
checked range overflow, truncation, and malformed surrogate sequences before
the fixed-width iterator (`crates/litchi-ppt/src/font/model.rs:646-671`). The
migration therefore does not turn an incomplete or malformed UTF-16 field into
a valid one, change error ordering, or consume padding as data.

Mutable fixed-width rewrites preserve the same bounded slices and element
order. The OfficeArt property readers and writers use exact
`instance * 6` ranges before iterating (`crates/litchi-ppt/src/slide_order.rs:
4353-4390`), so their remainder is empty by construction. Test and helper
decoders use the same complete-pair semantics and do not alter production
format behavior.

## ADR and scope checks

The migration remains compatible with the accepted constraints relevant here:

- ADR 0003's immutable snapshot and atomic publication boundaries are not
  touched; these are local parsing and encoding loops.
- ADR 0005's bounded positional I/O and measurement contract remain intact:
  checked ranges, limits, and tails are retained, and this change carries no
  performance claim.
- ADR 0006's preserve-by-default and fail-closed validation behavior remains
  intact; no malformed-input refusal was removed and no unknown bytes are
  reinterpreted.
- ADR 0008 and change 0425's strict non-iWork verification scope are respected;
  the review covers only the requested DOC/PPT downstream constant-width
  migration and does not touch iWork sources.
- ADR 0022's PPTX embedded-font ownership is outside this DOC/PPT change, but
  the legacy PPT EOT validation preconditions observed here remain strict.

## Result

The complete visible DOC/PPT constant chunk migration is source-consistent
with the baseline and has no review blocker.
