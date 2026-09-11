# Independent review

**Verdict: approved for the documented scope, 2026-09-11.**

Two independent review passes (`theme_family_reviewer` and
`theme_family_final_review`) approved the corrected source. The reviewed files
are bound by `review-source-hashes.json`; the final gates bind the complete
production/test input set separately.

Review verified direct namespace-aware ownership, the admitted native and
normative extension identifiers, inherited namespace closure, scalar edits
preserving current opaque content, source-checked publication, and bounded
resource behavior. XLSB publication uses the OPC source-XML/owned-element splice
seam and preserves bytes outside the root; it refuses stale source or a
replacement that changes those outer bytes.

Issues discovered during review were corrected before approval:

- Legal `<?` content in comments and CDATA survives scalar editing; only an
  actual leading XML declaration is refused during child embedding.
- Decoded namespace declarations reject empty prefixed bindings and invalid
  reserved XML/XMLNS bindings, including default and escaped aliases.
- Namespace scope is capped at 257 active bindings; scope copies are limited
  to recognized candidates, and inherited declarations are removed in one
  bounded pass. Insertion containers use exact output-size preflight checks.
- A validated standalone UTF-8 BOM is stripped before child insertion. The
  reviewer's corrected probe adds and reopens the family with zero embedded
  BOMs.
- Direct owned extension containers reject non-whitespace text, CDATA, and
  character references; XML whitespace and opaque foreign descendants remain
  accepted.

Empty `ext`/`extLst` wrappers deliberately remain after family removal. This
preserves source markup and is documented behavior. Complete MCE branch editing,
full Theme rendering, durable family patches, other package hosts, and native
Office acceptance are outside the approved scope.

The final reviewer reran 22 complete-part and 14 XLSB family tests successfully.
A temporary tester report counted 23 complete-part cases before a duplicate
regression was removed; final gate receipts are authoritative. Review probe
files and directories were removed by their owner.
