# 0362ul69a — Formula raw-policy review return

Date: 2026-10-25. Parent: 0362ul69. Status: source repairs complete; the
coordinator still owns Cargo, rustdoc and performance validation.

The RETURN0499 review found four admission/correctness gaps in the initial
Formula policy migration. This follow-up closes them without changing the
MTEF or OMML wire spelling or semantic AST conversion:

- MTEF preflight now charges the exact eight-byte delta for a top-level
  `LineBreak` pile. A bare one-break stream is 20 bytes; a 19-byte output
  policy is refused before reservation. Preflight depth uses the same exact
  boundary as `NodeWriter::enter`, so an exact nested depth cannot reserve and
  then emit partial framing. Boundary tests cover both cases.
- MTEF lookup spellings and codepoint/fallback spellings are checked before
  their retained `String` is created. Embellishment replacement computes the
  projected result before `replace` allocates it, including multiple
  placeholders. Focused tests cover a mapped `\\alpha ` spelling and a dot
  embellishment at their string ceilings.
- Template substitution now carries the selected MTEF policy. It checks the
  aggregate rendered leaf spelling before every destination growth, and checks
  the final derived spelling before the fallback AST text clone. This closes
  the unchecked child concatenation path while preserving structured template
  recognition.
- OMML `TextBuffer` admission is cumulative across Text and CDATA events.
  Trimming is inspected before an owned trim can be allocated, then the
  complete retained buffer is checked before append. A split Text/CDATA/Text
  regression proves the exact three-byte result and the two-byte refusal.

The review correction that predefined/numeric quick-xml entities cannot expand
beyond their raw event bytes is retained: the existing raw event precheck and
post-decode check remain, with no redundant entity-length pass added. This
keeps the hot OMML text path bounded without a needless decode scan.

Source-only checks for this return were direct `rustfmt --edition 2024` over
the changed Formula sources, `git diff --check`, and tracked-source searches.
No Cargo/native/rustdoc command was run in the leaf, and no binary, large
evidence artifact or upstream checkout path is part of the change.
