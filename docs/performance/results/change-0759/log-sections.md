# Log sections for change 0759

## For `HOTSPOTS.md`

## 0759 — the spec-gap branch merged; three costs it brings

[0759](0759-spec-gap-branch-merge.md) merges `feat/spec-gap-implementation`
(`a67a38abf2`) into this branch as merge commit `f592ecc1b0`. The merge keeps
the incoming safety checks, and three of them add work that later records
should measure:
- **ZIP admission.** It reads one 8-byte local-header prefix per member before
  any part is read, so positional and range sources pay one small read per
  member. The 0623 and 0632 read savings otherwise stand, and the test shapes
  are unchanged.
- **Whole-package byte ceilings.** Explicit Custom Data, pivot, SVG-lifecycle,
  InkAction, DOCX ink and effects operations decode every part to charge
  inflated bytes. A declared-size accessor on `PartMetadata` would make them
  lazy.
- **Incoming open-path checks.** ODG and OPC opens now run them; our records
  0502 and 0504–0507 were measured before them.

The DOC fresh-writer corpus has a new identity (Dop2002 fix), so
`doc_fresh_write_to` needs re-baselining before it is compared again.
[Evidence](results/change-0759/README.md).

## For `REPORT.md`

## 0759 — the spec-gap branch merged, with 20 recorded judgements

[0759](0759-spec-gap-branch-merge.md) merges 624 spec-gap commits into this
branch and resolves 65 conflicts.
- **Correctness-first choices:**
  - incoming ZIP admission (encryption and entry count) on every OPC ingress;
  - stricter CFB, DOC and ODG checks;
  - caller XML audited with `VerifiedSource`;
  - Reuse layouts that fall back when directory metadata changed.
- **Performance work kept:** per-member Deflate (0618, ADR 0031),
  decode-on-use for lazy parts (ADR 0030), our pristine publication proofs,
  the 0744 reduced readback and the 0754 tracker scanner.
- **Gates:** 13,221 tests pass in the 14 gate crates, and 5,794 in the ODF,
  RTF, crypto, VBA and XLDM crates.
- **Pre-existing failures carried in, each proved on a tip:** two
  incoming-branch harness tests and the non-iWork inventory. No performance
  claim. [Evidence](results/change-0759/README.md).

## For `GOAL_AUDIT.md`

## 0759 — spec-gap merge: both branches' guarantees kept, three items for the owner

[0759](0759-spec-gap-branch-merge.md) integrates the spec-gap branch under
0652's rule: correctness and safety first, and accepted ADRs bind.
- **Kept:** every feature, fix, test and validation from both sides, except
  two incoming `DeflateWorkspace` tests for a mechanism ADR 0031 and 0618
  exclude. Test inventories were compared name by name.
- **For the owner:**
  - re-baselining the DOC fresh-writer corpus identity;
  - registering `litchi-xldm` in the non-iWork gate;
  - which refusal message the incoming PPTX harness test intends.
- **Still open:** coalescing the admission reads, and declared sizes for
  whole-package ceilings.

The non-iWork goal remains active. [Evidence](results/change-0759/README.md).
