# 0362ul94o — XLSB shared-error match repair

This source-only follow-up starts from the verified `11c3` tip
`11c3f2fc8c26f3480c887422d49ea35e551dbc68` (tree
`00295793e7347cedd2b43bf3bb0918f17888e933`).  The leaf clippy row stopped at
the XLSB host error conversion because its wildcard arm hid known shared
error variants.  This repair makes no acceptance claim and does not rerun the
gate.

The XLSB `From<litchi_ooxml_common::Error>` conversion now lists every known
non-promoted common variant explicitly, then retains a final wildcard arm for
future `#[non_exhaustive]` additions.  Refused, patch-conflict, signed,
source-changed, and internal errors keep their existing promoted typed arms;
all other current and future variants still become `Error::Common` unchanged.
The VBA arm remains feature-gated.  A source-only audit also checked the
equivalent DOCX, PPTX, and XLSX conversions; they retain their existing
fallbacks because this local wildcard-deny lint is specific to XLSB.

Only rustfmt, diff, and repository size/binary guards are run for this repair;
Cargo, rustdoc, compiler, native, and performance execution remain pending.
