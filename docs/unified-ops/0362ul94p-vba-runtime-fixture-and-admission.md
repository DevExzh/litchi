# 0362ul94p — VBA runtime fixture and admission repair

This source-only repair starts from the composed OOXML runtime candidate
`3774fe4374dfb66e74b98b1ae742672ed6cad37f`.  The coordinator's runtime
receipt reached 511 passing tests and 17 failures.  The owned ten failures
are recorded in
`(deleted-coord-artifact)/0000.log`
(222,716 bytes, SHA-256
`668db8f2afc172f4abacf8ce06d36934cc1ca6a89d3ee7c404cc805abe5a8153`).

Eight VBA source-publication tests used the literal bytes `workbook` or
`document` for their main XML part.  Those are intentionally rejected by the
OPC source-publication boundary as character data outside a document element,
so the tests never reached the graph, replay, or limit assertion they were
intended to exercise.  The shared Excel and Word test-package fixtures now
use minimal, correctly namespaced document roots; the dangling-reference
fixture's auxiliary XML owner is valid as well.  This changes only test input,
not the production refusal for malformed source XML.

The other two failures exposed a production staging omission.  `remove()`
performed recipient metadata admission but discarded its `RecipientAdmission`
summary.  A subsequent changed commit therefore reported an internal
missing-summary error before publication.  `remove_inner` now retains the
summary under the same first-admission rule as `store_inner`; reset and the
no-op proof path remain unchanged.  Signed changes still reach the existing
typed signed refusal, and admission still occurs before graph staging.

The VBA source and its DOCX, PPTX, XLSX, and XLSB consumers were audited in
this leaf; no consumer API or limit policy was changed.  Only source diff and
repository-size/binary guards are run here.  Cargo, rustdoc, compiler,
native, and performance execution remain pending for the coordinator's
runtime rerun.
