DOCX Ink profile validation
==========================

Implements the MS-ODRAWXML EMMA placement boundary and effective brush defaults while retaining original lexical source values. Recognized numeric and boolean properties use XML Schema whitespace rules; normative length units are supported. Ignored wrappers remain opaque. Full trace and EMMA semantic closure are separate work.

The original v1 full gate passed 1,669 ordinary tests and 77 doctests. V2 passed 188 focused Ink tests after units and whitespace corrections. V3 changes only DrawingML authoring readback and passed 14 authoring tests, including finish/reopen with XML whitespace. Strict all-target Clippy, warning-denied documentation and scoped formatting passed for the final source. These are incremental gates, not a claim of a full v3 suite rerun.

Independent review by calc_chain_review cleared v3. Generic custom-property attribute whitespace is a separate pre-existing authoring issue; this change compares recognized typed values using their schema normalization rules.

Historical v1/v2 receipts are retained unchanged under history-v2, including their original absolute provenance paths. Current source and root logs are identified by the v3 manifests.
