# 0834 — final results review

This review covers the retained v3 qualification and custody packet. It is an
evidence review, not a performance claim or formal-capture admission.

The final reader replay is internally consistent: 12 reports and 25 samples
are retained across the diagnostic, partial v2, and final v3 stages. The final
v3 stage has 16 validated samples, no failed rows, and `qualification_valid:
true`; formal reports and formal samples remain zero. The earlier v2 failures
remain retained as failure evidence rather than being rewritten.

The final PPTX source lifecycle report demonstrates the intended bounded proof:

* warm uses the pinned 17,017,139-byte source and its pinned hash;
* cold-verified uses the 17,018,880-byte aligned source, page size 4096,
  EOCD offset 17,017,117, and 1,741 bytes of comment and zero-suffix padding;
* the exact 65,536-byte tail read starts at 16,953,344 and occurs once in the
  open phase; its 16,230 bytes of unselected-slide overlap remain visible;
* the semantic query fully covers the 522-byte selected slide with no
  unselected-slide or media overlap; classification is
  `selected-slide-only:aligned-eocd-tail-metadata-probe;target-slide-no-unselected-or-media-overlap`;
* the cold verifier is eligible, observes zero resident bytes before the
  operation, 282,624 resident bytes afterward, and a positive 282,624-byte
  process `read_bytes` delta. Its claim scope explicitly excludes a physical
  media assertion.

The v3 quality gates pass for formatting, check, tests, Clippy, documentation,
and boundaries. The focused aligned-ZIP suite reports 12 passed tests. Reader
mutation evidence rejects all 10 diagnostic mutations, all 11 PPTX mutations,
and all 10 OPC mutations. The final helper freeze includes the raw-span fix:
existing `ZipArchive::get_entry` descriptor validation is retained, and
descriptor-bearing member spans are compared through the next local header or
central-directory boundary. The v3 freeze records the repaired
`filesystem.rs`, `aligned_zip.rs`, and harness README descriptors.

No blocker remains for the bounded 0834 repair. The scope stays in the
filesystem harness and its helper; production readers, writers, timers, DOCX
behavior, and iWork remain outside the change. The evidence supports sealing
the repaired packet. Formal capture belongs to the next baseline batch using
the committed harness and its own measurement admission checks.
