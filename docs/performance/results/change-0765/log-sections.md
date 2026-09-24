# Log paragraphs for change 0765

Three blocks, one per shared log. Each is written to be prepended as the newest
section of its file.

## For `HOTSPOTS.md`

## 0765 — every quick-xml position in the OOXML crates now goes through one reader origin

[0765](0765-bom-prefixed-part-editing.md) is a correctness fix, not a performance change. quick-xml 0.41 drops a leading UTF-8 byte-order mark without counting it in `buffer_position`, so every span the OOXML crates took from a reader was three bytes early in a marked part — the open follow-up of 0744 and 0755. A read-only survey classified all 430 position reads; 162 spliced and 55 sliced original bytes at the shifted offsets. The new `litchi_core::xml::ReaderOrigin`, built from each reader's exact input, now converts 364 reads across litchi-opc, -ooxml-common, -drawingml, -xlsx, -pptx, -docx, -xlsb and three neighbours, and replaces the ad hoc `+3` sites. The 0744 lane now admits marked worksheets, and the XLSX and PPTX compactors carry the mark. The control cases are flat: paired changes from −1.7% to +0.6% (confirmation), whole-process instructions within +0.32%, no >5% flag. [Evidence](results/change-0765/README.md).

## For `REPORT.md`

## 0765 — byte-order-marked parts read, edit and save like unmarked ones

[0765](0765-bom-prefixed-part-editing.md) is retained with `performance_claim: none`. Marked-vs-unmarked differential suites for OPC, XLSX, PPTX and DOCX run each scenario on a package and on its twin with marked XML members, compact and indented, and require identical results and outputs that differ only by marks. 16 of their 22 tests fail on the base, and so does a DrawingML theme test; at base a marked, indented theme was published with `</a:clrScheme>me>`, well-formed and saved. Everything passes on the branch. Marks are kept on every route that derives a part from its source bytes and dropped only from parts regenerated from a model. Public spans (`Shape::span`, OMML ranges, `OwnedXmlPart` tag ranges, ink spans) now count the mark, a documented breaking change. On `xlsx_first_cell`, `pptx_semantic_one_edit_save` and `docx_semantic_one_edit_save` the paired changes are −1.69%, +0.62% and +0.55%, within layout noise, with instructions within +0.32%. [Evidence](results/change-0765/README.md).

## For `GOAL_AUDIT.md`

## 0765 — preservation of producer marks and exact spans for marked parts

[0765](0765-bom-prefixed-part-editing.md) closes a defect class that ADR 0006's preservation default and ADR 0003's source-checked edits both touch. On parts that begin with a UTF-8 byte-order mark — Word writes them, and 17 members of three repository fixtures carry one — edits were refused, reads returned shifted bytes, and indented parts could be published silently corrupted. One shared helper, `ReaderOrigin`, is pinned to the linked quick-xml by a contract test and now converts every position read in the OOXML crates. The mark itself is preserved on every source-derived route, compaction included, following 0650's precedent. The 0744 lane accepts marked worksheets; the 0747 window proof and 0754's scanner and publication proof were already mark-aware and are now exercised end to end. [Evidence](results/change-0765/README.md).
