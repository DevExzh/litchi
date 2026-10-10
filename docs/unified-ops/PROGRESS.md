# Unified operations API — progress log

Driven by `docs/UNIFIED_OPS_GOAL.md`; specification `docs/UNIFIED_OPS_API_DESIGN.md`.

## Records

| No. | Title | Status |
|-----|-------|--------|
| 0166 | [XLSB structural diagnostics](0166-xlsb-structural-diagnostics.md) | Canonical changed/no-op introspection and exact signed same-name replay; 1,067 default / 1,078 full-feature native and 489 facade tests/doctests passed; Clippy/rustdoc/glossary/boundaries passed; scratch cleaned; owner migration audited; full goal remains open |
| 0165 | [XLSB structural transaction names](0165-xlsb-structural-transaction-names.md) | Module-local structural transaction family without old aliases; 1,066 default / 1,077 full-feature native and 489 facade tests/doctests passed; Clippy, rustdoc, glossary, boundaries and parity passed; temporary artifacts cleaned; native owner/commit contract, staging costs and unified XLSB writes remain open |
| 0164 | [PPTX slide creation properties](0164-pptx-slide-creation-properties.md) | Slide add accepts name/hidden with strict aggregate admission, one native edit, source/eager replay and batch coverage; 1,125 facade tests/doctests and ten native checks passed; Clippy/docs/minimal/boundaries/parity passed; ten paired cases equivalent (source exact, eager decoded members), ratios 0.995–1.009; artifacts cleaned; full goal remains open |
| 0163 | [Combined PPTX slide properties](0163-pptx-combined-slide-properties.md) | Source primary combined setter and unified eager/source adoption avoid duplicate staging; 23 focused tests/doctests, Clippy and full-feature glossary passed; ten paired combined-native cases exact, one-shape ratios 1.002 eager / 1.029 source; artifacts cleaned; full goal remains open |
| 0162 | [PPTX slide properties](0162-pptx-slide-properties.md) | Eager/source native and unified slide name/hidden writes with strict atomic admission, exact no-ops and durable replay; 1,717 native and 1,122 facade tests/doctests plus final 34 focused checks passed; Clippy/docs/glossary/boundaries/parity passed; ten paired cases exact, one-shape ratios 1.042 eager / 1.014 source; 0159–0162 committed once Git metadata became writable; temporary artifacts cleaned; full goal remains open |
| 0161 | [PPTX shape removal](0161-pptx-shape-removal.md) | Native source/eager and unified top-level removal with reference guards, explicit durable v4 grammar and reversible topology; 1,117 facade and 1,708 native tests/doctests plus final 20 removal checks passed; Clippy/docs/glossary/boundaries/parity passed; ten paired cases exact, one-shape ratios 1.038 eager / 1.022 source; larger eager overhead remains a follow-up; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0160 | [Publication resource errors](0160-publication-resource-errors.md) | XML publication budgets retain `limit_exceeded` and allocation failures retain `internal_error`; native source chain preserved; 1,111 facade and 164 XML tests/doctests, Clippy/rustdoc/minimal build/boundaries/parity passed; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0159 | [PPTX fanout publication](0159-pptx-fanout-publication.md) | Same-slide selectors share native chunk publication with byte/target admission; 1,110 facade tests/doctests, Clippy/rustdoc/boundaries/parity checks passed; all 20 paired cases equivalent, 10k ratios 1.005–1.238 and unified 10k/1k growth 9.749–9.973x; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0158 | [PPTX chunk publication](0158-pptx-chunk-publication.md) | Native multi-chunk publication and facade adoption; 1,688 native and 1,108 facade tests/doctests passed; 1k/10k ratios below 2x and unified 10k/1k growth 9.66–10.05x; Clippy/glossary/boundaries passed; implementation/evidence committed concurrently; artifacts cleaned; own staging blocked by read-only Git metadata; full goal remains open |
| 0157 | [PPTX slide-root admission](0157-pptx-slide-root-admission.md) | Distinct slide/notes node limits and preserved typed resource errors; 1,665 native and 1,107 facade tests/doctests passed; eager 10k workloads pass with 1.016–1.022 ratios; Clippy/glossary/boundaries passed; artifacts cleaned; commit blocked by read-only Git metadata; linear scaling and full goal remain open |
| 0156 | [PPTX command coalescing](0156-pptx-command-coalescing.md) | 1,000 direct shape commands share one native commit; 1,107 combined tests/doctests plus Clippy passed; 1k paired ratios 1.005–1.224; linear scaling fails and eager 10k exposes masked native node limit; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0155 | [One eager edit per PPTX setter](0155-pptx-single-edit.md) | Removed discarded native edit; paired eager one-shape overhead is 4–5% in the measured generated corpus; 1,097 facade tests/doctests and Clippy passed; artifacts cleaned; commit blocked by read-only Git metadata; command coalescing, wider performance acceptance and full goal remain open |
| 0154 | [PPTX selector fanout](0154-pptx-selector-fanout.md) | Atomic text/geometry selection beyond 256 targets; 1,096 facade tests/doctests passed; paired 1–1,000-shape evidence finds eager small-write overhead above budget; Clippy passed; artifacts cleaned; commit blocked by read-only Git metadata; command batching, performance acceptance and full goal remain open |
| 0153 | [Atomic PPTX text and geometry writes](0153-shape-geometry-writes.md) | Native eager/source and unified combined text/bounds edits with durable replay; 1,663 native and 1,092 facade tests/doctests, Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; broader CRUD/performance/full goal remain open |
| 0152 | [Text-box output admission](0152-text-box-output-admission.md) | Caller output limits checked before managed memory admission; 11 native and six facade tests plus Clippy passed; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0151 | [Source-backed text-box addition](0151-source-text-box-addition.md) | Native and unified source text-box creation, canonical durable suffix proof and reference-aware IDs; 1,642 native and 1,090 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; broader shape/format CRUD and full goal remain open |
| 0150 | [Source-backed slide insertion](0150-source-slide-insertion.md) | Native and unified source slide insertion, empty/disjoint durable topology and mixed lexical relationship edits; 1,144 OPC, 1,630 PPTX and 1,089 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; source shape creation and full goal remain open |
| 0149 | [Unified source slide removal](0149-unified-source-slide-removal.md) | Source File/query/transaction/batch removal and effective catalog/related reads; 1,614 native and 1,081 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; source insertion, broader CRUD and full goal remain open |
| 0148 | [Primary source slide removal](0148-primary-source-slide-removal.md) | Readable structural revisions and durable v3 replay/inverse after reopening; 1,614 native and 1,078 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; facade source removal and full goal remain open |
| 0147 | [Exact prepared metadata images](0147-prepared-metadata-images.md) | Bounded source/changed/appended/omitted metadata handles retain managed admissions; 1,141 OPC, 1,589 native PPTX and 1,078 combined facade tests/doctests passed; OPC Clippy/docs and crate boundaries passed; commit blocked by read-only Git metadata; primary topology/durable integration and full goal remain open |
| 0146 | [Source-native slide removal plan](0146-source-slide-removal-plan.md) | Lazy OPC removal with strict ownership/metadata guards and exact root splicing; 1,589 native and 1,078 combined facade tests/doctests, Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; primary durable source removal and full goal remain open |
| 0145 | [Guarded source topology metadata](0145-source-topology-metadata.md) | Source-native exact relationship/content-type XML restoration behind physical and typed guards; 1,135 OPC, 1,583 native PPTX and 1,078 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; primary source removal and full goal remain open |
| 0144 | [Source PPTX removal prerequisites](0144-source-removal-prerequisites.md) | Shared lexical binding removal and source layout validation; physical XML limits enforced on master/layout/theme; 1,583 native and 1,078 combined tests/doctests plus Clippy/docs/glossary passed; boundaries passed and temporary artifacts cleaned, commit blocked by read-only Git metadata; primary source removal and full goal remain open |
| 0143 | [PPTX modeled order metadata](0143-pptx-modeled-order-metadata.md) | Stock source Move/Swap and typed presentation extensions preserve exact metadata; 1,578 native, 562 PPTX facade and 1,078 combined tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; source topology and full goal open |
| 0142 | [Primary source PPTX slide ordering](0142-pptx-primary-source-order.md) | Source Move/Swap composed with text, logical read mapping, stable bindings and durable v2 proofs; 1,572 native, 561 PPTX facade and 1,077 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; source topology, positional metadata and full goal open |
| 0141 | [Composable source PPTX slide ordering](0141-pptx-source-order-composition.md) | Native staged moves/swaps, checked positions, exact cancellation and managed root reservations; 1,557 native, 559 PPTX facade and 1,075 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; primary catalog/durable integration and full goal open |
| 0140 | [Unified PPTX slide removal](0140-pptx-unified-slide-removal.md) | Eager path/scoped-query removal, native dependency/final-slide guards and typed refusal mapping; 558 PPTX and 1,074 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; richer dependencies/source topology and full goal open |
| 0139 | [Unified PPTX slide ordering](0139-pptx-unified-slide-order.md) | Eager move/swap through File, transactions and batches; exact placement/no-op/replay/undo and pre-promotion refusal; 548 PPTX and 1,064 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; removal/source topology and full goal open |
| 0138 | [Unified PPTX text-box addition](0138-pptx-unified-text-box-addition.md) | Eager shape/textbox creation, strict geometry and native CR preservation; 987 native, 543 PPTX facade and 1,059 combined tests/doctests passed; Clippy/docs/glossary passed; commit blocked by read-only Git metadata; source structural creation and full goal open |
| 0137 | [Unified PPTX slide addition](0137-pptx-unified-slide-addition.md) | File/transaction/batch add for created and eager owners, native preservation/proofs and explicit source-owner capability gap; 538 PPTX and 1,054 combined facade tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; source topology, broader CRUD and full goal open |
| 0136 | [Native PPTX slide insertion](0136-pptx-native-slide-insertion.md) | Checked position/exact layout insertion, durable replay/inverse and topology dependency guards; 1,001 native tests, 20 native doctests and 533 facade tests/doctests passed; Clippy/docs/glossary passed; commit blocked by read-only Git metadata; facade CRUD and full goal open |
| 0135 | [PPTX part relationship metadata resolution](0135-pptx-part-metadata-three-way.md) | Physical part choices retain matching graph/XML; unrelated content-type and owner edits remain automatic; 986 native and 533 facade tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; type/topology projection and full goal open |
| 0134 | [PPTX package-root three-way merge](0134-pptx-root-three-way.md) | Compatible root IDs merge automatically; divergent graph/XML candidates resolve with Base/Left/Right and durable inverse preservation; 983 native and 533 facade tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; part/content-type conflict resolution and full goal open |
| 0133 | [PPTX durable metadata replay](0133-pptx-durable-metadata-replay.md) | LPTX0003 per-key metadata replay, exact forward/inverse publication and deterministic disjoint joins; 978 native, 1,126 OPC and 533 facade tests/doctests passed; Clippy/docs/glossary passed; commit blocked by read-only Git metadata; three-way metadata conflict resolution and full goal open |
| 0132 | [Guarded OPC content-type fragments](0132-opc-content-type-fragments.md) | Native-key Default/Override transport, shared metadata planner and stricter no-op revalidation; 1,126 OPC tests/doctests, 967 native PPTX tests, Clippy/docs passed; commit blocked by read-only Git metadata; durable metadata wire integration and full goal open |
| 0131 | [Guarded OPC relationship fragments](0131-opc-relationship-fragments.md) | Bounded lexical entry transport with source/context guards and exact inverse restoration; full OPC suite, 13 final focused tests, Clippy/docs and native PPTX check passed; commit blocked by read-only Git metadata; content-type transport, native wire integration and full goal open |
| 0130 | [Semantic PPTX slide-order selectors](0130-pptx-slide-order-selectors.md) | Canonical slide selector, semantic move/swap/reorder and private raw-ID lowering; 982 native and 533 facade tests/doctests, harness, Clippy and glossary passed; commit blocked by read-only Git metadata; metadata transport and full goal open |
| 0129 | [Guarded OPC metadata restoration](0129-opc-guarded-metadata-restoration.md) | Exact-source raw metadata installation with bounded parsing/audit and physical member presence; 1,103 OPC tests/doctests, 967 native PPTX tests and Clippy passed; commit blocked by read-only Git metadata; native durable wire integration and full goal open |
| 0128 | [PPTX empty slide-list prerequisite](0128-pptx-empty-slide-list.md) | Empty-list lexical insertion and bounded IDs; 967 native tests plus metadata replay diagnostic and 99 integration tests and Clippy passed; durable forward metadata preservation remains prerequisite to public insertion; commit blocked by read-only Git metadata; full goal open |
| 0127 | PPTX canonical constructor | Public package writer/font setters removed; native 1,515 and combined facade 1,081 tests/doctests passed; direct new-package edits/replay/inverse verified; remaining mutators/full goal open; commit blocked |
| 0126 | PPTX builder consumer migration | Public authoring consumers migrated; 12 unified write + 10 harness tests, 21/24 root and both native smoke examples passed; 3 existing codec/adapter failures recorded; native constructor/full goal open; commit blocked |
| 0125 | PPTX native builder and unified creation | Direct canonical native owner, typed option gaps and creation capabilities; 1,081 combined facade tests/doctests and native builder regression passed; Clippy/both glossary profiles passed; commit blocked by read-only Git metadata; full goal remains open |
| 0124 | PPTX immutable encrypted publication | Retained-profile re-encryption integrated with common atomic save; 531 facade tests/doctests, 13 native encryption tests and final public regressions passed; Clippy/glossary passed; commit blocked by read-only Git metadata; full goal remains open |
| 0123 | PPTX unified shape-text writes and publication | Native immutable setters, transactions/batches and atomic save integrated; 962 native + 1,076 combined facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; broader writes and full goal remain open |
| 0122 | PPTX committed source revision readback | 26 native primary tests, 17 native doctests and 519 facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; facade writes and full goal remain open |
| 0121 | Shared immutable PPTX native promotion | 518 facade tests/doctests and final signature regression passed; Clippy/glossary passed; commit blocked by read-only Git metadata; facade writes and full goal remain open |
| 0120 | PPTX durable source shape-text patches | 960 encrypted native tests, 17 native doctests and 516 facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; full goal open |
| 0119 | PPTX immutable source package ownership | 92 source, 8 owned-scene, 17 native doctests and 516 facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; full goal open |
| 0118 | PPTX committed package and immutable patch replay | 210 encrypted native opened-editor, 22 policy/marked-XML, 16 native doctests and 516 facade tests/doctests passed; consumers/features/glossary and Clippy passed; commit blocked by read-only Git metadata; full goal open |
| 0117 | PPTX patch admission and visible payload ownership | 203 default + 203 encrypted native opened-editor, 22 policy/marked-XML and 516 facade tests/doctests passed; Clippy clean; commit blocked; package-carrying commits and full goal remain open |
| 0116 | PPTX primary Edit entry and drop semantics | 193 native opened-editor, 22 policy/marked-XML, one native example and 516 facade tests/doctests passed; glossary, Clippy, harness and public example checks passed; commit blocked; full goal open |
| 0115 | XLSX border writes and typed local formatting | 1,012 combined tests/doctests, 90 stylesheet codec, 64 native cell-edit and 3 state tests passed; native examples, capabilities, Clippy and XLSX glossary passed; commit blocked; full goal open |
| 0114 | XLSX gradient-fill writes | 1,001 combined tests/doctests, 81 stylesheet codec, 57 native cell-edit and 78 style tests passed; canonical native replacement, bounded multi-stop input and typed join-budget recovery; examples, capabilities, Clippy and XLSX glossary passed; commit blocked; full goal open |
| 0113 | XLSX solid and pattern-fill writes | 995 combined tests/doctests, 75 stylesheet codec and 50 native cell-edit tests passed; typed native fill changes, facade foreground/pattern/background policy, Clippy/native example, both capability profiles and XLSX glossary check passed; commit blocked; full goal open |
| 0112 | XLSX number-format writes | 988 combined tests/doctests, 66 stylesheet codec, 44 native cell-edit and 75 style tests passed; native resource reuse/allocation and unified aliases; Clippy/native example clean; commit blocked; full goal open |
| 0111 | XLSX protection and remaining alignment attributes | 982 combined tests/doctests, 56 stylesheet, 39 native cell-edit and 73 native style checks passed; five setters and optional justification readback; 50 final facade checks and Clippy clean; commit blocked; full goal open |
| 0110 | XLSX alignment writes | 975 combined + 45 final facade cell tests/doctests, 48 stylesheet and 37 native cell-edit checks passed; Clippy and native example clean; shared font/alignment lowering and seven setters; commit blocked; full goal open |
| 0109 | XLSX remaining typed font properties | 967 final DOCX/XLSX/XLSB/encryption tests/doctests, 38 stylesheet, 38 font-focused and 39 final cell-write checks passed; seven font read/write properties, checked integer bounds and explicit XLSB gaps; Clippy clean; commit blocked; full goal open |
| 0108 | XLSX font color writes | 929 combined + 34 final cell tests/doctests, 32 stylesheet and 32 font-focused checks passed; typed RGB/theme/tint, lexical preservation and ordered insertion; Clippy clean; full goal open; commit blocked |
| 0107 | XLSX font name, size and script writes | 924 combined + 29 final cell tests/doctests, 26 stylesheet and 28 font-focused checks passed; Clippy clean; owned names, checked size/script, bounded quote-safe writes; full goal open; commit blocked |
| 0106 | XLSX strike-through and underline writes | 919 combined + 49 final write tests/doctests, 20 stylesheet and 22 font-focused checks passed; Clippy clean; native strike/underline, aliases and strict XML envelope; full goal open; commit blocked |
| 0105 | Reuse explicit XLSX merge-list planning | 914 combined tests/doctests and 18 focused merge tests passed; Clippy clean; per-worksheet admission reuse retains selected-node limits and atomic companions; full performance/coverage goal open; commit blocked |
| 0104 | XLSX merge operations and range readback | 911 combined + 419 minimal tests/doctests, 37 native + 21 source merge tests and final 15 unified regressions passed; Clippy clean; bounded native merges, atomic companions and table guards; broader coverage/full goal open; commit blocked |
| 0103 | XLSX row and column layout | 895 combined + 419 minimal tests/doctests and 14 delta tests passed; Clippy clean; native height/width/visibility/outline/collapse assignments with bounded selectors; broader writes/full goal open; commit blocked |
| 0102 | XLSX worksheet lifecycle | 883 combined + 419 minimal tests/doctests and final 34 targeted regressions passed; Clippy clean; worksheet add/remove, tab rename/visibility/move/swap and native patch integration; broader coverage/full goal open; commit blocked |
| 0101 | XLSX distinct-cell batch coalescing | 870 combined + 419 minimal tests/doctests, 3 native state tests and 9 native doctests passed; Clippy clean; source/eager 1k scalar batches within 2× native in scoped ABBA runs; broader coalescing/full goal open; commit blocked |
| 0100 | Preserve ignorable XLSX stylesheet extensions | 16 codec + 11 native font + 859 combined tests/doctests and 14 delta tests passed; Clippy clean; bounded MCE admission, opaque preservation and font insertion order verified; broader formatting/full goal open; commit blocked |
| 0099 | XLSX native and unified font assignments | 2,254 native + 858 combined + 419 minimal tests/doctests, final 11 codec and 14 delta tests passed; Clippy clean; lossless bold/italic resources, range writes, joins, replay and inverse verified; stylesheet profiles/compaction and full goal open; commit blocked |
| 0098 | XLSX range and scoped-selector writes | 855 combined + 419 minimal tests/doctests, 14 target-write regressions and 14 delta tests passed; Clippy clean; one native commit, distinct-cell/staging budgets and early rectangle admission verified; structural/coalesced writes and full goal open; commit blocked |
| 0097 | XLSX cell writes and native transaction history | 2,233 native + 20 final native + 841 combined + 419 minimal tests/doctests and 14 delta tests passed; Clippy clean; direct cell writes, rollback, replay, opaque preservation and protection reports verified; range/structural/coalesced writes and full goal open; commit blocked |
| 0096 | Immutable XLSX source promotion | 561 XLSX encryption + 817 combined tests/doctests and final concurrent regression passed; Clippy clean; exact source, successful promotion reuse, retry and signature isolation verified; cell writes/full goal open; commit blocked |
| 0095 | XLSX native creation and publication | 814 combined DOCX/XLSX + 558 XLSX encryption + 418 minimal tests/doctests and 14 delta tests passed; strict output/origin/signature policies verified; XLSX mutations and managed publication/full goal open; commit blocked |
| 0094 | DOCX cached-result field invalidation | 2,282 native + 703 DOCX + 418 minimal tests/doctests and 14 delta tests passed; exact dirty-marker replay/inverse and atomic rollback verified; additional namespace replay regression and Clippy passed; full goal open; commit blocked |
| 0093 | DOCX command paragraph-add splitting | 700 DOCX + 418 minimal tests/doctests, 39 final batch and 14 delta tests passed; Clippy clean; atomic LF splitting/VT inline controls verified; full goal open; commit blocked |
| 0092 | Literal first-node DOCX command run text | 2,277 native + 697 DOCX + 418 minimal tests/doctests; 14 delta tests, Clippy and replay review passed; paragraph splitting/field cache/managed edits/full goal open; commit blocked |
| 0091 | Reuse DOCX body during write admission | 692 DOCX tests/doctests and Clippy passed; controlled 1k fanout at 1.13× eager / 1.10–1.11× source native; small-write/full goal open; commit blocked |
| 0090 | Borrow DOCX immediate write assignments | 691 DOCX tests/doctests and Clippy passed; deferred plans retain ownership, paired timing shows no clear improvement; full goal open; commit blocked |
| 0089 | Paired DOCX paragraph text fanout baseline | Two 9-round ABBA invocations, byte/ledger equivalence and Clippy passed; small-write overhead exceeds review trigger, 10k typed refusal; full goal open; commit blocked |
| 0088 | DOCX paragraph selector command text controls | 690 DOCX + 418 minimal tests/doctests, 31 batch and 14 delta tests passed; single-pass target admission, Clippy clean; broader text/full goal open; commit blocked |
| 0087 | Native bulk dispatch for DOCX paragraph text fanout | 688 unified DOCX tests/doctests and 3 final fanout regressions passed; Clippy clean; eligible fanout uses native bulk, performance/full goal open; commit blocked |
| 0086 | Native DOCX bulk paragraph inline controls | 1,397 native + 12 bulk integration + 685 unified tests/doctests passed; managed publication/inverse and resource release verified, Clippy clean; full goal open; commit blocked |
| 0085 | Direct DOCX paragraph inline text controls | 1,397 native + 9 managed + 685 DOCX + 418 minimal tests/doctests; Clippy/delta checks passed; direct paragraph command controls integrated, broader text/full goal open; commit blocked |
| 0084 | DOCX command run-add text controls | 683 DOCX tests/doctests, 32 focused batch/capability and 14 delta tests passed; Clippy clean; input run controls integrated, other text contexts/full goal open; commit blocked |
| 0083 | Decode typed DOCX scalar controls in scope | 1,397 native + 19 source/content + 680 unified tests/doctests passed; 3k typed controls need zero owner scans; malformed scalar payloads refused; Clippy clean; full goal open, commit blocked |
| 0082 | Reuse DOCX run namespace context | 1,395 native + 19 source/content + 680 unified tests/doctests passed; zero owner scans for 1k captured run text reads, Clippy clean; broader context/performance/full goal open; commit blocked by read-only Git metadata |
| 0081 | Contextual DOCX paragraph/run text reads | 1,393 native + 5 content + 10 tracker + 39 source/streaming + 680 unified tests/doctests passed; Clippy clean; context reuse, text policy and full goal open; commit blocked by read-only Git metadata |
| 0080 | Namespace-correct DOCX typed run content | 1,397 final native unit/integration + 679 unified tests/doctests; fixture targets recovered (36 tests), Clippy clean; text-policy completion remains open; commit blocked by read-only Git metadata |
| 0079 | Shared formatting runs for DOCX closure edits | 678 DOCX +418 minimal tests/doctests and final19 transaction tests passed; typed targets, staged reads, poison and best-effort isolation verified; wider coalescing/full goal open; commit blocked by read-only Git metadata |
| 0078 | Paired DOCX batch scaling and target reuse | 452 focused tests + final 25 batch regressions passed; repeated-target 1k batch measures 1.08–1.10× leaf across eager/source; broader performance acceptance open; commit blocked by read-only Git metadata |
| 0077 | Atomic DOCX direct-path formatting runs | 669 DOCX + 418 minimal tests/doctests, final 24 batch regressions and 2 bisection probes passed; native-capacity splitting and retained XML reduction verified; performance/wider coalescing open; commit blocked by read-only Git metadata |
| 0076 | DOCX atomic and best-effort command batches | 1,039 all-feature + 665 final DOCX tests/doctests and 26 final batch/capability regressions passed; atomic/best-effort core integrated; coalescing/full command coverage open; commit blocked by read-only Git metadata |
| 0075 | Bounded DOCX explicit transactions | 1,021 all-feature + 417 minimal tests/doctests and final 24 transaction/capability regressions passed; owner-bound replay/limits verified; coalescing/batch remain open; commit blocked by read-only Git metadata |
| 0074 | Complete internal DOCX operation commits | 1,002 all-feature + 413 minimal tests/doctests; patch/diagnostics retained across all five verbs, replay/inverse/no-ops verified and empty-target admission fixed; public transactions/coalescing remain open; commit blocked by read-only Git metadata |
| 0073 | Bounded OfficeCLI batch payload ingress | 994 all-feature + 413 minimal tests/doctests; complete transport vocabulary, caller-selected byte/item limits and typed redacted failures verified; executor/coalescing remains open; commit blocked by read-only Git metadata |
| 0072 | DOCX creation and atomic path saving | 970 all-feature ops + 596 DOCX-only + 389 minimal + 1,097 OPC tests/doctests; typed creation, origin retention, guarded overwrite policy and encrypted save round trips verified; broader writes/batches remain open; commit blocked by read-only Git metadata |
| 0071 | Immutable DOCX structural operations | 946 all-feature ops + 2,257 DOCX leaf + 384 minimal tests/doctests; paragraph/run add/remove/move/swap, exact slot preservation and add registry verified; batches/path saving remain open; commit blocked by read-only Git metadata |
| 0070 | Immutable DOCX set and stream publication | 917 all-feature ops + 2,252 DOCX leaf + 1,093 OPC tests/doctests; 68 setters and exact eager/source no-ops verified; structural writes/batches/path saving open; commit blocked by read-only Git metadata |
| 0069 | Property inputs, strict registry admission and error attribution | M2 W1/W2 foundation: 529 DOCX + 359 minimal tests/doctests; File setters/batches pending; commit blocked by read-only Git metadata |
| 0068 | XLSB unified read adapter and indexed scalar fallback | R4d integrated: 1,076 leaf + 418 ops + 468 XLSX regression tests/doctests; extended gaps/write/performance open; commit blocked by read-only Git metadata |
| 0067 | XLSB retained grids and checked sparse reads | L-XLSB-R2 verified: 1,072 leaf + 401 ops tests/doctests; R4d pending; commit blocked by read-only Git metadata |
| 0066 | XLSB semantic tab catalog | L-XLSB-R1 verified: 1,055 leaf + 400 ops tests/doctests; R2/R4d pending; commit blocked by read-only Git metadata |
| 0065 | XLSX trait read boundaries and capability publication | 468 tests/doctests and Clippy passed; infallible metadata migration open; commit blocked by read-only Git metadata |
| 0064 | PPTX unified read adapter | R4c integrated; 2,310 tests/doctests passed, 3 ignored; extended gaps/performance open; commit blocked by read-only Git metadata |
| 0063 | Source-backed PPTX modern comments | R9 read foundation verified; unified PPTX dispatch pending; commit blocked |
| 0062 | Source-backed PPTX presentation catalog | R4 read foundation verified; full goal active; commit blocked |
| 0061 | Source-backed PPTX related parts | R3 read foundation verified; modern comments and full charts pending; commit blocked |
| 0060 | Shared DrawingML text bodies | R7 codec/property projection verified; full policy migration pending; commit blocked |
| 0059 | PPTX lazy shape properties | R6 properties/raw identity verified; commit blocked |
| 0058 | Source-backed PPTX scenes and bounded retention | R1/R2 source read/cache verified; R3/R6/R7 and R4c pending; commit blocked by read-only Git metadata |
| 0057 | Independently retained PPTX scenes and token classification | R2 eager foundation and R10 verified; source memo and R4c pending; commit blocked by read-only Git metadata |
| 0056 | XLSX unified read adapter and File dispatch | R4b integrated and verified; broader M1 and performance acceptance pending; commit blocked by read-only Git metadata |
| 0055 | XLSX readable formatting and source table access | L-XLSX-R3/R4 integrated and verified; adapter and performance acceptance pending; commit blocked |
| 0054 | XLSX retained grid, metadata projection and early-stop leaf reads | L-XLSX-R1/R2/R6 integrated; formatting/tables, adapter and performance acceptance pending; commit blocked |
| 0053 | Capability reports and generated OfficeCLI comparison | DOCX discovery integrated; parity/registry incomplete; commit blocked by read-only Git metadata |
| 0036 | OPC source-limit kinds and canonical errors | physical kind/error fields migrated; broader limits policy pending |
| 0035 | OPC/OOXML source-limit vocabulary | shared physical profile and builder renamed; broader limits policy pending |
| 0034 | Effective source DOCX text streaming | retained main/donor output integrated; specialized routes pending |
| 0033 | Effective source DOCX metadata reads | shared property graph/decoder integrated; remaining specialized routes pending |
| 0032 | Typed primary DOCX transaction limits | typed kind/count fields migrated; configurable limits and broader policy pending |
| 0031 | Canonical source DOCX package-owning Edit/Commit | primary owner migrated; specialized families, limits and performance pending |
| 0030 | Immutable retained DOCX semantic apply | main-document apply integrated; independent managed donor snapshots refused |
| 0029 | Source DOCX effective owner and physical-operation boundary | effective reads/publication integrated; specialized routes pending |
| 0028 | Effective OPC owning conversion | explicit materialization integrated |
| 0027 | DOCX patch predicate and content-free Debug | primary patch migrated; broader diagnostics pending |
| 0026 | Owned reads and physical identity access for retained OPC candidates | effective read adapter foundation; integrated in 0029–0031 |
| 0025 | Immutable XML edits of retained effective OPC candidates | flat chained source-Part candidates; primary owner integrated in 0031 |
| 0024 | Revision-bound source XML edits over retained payloads | per-part chaining foundation; package overlay revisions pending |
| 0023 | Retained OPC prepared topology and shared ZIP appended members | source-overlay ownership foundation; chained commits integrated in 0031 |
| 0022 | Eager DOCX package-bound Edit/Commit and removal of mutating publication doors | eager primary commit migrated; source joined in 0031; history and full W2 pending |
| 0021 | Shared source-backed DOCX owners and borrowing stream publication | sharing prerequisite complete; primary package-bound triad integrated in 0031 |
| 0020 | Cheap eager DOCX package shares and immutable primary patch application | eager apply migrated; package-owning commit W2 work pending |
| 0019 | Owned eager DOCX read projection and shared physical graph | projection migrated; immutable package and commit W2 work pending |
| 0018 | DOCX primary edit entry and focused error/result names; minimal/full glossary bans | names migrated; package-owned commit/projection W2 work pending |
| 0017 | W0 rustdoc API glossary gate, scoped rules, six-type family checks and CI | complete; profiles for leaf migration waves pending |
| 0016 | Rebase onto current base; adopt API glossary as ADR 0034; preserve shared fingerprint memo | kickoff; W0 checker and leaf migrations pending |
| 0001 | Owner override: read + write for every claimed format; tiers OOXML → OLE → (RTF, ODF, iWork) | recorded in the goal (2026-10-04) |
| 0002 | Leaf write-capability inventory (deep: OOXML + OLE; brief: tier 3) | merged (`cfe7681627`) |
| 0003 | M0a facade migration plan (flat `File`, `litchi::open`) | merged (`45a932ea2e`) |
| 0004 | ADR 0033 draft + `litchi-ops` scaffold + gate registration | merged (`af70966a6d`) |
| 0005 | M0b path/selector grammar in `litchi-core::selector` | merged (`9f8e49b181`) |
| 0006 | Owner accepts ADR 0033 | done |
| 0007 | Coordinator defaults for 0003's open questions | adopted, awaiting owner |
| M0a | Steps S1–S10 on `uo/m0a` (plan 0003 §4) | merged (`1927763795`); S10 gate passes with E1/E2 awaiting owner (F1, F2) |
| 0008 | Diagnosis: source-backed DOCX edits refuse every Word-authored fixture (MCE) — M2 blocker | merged (`e9c3868109`); fix staged as 0011 |
| 0009 | Design: editing a shared snapshot (R4) — M2 prerequisite | done (`9b3cab60f6`); staged leaf plan S0–S5 |
| 0010 | M1 unified-read plan (node view, path resolution, query, view modes, capabilities) | done (`31bd660497`); steps R1–R11 with leaf waves L1/L2 |
| M1 | Steps on `uo/m1` (plan 0010 §10); merges after M0a | in progress |
| 0011 | DOCX MCE admission, stages 1–3 of 0008 | merged into `uo/m0a` (`53f938e83c`); 19/19 Word fixtures editable source-backed; eager owner gains the same refusals (breaking) |
| 0013 | Signature policy: `Signed::unsign`, builder refusals, pptx/xlsb silent-strip fixes (M0a S6b) | merged into `uo/m0a` (`f30c521c90`) |
| 0015 | M2 write-kernel plan (verbs, registry, per-format handlers, batch executor, leaf prerequisites, save/create) | done (`de8d557c56`); steps W1–W10 and leaf prerequisites L-*-W* |
| M2 | Steps on `uo/m2` (plan 0015 §11); merges after M0a and M1 | in progress |
| 0014 | Implicit `unsign()` sites: pptx and shared ooxml-common (0014a), xlsx/docx and XLSB task panes/Ribbon (0014b); 0013 F1, F1b, F2 | merged into `uo/m0a`; silent strips now refuse until explicit `unsign` (breaking) |
| 0012 | litchi-opc: cheap `Clone` for `SourceBackedPackage` (0009 stage S0) | merged into `uo/m0a` (`0cf1f43403`); clone 6 reference counts, no allocation; open unchanged |

## Awaiting owner confirmation

- M0a S10 (`docs/unified-ops/s10-measurement.md` §9). M0a landed on the strength of two explanations. If either is rejected, the chosen alternative becomes a harness follow-up; M0a is not reverted.
  - F1: accept E1. Seven rows over their trigger time one call in a fresh child process, so they measure the harness executable. Standalone probes give the same first-call time at both revisions (XLSX open 187.6 vs 187.9 µs); Callgrind shows +0.75% instructions. Alternatives: an untimed warm-up call in the child (new route names), or standalone binaries for first calls.
  - F2: accept E2. The RTF open row depends on the allocator state the untimed projection leaves; with malloc trimming off, the unified open is 0.98× the pre-M0a one. Alternatives: fixed malloc tunables for RTF children, or bare-loop variants.

- Record 0015 decisions D1–D19 and open questions Q1–Q14 (provisional defaults). Highlights:
  - Write verbs take `&self` and return a new `File`; `edit` returns a `Commit`.
  - Batches run sequentially, grouped into same-family runs, and the batch result must equal its items applied one at a time.
  - Writes refuse invalid values and unknown keys, with suggestions. OfficeCLI's apply-the-rest-and-warn behaviour is only a command-channel option.
  - After `Signed::unsign`, a save writes the unsigned copy. Decrypted files save re-encrypted when given the password, or refuse; plaintext needs `into_plaintext()`.
  - XLSX moves to the eager engine on its first write; PPTX text-only runs stay on the source-backed editor.

- Record 0010 open questions (§11.3), the two that most need the owner:
  - Q3: regex values (`~= r"…"`) are refused with `unsupported_capability` until the owner approves adding the `regex` crate.
  - Q6: sheet names match case-insensitively in every workbook format. This changes XLSB and ODS.
  - The other eleven defaults are listed in §11.3.

- Record 0009 (R4): a shared handle shares the leaf package instead of moving it. Write verbs take `&self` and return a new handle; there are no consuming verbs on `File`. Option B (keep the reader and re-open) and option C (consuming verbs with a sole-owner fast path) are rejected.

- Record 0007 defaults: flat `File` (F2), `limit_exceeded`/`cancelled` codes (Q1), iWork → `unsupported_format` until tier 3 (Q2), `retain_source = true` (Q3), `Signed` wraps every signed input (Q4).

- `OpenOptions` shape: one struct with common knobs plus optional per-format sub-options.
- Verb names: `apply`; find/replace as `find=`/`replace=` props on `set`; `merge` = template merge.

## 0004 — log

- Drafted ADR 0033 (Proposed) and listed it in the ADR index; ADR 0024 inventory amended (66 packages / 245 declarations).
- Scaffolded inert `litchi-ops` (edge to `litchi-core` only); registered in `crate_boundaries.json` and `non_iwork_gate.py` (bulk 46 → 47); gate tests updated and green.
- Known: workspace `cargo fmt --all --check` fails on base-branch files (soapberry-zip, iwa-archive) unrelated to this program.
- Follow-up: the first iWork leaf edge from `litchi-ops` (tier 3) requires reclassifying it out of the bulk gate, like the facade.

## 0002 — log

- Code-verified inventory; facade is read-only for every format and drops/never holds the editable leaf owner for doc, ppt, rtf; xlsx/xlsb are held only in the non-editable lazy form.
- No `Locked`/`Signed` types exist. Changed commits silently strip signatures in pptx `opened::Transaction`, several xlsb families, odf-formula, ODS legacy mutators, ODP in-place mutators (ADR 0006 concern; fix in the signature-policy prerequisite).
- doc/ppt `Writer::set_property` never emitted; xls `Writer` writes no property streams. Feature-matrix overclaims: pptx hyperlink write, docx app.xml write.
- Stream order: T1 xlsx > docx > pptx > xlsb; T2 xls > ppt > doc; T3 per record. Cross-cutting prerequisites: facade atomic writer, signature/encryption policy, commit-shape adapter, editable owner retention in M0a.

## 0003 — log

- Plan for M0a: a flat `File` enum with one variant per format, plus `Locked`/`Signed`, each variant one `Arc` handle (16 bytes, Send + Sync). The editable owner is built at the first edit, not at open. Leaf constructors L-PPTX-1, L-XLSX-1, L-XLSB-1, L-DOC-1, L-XLS-1 and L-PPT-1 avoid indexing the container twice.
- Gate: `non_iwork_gate.py` forbids any edge from a bulk package to an iWork crate. So `litchi-ops` stays free of iWork until M6, and the gate must be redesigned first (R1).
- Findings against the design: F1 (§16.2's "mutable owners" for pptx/xlsx are read-only), F2 (flat enum), F14 (Python `FileFormat` drift, including Odb → Doc).

## 0005 — log

- `litchi-core::selector` grammar: paths (1-based → `Position`, `last()`, quoted segments), predicates with `and`/`or`/`not`, header-row predicates, CSS-like queries (`:contains`, `:has`, `:not`, combinators, unions, sheet scopes, ranges), story axes, typed `ParseError` with offset and code, explicit limits.
- 950 OfficeCLI paths and 132 selectors parse and round-trip. Typical paths and simple queries parse with zero allocations (~57 ns per path, ~120 ns per selector, release build, single machine). There are 30 recorded OfficeCLI deltas: element aliases and bare-key meaning move to M1 resolution, and the unscoped-mutation guard moves to the verb layer.
- Open: the `selector_grammar` fuzz target compiles but has not been run (no nightly toolchain or cargo-fuzz on the host).

## 0008 — log

- None of the 19 DOCX fixtures uses a content-changing MCE construct. 13 only declare the namespace (Word 2007 `xmlns:ve`) and 6 carry one root `mc:Ignorable`.
- The source-backed guard refuses them because it compares buffers, and the MCE codec rewrites any part that mentions the namespace URI. The guard also has a hole: a namespace URI written with a character escape slips past both the codec and the guard.
- The eager owner keeps every `mc` construct byte for byte, but it has no MCE or protection check. It accepts a wrong-target edit under body-level `AlternateContent`, accepts `MustUnderstand` documents, and publishes edits to 3 protected fixtures.
- `to_owned_package()` rebuilds the ZIP layout (M0a `SourceMode::Eager` uses it). The execution-context route refuses even reads of all 19 files.
- Reads and edits count paragraphs differently: reads include table cells, edits count only top-level body paragraphs. M1 has to choose for `/body/p[N]`.
- Recommended: a staged admission rule (declared-only, 13/19; then Ignorable-only, 19/19) applied to both owners and the execution-context route, with one protection policy. Run-level `AlternateContent` comes later with a post-commit check. No silent eager fallback.

## 0009 — log

- No leaf guard needs a unique owner. The OOXML publishers take `self` by API choice but only borrow inside, and every share of one opened package carries the same lineage token. The packages cannot be cloned only because their index, catalog and cache are owned inline.
- Probe: the xlsx shared-handle fallback in `Xlsx::into_editor` breaks exact no-ops on all 4 real-producer fixtures (0 of 10–13 ZIP members kept byte-identical) and makes output depend on whether a second handle exists. It must be retired before M2.
- No OOXML source-backed commit can be chained before publishing, so M2 publishes into memory and re-opens at most once per family boundary until overlay chaining (option D) lands.
- Staged leaf plan: S0 litchi-opc cheap `Clone`; S1 xlsx (ops only); S2 docx `&self` publishers; S3 pptx; S4 xlsb non-mutating apply; S5 OLE needs nothing for sharing.

## 0010 — log

- M1 plan:
  - A borrowed node view serialized straight to OfficeCLI's `DocumentNode` JSON, built from static per-element property tables (no per-node map).
  - Leaf-owned evictable indexes.
  - `/body/p[N]` is OfficeCLI's direct-body paragraph, counted on the MCE-visible view. M2 maps it to the edit transaction's numbering under 0011's admission check.
- The default source-backed owners answer little today: DOCX exposes only the main story, PPTX only slide text, XLSX only cells. About 35 core leaf read requests plus 4 correctness fixes are listed in §4.11 (L1 tier 1, L2 tier 2).
- Steps R1–R11. R1–R3 and L1 start before M0a lands, on `uo/m1`; R5 waits for S6, R8 for S9. M1 merges into trunk after M0a. 17 code findings (F-1…F-17) and 21 OfficeCLI deltas.

## M0a — log

- Landed as one merge of `uo/m0a` (`1927763795`). Step detail: `docs/unified-ops/m0a-log.md`.
- `litchi::open`, `open_bytes`, `open_with` and `open_prepared` return the flat, non-exhaustive `File`: 16 bytes, one `Arc` handle per format, plus `Locked` and `Signed`. Detection lives in `litchi_ops::detect`.
- One pointer-sized `Error` with an OfficeCLI-aligned `Code` and a typed `Source`.
- Deleted (S9, −27k lines):
  - the facade's `Document`, `Presentation`, `sheet::Workbook` and sheet openers;
  - the old Markdown adapter;
  - the `__migration` seam.
- The facade, examples, tests, Python bindings and perf harness all use `File` (S7, S8a–d).
- Breaking changes:
  - Changed commits that silently stripped a signature now refuse until `Signed::unsign` (0013, 0014a/b).
  - The eager DOCX owner refuses what the source-backed one refuses (0011); 19/19 Word fixtures are editable source-backed.
- `SourceBackedPackage` clones cheaply (0012).
- S10 gate: passes on the plan's terms.
  - 91/99 facade rows and all 47 kept predecessors pass the raw rule; the other 8 are E1/E2 (above).
  - R5: DOC and the PPT byte route keep about one input size alive after open, and `retain_source(false)` removes it.
  - Follow-ups F3 (drop the whole-child RSS tripwire, which measures the harness's `git` child) and F6 (captures off the shared tmpfs).
- Known failures inherited from the base, unchanged:
  - workspace fmt (iwa-archive, soapberry-zip);
  - 18 facade feature-subset sets;
  - the `keynote_chart_title` fuzz target;
  - rustdoc without features.

## 0015 — log

- M2 plan:
  - Verbs take `&self`: the sugar verbs return a `File`, `edit` returns a `Commit`, and `batch` is atomic or best-effort with per-item results.
  - The registry grows from M1's `PropSpec` and refuses rather than repairs.
  - Batches run sequentially, in runs, with barriers and fusion, and rebase only at family boundaries.
- Critical path: no tier-1 leaf can format existing content today, so the parity engines come first: L-DOCX-W5 (DOCX formatting), L-XLSX-W2 (style allocation) and L-PPTX-W3/W5. They follow 0009 S1–S3 (sharing) and the DOCX run, cell and story addressing.
- Leaf measurements (leaf APIs only):
  - Coalescing is essential: 1k chained XLSX commits took 0.18–12.6 s, against 0.5–7.8 ms in one transaction.
  - The XLSX source-backed cell editor admits 45 of 137 corpus workbooks; the eager route admits 132.
  - PPTX source-backed text edits are 7–18× cheaper than the eager transaction.
  - A DOCX reopen between runs costs 9.4 ms on 10k paragraphs.
  - An XLSX commit rewrites the whole sheet: 53.5 ms for one cell on a 75k-cell sheet.

## Awaiting owner decision

- ~~ADR 0033 acceptance~~ — accepted 2026-10-04 (record 0006).

## 0016 — restart and glossary adoption (2026-10-08)

- Rebased `feat/unified-ops-api` from `dbe66611f1` onto base `32ad64e320`;
  rebased tip before this record is `61630e5635`. Existing M0a/M0b work is
  retained. Historical hashes above are evidence references, not current tips.
- Adopted the owner-supplied glossary/default-policy draft as ADR 0034;
  preserved the existing accepted operations ADR 0033 and its citations.
- Resolved the OPC cheap-clone/fingerprint-memo conflict by sharing the base
  memo through package clones; added regression coverage for clone survival.
- W0 next: rustdoc JSON checker with scoped rules, tests and CI integration.
  Then DOCX → PPTX → XLSX shape waves with simultaneous consumer migrations,
  followed by the remaining format waves and M1/M2 completion. iWork requests
  are recorded for its owners; no iWork implementation changes here.
- No performance, parity, all-format CRUD or milestone-completion claim is
  made by this kickoff. Verification and remaining gates are in record 0016.

## 0017 — W0 glossary enforcement

- Added the rustdoc public-API checker, explicit feature-profile policy, real
  compiler fixtures and CI integration. W0 inspects `litchi-core` with empty
  bans/families; pending API migrations are not claimed conformant.
- Scoped bans cover public names, aliases/methods and public signature type
  names. Registered transaction modules require all six canonical type names.
- 19 tests and the current-tree W0 profile pass. Review findings and rustdoc
  limitations are recorded in 0017; unresolved external globs fail closed.
- Next: DOCX W2 migration with consumer updates and behavior gates. W1 iWork
  ownership and the original all-format CRUD goal remain unchanged.

## 0018 — DOCX canonical entry/error names

- Eager/source `Package::edit()` replaces `edit_document`; main-document
  errors/results now live at `document::{Error, Result}`. Existing edit,
  commit, patch and publication semantics are retained for the ownership step.
- Active consumers, examples, tests and current leaf docs migrated together;
  historical benchmark records keep their original source/API names.
- Minimal/full DOCX glossary profiles enforce removal without public aliases.
- Verification and outstanding W2 ownership work are recorded in 0018.

## 0019 — owned eager DOCX projection

- `Package::document()` returns owned `Document`, retaining a shared physical
  snapshot. It remains usable after package drop and stable across later edits.
- Existing eager mutation paths detach through copy-on-write; clean saves
  preserve the shared graph. Main-part payloads and paragraph indexes remain
  shared, with related-part resolution tied to the retained graph.
- Updated dependent Rust types, scoped style borrows, runnable documentation,
  glossary status and ownership/rollback/no-op regression coverage.
- This completes the eager projection shape, not the immutable package or
  package-owning commit contract. Evidence and remaining work are in 0019.

## 0020 — immutable eager DOCX patch application

- Eager `Package::apply(&patch)` returns a new package and replaces the removed
  mutating `apply_document_patch`. Active consumers and the glossary gate move
  together; the existing validation/publication engine stays private.
- Package clones share the physical graph, metadata and legacy writer state.
  Mutations detach state; failed saves retain edit intent on the affected clone.
- Regression coverage includes immutable apply/inverse/stale rejection, no-op
  sharing, writer/metadata isolation, signatures and encrypted-source policy.
- This is the publication prerequisite for package-bound `Edit`/`Commit`.
  Those shapes, source-backed package publication and writer removal remain
  required; evidence and limits are recorded in 0020.

## 0021 — shared source-backed DOCX publication

- Source-backed `Package` clones share source identity, lazy caches, admission,
  tolerance reports and execution reservations without reopening the archive.
- All 32 package stream publishers borrow their owner. Retained owners support
  repeated publication and retry after sink failure; source-version checks stay
  active. Cache reservations persist until the final retaining owner drops.
- Concurrent clone/no-op and metadata-report tests cover shared state; existing
  managed tests explicitly release the retained source before zero-budget checks.
- This enables retained source publication; package-bound `Edit`/`Commit`, source
  overlays and canonical immutable publication remain required. See record 0021.

## 0022 — eager DOCX package-bound commit

- Canonical `document::Edit` retains its complete eager source. Its commit
  validates and publishes a new `Package` with the patch and diagnostics triad.
- Removed eager `publish_document_edit`/`publish_document_commit`; migrated
  consumers and enforced removal in minimal/full glossary profiles.
- Explicit nominal `document::snapshot` types retain the raw XML engine for
  standalone/source/history callers until their owner migration. Root commits
  cannot carry only a snapshot or an optional owner.
- Tests cover retained related parts, source drop, forked edits, chaining,
  inverse, no-op sharing, later source writer changes and security refusals.
- Source-overlay owners, snapshot composition/history, full package lineage,
  limits and writer removal remain required. Evidence and limits are in 0022.

## 0023 — retained source-backed topology publication

- OPC `PreparedTopology` now cheaply clones its complete prepared candidate
  and reservations. Publication borrows it, enabling retry and repeated output.
- ZIP append plans retain shared generated-member handles; publication does
  not deep-copy prepared relationship XML, member names or appended payloads.
- Added metadata admissions and tests for concurrent clones, retry, source and
  transfer rechecks, candidate readback and final-owner reservation release.
- This is the retained topology foundation, not chained source-package commit.
  Logical overlay revisions and revision-bound XML proofs remain required;
  the original physical source proof is deliberately unchanged. See 0023.

## 0024 — revision-bound source XML edits

- `SourceXmlPart::begin_revision()` establishes a fresh logical edit base over
  the current validated payload, without reopening the ZIP or rescanning XML.
- Range proofs bind to that revision, physical source, part and content type;
  sibling and stale proofs are refused even when their bytes match.
- Physical replacement authority remains the original archive. Managed proofs
  retain payload admissions; sequential revisions do not retain an ancestor chain.
- Chained edits, inverse exact-artifact publication, source fences and managed
  limits are covered. Package-wide effective revisions and canonical source
  commits remain required. See record 0024.

## 0025 — immutable edits of retained effective candidates

- `PreparedTopology::edit_xml` applies a checked edit to effective source XML,
  returning a new flat candidate over the same physical archive.
- Prior additions, omissions, relationships and content types survive; logical
  no-ops preserve the current candidate, and exact inverses restore source bytes.
- Effective parts retain XML provenance. Unchanged metadata and transfer guards
  share ownership; old candidate vectors and intermediate payloads can release.
- Revisions enforce physical/transfer fences, aggregate limits, operation bounds,
  cancellation and managed admission. Swapped callback transactions are refused.
- Added/unproven XML targets, general topology application and canonical source
  DOCX commits remain open. Record 0025 covers evidence and performance limits.

## 0026 — owned effective Part and payload reads

- Retained candidates expose owned lazy Part views and payload handles, including
  unchanged candidates, plus effective iteration and main-document resolution.
- Added/removed Parts, effective relationships and changed payloads route through
  the captured catalog. Views survive owner drop and remain stable across edits.
- New captures recheck base/transfer sources and execution contexts; captured
  metadata and bytes remain immutable readable snapshots. Payloads retain their
  managed admissions and the complete candidate, without a detachable Arc.
- XML proof capture is shared with `edit_xml`. Narrow physical identity, artifact,
  limits and execution accessors avoid requiring a parallel raw package owner.
- Canonical source DOCX integration still needs one effective read/publish adapter
  and logical cache keys. Record 0026 documents those requirements and evidence.

## 0027 — DOCX patch predicate and Debug policy

- Primary main-document `Patch::is_noop()` replaces `changed()` without an alias;
  commit `Diagnostics::changed()` remains. Publishers and harness consumers are
  migrated, and both DOCX glossary profiles enforce the new name.
- Main-document snapshots, patches, operations, transfers, replacement inputs
  and local error/composition wrappers use content-free Debug formatting.
- Regression tests cover compact/pretty output and byte-vector disclosure.
  Display and shared-core diagnostics remain separate open policy work.
- Record 0027 contains verification and the remaining canonical owner scope.

## 0028 — explicit effective graph materialization

- Retained OPC candidates can materialize their effective graph into an owning
  package, preserving edits, additions/removals, relationships and content types.
- Conversion shares retained payloads, checks base/transfer fences and refuses
  managed base or donor escape before payload reads. Signature and ingress
  repair policy survives conversion.
- This supplies DOCX's future effective `to_owned_package` route. Canonical
  source ownership and logical caches remain open; record 0028 also records
  the empty-topology/read-ahead migration hazard.

## 0029 — source DOCX effective-owner boundary

- Same-lineage retained topologies can be adopted into immutable DOCX owners.
  Main-document views, explicit owning conversion and publication use the
  effective graph; logical owners receive fresh semantic caches.
- Pinned effective views retain payload admissions and donor execution guards.
  Queries observe cancellation while captured bytes survive source changes.
- All unmigrated physical operations refuse effective owners, including no-op
  requests. The raw package accessor is fallible; the unified consumer is adapted.
- Canonical source semantic Edit/Commit and specialized effective routes remain
  open. Record 0029 documents scope, verification and current allocation limits.

## 0030 — immutable retained semantic apply

- Source DOCX `Package::apply` returns a new effective owner without ZIP
  serialization/reopening. Snapshots and raw edits read effective main XML.
- Exact source proofs and effective settings policy guard changes; unrelated
  graph edits survive, no-ops preserve signatures, and inverses restore bytes.
- OPC full-document replacement/provenance primitives retain managed admission
  and source fences. Revision direct/replay edits preserve unmanaged identity.
- Canonical package-owning source Edit/Commit, specialized effective routes and
  allocation/performance work remain open. Record 0030 documents verification
  and the exact-byte (rather than strict logical-version) precondition boundary.


## 0031 — canonical source package-owning commit

- Source `Package::edit()` returns `document::Edit<source_backed::Package>`;
  commit returns the new effective owner through the canonical package accessors.
- Generic Edit/Commit share eager/source staging and diagnostics with private
  construction and specialized commit policy. Unified editor consumers migrate.
- Legacy stream/history harnesses explicitly request raw XML transactions.
- Canonical MCE coverage exposed and fixed escaped-namespace compatibility:
  unmanaged full-document apply uses the existing publication audit, while
  managed edits keep typed splice authority and admissions.
- Earlier pending primary-source-triad notes are superseded by this record.
  Specialized families, history, limits, performance and full acceptance remain
  open. See 0031 for scope and verification.


## 0032 — typed primary transaction limits

- `document::LimitKind` and `Error::Limit { kind, observed, maximum }` replace
  free-form resource strings and legacy count field names across primary paths.
- Durable aggregate transfer bytes now fail with a typed limit before payload
  copying; dynamic count/path helpers carry typed dimensions.
- DOCX glossary profiles ban the retired fields at their precise variant scope.
  Public boundary/rollback and durable overflow tests accompany the migration.
- Caller-selected semantic limits, other error families, policy vocabulary and
  full family registration remain open. See 0032 for verification and scope.


## 0033 — metadata on retained source owners

- Source DOCX metadata now follows effective property payloads and relationships,
  including canonical commit owners, additions/redirection and removal.
- A common OOXML prepared-graph adapter reuses existing property validation,
  repairs and decoding. Only the selected payload is decoded.
- Base/donor fences cover success and errors; repair reports stay per-owner and
  deduplicate repeated reads. Managed donor cancellation/release is tested.
- Other specialized reads and metadata writes remain pending. Record 0033
  documents verification, allocation limits and the remaining full-goal scope.


## 0034 — text streaming on retained source owners

- Source DOCX text output follows effective main replacements and relationship
  redirection, including canonical commit owners and managed donor payloads.
- Payload handles stay retained through output and final source checks; declared
  size is checked before capture. Partial-output errors keep exact progress.
- Canonical managed/unmanaged commits, donor cancellation, early limits and MCE
  source-failure precedence have integration coverage.
- Specialized reads/transactions, configurable limits and the wider goal remain
  open. Record 0034 documents verification and allocation boundaries.


## 0035 — shared OPC/OOXML SourceLimits vocabulary

- Renamed the physical profile and builder to SourceLimits/SourceLimitsBuilder
  throughout OPC, OOXML re-exports, unified consumers and active tools, without
  retaining aliases. Defaults, admission and checked builders are unchanged.
- Added minimal/full OPC, PPTX, XLSX and XLSB glossary profiles. Exact-scope bans
  distinguish XLSB's root re-export from unrelated semantic ReadLimits types.
- Broader tests caught an empty-plan publication regression: the one-shot
  path now preserves its exact-copy budget without allocating a retained owner.
- Kind/error names, options envelopes and other limit migrations remain open.
  Record 0035 documents checks and the unrelated existing iWork example error
  encountered by the exploratory workspace all-targets build.


## 0036 — typed physical source-limit errors

- OPC and OOXML expose SourceLimitKind; primary OPC failures use Limit with
  kind/observed/maximum and InvalidLimit with kind/value/maximum.
- Builder failures carry scalar or cross-field bounds. The members/entries
  inconsistency now identifies the excessive member dimension correctly.
- ZIP mapping, cached failures, unified classification and consumer assertions
  retain typed numeric evidence. Glossary profiles ban the retired vocabulary.
- Remaining configuration/dimension and semantic-limit work is documented in
  record 0036; the wider goal remains active.


## 0037 — canonical physical source dimensions

- Shared OPC/OOXML limits use members, entry_bytes, aggregate_bytes, depth and
  events with matching max_* methods and SourceLimitKind variants; no aliases.
- Registered all 24 OPC dimensions and their distinct meanings in the glossary.
  Defaults, admission rules and owner-local semantic profiles are unchanged.
- Migrated consumers, tests, active fuzz/performance tools and added exact-scope
  glossary bans. Record 0037 documents verification and remaining work.
- OpenOptions, semantic profiles and the wider unified goal remain active.


## 0038 — owning DOCX OpenOptions

- Added one crate-root OpenOptions envelope for owning DOCX ingress, canonical
  byte-slice constructors and configured password routes; removed the replaced
  limit-taking and ad-hoc constructors without aliases.
- Migrated unified decrypted adoption and consumers. The same source policy
  reaches clear OPC parsing while encryption mode remains retained.
- Added path/reader/bytes boundary and independent encryption/source limit
  checks; glossary rules ban the retired owning-Package methods.
- Record 0038 records evidence and remaining lazy-owner, semantic and crypto
  policy work. The broader unified goal remains active.


## 0039 — DOCX positional options and resource retention

- Replaced 17 path/positional constructor variants with canonical open and
  from_read_at options routes, using the existing root OpenOptions envelope.
- Cache, read-ahead and execution settings reach OPC admission and deferred
  reads; cloned options share budgets and cancellation. Path cancellation is
  checked before file opening.
- Owning routes explicitly refuse lazy policies before I/O, with typed unified
  classification. Sequential reader accounting and owner consolidation remain
  required; this does not declare the ingress family complete.
- Record 0039 documents migrated consumers, verification and remaining work.


## 0040 — managed sequential DOCX ingress

- Lazy reader and byte-slice ingress use OpenOptions; the old limits-only
  reader constructor is deleted and active consumers are migrated.
- Retained source buffers own their memory reservations through package clones
  and artifacts. Growth admits old/new peaks; reads, retries, copies and input
  counts obey caller execution policies, including the bounded EOF probe.
- Added focused failure/lifetime tests, a managed sequential fuzz target and a
  deterministic 1,024-mode smoke driver. Record 0040 documents results.
- Owner consolidation, semantic/crypto policy migration and the wider unified
  goal remain active; no complete milestone or performance gain is claimed.


## 0041 — lazy DOCX canonical publication and atomic save

- Replaced lazy Package::to_stream with write_to and migrated callers; added
  save (Full durability) and save_with_durability over retained effective graphs.
- Shared OPC atomic staging admits its fixed buffer to the destination owner's
  context and checks caller source state after file sync, before replacement.
  Pre-rename failures preserve the destination; post-rename sync errors retain
  their Committed classification.
- Record 0041 documents durability, preservation, refusal and cleanup tests.
- Owning/specialized output names, budgeted to_bytes, owner consolidation and
  the wider unified goal remain active. No complete family is claimed.


## 0042 — owning DOCX canonical stream output

- Owning Package now exposes write_to; to_stream is removed without an alias.
  Both DOCX owner implementations use the glossary's canonical stream verb.
- Migrated unit/integration tests, operations consumers, active benchmark code
  and the sequential-ingress fuzz smoke fixture producer. Historical evidence
  and low-level OPC names remain unchanged.
- The same serialization path retains encryption refusal, signature policy,
  property staging and retry behavior. Record 0042 lists verification.
- Owner consolidation, immutable owning publication, specialized APIs and the
  broader all-format unified goal remain active.


## 0043 — typed owning DOCX publication failures

- Removed the owning publisher's conversion of OPC failures into formatted I/O
  strings. Error::Opc now preserves typed causes and IncompleteOutput counts.
- Regression tests distinguish zero-byte and partial exact-source failures,
  prove byte-exact retry, and check typed partial-output errors while custom
  property staging rolls back and remains retryable.
- Record 0043 documents verification. Broader error normalization and the
  complete unified goal remain active.


## 0044 — owning DOCX bytes and encryption policy

- Added canonical to_bytes and explicit to_plain_bytes over the existing owning
  publisher. Signed no-op/refusal and pending-edit behavior use the same path.
- Adopted shared OOXML encryption provenance and typed EncryptionPolicy causes
  for ordinary output and missing retained modes; explicit plaintext preserves
  provenance. Unified classification keeps the refusal typed.
- Added feature-isolated, encryption-mode and byte-publication coverage; see
  record 0044. Managed lazy bytes, immutable publication, owner consolidation
  and the full unified-operations goal remain active.


## 0045 — typed DOCX direct property reads

- Added typed run, paragraph, table, row and cell property projections plus the
  positional run set and inline classification needed by the formatting work.
- Adapted the earlier M1 leaf candidate to inherited namespace bindings, strict
  attribute checks, content-free diagnostics and touched-view Debug output.
  Run and cell reads exclude nested owners and revision snapshots.
- Added bounded fuzz coverage and corpus/regression checks. Record 0045 lists
  source provenance, verification and remaining scope. M1/M2 and the complete
  all-format goal remain active; no candidate branch history was merged.


## 0046 — reversible DOCX formatting and body structure

- Adapted the earlier formatting-engine leaf into the current package-bound
  Edit API: typed run/paragraph properties and paragraph/run insert/remove/move.
  Exact patch inverses, durable replay and composition use the same operations.
- Threaded inherited namespaces through existing-fragment readers and splicers;
  added strict replay checks, opaque-comment preservation, namespace-checked
  dependency attributes, bounded generation and content-free Debug output.
- Corpus, owner publication, inverse and policy tests are recorded in 0046.
  This advances the DOCX write kernel; the property registry, managed formatting,
  caller semantic limits, performance acceptance and all-format goal remain open.


## 0047 — node projection and shared value coercion cores

- Added borrowed node views, typed paths, static property specifications and
  streaming JSON projection, with one shared length/color coercion model.
- Adapted earlier source-only candidates to explicit raw identity construction,
  content-free diagnostics, strict depth/non-finite refusal and Word twips.
- Added golden, allocation, minimal-feature and bounded parser fuzz checks;
  record 0047 lists validation and explicit parity/allocation limitations.
- Format adapters remain placeholders. Registry dispatch, File query/write APIs,
  M1/M2 integration and the complete all-format goal remain active.


## 0048 — bounded path resolution core

- Added IntoPath/PathArg, Excel notation, typed grammar errors and a static
  family-vocabulary resolver with relative navigation and get/find semantics.
- Adapted source-only candidate 85300f74e2 to raw-only native identities,
  content-free diagnostics/Debug, typed-input bounds and canonical Unicode
  name matching. Lowercase whole-column notation stays borrowed.
- Record 0048 documents validation, preserved typed causes and allocation/parity
  limitations. Concrete adapters, File navigation, query compilation, registry
  dispatch and the full M1/M2/all-format goal remain open.

### 0049 — DOCX indexed block prerequisite

Record [0049](0049-docx-block-index.md) adds typed direct-block navigation,
namespace-preserving values and generation memos, adapted to immutable eager and
effective source owners. Managed detached views refuse, including donor-backed
projections. The DOCX unified adapter and M1 closure remain pending; the complete
all-format goal remains active.

### 0050 — Query engine prerequisite

Record [0050](0050-query-engine.md) integrates bounded plan compilation and
streaming evaluation, with explicit raw identity admission and content-free
errors/Debug. Test-only Word, presentation and workbook adapters exercise the
engine. Concrete DOCX and File dispatch remain pending; M1 and the full goal
remain active.

### 0051 — DOCX read adapter integration

Record [0051](0051-docx-read-adapter.md) integrates DOCX nodes, navigation,
query binding, projections and internal document views. ADR 0034 adaptations
keep ordinary paths positional and native identities behind raw admission.
Source freshness, invalid-property refusal and feature-specific snapshots are
covered by regressions. File dispatch, M1 acceptance and the full goal remain
active.

### 0052 — File unified read entry points

Record [0052](0052-file-read-dispatch.md) connects DOCX navigation, lazy queries
and text/JSON views to File and the `litchi` facade. Walks now own their adapter,
with retained source checks and borrowed-plan reuse. Security wrappers and raw
identity admission have public dispatch coverage. Other adapters, capability
publication and M1 acceptance remain open. Git metadata is still read-only, so
records 0051–0052 remain uncommitted working-tree changes.

## Capability reports (record 0053)

File/facade capability discovery publishes static DOCX read metadata and
truthful protected/pending-format refusals. Ordinary reports hide native
selection fields and paths. Generated OfficeCLI deltas and CI snapshot-hash
checks expose remaining parity work. Complete gap triage, registry writes,
other adapters and full-goal verification remain pending. Changes in records
0051–0053 remain uncommitted because Git metadata is read-only.

## XLSX leaf read prerequisites (record 0054)

Focused `workbook::source::{Grid, CacheDiagnostics}`, bounded shared retention,
worksheet delegates and early-stop visits are integrated. Cold cell reads
retain streaming; metadata projection enforces caller physical XML limits.
R3/R4 formatting/table prerequisites and the concrete XLSX adapter remain
pending. Source/semantic-limit and owner vocabulary migrations are incomplete.
The full goal remains active; records 0051–0054 are uncommitted due to read-only
Git metadata.


## XLSX formatting and tables (record 0055)

Readable formatting, lineage-bound source style keys, source table access,
checked table bands and ambiguity-refusing Unicode name lookup are integrated.
Caller XML policy and content-free diagnostics passed regression checks. The concrete
XLSX adapter and performance acceptance remain pending. The full goal is active;
records 0051–0055 remain uncommitted because linked-worktree Git metadata is
read-only. See [0055](0055-xlsx-formatting-and-tables.md).


## XLSX unified read adapter (record 0056)

XLSX now participates in File navigation, lazy queries, typed node views and
text/annotated/outline/stats views. Ordinary canonical paths are positional;
native number-format predicates require explicit raw admission. Source fences,
Unicode ambiguity, caller limits and fixed per-walk allocation costs are
covered by regression checks. Operations tests, facade tests, feature builds,
Clippy, rustdoc, workspace compilation, glossary and boundaries pass. M1 acceptance,
remaining adapters and all later milestones remain open. See
[0056](0056-xlsx-unified-read-adapter.md).


## PPTX retained-scene foundation (record 0057)

Eager slides, layouts and masters expose independently retained scenes with
checked positional/name selection, typed ambiguity and content-free diagnostics.
Layout and placeholder token classifiers are integrated. Source-backed scene
caching, related parts and the unified PPTX adapter remain open. See
[0057](0057-owned-pptx-scenes.md); the full goal remains active.


## Source-backed PPTX scenes (record 0058)

Source slide scenes now use the independent owned projection and a bounded
weighted LRU, with managed no-retention, source/execution fences and physical
XML limits before MCE. Scene handles do not retain their presentation owner.
Related parts, richer property/text projection and unified PPTX dispatch remain
open. See [0058](0058-source-pptx-scenes.md); the full goal remains active.


## PPTX lazy shape properties (record 0059)

Non-visual properties and full DrawingML transforms are decoded on demand;
native payload relationships and spans use explicit raw access. Namespace,
ambiguity, retained-source parity and diagnostic redaction checks pass.
Formatting tab stops no longer appear as literal text. Related parts, R7 text
projection and unified PPTX dispatch remain open; the full goal is active.
See [0059](0059-pptx-shape-properties.md).


## Shared DrawingML text bodies (record 0060)

PPTX lazy body reads and SpreadsheetDrawing now share paragraph/run/property
decoding. Strict refusal, raw native identities, retained limits, redacted
diagnostics and atomic authoring guards replace the candidate's lossy legacy
policy. Namespace preparation has a root-inheritance fast path. Related parts,
unified PPTX dispatch and the full goal remain open.
See [0060](0060-shared-drawingml-text-bodies.md).


## Source PPTX related parts (record 0061)

Source slides now acquire detached layout, notes, legacy comments and basic
chart snapshots. Native identities remain under explicit raw access. Reads
apply source/execution fences, bounded preprocessing, strict relationship
validation and rollback of failed scene publication. Full PPTX tests and the
final targeted related-part regressions pass. Presentation metadata, full
chart/modern-comment coverage and unified PPTX dispatch remain pending.
See [0061](0061-source-pptx-related-parts.md); the full goal is active.


## Source PPTX presentation catalog (record 0062)

Source presentations now expose detached masters, layouts and sections in
authored order, plus explicit raw root access. Shared leaf codecs enforce
core namespaces, ID-list structure and schema-valid dimensions. Collection
reads bound accounted retained bytes and roll back failed scene publication.
Full PPTX tests, focused source tests, Clippy, rustdoc, workspace compilation
and both glossary profiles pass. Modern comments, full charts, unified PPTX
dispatch and performance acceptance remain pending; the full goal is active.
See [0062](0062-source-pptx-presentation-catalog.md).


## Source PPTX modern comments (record 0063)

Selected slides now acquire detached modern comments and replies with semantic
text and author names. The source read validates its selected author closure,
preserves inherited body namespaces, and keeps wire identities under explicit
raw access. Shared codecs accept valid XML character references and retain typed
resource refusals. Full PPTX tests, Clippy, rustdoc, workspace compilation and
both glossary profiles pass. Full charts, unified PPTX dispatch and performance
acceptance remain pending; the full goal is active.
See [0063](0063-source-pptx-modern-comments.md).
