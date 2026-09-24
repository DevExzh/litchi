# PPTX SVG attachment lifecycle requirements

The audit explicitly leaves attaching and detaching SVG from an existing raster
picture open after SVG replacement/retargeting. This batch addresses that host
operation; it does not substitute a standalone extension writer for package
publication.

The normative vocabulary is MS-ODRAWXML §2.26 and its §5.24 schema. Section 1.3
specifies the embedded PNG compatibility fallback. The native extension GUID
`{96DAC541-7B7A-43D3-8B79-37D633B846F1}` is corroborated by the three retained
LibreOffice corpus references in `corpus-scan.json`. Those are source fixtures,
not evidence of an Office application accepting newly generated output.

| Requirement | Required evidence |
|---|---|
| Select an existing picture without raw IDs in the ordinary attach call | Public downstream example using slide/picture positions and automatic names |
| Preserve caller payload on rejected insertion | Borrowed ordinary input and retained advanced request on failure; validate limits before copying |
| Represent attached and detached states honestly | Optional SVG snapshot state and checked state transitions |
| Validate the changed dependency closure at commit | Ordinary fallible `commit()` using an editor-borrowed edit; malformed closure refusal before publication, with publication rebuilding its plan after a freshness check |
| Preserve PNG fallback and unrelated XML/media | Exact source/member assertions through save/reopen |
| Publish slide markup, SVG relationship, and media atomically | OPC topology plan, reopened graph checks, failure-before-output tests |
| Detach only the selected owner | Shared relationship-ID and shared media tests, including character-reference-encoded IDs on another owner |
| Preserve unknown extension children, comments, namespace and quote spelling | Opaque-payload and inherited-namespace regressions, including expanding an attributed self-closing extLst |
| Derive edit ranges from raw source | Unrelated MCE transition control and ambiguous picture-MCE refusal |
| Recognize the native owner by resolved namespaces and token-normalized URI | Prefix/default namespace, foreign URI, duplicate-owner, and URI whitespace cases |
| Keep source authorization and exact in-memory inverse | Stale-source tests, inverse application, explicit saved-artifact scope |
| Reject unsafe graph/signature/limit changes before writing | Counting/failing sink and finite-limit tests; exact precharge for self-closing blip/container expansion and namespace-complete fragments |
| Preserve producer-formatted relationship XML | OPC removal regression with declarations, whitespace, comments, and remaining relationships |
| Emit valid supported markup | Offline complete-slide ECMA XSD plus direct SVG MS-ODRAWXML XSD |
| Make resource behavior inspectable | Source-bound fresh-process allocation/timing samples, recomputed reports, many-picture scaling, and large shared namespace contexts with distinct local declarations |

ADR 0003 governs immutable snapshots, atomic commits, dependency closure and
reversible patches; ADR 0004 governs typed semantic verbs; ADR 0005 requires
finite budgets, source sharing and measured claims; ADR 0006 requires preservation,
inert external content and explicit MCE/signature policy. ADR 0011 keeps physical
archive operations in OPC. ADR 0007 requires rejected ownership-consuming
insertions to retain the rejected value; borrowed attachment requests satisfy
that requirement without an error wrapper carrying a large payload. The new host layer must reuse those boundaries.

The existing replacement API remains available. New picture creation/reordering,
linked-resource fetching, SVG rendering/rasterization, native application
acceptance, and durable cross-process patch exchange are separate work.

Final approval requires current production hashes, compile-first gates with no
warning suppressions, independent review, focused and crate-wide tests, and
verifiable evidence. Preparation files or a green older revision do not establish
completion.
