# 0437 integrated ODP gate: read-only audit

Scope: source inspection only. No build, test, benchmark, or runtime command was run.

## Contract that is enforced

The integrated content_structure gate is a strict fixed-profile grammar for the fresh buffered ODP corpus. It requires the fixed root namespace bindings and office:version, the XML declaration values when a declaration is present, the document prelude order, one dp1 drawing-page style, and one body containing one presentation. presentation_seen now rejects a second presentation. Every expected page has draw:name=pageN, draw:style-name=dp1, and draw:master-page-name=Default, followed by exactly two ordered frames. Frame geometry, draw/presentation styles, layer, class, one text box, one paragraph, and paragraph styles are all checked.

The gate rejects foreign elements/attributes, extra namespace bindings, comments, processing instructions, CDATA, DTD, unsupported entity references, text outside the expected paragraph, extra pages/shapes/paragraphs, and incomplete nesting. Paragraph text is decoded before comparison with the title/body fixture. The archive-side checks separately pin the five-member set, mimetype bytes, per-member compression, manifest bindings, styles/meta hashes, the reopened Presentation projection, and the semantic digest. The 20 focused mutations exercise the principal XML structure, page metadata, text, entity, declaration, namespace, and duplicate-element failures.

## Findings requiring claim alignment

1. **ODF ZIP placement is not independently checked.** inspect_odp_buffered_archive compares file_names() as a set and checks compression, but it does not require mimetype to be the first ZIP member/local header or inspect the local-header restrictions that belong to the ODF mimetype rule. A package with the same five names and compression could therefore pass after reordering or changing local-header metadata. If the evidence claims the complete ODF package contract, add an independent first-entry/local-header check; otherwise scope the claim to member set, bytes, manifest, and compression. The current tests do not mutate ZIP order or the mimetype local header.

2. **The XML declaration is accepted as absent.** declaration_seen is used to reject duplicates/misplacement, but EOF does not require it. The producer emits the fixed declaration and the mutation test only changes its version. Either require exactly one declaration for the fixed lexical profile or explicitly document that declaration omission is an accepted semantic variant.

3. **The text oracle shares fixture functions with the producer.** The gate parses XML independently, but expected count/title/body and the medium/large semantic digest are derived from odp_buffered_slide_count, odp_buffered_title, and odp_buffered_body, which are also called by odp_buffered_bytes. The fixed tiny digest and byte/count assertions provide a small independent anchor; they do not independently pin the larger shapes. For a source-independent producer oracle, pin each formal shape's projection hash/counts (or retain the shared functions as the declared corpus specification and state that limitation).

4. **Digest and input-byte definitions differ intentionally but need explicit schema prose.** The semantic digest hashes each title/body with length prefixes and no inter-slide separators, while semantic_input_bytes includes one newline within each slide and two newlines between slides. This is internally consistent, but the report/manifest must describe the two definitions so a portable verifier does not treat the digest byte stream as the measured input projection.

5. **The formal size claim is 64/4,096/8,192 slides.** The 32,768-slide test is a refusal oracle for the default XML attribute limit, not a measured formal shape. Reports and protocol text must keep that refusal evidence separate from performance, allocation, and throughput claims.

## Useful mutation coverage additions

The existing mutations are meaningful. For complete fixed-profile regression coverage, add (or retain as explicit review items) deletion of the XML declaration, missing/extra default page attributes, draw:text-style-name/body P2 changes, body geometry changes (height, y, or presentation:class), missing or duplicated page elements, and ZIP mimetype order/local-header mutations. No source change was made by this audit.
