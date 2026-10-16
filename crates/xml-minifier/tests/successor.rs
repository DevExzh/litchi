//! `VerifiedSource::verify_successor` returns exactly what `verify_source`
//! returns for the replacement, and proves an edit between two recorded
//! content boundaries by scanning only the edited run (record 0313).
//!
//! Every test compares the verdict with the complete audit, including the
//! error and its offset; debug builds also re-derive every splice proof from
//! a complete audit inside the library.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "each test states one fixed expected verdict"
)]

use core::ops::Range;
use std::sync::Arc;

use xml_minifier::audit::{
    IndexDepth, Limits, ReplacementError, ReplacementProof, VerifiedSource, verify_source,
    verify_source_replacement,
};

const HEAD: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document xmlns:w=\"urn:w\" xmlns:r=\"urn:r\"><w:body>";
const TAIL: &str = "<w:sectPr/></w:body></w:document>";

fn paragraph(index: usize) -> String {
    format!(
        "<w:p w:rsidR=\"00A{index:03}\"><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">Paragraph {index} holds </w:t></w:r><w:r><w:t>words &amp; more.</w:t></w:r></w:p>"
    )
}

/// A document of `count` paragraphs, the shape `docx_large_bench` edits.
fn document(count: usize) -> String {
    let mut text = String::from(HEAD);
    for index in 0..count {
        text.push_str(&paragraph(index));
    }
    text.push_str(TAIL);
    text
}

fn shared(text: &[u8]) -> Arc<Vec<u8>> {
    Arc::new(text.to_vec())
}

/// Run the successor audit and require its verdict to be the complete
/// audit's, error and offset included.
fn checked(base: &VerifiedSource, replacement: &[u8], limits: Limits) -> Option<ReplacementProof> {
    let actual = base.verify_successor(shared(replacement), limits);
    let expected = verify_source(replacement, limits);
    match (&actual, &expected) {
        (Ok((proof, _how)), Ok(_)) => {
            assert!(proof.covers(proof.bytes().as_slice(), limits));
            assert_eq!(proof.bytes().as_slice(), replacement);
        },
        (Err(actual), Err(expected)) => assert_eq!(actual, expected),
        _ => panic!(
            "successor verdict {actual:?} differs from the complete audit {expected:?}\nreplacement: {}",
            String::from_utf8_lossy(replacement)
        ),
    }
    actual.ok().map(|(_proof, how)| how)
}

fn indexed(text: &str) -> VerifiedSource {
    VerifiedSource::verify_indexed(
        shared(text.as_bytes()),
        Limits::default(),
        IndexDepth::default(),
    )
    .unwrap()
}

fn splice(original: Range<usize>, replacement: Range<usize>) -> ReplacementProof {
    ReplacementProof::Splice {
        original,
        replacement,
    }
}

fn at(document: &str, needle: &str) -> usize {
    document.find(needle).expect("fixture needle")
}

/// `document` with `fragment` inserted before paragraph `index`.
fn insert_before(document: &str, index: usize, fragment: &str) -> String {
    let offset = at(document, &paragraph(index));
    format!("{}{fragment}{}", &document[..offset], &document[offset..])
}

#[test]
fn indexed_proofs_have_the_source_verdict() {
    let limits = Limits::default();
    let valid = document(3);
    let proof =
        VerifiedSource::verify_indexed(shared(valid.as_bytes()), limits, IndexDepth::default())
            .unwrap();
    assert!(proof.covers(proof.bytes().as_slice(), limits));
    assert!(format!("{proof:?}").contains("boundaries: Some("));
    for invalid in [
        "<w:p>",
        "<x:p/>",
        "<!DOCTYPE d><d/>",
        "<a/><b/>",
        "<a b=\"<\"/>",
    ] {
        assert_eq!(
            VerifiedSource::verify_indexed(
                shared(invalid.as_bytes()),
                limits,
                IndexDepth::default()
            )
            .unwrap_err(),
            verify_source(invalid.as_bytes(), limits).unwrap_err(),
            "{invalid}"
        );
    }
}

#[test]
fn a_paragraph_insertion_is_proved_by_the_paragraphs_it_touches() {
    let original = document(20);
    let base = indexed(&original);
    let inserted = "<w:p><w:r><w:t>Inserted</w:t></w:r></w:p>";
    let replacement = insert_before(&original, 10, inserted);
    let proof = checked(&base, replacement.as_bytes(), Limits::default()).unwrap();
    // The common prefix runs into paragraph 10 (`<w:p`); measured again from
    // the boundary before it, the run is empty and only the new paragraph is
    // scanned.
    let start = at(&original, &paragraph(10));
    assert_eq!(
        proof,
        splice(start..start, start..start + inserted.len()),
        "{proof:?}"
    );

    // Another element.
    let replacement = insert_before(&original, 10, "<w:tbl/>");
    assert_eq!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        Some(splice(start..start, start..start + "<w:tbl/>".len()))
    );

    // At the end of the body and as the first child of the document element.
    let sect = at(&original, "<w:sectPr/>");
    let replacement = format!("{}<w:p/>{}", &original[..sect], &original[sect..]);
    assert!(matches!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        Some(ReplacementProof::Splice { .. })
    ));
    let body = at(&original, "<w:body>");
    let replacement = format!("{}<w:background/>{}", &original[..body], &original[body..]);
    assert_eq!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        Some(splice(body..body, body..body + "<w:background/>".len()))
    );
}

#[test]
fn text_and_attribute_edits_are_proved_by_their_paragraph() {
    let original = document(20);
    let base = indexed(&original);
    let start = at(&original, &paragraph(7));
    let end = start + paragraph(7).len();
    for (from, to) in [
        ("Paragraph 7 holds", "Paragraph seven now holds"),
        ("w:rsidR=\"00A007\"", "w:rsidR=\"00B007\""),
        (
            "w:rsidR=\"00A007\"",
            "w:rsidR=\"00A007\" w:rsidRDefault=\"00A0\"",
        ),
        // `>` is allowed in an attribute value and in character data.
        ("w:rsidR=\"00A007\"", "w:rsidR=\"a>b\""),
        ("Paragraph 7 holds", "Paragraph 7 > holds"),
    ] {
        let edited = paragraph(7).replace(from, to);
        let replacement = original.replacen(&paragraph(7), &edited, 1);
        assert_eq!(
            checked(&base, replacement.as_bytes(), Limits::default()),
            Some(splice(start..end, start..start + edited.len())),
            "{edited}"
        );
    }
}

#[test]
fn a_boundary_inside_an_attribute_or_tag_never_hides_a_defect() {
    let original = document(20);
    let base = indexed(&original);
    let target = paragraph(7);
    let value = at(&target, "00A007");
    let name = at(&target, "rsidR");
    for edited in [
        // Inside an attribute value.
        format!("{}\"{}", &target[..value], &target[value..]),
        format!("{}<{}", &target[..value], &target[value..]),
        format!("{}&bogus;{}", &target[..value], &target[value..]),
        format!("{}\u{1}{}", &target[..value], &target[value..]),
        // Inside an attribute name, and a duplicate attribute.
        format!("{}:{}", &target[..name], &target[name..]),
        target.replace("w:rsidR=\"00A007\"", "w:rsidR=\"1\" w:rsidR=\"2\""),
        // Splitting a tag.
        target.replace("<w:r><w:rPr>", "<w:r <w:rPr>"),
        target.replace("<w:i/>", "<w:i/"),
    ] {
        let replacement = original.replacen(&target, &edited, 1);
        assert_eq!(
            checked(&base, replacement.as_bytes(), Limits::default()),
            None,
            "{edited}"
        );
    }
}

#[test]
fn unbalanced_fragments_are_refused_with_the_complete_audit_error() {
    let original = document(20);
    let base = indexed(&original);
    for fragment in [
        "<w:p>",
        "</w:p>",
        "<w:p></w:r>",
        "<w:p><w:r></w:p></w:r>",
        "</w:body></w:document><w:document xmlns:w=\"urn:w\"><w:body>",
        // A second document element, a declaration and a DOCTYPE.
        "<?xml version=\"1.0\"?>",
        "<!DOCTYPE w:document>",
    ] {
        let replacement = insert_before(&original, 10, fragment);
        assert_eq!(
            checked(&base, replacement.as_bytes(), Limits::default()),
            None,
            "{fragment}"
        );
    }
    // Balanced in count but closing the parent and reopening it: well-formed
    // (two bodies), and never proved by a run, which may not close its
    // parent.
    for fragment in ["</w:body><w:body>", "</w:body><w:p/><w:body>"] {
        let replacement = insert_before(&original, 10, fragment);
        assert_eq!(
            checked(&base, replacement.as_bytes(), Limits::default()),
            Some(ReplacementProof::Complete),
            "{fragment}"
        );
    }
}

#[test]
fn namespace_constraints_are_checked_in_the_parents_scope() {
    // The body declares a prefix of its own; one paragraph declares another
    // for itself only.
    let original = format!(
        "<w:document xmlns:w=\"urn:w\"><w:body xmlns:b=\"urn:b\">{}<w:p xmlns:s=\"urn:s\"><s:x/></w:p>{}{}</w:body></w:document>",
        paragraph(0),
        paragraph(1),
        paragraph(2)
    );
    let base = indexed(&original);
    for (fragment, passes) in [
        // In scope from the document element and from the body.
        ("<w:p><b:x/></w:p>", true),
        // Redeclarations are legal and scoped to the fragment.
        ("<w:p xmlns:w=\"urn:other\"><w:r/></w:p>", true),
        ("<w:p xmlns:s=\"urn:s2\"><s:y/></w:p>", true),
        ("<w:p xmlns=\"urn:default\"><x/></w:p>", true),
        // Declared only inside an earlier sibling: out of scope here.
        ("<w:p><s:x/></w:p>", false),
        ("<q:p/>", false),
        ("<w:p q:a=\"1\"/>", false),
        // Reserved prefixes and names, and undeclaring a prefix.
        ("<w:p xmlns:xml=\"urn:not-xml\"/>", false),
        ("<w:p xmlns:xmlns=\"urn:x\"/>", false),
        ("<w:p xmlns:x=\"http://www.w3.org/2000/xmlns/\"/>", false),
        ("<w:p xmlns:w=\"\"/>", false),
        // Two prefixes bound to one namespace name make one expanded name.
        (
            "<w:p xmlns:a=\"urn:u\" xmlns:c=\"urn:u\" a:v=\"1\" c:v=\"2\"/>",
            false,
        ),
        // `b` is bound to the body's namespace, as `w2` is here.
        ("<w:p xmlns:w2=\"urn:b\" b:v=\"1\" w2:v=\"2\"/>", false),
    ] {
        let offset = at(&original, &paragraph(2));
        let replacement = format!("{}{fragment}{}", &original[..offset], &original[offset..]);
        let proof = checked(&base, replacement.as_bytes(), Limits::default());
        assert_eq!(proof.is_some(), passes, "{fragment}: {proof:?}");
        if passes {
            assert!(
                matches!(proof, Some(ReplacementProof::Splice { .. })),
                "{fragment}: {proof:?}"
            );
        }
    }
}

#[test]
fn references_cdata_comments_and_characters_are_checked() {
    let original = document(20);
    let base = indexed(&original);
    for (fragment, passes) in [
        (
            "<w:p><w:t>&amp;&lt;&gt;&quot;&apos;&#65;&#x10FFFF;</w:t></w:p>",
            true,
        ),
        ("<w:p><w:t>&nbsp;</w:t></w:p>", false),
        ("<w:p><w:t>&#0;</w:t></w:p>", false),
        ("<w:p><w:t>&#xD800;</w:t></w:p>", false),
        ("<w:p><w:t>&amp</w:t></w:p>", false),
        ("<w:p><w:t><![CDATA[a<b&c]]></w:t></w:p>", true),
        ("<![CDATA[body-level]]>", true),
        ("<w:p><w:t>x]]>y</w:t></w:p>", false),
        ("<w:p><![CDATA[x]]>]]></w:p>", false),
        ("<w:p><![CDATA[unterminated</w:p>", false),
        ("<!-- a comment -->", true),
        ("<!-- a -- b -->", false),
        ("<!-- ends with --->", false),
        ("<?pi data?>", true),
        ("<?xml-stylesheet x?>", true),
        ("<?xml data?>", false),
        ("<w:p><w:t>\u{1}</w:t></w:p>", false),
        ("<w:p><w:t>\u{FFFE}</w:t></w:p>", false),
        ("<w:p a=\"\u{1}\"/>", false),
    ] {
        let replacement = insert_before(&original, 10, fragment);
        let proof = checked(&base, replacement.as_bytes(), Limits::default());
        assert_eq!(proof.is_some(), passes, "{fragment}: {proof:?}");
    }

    // Invalid UTF-8 inside the fragment.
    let mut replacement = insert_before(&original, 10, "<w:p>XX</w:p>").into_bytes();
    let offset = at(&String::from_utf8_lossy(&replacement), "XX");
    replacement[offset] = 0xC3;
    replacement[offset + 1] = 0x28;
    assert_eq!(checked(&base, &replacement, Limits::default()), None);
}

#[test]
fn fragments_that_do_not_begin_and_end_with_markup_are_audited_completely() {
    let original = document(20);
    let base = indexed(&original);
    for fragment in ["text<w:p/>", "text", " ", "&amp;"] {
        let replacement = insert_before(&original, 10, fragment);
        assert!(
            !matches!(
                checked(&base, replacement.as_bytes(), Limits::default()),
                Some(ReplacementProof::Splice { .. })
            ),
            "{fragment}"
        );
    }
    // Text at the end of the fragment is covered by the sibling after it.
    let replacement = insert_before(&original, 10, "<w:p/>text");
    let start = at(&original, &paragraph(10));
    assert_eq!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        Some(splice(
            start..start + paragraph(10).len(),
            start..start + paragraph(10).len() + "<w:p/>text".len()
        ))
    );
    // Removing a run is audited completely: text around it would merge. Here
    // the merged text holds `]]>`, which neither side held.
    let original = "<r><a>x]]<b/>></a><pad>padding that keeps the run small</pad></r>";
    let base = indexed(original);
    let removed = original.replace("<b/>", "");
    assert_eq!(checked(&base, removed.as_bytes(), Limits::default()), None);
    let removed = original.replace("<b/>", "<!---->");
    assert!(matches!(
        checked(&base, removed.as_bytes(), Limits::default()),
        Some(ReplacementProof::Splice { .. })
    ));
}

#[test]
fn edits_outside_the_document_element_are_audited_completely() {
    let original = document(20);
    let base = indexed(&original);
    for replacement in [
        format!("{original}<w:p/>"),
        format!("{original}text"),
        format!("{original}<!-- trailing -->"),
        original.replacen("\r\n", "\n", 1),
        original.replacen("standalone=\"yes\"", "standalone=\"no\"", 1),
        original.replacen("<w:document ", "<w:document x=\"1\" ", 1),
        original.replacen("</w:document>", "</w:document >", 1),
    ] {
        let how = checked(&base, replacement.as_bytes(), Limits::default());
        assert!(
            !matches!(how, Some(ReplacementProof::Splice { .. })),
            "{replacement}: {how:?}"
        );
    }
}

#[test]
fn distant_edits_need_a_run_of_at_most_half_the_replacement() {
    let original = document(20);
    let base = indexed(&original);
    // Two distant edits make one run spanning most of the body.
    let replacement = original
        .replacen("Paragraph 1 holds", "Paragraph one holds", 1)
        .replacen("Paragraph 18 holds", "Paragraph eighteen holds", 1);
    assert_eq!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        Some(ReplacementProof::Complete)
    );
    // A defect far from a valid insertion is still found.
    let replacement = insert_before(&original, 10, "<w:p/>").replacen("</w:r></w:p>", "</w:p>", 1);
    assert_eq!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        None
    );
}

#[test]
fn every_budget_is_re_totalled_for_the_whole_replacement() {
    let original = document(20);
    let fragment =
        "<w:p><w:r><w:rPr><w:b><w:x a=\"1\" b=\"2\"/></w:b></w:rPr><w:t>deep</w:t></w:r></w:p>";
    let replacement = insert_before(&original, 10, fragment);
    let report = verify_source(original.as_bytes(), Limits::default()).unwrap();
    let after = verify_source(replacement.as_bytes(), Limits::default()).unwrap();
    let limits =
        |configure: &dyn Fn(xml_minifier::audit::Builder) -> xml_minifier::audit::Builder| {
            configure(Limits::builder()).build()
        };
    for (name, tight) in [
        ("events", limits(&|b| b.events(after.events() - 1).unwrap())),
        (
            "attributes",
            limits(&|b| b.attributes(after.attributes() - 1).unwrap()),
        ),
        (
            "text",
            limits(&|b| b.text_bytes(after.text_bytes() - 1).unwrap()),
        ),
        (
            "bytes",
            limits(&|b| b.bytes(replacement.len() - 1).unwrap()),
        ),
        ("depth", limits(&|b| b.depth(report.max_depth()).unwrap())),
    ] {
        let base = VerifiedSource::verify_indexed(
            shared(original.as_bytes()),
            tight,
            IndexDepth::default(),
        )
        .unwrap();
        assert_eq!(
            checked(&base, replacement.as_bytes(), tight),
            None,
            "{name}"
        );
    }
    // Under the replacement's exact totals the run still proves it.
    let exact = Limits::builder()
        .events(after.events())
        .unwrap()
        .attributes(after.attributes())
        .unwrap()
        .text_bytes(after.text_bytes())
        .unwrap()
        .bytes(replacement.len())
        .unwrap()
        .depth(after.max_depth())
        .unwrap()
        .build();
    let base =
        VerifiedSource::verify_indexed(shared(original.as_bytes()), exact, IndexDepth::default())
            .unwrap();
    assert!(matches!(
        checked(&base, replacement.as_bytes(), exact),
        Some(ReplacementProof::Splice { .. })
    ));
}

#[test]
fn proofs_without_boundaries_or_under_other_limits_audit_completely() {
    let original = document(20);
    let replacement = insert_before(&original, 10, "<w:tbl/>");
    // A plain proof records nothing.
    let plain = VerifiedSource::verify(shared(original.as_bytes()), Limits::default()).unwrap();
    assert_eq!(
        checked(&plain, replacement.as_bytes(), Limits::default()),
        Some(ReplacementProof::Complete)
    );
    // Other limits than the base proof's.
    let other = Limits::builder()
        .events(Limits::default().max_events() - 1)
        .unwrap()
        .build();
    assert_ne!(other, Limits::default());
    let base = indexed(&original);
    assert_eq!(
        checked(&base, replacement.as_bytes(), other),
        Some(ReplacementProof::Complete)
    );
    // A successor of an indexed proof keeps boundaries of its own (record
    // 0317); one of a plain proof keeps none.
    let (successor, _how) = base
        .verify_successor(shared(replacement.as_bytes()), Limits::default())
        .unwrap();
    assert!(format!("{successor:?}").contains("boundaries: Some("));
    let (successor, _how) = plain
        .verify_successor(shared(replacement.as_bytes()), Limits::default())
        .unwrap();
    assert!(format!("{successor:?}").contains("boundaries: None"));
    // Identical bytes have the base's verdict.
    assert_eq!(
        checked(&base, original.as_bytes(), Limits::default()),
        Some(ReplacementProof::Identical)
    );
}

#[test]
fn an_index_over_its_bound_is_not_recorded() {
    // One boundary per four bytes exceeds one per 32.
    let mut flat = String::from("<r>");
    for _ in 0..10_000 {
        flat.push_str("<a/>");
    }
    flat.push_str("</r>");
    let base = indexed(&flat);
    assert!(format!("{base:?}").contains("boundaries: None"));
    let replacement = flat.replacen("<a/>", "<b/>", 1);
    assert_eq!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        Some(ReplacementProof::Complete)
    );
}

#[test]
fn a_byte_order_mark_keeps_physical_offsets() {
    let original = format!("\u{feff}{}", document(20));
    let base = indexed(&original);
    let replacement = insert_before(&original, 10, "<w:tbl/>");
    let start = at(&original, &paragraph(10));
    assert_eq!(
        checked(&base, replacement.as_bytes(), Limits::default()),
        Some(splice(start..start, start..start + "<w:tbl/>".len()))
    );
}

/// `docx_large_bench`'s paragraph shape, which no element window proves an
/// insertion into.
fn plain(count: usize) -> String {
    let mut text = String::from(HEAD);
    for index in 0..count {
        text.push_str(&format!("<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">Paragraph {index} holds </w:t></w:r><w:r><w:t>a few ordinary words of body text.</w:t></w:r></w:p>"));
    }
    text.push_str(TAIL);
    text
}

#[test]
fn the_indexed_pair_audit_has_the_pair_verdict() {
    let limits = Limits::default();
    let original = plain(20);
    let pair = |replacement: &str| {
        let actual = VerifiedSource::verify_indexed_replacement(
            shared(original.as_bytes()),
            shared(replacement.as_bytes()),
            limits,
            IndexDepth::default(),
        );
        let expected =
            verify_source_replacement(original.as_bytes(), replacement.as_bytes(), limits);
        match (&actual, &expected) {
            (Ok((base, proof, _how)), Ok(_)) => {
                assert_eq!(base.bytes().as_slice(), original.as_bytes());
                assert!(format!("{base:?}").contains("boundaries: Some("));
                assert_eq!(proof.bytes().as_slice(), replacement.as_bytes());
            },
            (Err(actual), Err(expected)) => assert_eq!(actual, expected),
            _ => panic!("{actual:?} against {expected:?}"),
        }
        actual.ok().map(|(_base, _proof, how)| how)
    };
    // The element window stays the proof where it exists.
    let replacement = original.replacen("Paragraph 3 holds", "Paragraph 3 held", 1);
    assert!(matches!(
        pair(&replacement),
        Some(ReplacementProof::Window { .. })
    ));
    // An insertion has none: the pair audit scanned the replacement
    // completely, the indexed pair proves it from the same scan.
    let next = at(
        &original,
        "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">Paragraph 10 ",
    );
    let replacement = format!(
        "{}<w:p><w:r><w:t>Inserted</w:t></w:r></w:p>{}",
        &original[..next],
        &original[next..]
    );
    assert_eq!(
        verify_source_replacement(original.as_bytes(), replacement.as_bytes(), limits),
        Ok(ReplacementProof::Complete)
    );
    assert!(matches!(
        pair(&replacement),
        Some(ReplacementProof::Splice { .. })
    ));
    assert_eq!(pair(&original), Some(ReplacementProof::Identical));
    // Either side's failure, in the pair's order.
    assert!(matches!(
        VerifiedSource::verify_indexed_replacement(
            shared(b"<w:p>"),
            shared(original.as_bytes()),
            limits,
            IndexDepth::default(),
        ),
        Err(ReplacementError::Original(_))
    ));
    let broken = format!("{}<w:p>{}", &original[..next], &original[next..]);
    assert_eq!(pair(&broken), None);
}

/// A small deterministic generator, so a failure names its seed.
struct Lcg(u64);

impl Lcg {
    fn next(&mut self) -> usize {
        self.0 = self
            .0
            .wrapping_mul(6_364_136_223_846_793_005)
            .wrapping_add(1_442_695_040_888_963_407);
        (self.0 >> 33) as usize
    }
}

#[test]
fn random_edits_have_the_complete_audit_verdict() {
    let original = document(40);
    let base = indexed(&original);
    let pieces = [
        "<w:p/>",
        "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
        "<w:p>",
        "</w:p>",
        "</w:body><w:body>",
        "<q:p/>",
        "<w:p xmlns:q=\"urn:q\"><q:x/></w:p>",
        "&amp;",
        "&bogus;",
        "<![CDATA[c]]>",
        "]]>",
        "<!-- c -->",
        "\"",
        "'",
        "<",
        ">",
        "=",
        " ",
        "x",
        "\u{1}",
        "é",
    ];
    let mut random = Lcg(0x0313);
    let mut splices = 0;
    for round in 0..3_000 {
        let mut text = original.clone().into_bytes();
        for _ in 0..1 + random.next() % 2 {
            let offset = random.next() % (text.len() + 1);
            match random.next() % 3 {
                0 => {
                    let piece = pieces[random.next() % pieces.len()].as_bytes();
                    text.splice(offset..offset, piece.iter().copied());
                },
                1 if offset < text.len() => {
                    let end = (offset + 1 + random.next() % 8).min(text.len());
                    text.drain(offset..end);
                },
                _ if offset < text.len() => text[offset] = b"<>/\"=a &;"[random.next() % 9],
                _ => {},
            }
        }
        let how = checked(&base, &text, Limits::default());
        if matches!(how, Some(ReplacementProof::Splice { .. })) {
            splices += 1;
        }
        let _ = round;
    }
    // Most valid edits are proved by a run.
    assert!(splices > 100, "{splices}");
}
