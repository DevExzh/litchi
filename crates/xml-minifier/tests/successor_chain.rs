//! Chains of successor proofs and depth-3 indexes (record 0317).
//!
//! A successor proved from recorded boundaries carries boundaries derived for
//! its own bytes, so a chain of edits is proved edit by edit. Every test
//! compares each verdict with the complete audit, error and offset included;
//! debug builds also compare every derived index with the index a recording
//! audit of the replacement makes, inside the library.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "each test states one fixed expected verdict"
)]

use core::ops::Range;
use std::sync::Arc;

use xml_minifier::audit::{IndexDepth, Limits, ReplacementProof, VerifiedSource, verify_source};

const HEAD: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document xmlns:w=\"urn:w\" xmlns:r=\"urn:r\"><w:body>";
const TAIL: &str = "<w:sectPr/></w:body></w:document>";

fn paragraph(index: usize) -> String {
    format!(
        "<w:p w:rsidR=\"00A{index:03}\"><w:r><w:t xml:space=\"preserve\">Paragraph {index} holds </w:t></w:r><w:r><w:t>words &amp; more.</w:t></w:r></w:p>"
    )
}

fn document(count: usize) -> String {
    let mut text = String::from(HEAD);
    for index in 0..count {
        text.push_str(&paragraph(index));
    }
    text.push_str(TAIL);
    text
}

const SLIDE_HEAD: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<p:sld xmlns:a=\"urn:a\" xmlns:r=\"urn:r\" xmlns:p=\"urn:p\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";
const SLIDE_TAIL: &str =
    "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";

fn shape(index: usize, text: &str) -> String {
    format!(
        "<p:sp><p:nvSpPr><p:cNvPr id=\"{}\" name=\"TextBox {index}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"{}\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>",
        index + 2,
        index * 100,
    )
}

/// A slide of `count` text boxes, the shape the PPTX legs edit.
fn slide(count: usize) -> String {
    let mut text = String::from(SLIDE_HEAD);
    for index in 0..count {
        text.push_str(&shape(index, &format!("Shape {index}")));
    }
    text.push_str(SLIDE_TAIL);
    text
}

fn shared(text: &[u8]) -> Arc<Vec<u8>> {
    Arc::new(text.to_vec())
}

fn indexed(text: &str, depth: IndexDepth) -> VerifiedSource {
    VerifiedSource::verify_indexed(shared(text.as_bytes()), Limits::default(), depth).unwrap()
}

/// How many boundaries a proof keeps, from its `Debug` form.
fn boundaries(proof: &VerifiedSource) -> Option<usize> {
    let debug = format!("{proof:?}");
    let (_, rest) = debug.split_once("boundaries: Some(")?;
    rest.split(')').next()?.parse().ok()
}

/// Prove `replacement` from `base` and require the complete audit's verdict.
/// A successor that passes must keep exactly as many boundaries as a fresh
/// recording audit at the base's depth, unless it keeps none.
fn step(
    base: &VerifiedSource,
    replacement: &[u8],
    limits: Limits,
) -> Option<(VerifiedSource, ReplacementProof)> {
    let actual = base.verify_successor(shared(replacement), limits);
    let expected = verify_source(replacement, limits);
    match (&actual, &expected) {
        (Ok((proof, _how)), Ok(_)) => {
            assert_eq!(proof.bytes().as_slice(), replacement);
            assert!(proof.covers(proof.bytes().as_slice(), limits));
            if let (Some(kept), Some(depth)) = (boundaries(proof), base.index_depth()) {
                let fresh =
                    VerifiedSource::verify_indexed(shared(replacement), limits, depth).unwrap();
                assert_eq!(Some(kept), boundaries(&fresh));
                assert_eq!(proof.index_depth(), Some(depth));
            }
        },
        (Err(actual), Err(expected)) => assert_eq!(actual, expected),
        _ => panic!(
            "successor verdict {actual:?} differs from the complete audit {expected:?}\nreplacement: {}",
            String::from_utf8_lossy(replacement)
        ),
    }
    actual.ok()
}

fn splice(original: Range<usize>, replacement: Range<usize>) -> ReplacementProof {
    ReplacementProof::Splice {
        original,
        replacement,
    }
}

fn at(text: &str, needle: &str) -> usize {
    text.find(needle).expect("fixture needle")
}

#[test]
fn chained_edits_each_scan_only_their_own_paragraph() {
    // Record 0313 §4.1: two edits far apart used to fall back to the full
    // audit, because the second was proved from the first's base.
    let original = document(40);
    let base = indexed(&original, IndexDepth::default());
    let first = original.replacen("Paragraph 2 holds", "Paragraph 2 held", 1);
    let (first_proof, how) = step(&base, first.as_bytes(), Limits::default()).unwrap();
    let start = at(&first, &paragraph(2).replace("holds", "held"));
    let len = paragraph(2).len() - 1;
    assert_eq!(how, splice(start..start + len + 1, start..start + len));
    assert!(boundaries(&first_proof).is_some());

    // The second edit is 36 paragraphs away: proved from the first edit's
    // derived boundaries by its own paragraph alone.
    let second = first.replacen("Paragraph 38 holds", "Paragraph 38 holds now", 1);
    let (second_proof, how) = step(&first_proof, second.as_bytes(), Limits::default()).unwrap();
    let start = at(&second, "<w:p w:rsidR=\"00A038\">");
    let len = paragraph(38).len();
    assert_eq!(how, splice(start..start + len, start..start + len + 4));

    // An insertion between them, then an edit inside the inserted paragraph.
    let next = at(&second, "<w:p w:rsidR=\"00A020\">");
    let inserted = "<w:p><w:r><w:t>Inserted</w:t></w:r></w:p>";
    let third = format!("{}{inserted}{}", &second[..next], &second[next..]);
    let (third_proof, how) = step(&second_proof, third.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, splice(next..next, next..next + inserted.len()));
    let fourth = third.replacen("Inserted", "Inserted again", 1);
    let (fourth_proof, how) = step(&third_proof, fourth.as_bytes(), Limits::default()).unwrap();
    assert_eq!(
        how,
        splice(next..next + inserted.len(), next..next + inserted.len() + 6)
    );

    // A removal is proved by the run that also holds the next paragraph.
    let removed = fourth.replacen(&paragraph(10), "", 1);
    let (removed_proof, how) = step(&fourth_proof, removed.as_bytes(), Limits::default()).unwrap();
    assert!(matches!(how, ReplacementProof::Splice { .. }), "{how:?}");
    // An edit of the document element's start tag lies outside every run,
    // so it is audited completely; that audit records its own boundaries
    // and the chain goes on.
    let fifth = removed.replacen(
        "xmlns:r=\"urn:r\"",
        "xmlns:r=\"urn:r\" xmlns:s=\"urn:s\"",
        1,
    );
    let (fifth_proof, how) = step(&removed_proof, fifth.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, ReplacementProof::Complete);
    assert!(boundaries(&fifth_proof).is_some());
    let sixth = fifth.replacen("Paragraph 30 holds", "Paragraph 30 held", 1);
    let (_, how) = step(&fifth_proof, sixth.as_bytes(), Limits::default()).unwrap();
    assert!(matches!(how, ReplacementProof::Splice { .. }), "{how:?}");

    // Identical bytes share the boundaries.
    let (same, how) = step(&fifth_proof, fifth.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, ReplacementProof::Identical);
    assert_eq!(boundaries(&same), boundaries(&fifth_proof));
}

#[test]
fn a_chain_refuses_every_defect_with_the_complete_audit_error() {
    let original = document(20);
    let base = indexed(&original, IndexDepth::default());
    let (edited, _how) = step(
        &base,
        original
            .replacen("Paragraph 3 holds", "Paragraph 3 held", 1)
            .as_bytes(),
        Limits::default(),
    )
    .unwrap();
    let text = std::str::from_utf8(edited.bytes()).unwrap().to_owned();
    let next = at(&text, "<w:p w:rsidR=\"00A015\">");
    for defect in [
        "<w:p>",
        "</w:body>",
        "<q:p/>",
        "<w:p a=\"1\" a=\"2\"/>",
        "<w:p>&bogus;</w:p>",
        "<w:p>]]></w:p>",
        "<w:p>\u{1}</w:p>",
        "<?xml version=\"1.0\"?>",
        "<w:p xmlns:xmlns=\"urn:x\"/>",
        "<w:p xmlns:q=\"\"/>",
        "<!DOCTYPE d>",
    ] {
        let broken = format!("{}{defect}{}", &text[..next], &text[next..]);
        assert!(
            step(&edited, broken.as_bytes(), Limits::default()).is_none(),
            "{defect}"
        );
    }
}

#[test]
fn bindings_declared_inside_a_window_scope_the_next_edit() {
    // A new child of the document element declares its own prefix; its
    // content is a derived parent with a scope of its own.
    let original = document(20);
    let base = indexed(&original, IndexDepth::default());
    let body = at(&original, "<w:body>");
    let section = "<w:ext xmlns:q=\"urn:q\"><q:a/><q:b/></w:ext>";
    let first = format!("{}{section}{}", &original[..body], &original[body..]);
    let (first_proof, how) = step(&base, first.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, splice(body..body, body..body + section.len()));
    // Inside it, the prefix is bound.
    let inner = body + "<w:ext xmlns:q=\"urn:q\"><q:a/>".len();
    let second = format!("{}<q:c/>{}", &first[..inner], &first[inner..]);
    let (_, how) = step(&first_proof, second.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, splice(inner..inner, inner..inner + "<q:c/>".len()));
    // Outside it, it is not.
    let next = at(&first, "<w:p w:rsidR=\"00A010\">");
    let outside = format!("{}<w:p><q:c/></w:p>{}", &first[..next], &first[next..]);
    assert!(step(&first_proof, outside.as_bytes(), Limits::default()).is_none());
    // Redeclaring the prefix inside, then using it, passes.
    let rebound = format!(
        "{}<q:d xmlns:q=\"urn:other\"><q:e/></q:d>{}",
        &first[..inner],
        &first[inner..]
    );
    assert!(step(&first_proof, rebound.as_bytes(), Limits::default()).is_some());
}

#[test]
fn every_budget_is_re_totalled_along_a_chain() {
    let original = document(30);
    let exact = verify_source(original.as_bytes(), Limits::default()).unwrap();
    // Room for exactly three more inserted paragraphs' events.
    let inserted = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
    let per_insert = 7;
    let limits = Limits::builder()
        .events(exact.events() + 3 * per_insert)
        .unwrap()
        .build();
    let mut current =
        VerifiedSource::verify_indexed(shared(original.as_bytes()), limits, IndexDepth::default())
            .unwrap();
    let mut text = original;
    for round in 0..4 {
        let next = at(&text, "<w:p w:rsidR=\"00A005\">");
        let replacement = format!("{}{inserted}{}", &text[..next], &text[next..]);
        match step(&current, replacement.as_bytes(), limits) {
            Some((proof, how)) => {
                assert!(round < 3, "round {round} passed");
                assert!(matches!(how, ReplacementProof::Splice { .. }));
                current = proof;
                text = replacement;
            },
            None => assert_eq!(round, 3),
        }
    }
}

#[test]
fn a_derived_index_over_its_bound_is_dropped() {
    let original = document(20);
    let base = indexed(&original, IndexDepth::default());
    // Many tiny body children: more than one boundary per 32 bytes.
    let next = at(&original, "<w:p w:rsidR=\"00A010\">");
    let flood = "<w:x/>".repeat(1_000);
    let flooded = format!("{}{flood}{}", &original[..next], &original[next..]);
    // The window is over half the replacement, so it is audited completely,
    // and that audit's own index is over the bound too.
    let (proof, how) = step(&base, flooded.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, ReplacementProof::Complete);
    assert_eq!(boundaries(&proof), None);
    // A splice whose derived index would exceed the bound keeps none, and
    // the next edit is audited completely.
    let original = document(400);
    let base = indexed(&original, IndexDepth::default());
    let allowed = original.len() / 32 + 64;
    let recorded = boundaries(&base).unwrap();
    // Each added child adds one boundary and six bytes, which raise the
    // allowance by 6/32 of a boundary.
    let extra = (allowed - recorded) * 32 / 26 + 50;
    let next = at(&original, "<w:p w:rsidR=\"00A200\">");
    let flood = "<w:x/>".repeat(extra);
    let flooded = format!("{}{flood}{}", &original[..next], &original[next..]);
    let (proof, how) = step(&base, flooded.as_bytes(), Limits::default()).unwrap();
    assert!(matches!(how, ReplacementProof::Splice { .. }), "{how:?}");
    assert_eq!(boundaries(&proof), None);
    let again = flooded.replacen("Paragraph 3 holds", "Paragraph 3 held", 1);
    let (_, how) = step(&proof, again.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, ReplacementProof::Complete);
}

#[test]
fn a_byte_order_mark_keeps_physical_offsets_along_a_chain() {
    let original = format!("\u{feff}{}", document(20));
    let base = indexed(&original, IndexDepth::default());
    let first = original.replacen("Paragraph 4 holds", "Paragraph 4 held", 1);
    let (proof, _how) = step(&base, first.as_bytes(), Limits::default()).unwrap();
    let next = at(&first, "<w:p w:rsidR=\"00A017\">");
    let second = format!("{}<w:tbl/>{}", &first[..next], &first[next..]);
    let (_, how) = step(&proof, second.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, splice(next..next, next..next + "<w:tbl/>".len()));
}

#[test]
fn depth_three_proves_slide_shape_edits() {
    let original = slide(200);
    let deep = indexed(&original, IndexDepth::GRANDCHILDREN);
    let shallow = indexed(&original, IndexDepth::CHILDREN);
    assert_eq!(deep.index_depth(), Some(IndexDepth::GRANDCHILDREN));
    assert!(boundaries(&deep).unwrap() > boundaries(&shallow).unwrap());
    // A shape text set: the run is the one shape.
    let set = original.replacen(">Shape 120<", ">Shape 120 edited<", 1);
    let start = at(&original, &shape(120, "Shape 120"));
    let len = shape(120, "Shape 120").len();
    let (set_proof, how) = step(&deep, set.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, splice(start..start + len, start..start + len + 7));
    // At depth two the run is the whole shape tree, which is audited
    // completely.
    let (_, how) = step(&shallow, set.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, ReplacementProof::Complete);
    // A shape add after the last shape, chained on the set.
    let end = at(&set, "</p:spTree>");
    let added = shape(200, "New");
    let add = format!("{}{added}{}", &set[..end], &set[end..]);
    let (add_proof, how) = step(&set_proof, add.as_bytes(), Limits::default()).unwrap();
    assert_eq!(how, splice(end..end, end..end + added.len()));
    // And a set of a shape far from both.
    let far = add.replacen(">Shape 3<", ">Shape three<", 1);
    let (_, how) = step(&add_proof, far.as_bytes(), Limits::default()).unwrap();
    assert!(matches!(how, ReplacementProof::Splice { .. }), "{how:?}");
    // A shape that uses an unbound prefix is refused as the complete audit
    // refuses it.
    let bad = format!("{}<p:sp><z:x/></p:sp>{}", &set[..end], &set[end..]);
    assert!(step(&set_proof, bad.as_bytes(), Limits::default()).is_none());
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

const PIECES: [&str; 24] = [
    "<w:p/>",
    "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
    "<p:sp><p:txBody><a:p/></p:txBody></p:sp>",
    "<w:p>",
    "</w:p>",
    "</p:sp>",
    "</w:body><w:body>",
    "<q:p/>",
    "<w:p xmlns:q=\"urn:q\"><q:x/></w:p>",
    "<w:x xmlns:q=\"urn:q\"><q:y/></w:x>",
    "&amp;",
    "&bogus;",
    "<![CDATA[c]]>",
    "]]>",
    "<!-- c -->",
    "\"",
    "<",
    ">",
    "=",
    " ",
    "x",
    "\u{1}",
    "é",
    "<?pi x?>",
];

/// Random chains of edits: every accepted successor becomes the next base,
/// so the derived indexes are exercised edit after edit.
fn random_chains(original: &str, depth: IndexDepth, seed: u64, rounds: usize) -> usize {
    let mut random = Lcg(seed);
    let mut base = indexed(original, depth);
    let mut text = original.as_bytes().to_vec();
    let mut splices = 0;
    for _ in 0..rounds {
        let mut next = text.clone();
        for _ in 0..1 + random.next() % 2 {
            let offset = random.next() % (next.len() + 1);
            match random.next() % 4 {
                0 | 3 => {
                    let piece = PIECES[random.next() % PIECES.len()].as_bytes();
                    next.splice(offset..offset, piece.iter().copied());
                },
                1 if offset < next.len() => {
                    let end = (offset + 1 + random.next() % 8).min(next.len());
                    next.drain(offset..end);
                },
                _ if offset < next.len() => next[offset] = b"<>/\"=a &;"[random.next() % 9],
                _ => {},
            }
        }
        if let Some((proof, how)) = step(&base, &next, Limits::default()) {
            if matches!(how, ReplacementProof::Splice { .. }) {
                splices += 1;
            }
            base = proof;
            text = next;
        }
    }
    splices
}

#[test]
fn random_chains_have_the_complete_audit_verdict() {
    let splices = random_chains(&document(40), IndexDepth::default(), 0x0317, 2_000);
    assert!(splices > 100, "{splices}");
}

#[test]
fn random_chains_on_a_slide_at_depth_three_have_the_complete_audit_verdict() {
    let splices = random_chains(&slide(30), IndexDepth::GRANDCHILDREN, 0x0317_0003, 2_000);
    assert!(splices > 100, "{splices}");
}
