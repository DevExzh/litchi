//! Regression coverage for the layered reference-mark owner.

use super::MAX_DEPTH;
use super::codec::parse_reference_marks;
use super::model::ReferenceMark;
use super::writing::{
    ReferenceMarkFragments, insert_reference_mark_xml, remove_reference_mark_xml,
    replace_reference_mark_xml,
};

const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

#[test]
fn parses_point_and_range_reference_marks_with_exact_positions_and_text() {
    let xml = format!(
        r#"<x:text xmlns:x="{TEXT}"><x:p>ab<x:reference-mark x:name="point"/>c<x:reference-mark-start x:name="range"/>D&amp;<x:span>E</x:span><x:s x:c="2"/></x:p><x:p>F<![CDATA[!]]><x:reference-mark-end x:name="range"/>z</x:p></x:text>"#
    );
    let marks = parse_reference_marks(&xml).unwrap();
    assert_eq!(marks.len(), 2);
    assert_eq!(marks[0].name(), "point");
    assert!(!marks[0].is_range());
    assert_eq!(marks[0].start(), Some((0, 2)));
    assert_eq!(marks[0].end(), Some((0, 2)));
    assert!(marks[0].text().is_empty());
    assert_eq!(marks[1].name(), "range");
    assert!(marks[1].is_range());
    assert_eq!(marks[1].start(), Some((0, 3)));
    assert_eq!(marks[1].end(), Some((1, 2)));
    assert_eq!(marks[1].text(), "D&E  \nF!");
}

#[test]
fn reference_marks_reject_missing_duplicate_unmatched_and_nonempty_markers() {
    let missing = format!(r#"<x:reference-mark xmlns:x="{TEXT}"/>"#);
    assert!(parse_reference_marks(&missing).is_err());
    let duplicate = format!(
        r#"<x:p xmlns:x="{TEXT}"><x:reference-mark-start x:name="a"/><x:reference-mark-start x:name="a"/></x:p>"#
    );
    assert!(parse_reference_marks(&duplicate).is_err());
    let unmatched = format!(r#"<x:reference-mark-end xmlns:x="{TEXT}" x:name="a"/>"#);
    assert!(parse_reference_marks(&unmatched).is_err());
    let unclosed = format!(r#"<x:reference-mark-start xmlns:x="{TEXT}" x:name="a"/>"#);
    assert!(parse_reference_marks(&unclosed).is_err());
    let nonempty =
        format!(r#"<x:reference-mark xmlns:x="{TEXT}" x:name="a">bad</x:reference-mark>"#);
    assert!(parse_reference_marks(&nonempty).is_err());
    let aliases =
        format!(r#"<x:reference-mark xmlns:x="{TEXT}" xmlns:y="{TEXT}" x:name="a" y:name="b"/>"#);
    assert!(parse_reference_marks(&aliases).is_err());
    assert!(parse_reference_marks("<x:reference-mark>").is_err());
}

#[test]
fn reference_marks_enforce_nesting_bound() {
    let mut xml = format!(r#"<x:p xmlns:x="{TEXT}">"#);
    for _ in 0..MAX_DEPTH {
        xml.push_str("<x:span>");
    }
    for _ in 0..MAX_DEPTH {
        xml.push_str("</x:span>");
    }
    xml.push_str("</x:p>");
    assert!(parse_reference_marks(&xml).is_err());
}

fn wrapped(body: &str) -> String {
    format!(
        r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{TEXT}" xmlns:text="{TEXT}">{body}</o:text>"#
    )
}

#[test]
fn canonical_point_and_range_fragments_round_trip() {
    let point = ReferenceMark::point("point<&\"");
    assert_eq!(
        point.to_xml_fragments().unwrap(),
        ReferenceMarkFragments::Point(
            r#"<text:reference-mark text:name="point&lt;&amp;&quot;"/>"#.to_string()
        )
    );
    let range = ReferenceMark::range("range");
    assert_eq!(
        range.to_xml_fragments().unwrap(),
        ReferenceMarkFragments::Range {
            start: r#"<text:reference-mark-start text:name="range"/>"#.to_string(),
            end: r#"<text:reference-mark-end text:name="range"/>"#.to_string(),
        }
    );
    let inserted = insert_reference_mark_xml(&wrapped("<t:p>payload</t:p>"), 0, &range).unwrap();
    let parsed = parse_reference_marks(&inserted).unwrap();
    assert_eq!(parsed.len(), 1);
    assert_eq!(parsed[0].name(), "range");
    assert_eq!(parsed[0].text(), "payload");
}

#[test]
fn lossless_insert_replace_remove_preserves_unrelated_and_enclosed_xml() {
    let source = wrapped(
        r#"<t:p t:style-name="P">A&amp;<t:span t:style-name="S">B</t:span><!--keep--></t:p><t:p/>"#,
    );
    let with_range = insert_reference_mark_xml(&source, 0, &ReferenceMark::range("r1")).unwrap();
    assert!(with_range.contains(r#"<text:reference-mark-start text:name="r1"/>A&amp;<t:span t:style-name="S">B</t:span><!--keep--><text:reference-mark-end text:name="r1"/>"#));
    let replaced = replace_reference_mark_xml(&with_range, 0, &ReferenceMark::point("p2")).unwrap();
    assert!(replaced.contains(r#"<text:reference-mark text:name="p2"/>A&amp;<t:span t:style-name="S">B</t:span><!--keep-->"#));
    let removed = remove_reference_mark_xml(&replaced, 0).unwrap();
    assert_eq!(removed, source);
    let empty = insert_reference_mark_xml(&source, 1, &ReferenceMark::point("empty")).unwrap();
    assert!(empty.contains(r#"<t:p><text:reference-mark text:name="empty"/></t:p>"#));
}

#[test]
fn hostile_namespaces_identity_content_and_resources_are_rejected() {
    for body in [
        r#"<t:p><t:reference-mark t:name="x" u:extra="1"/></t:p>"#,
        r#"<t:p><t:reference-mark u:name="x"/></t:p>"#,
        r#"<t:p><t:reference-mark t:name="x"/><t:reference-mark t:name="x"/></t:p>"#,
        r#"<t:p><t:reference-mark t:name="x"/><t:reference-mark-start t:name="x"/><t:reference-mark-end t:name="x"/></t:p>"#,
        r#"<t:p><t:reference-mark t:name="x">bad</t:reference-mark></t:p>"#,
        r"<!DOCTYPE x><t:p/>",
    ] {
        let xml = format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{TEXT}" xmlns:u="urn:hostile">{body}</o:text>"#
        );
        assert!(
            insert_reference_mark_xml(&xml, 0, &ReferenceMark::point("new")).is_err(),
            "accepted {body}"
        );
    }
    assert!(
        ReferenceMark::point("x".repeat(65_537))
            .to_xml_fragments()
            .is_err()
    );
    assert!(
        ReferenceMark::point("bad\0name")
            .to_xml_fragments()
            .is_err()
    );
}

#[test]
fn producer_shaped_point_field_and_overlapping_ranges_round_trip() {
    // LibreOffice/Zotero emits long metadata-bearing names and adjacent range markers.
    let name = r"ZOTERO_ITEM CSL_CITATION {&quot;citationID&quot;:&quot;abc&quot;} RNDxyz";
    let xml = wrapped(&format!(
        r#"<t:p><t:reference-mark-start t:name="{name}"/>(<t:span t:style-name="T1">Author</t:span>, 2026)<t:reference-mark-start t:name="second"/> tail<t:reference-mark-end t:name="{name}"/><t:reference-mark-end t:name="second"/></t:p><t:p><t:reference-mark t:name="anchor"/><t:reference-ref t:reference-format="page" t:ref-name="anchor">1</t:reference-ref></t:p>"#
    ));
    let marks = parse_reference_marks(&xml).unwrap();
    assert_eq!(marks.len(), 3);
    assert_eq!(marks[0].text(), "(Author, 2026) tail");
    assert_eq!(marks[1].text(), " tail");
    let replaced =
        replace_reference_mark_xml(&xml, 2, &ReferenceMark::point("odfpy-anchor")).unwrap();
    assert!(replaced.contains(
        r#"<t:reference-ref t:reference-format="page" t:ref-name="anchor">1</t:reference-ref>"#
    ));
    assert_eq!(parse_reference_marks(&replaced).unwrap().len(), 3);
}
