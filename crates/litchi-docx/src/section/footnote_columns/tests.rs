//! Focused codec, package, snapshot, and atomic-edit coverage.

use super::*;
use crate::section::Section;
use litchi_opc::PackURI;
use litchi_opc::part::BlobPart;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W12: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn section_xml(body: &str) -> String {
    format!(
        r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12">{body}</w:sectPr>"#
    )
}

fn layout(columns: i32) -> Layout {
    Layout::new(columns).expect("valid footnote column count")
}

#[test]
fn reads_zero_and_preserves_absence() {
    let absent = Snapshot::from_xml(section_xml(r#"<w:pgSz w:w="12240"/>"#).into_bytes())
        .expect("valid section");
    assert_eq!(absent.layout(), None);

    let zero = Snapshot::from_xml(
        section_xml(r#"<w12:footnoteColumns w:val="0"/><w:pgSz w:w="12240"/>"#).into_bytes(),
    )
    .expect("valid extension");
    assert_eq!(zero.layout(), Some(layout(0)));
    assert!(zero.layout().expect("zero layout").follows_page_layout());
}

#[test]
fn transaction_is_lossless_and_inverse_is_atomic() {
    let source = Snapshot::from_xml(
        section_xml(
            r#"
                <x:before xmlns:x="urn:opaque" x:keep="1"/>
                <w12:footnoteColumns w:val="2"/>
                <x:after xmlns:x="urn:opaque"/>
            "#,
        )
        .into_bytes(),
    )
    .expect("valid section");
    let mut edit = source.edit();
    edit.set_layout(Some(layout(4))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    assert_eq!(source.layout(), Some(layout(2)));
    assert_eq!(commit.snapshot().layout(), Some(layout(4)));
    let changed = std::str::from_utf8(commit.snapshot().xml_bytes()).expect("UTF-8 XML");
    assert!(changed.contains("x:before"));
    assert!(changed.contains("x:after"));
    assert_eq!(commit.patch().before(), Some(layout(2)));
    assert_eq!(commit.patch().after(), Some(layout(4)));

    let restored = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("inverse");
    assert_eq!(restored.layout(), Some(layout(2)));

    let stale = Snapshot::from_xml(section_xml("").into_bytes()).expect("valid section");
    assert!(commit.patch().apply(&stale).is_err());
}

#[test]
fn section_facade_publishes_one_atomic_edit() {
    let mut section = Section::from_xml_bytes(
        section_xml(r#"<x:keep xmlns:x="urn:opaque"/><w:pgSz w:w="12240"/>"#).into_bytes(),
    )
    .expect("valid section");
    assert_eq!(section.footnote_columns().expect("read"), None);
    section.set_footnote_columns(Some(layout(3))).expect("set");
    let xml = section.to_xml().expect("write");
    assert!(xml.contains(r#"xmlns:w12="http://schemas.microsoft.com/office/word/2012/wordml""#));
    assert!(xml.contains(r#"mc:Ignorable="w12""#));
    assert!(xml.contains(r#"w12:footnoteColumns w:val="3""#));
    assert!(xml.contains("x:keep"));
    assert_eq!(section.footnote_columns().expect("read"), Some(layout(3)));

    section.clear_footnote_columns().expect("clear");
    assert_eq!(section.footnote_columns().expect("read"), None);
    assert!(!section.to_xml().expect("write").contains("footnoteColumns"));
}

#[test]
fn package_discovery_reads_all_sections() {
    let part = BlobPart::new(
        PackURI::new("/word/document.xml").expect("URI"),
        "application/xml".into(),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w:body><w:sectPr><w12:footnoteColumns w:val="2"/></w:sectPr><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:body></w:document>"#
        )
        .into_bytes(),
    );
    let sections = parse_part(&part).expect("document sections");
    assert_eq!(sections.len(), 2);
    assert_eq!(sections[0].layout(), Some(layout(2)));
    assert_eq!(sections[1].layout(), None);
}

#[test]
fn malformed_extension_and_missing_ignorable_are_rejected() {
    assert!(
        Snapshot::from_xml(section_xml(r#"<w12:footnoteColumns w:val="-1"/>"#).into_bytes())
            .is_err()
    );
    assert!(Snapshot::from_xml(
        format!(
            r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}"><w12:footnoteColumns w:val="2"/></w:sectPr>"#
        )
        .into_bytes()
    )
    .is_err());
    assert!(
        Snapshot::from_xml(
            section_xml(
                r#"<w12:footnoteColumns w:val="2"><x:child xmlns:x="urn:x"/></w12:footnoteColumns>"#
            )
            .into_bytes()
        )
        .is_err()
    );
}

#[test]
fn mutation_merges_existing_ignorable_tokens_and_custom_prefixes() {
    let xml = format!(
        r#"<q:sectPr xmlns:q="{W}" xmlns:x12="{W12}" xmlns:m="{MC}" m:Ignorable="vml x12"><x12:footnoteColumns q:val="2"/><x:opaque xmlns:x="urn:opaque"/></q:sectPr>"#
    );
    let source = Snapshot::from_xml(xml.into_bytes()).expect("valid custom-prefix section");
    let mut edit = source.edit();
    edit.set_layout(Some(layout(6))).expect("valid edit");
    let changed = edit.commit().expect("commit");
    let output = std::str::from_utf8(changed.snapshot().xml_bytes()).expect("UTF-8 XML");
    assert!(output.contains(r#"m:Ignorable="vml x12""#));
    assert!(output.contains(r#"<x12:footnoteColumns q:val="6"/>"#));
    assert!(output.contains("x:opaque"));
}

#[test]
fn rejects_overdeep_section_fragments() {
    let nested = format!(
        "<x:opaque xmlns:x=\"urn:x\">{} </x:opaque>",
        "<x:n>".repeat(validation::MAX_XML_DEPTH) + &"</x:n>".repeat(validation::MAX_XML_DEPTH)
    );
    assert!(Snapshot::from_xml(section_xml(&nested).into_bytes()).is_err());
}
