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
fn package_discovery_retains_inherited_namespace_context() {
    let source = format!(
        r#"<q:document xmlns:q="{W}" xmlns:x12="{W12}" xmlns:m="{MC}" m:Ignorable="x12"><q:body><q:sectPr><x12:footnoteColumns q:val="02"/></q:sectPr></q:body></q:document>"#
    );
    let part = BlobPart::new(
        PackURI::new("/word/document.xml").expect("URI"),
        "application/xml".into(),
        source.into_bytes(),
    );
    let sections = parse_part(&part).expect("document sections");
    assert_eq!(sections.len(), 1);
    assert_eq!(sections[0].layout(), Some(layout(2)));
    assert_eq!(
        sections[0].xml_bytes(),
        br#"<q:sectPr><x12:footnoteColumns q:val="02"/></q:sectPr>"#
    );

    let mut edit = sections[0].edit();
    edit.set_layout(Some(layout(6))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    assert_eq!(
        commit.snapshot().xml_bytes(),
        br#"<q:sectPr><x12:footnoteColumns q:val="6"/></q:sectPr>"#
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("inverse");
    assert_eq!(
        restored.xml_bytes(),
        br#"<q:sectPr><x12:footnoteColumns q:val="02"/></q:sectPr>"#
    );
}

#[test]
fn package_discovery_rejects_unignorable_inherited_extension() {
    let part = BlobPart::new(
        PackURI::new("/word/document.xml").expect("URI"),
        "application/xml".into(),
        format!(
            r#"<q:document xmlns:q="{W}" xmlns:x12="{W12}"><q:body><q:sectPr><x12:footnoteColumns q:val="2"/></q:sectPr></q:body></q:document>"#
        )
        .into_bytes(),
    );
    assert!(parse_part(&part).is_err());
}

#[test]
fn value_edits_preserve_extension_attributes_and_lexical_source() {
    let source = Snapshot::from_xml(
        format!(
            r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w12:footnoteColumns foo:extra="keep" w:val='02' xmlns:foo="urn:opaque"  /></w:sectPr>"#
        )
        .into_bytes(),
    )
    .expect("valid section");
    let mut edit = source.edit();
    edit.set_layout(Some(layout(3))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    assert_eq!(
        commit.snapshot().xml_bytes(),
        format!(
            r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w12:footnoteColumns foo:extra="keep" w:val='3' xmlns:foo="urn:opaque"  /></w:sectPr>"#
        )
        .as_bytes()
    );

    let no_op =
        Snapshot::from_xml(section_xml(r#"<w12:footnoteColumns w:val="02"/>"#).into_bytes())
            .expect("valid section");
    let mut no_op_edit = no_op.edit();
    no_op_edit
        .set_layout(Some(layout(2)))
        .expect("valid no-op edit");
    let no_op_commit = no_op_edit.commit().expect("no-op commit");
    assert_eq!(no_op_commit.snapshot().xml_bytes(), no_op.xml_bytes());
    assert_eq!(
        no_op_commit
            .patch()
            .apply(&no_op)
            .expect("no-op patch")
            .xml_bytes(),
        no_op.xml_bytes()
    );
}

#[test]
fn patches_require_exact_source_bytes() {
    let source = Snapshot::from_xml(
        section_xml(r#"<x:opaque xmlns:x="urn:first"/><w12:footnoteColumns w:val="2"/>"#)
            .into_bytes(),
    )
    .expect("valid source");
    let same_value_different_source = Snapshot::from_xml(
        section_xml(r#"<x:opaque xmlns:x="urn:second"/><w12:footnoteColumns w:val="2"/>"#)
            .into_bytes(),
    )
    .expect("valid alternate source");
    let mut edit = source.edit();
    edit.set_layout(Some(layout(4))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    assert!(commit.patch().apply(&same_value_different_source).is_err());
}

#[test]
fn rejects_values_outside_schema_int_bounds() {
    assert!(
        Snapshot::from_xml(
            section_xml(r#"<w12:footnoteColumns w:val="2147483648"/>"#).into_bytes()
        )
        .is_err()
    );
    assert!(
        Snapshot::from_xml(
            section_xml(r#"<w12:footnoteColumns w:val="-2147483649"/>"#).into_bytes()
        )
        .is_err()
    );
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
fn inserts_before_section_change_without_rewriting_opaque_trivia() {
    let source = section_xml(
        r#"
            <?keep processing?>
            <!-- preserve before the extension -->
            <x:opaque xmlns:x="urn:opaque" x:before="1"/>
            <![CDATA[ preserve cdata trivia ]]>
            <w:sectPrChange w:id="7">
                <x:opaque xmlns:x="urn:opaque" x:inside="1"/>
            </w:sectPrChange>
            <!-- preserve after the barrier -->
        "#,
    );
    let snapshot = Snapshot::from_xml(source.clone().into_bytes()).expect("valid section");
    let mut edit = snapshot.edit();
    edit.set_layout(Some(layout(4))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    let output = String::from_utf8(commit.snapshot().xml_bytes().to_vec()).expect("UTF-8 XML");
    let inserted = r#"<w12:footnoteColumns w:val="4"/>"#;
    let inserted_at = output.find(inserted).expect("inserted footnote layout");
    let barrier_at = output
        .find("<w:sectPrChange")
        .expect("section change barrier");
    assert!(inserted_at < barrier_at);
    assert!(output.contains("<?keep processing?>"));
    assert!(output.contains("<!-- preserve before the extension -->"));
    assert!(output.contains("<![CDATA[ preserve cdata trivia ]]>"));
    assert!(output.contains(r#"<x:opaque xmlns:x="urn:opaque" x:inside="1"/>"#));
    assert!(output.contains("<!-- preserve after the barrier -->"));
    assert_eq!(output.replacen(inserted, "", 1), source);
}

#[test]
fn selects_matching_aliases_when_preferred_namespace_prefixes_collide() {
    let source = format!(
        r#"<w:sectPr xmlns:w="{W}" xmlns:w12="urn:occupied-word12" xmlns:x12="{W12}" xmlns:mc="urn:occupied-mc" xmlns:c="{MC}" c:Ignorable="opaque"><w:sectPrChange w:id="7"/></w:sectPr>"#
    );
    let snapshot = Snapshot::from_xml(source.into_bytes()).expect("valid aliased section");
    let mut edit = snapshot.edit();
    edit.set_layout(Some(layout(5))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    let output = std::str::from_utf8(commit.snapshot().xml_bytes()).expect("UTF-8 XML");

    assert!(output.contains(r#"<x12:footnoteColumns w:val="5"/>"#));
    assert!(!output.contains(r#"<w12:footnoteColumns"#));
    assert!(output.contains(r#"xmlns:w12="urn:occupied-word12""#));
    assert!(output.contains(r#"xmlns:mc="urn:occupied-mc""#));
    assert!(output.contains(r#"c:Ignorable="opaque x12""#));
    assert!(output.find("<x12:footnoteColumns").unwrap() < output.find("<w:sectPrChange").unwrap());
}

#[test]
fn rejects_duplicate_expanded_mc_ignorable_aliases() {
    let fragment = format!(
        r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" xmlns:c="{MC}" mc:Ignorable="w12" c:Ignorable="w12"><w12:footnoteColumns w:val="2"/></w:sectPr>"#
    );
    assert!(Snapshot::from_xml(fragment.clone().into_bytes()).is_err());

    let part = BlobPart::new(
        PackURI::new("/word/document.xml").expect("URI"),
        "application/xml".into(),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" xmlns:c="{MC}" mc:Ignorable="w12" c:Ignorable="w12"><w:body>{fragment}</w:body></w:document>"#
        )
        .into_bytes(),
    );
    assert!(parse_part(&part).is_err());
}

#[test]
fn unresolved_word_prefix_still_places_the_extension_before_section_change() {
    let source = r#"<w:sectPr><w:sectPrChange w:id="7"/></w:sectPr>"#;
    let snapshot = Snapshot::from_xml(source.as_bytes().to_vec()).expect("detached section");
    let mut edit = snapshot.edit();
    edit.set_layout(Some(layout(3))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    assert_eq!(
        commit.snapshot().xml_bytes(),
        br#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w12="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w12"><w12:footnoteColumns w:val="3"/><w:sectPrChange w:id="7"/></w:sectPr>"#
    );
}

#[test]
fn insertion_refuses_unknown_or_post_change_word_children() {
    let unknown = Snapshot::from_xml(
        format!(r#"<w:sectPr xmlns:w="{W}"><w:future/></w:sectPr>"#).into_bytes(),
    )
    .expect("opaque Word child remains readable");
    let mut edit = unknown.edit();
    edit.set_layout(Some(layout(2))).expect("valid value");
    assert!(edit.commit().is_err());

    let post_change = Snapshot::from_xml(
        format!(r#"<w:sectPr xmlns:w="{W}"><w:sectPrChange/><w:pgSz/></w:sectPr>"#).into_bytes(),
    )
    .expect("source remains losslessly readable");
    let mut edit = post_change.edit();
    edit.set_layout(Some(layout(2))).expect("valid value");
    assert!(edit.commit().is_err());
}

#[test]
fn insertion_does_not_capture_opaque_descendant_prefixes() {
    let source = format!(
        r#"<w:sectPr xmlns:w="{W}"><x:opaque xmlns:x="urn:opaque"><w12:nested/></x:opaque></w:sectPr>"#
    );
    let snapshot = Snapshot::from_xml(source.into_bytes()).expect("opaque descendant");
    let mut edit = snapshot.edit();
    edit.set_layout(Some(layout(6))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    let output = std::str::from_utf8(commit.snapshot().xml_bytes()).expect("UTF-8 XML");
    assert!(
        output.contains(r#"xmlns:w12_1="http://schemas.microsoft.com/office/word/2012/wordml""#)
    );
    assert!(output.contains(r#"<w12_1:footnoteColumns w:val="6"/>"#));
    assert!(output.contains("<w12:nested/>"));
}

#[test]
fn inherited_ignorable_tokens_are_not_duplicated_or_rebound_locally() {
    let part = BlobPart::new(
        PackURI::new("/word/document.xml").expect("URI"),
        "application/xml".into(),
        format!(
            r#"<q:document xmlns:q="{W}" xmlns:x12="{W12}" xmlns:m="{MC}" m:Ignorable="x12"><q:body><q:sectPr xmlns:x12="urn:locally-shadowed"/></q:body></q:document>"#
        )
        .into_bytes(),
    );
    let sections = parse_part(&part).expect("document sections");
    let mut edit = sections[0].edit();
    edit.set_layout(Some(layout(7))).expect("valid edit");
    let commit = edit.commit().expect("commit");
    let output = std::str::from_utf8(commit.snapshot().xml_bytes()).expect("UTF-8 XML");
    assert!(output.contains(r#"m:Ignorable="w12""#));
    assert!(!output.contains("x12 w12"));
    assert!(!output.contains("w12 w12"));
    assert!(output.contains(r#"<w12:footnoteColumns q:val="7"/>"#));
}

#[test]
fn rejects_overdeep_section_fragments() {
    let nested = format!(
        "<x:opaque xmlns:x=\"urn:x\">{} </x:opaque>",
        "<x:n>".repeat(validation::MAX_XML_DEPTH) + &"</x:n>".repeat(validation::MAX_XML_DEPTH)
    );
    assert!(Snapshot::from_xml(section_xml(&nested).into_bytes()).is_err());
}
