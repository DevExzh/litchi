use super::*;
use crate::Error;
use crate::leniency::{Defect, SectionDefect};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W12: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn invalid_format_text(result: crate::Result<Section>) -> String {
    result
        .expect_err("the section must be rejected")
        .to_string()
}

#[test]
fn inventory_reuse_keeps_public_semantic_laziness_and_error_order() {
    let semantic_only =
        format!(r#"<w:sectPr xmlns:w="{W}"><w:type w:val="not-a-section-break"/></w:sectPr>"#);
    let mut section = Section::from_xml_bytes(semantic_only.into_bytes())
        .expect("semantic decoding remains deferred by the public constructor");
    let semantic_error = section
        .validate()
        .expect_err("the first semantic access must expose the invalid section type")
        .to_string();
    assert!(semantic_error.contains("section type"), "{semantic_error}");

    let codec_only = format!(
        r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w12:footnoteColumns w:val="0"/></w:wrong>"#
    );
    let codec_and_footnote = format!(
        r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w12:footnoteColumns w:val="-1"/></w:wrong>"#
    );
    assert_eq!(
        invalid_format_text(Section::from_xml_bytes(codec_only.into_bytes())),
        invalid_format_text(Section::from_xml_bytes(codec_and_footnote.into_bytes())),
        "raw codec validation must win before focused footnote validation"
    );

    let footnote_before_semantic = format!(
        r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w12:footnoteColumns w:val="-1"/><w:type w:val="not-a-section-break"/></w:sectPr>"#
    );
    let error = invalid_format_text(Section::from_xml_bytes(
        footnote_before_semantic.into_bytes(),
    ));
    assert!(error.contains("footnote"), "{error}");
    assert!(
        !error.contains("section type"),
        "semantic decoding must remain deferred after focused validation: {error}"
    );
}

#[test]
fn inventory_reuse_matches_section_footnote_before_semantic_error() {
    let fragment = format!(
        r#"<w:sectPr xmlns:w="{W}" xmlns:w12="{W12}"><w12:footnoteColumns w:val="-1"/><w:type w:val="not-a-section-break"/></w:sectPr>"#
    );
    // PreserveElements keeps the unknown extension through the document
    // preprocessor so it reaches the private inventory codec route.
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12" mc:PreserveElements="w12:footnoteColumns"><w:body>{fragment}</w:body></w:document>"#
    );
    let inventory_error = Inventory::parse(document.as_bytes(), &Default::default())
        .expect_err("inventory must reject the retained malformed focused extension");

    let footnote_only = fragment.replace(r#"<w:type w:val="not-a-section-break"/>"#, "");
    let footnote_document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12" mc:PreserveElements="w12:footnoteColumns"><w:body>{footnote_only}</w:body></w:document>"#
    );
    let footnote_error = Inventory::parse(footnote_document.as_bytes(), &Default::default())
        .expect_err("the retained focused extension must fail on its own");
    assert_eq!(
        inventory_error.to_string(),
        footnote_error.to_string(),
        "the semantic error must not replace the focused error after the same MCE pipeline"
    );
    assert!(!inventory_error.to_string().contains("section type"));
}

#[test]
fn inventory_reuse_keeps_raw_depth_error_before_footnote_error() {
    let nested = format!("{}{}", "<x:n>".repeat(64), "</x:n>".repeat(64));
    let section = |footnote_value: &str| {
        format!(
            r#"<w:sectPr xmlns:w="{W}" xmlns:x="urn:fragment-test" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><x:n>{nested}</x:n><w12:footnoteColumns w:val="{footnote_value}"/></w:sectPr>"#
        )
    };
    let valid_footnote = section("0");
    let malformed_footnote = section("-1");
    let valid_document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12" mc:PreserveElements="w12:footnoteColumns"><w:body>{valid_footnote}</w:body></w:document>"#
    );
    let malformed_document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12" mc:PreserveElements="w12:footnoteColumns"><w:body>{malformed_footnote}</w:body></w:document>"#
    );
    let raw_error = Inventory::parse(valid_document.as_bytes(), &Default::default())
        .expect_err("the private raw codec depth bound must reject this fragment");
    let combined_error = Inventory::parse(malformed_document.as_bytes(), &Default::default())
        .expect_err("the malformed footnote is behind the raw depth failure");
    assert_eq!(
        raw_error.to_string(),
        combined_error.to_string(),
        "raw validation must win before focused footnote validation"
    );
    assert!(raw_error.to_string().contains("section XML nesting"));
}

#[test]
fn inventory_reuse_keeps_reference_preflight_before_section_and_footnote_errors() {
    let relationship_id = "relationship-id";
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{REL}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w:body><w:sectPr><w:headerReference w:type="default" r:id="{relationship_id}"/><w12:footnoteColumns w:val="-1"/><w:type w:val="not-a-section-break"/></w:sectPr></w:body></w:document>"#
    );
    let limits = Limits {
        max_reference_bytes: relationship_id.len() - 1,
        ..Limits::default()
    };
    match Inventory::parse(xml.as_bytes(), &limits) {
        Err(Error::SectionInventoryLimit {
            kind,
            maximum,
            actual,
        }) => {
            assert_eq!(kind, LimitKind::HeaderFooterReferenceBytes);
            assert_eq!(maximum, relationship_id.len() - 1);
            assert_eq!(actual, relationship_id.len());
        },
        Err(error) => panic!("preflight returned a later error: {error}"),
        Ok(_) => panic!("reference limit was not enforced"),
    }
}

#[test]
fn inventory_reuse_maps_tolerance_to_original_offsets_after_namespace_injection() {
    let xml = format!(
        r#"<q:document xmlns:q="{W}" xmlns:rel="{REL}"><q:body><q:sectPr><q:headerReference q:type="default" rel:id="bad id"/><q:cols q:equalWidth="0" q:num="3"><q:col q:space="720"/><q:col q:w="4000"/></q:cols></q:sectPr></q:body></q:document>"#
    );
    let inventory =
        Inventory::parse(xml.as_bytes(), &Default::default()).expect("aliased section is valid");
    assert_eq!(inventory.sections().len(), 1);
    assert_eq!(
        inventory.sections()[0].headers()[0].relationship_id,
        "bad id"
    );

    let defects = inventory
        .tolerance_report()
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect::<Vec<_>>();
    assert_eq!(
        defects,
        [
            (
                Defect::Section(SectionDefect::NonConformingRelationshipId),
                xml.find("<q:headerReference").expect("header offset") as u64,
            ),
            (
                Defect::Section(SectionDefect::ColumnWidthMissing),
                xml.find("<q:col q:space").expect("column offset") as u64,
            ),
            (
                Defect::Section(SectionDefect::ColumnCountMismatch),
                xml.find("<q:cols").expect("columns offset") as u64,
            ),
        ]
    );
}

#[test]
fn inventory_reuse_preserves_mce_filtering_and_footnote_controls() {
    let absent = format!(r#"<q:document xmlns:q="{W}"><q:body><q:sectPr/></q:body></q:document>"#);
    assert!(Inventory::parse(absent.as_bytes(), &Default::default()).is_ok());

    let mut zero = Section::from_xml_bytes(
        format!(
            r#"<q:sectPr xmlns:q="{W}" xmlns:x12="{W12}" xmlns:m="{MC}" m:Ignorable="x12"><x12:footnoteColumns q:val="0"/></q:sectPr>"#
        )
        .into_bytes(),
    )
    .expect("a detached explicit-zero extension is valid");
    assert!(
        zero.footnote_columns()
            .expect("read explicit-zero extension")
            .is_some_and(|layout| layout.columns() == 0),
        "explicit zero must remain distinct from absence"
    );

    let inherited = format!(
        r#"<q:document xmlns:q="{W}" xmlns:x12="{W12}" xmlns:m="{MC}" m:Ignorable="x12"><q:body><q:sectPr><x12:footnoteColumns q:val="0"/></q:sectPr></q:body></q:document>"#
    );
    let inventory = Inventory::parse(inherited.as_bytes(), &Default::default())
        .expect("MCE removes this unknown ignorable extension before inventory scanning");
    assert_eq!(inventory.sections().len(), 1);

    let non_ignorable = format!(
        r#"<q:document xmlns:q="{W}" xmlns:x12="{W12}"><q:body><q:sectPr><x12:footnoteColumns q:val="0"/></q:sectPr></q:body></q:document>"#
    );
    let error = Inventory::parse(non_ignorable.as_bytes(), &Default::default())
        .expect_err("a non-ignorable extension must reach focused validation");
    assert!(
        error
            .to_string()
            .contains("mc:Ignorable does not include the footnoteColumns namespace"),
        "unexpected non-ignorable extension result: {error}"
    );
}
