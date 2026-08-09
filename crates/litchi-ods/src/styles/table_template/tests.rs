use super::{
    codec::parse_parts,
    semantic::{Axis, Template},
};
use litchi_core::Result;

const PREFIX: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:styles>"#;
const SUFFIX: &str = "</office:styles></office:document-styles>";

fn parse(fragment: &str) -> Result<Vec<Template>> {
    parse_parts(&[&format!("{PREFIX}{fragment}{SUFFIX}")])
}

#[test]
fn parses_complete_templates_and_legacy_axes() {
    let templates = parse(
        r#"<table:table-template table:name="Bands &amp; Body" text:first-row-start-column="row" table:first-row-end-column="column" table:use-first-row-styles="true" table:use-banding-rows-styles="1"><table:first-row table:style-name="Header" table:paragraph-style-name="HeaderP"/><table:body table:style-name="Body"/><table:even-rows table:style-name="Even"/><table:odd-rows table:style-name="Odd"/><table:background table:style-name="Background"/></table:table-template>"#,
    )
    .expect("test fixture or operation should succeed");
    let template = &templates[0];
    assert_eq!(template.name, "Bands & Body");
    assert_eq!(template.first_row_start_column, Some(Axis::Row));
    assert_eq!(template.first_row_end_column, Some(Axis::Column));
    assert_eq!(template.use_first_row_styles, Some(true));
    assert_eq!(
        template
            .first_row
            .as_ref()
            .expect("test fixture or operation should succeed")
            .paragraph_style_name
            .as_deref(),
        Some("HeaderP")
    );
    assert_eq!(
        template
            .background
            .as_ref()
            .expect("test fixture or operation should succeed")
            .style_name,
        "Background"
    );
}

#[test]
fn rejects_invalid_locations_shapes_and_duplicates() {
    for fragment in [
        r#"<table:table-template table:name="Missing"/>"#,
        r#"<table:table-template table:name="Partial"><table:even-rows table:style-name="E"/></table:table-template>"#,
        r#"<table:table-template table:name="MissingStyle"><table:body/></table:table-template>"#,
        r#"<table:table-template table:name="Duplicate"><table:body table:style-name="A"/><table:body table:style-name="B"/></table:table-template>"#,
        r#"<table:table-template table:name="Text"><table:body table:style-name="A">bad</table:body></table:table-template>"#,
        r#"<table:table-template table:name="Bool" table:use-first-row-styles="yes"><table:body table:style-name="A"/></table:table-template>"#,
    ] {
        assert!(parse(fragment).is_err(), "accepted {fragment}");
    }
    let duplicate = format!(
        "{PREFIX}<table:table-template table:name=\"A\"><table:body table:style-name=\"A\"/></table:table-template><table:table-template table:name=\"A\"><table:body table:style-name=\"B\"/></table:table-template>{SUFFIX}"
    );
    assert!(parse_parts(&[&duplicate]).is_err());
    assert!(parse_parts(&[r#"<table:table-template xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" table:name="A"><table:body table:style-name="A"/></table:table-template>"#]).is_err());
}

#[test]
fn parses_libreoffice_table_style_catalog() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/res/xml/tablestyles.xml");
    let xml = std::fs::read_to_string(path).expect("test fixture or operation should succeed");
    let templates = parse_parts(&[&xml]).expect("test fixture or operation should succeed");
    assert!(templates.len() > 10);
    let default = templates
        .iter()
        .find(|template| template.name == "Default Style")
        .expect("test fixture or operation should succeed");
    assert_eq!(
        default
            .body
            .as_ref()
            .expect("test fixture or operation should succeed")
            .style_name,
        "Default-Style.body"
    );
    assert!(default.background.is_some());
}

#[test]
fn validates_and_round_trips_deterministic_template_xml() {
    let mut template = parse(
        r#"<table:table-template table:name="A &amp; B" table:use-banding-columns-styles="false"><table:first-row table:style-name="Head&amp;" table:paragraph-style-name="P&amp;"/><table:even-columns table:style-name="Even"/><table:odd-columns table:style-name="Odd"/></table:table-template>"#,
    )
    .expect("test fixture or operation should succeed")
    .remove(0);
    let xml = template
        .to_xml()
        .expect("test fixture or operation should succeed");
    assert!(xml.contains(r#"table:name="A &amp; B""#));
    assert!(xml.contains(r#"table:style-name="Head&amp;""#));
    let reparsed = parse(&xml)
        .expect("test fixture or operation should succeed")
        .remove(0);
    assert_eq!(reparsed, template);

    template.odd_columns = None;
    let mut untouched = String::from("prefix");
    assert!(template.write_xml(&mut untouched).is_err());
    assert_eq!(untouched, "prefix");
}

#[test]
fn exposes_typed_region_accessors() {
    use super::semantic::{Region, Style};
    let template = Template::new("Bands").with_region(
        Region::Body,
        Style::new("Body").with_paragraph_style("BodyText"),
    );
    assert_eq!(
        template
            .region(Region::Body)
            .expect("test fixture or operation should succeed")
            .style_name,
        "Body"
    );
    assert_eq!(
        template
            .region(Region::Body)
            .expect("test fixture or operation should succeed")
            .paragraph_style_name
            .as_deref(),
        Some("BodyText")
    );
    assert!(template.region(Region::Background).is_none());
}
