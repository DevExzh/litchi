#![allow(
    clippy::default_trait_access,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::{
    Presentation, constants,
    handout_master::{Child, ChildKind, Master},
};
use std::path::PathBuf;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

fn package(styles: &str) -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><office:automatic-styles/><office:body><office:presentation><draw:page draw:name="Slide1"/></office:presentation></office:body></office:document-content>"#
    );
    let mut writer = litchi_odp::core::PackageWriter::new();
    writer.set_mimetype(constants::ODF_PRESENTATION).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    writer.finish().unwrap()
}

fn styles_xml(include_layout: bool, handout: &str) -> String {
    let layout = if include_layout {
        {
            r#"<style:presentation-page-layout style:name="handout-layout"><presentation:placeholder presentation:object="handout" svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/></style:presentation-page-layout>"#.to_string()
        }
    } else {
        Default::default()
    };
    format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:draw="{DRAW}" xmlns:presentation="{PRESENTATION}" xmlns:svg="{SVG}"><office:styles>{layout}</office:styles><office:automatic-styles><style:page-layout style:name="physical"/><style:style style:name="drawing" style:family="drawing-page"/></office:automatic-styles><office:master-styles>{handout}</office:master-styles></office:document-styles>"#
    )
}

#[test]
fn handout_master_facade_and_package_round_trip() {
    let mut presentation = Presentation::from_bytes(package(&styles_xml(true, ""))).unwrap();
    let mut master = Master::new("physical").unwrap();
    master.presentation_page_layout_name = Some("handout-layout".to_string());
    master.drawing_style_name = Some("drawing".to_string());
    master
        .push_child(Child::new(
            ChildKind::Shape,
            "<draw:rect draw:id=\"shape-1\"/>",
        ))
        .unwrap();

    presentation.set_handout_master(&master).unwrap();
    let bytes = presentation.to_bytes().unwrap();
    let reopened = Presentation::from_bytes(bytes).unwrap();
    let parsed = reopened.handout_master().unwrap().unwrap();

    assert_eq!(parsed, master);
    assert_eq!(parsed.children[0].kind, ChildKind::Shape);
    assert_eq!(
        Master::from_xml_fragment(&parsed.to_xml_fragment().unwrap()).unwrap(),
        master
    );
    let resolved = reopened.resolved_handout_master().unwrap().unwrap();
    assert_eq!(resolved.presentation_layout.unwrap().name, "handout-layout");
}

#[test]
fn handout_master_rejects_malformed_xml_and_missing_inheritance() {
    let missing_required = format!(r#"<style:handout-master xmlns:style="{STYLE}"/>"#);
    assert!(Master::from_xml_fragment(&missing_required).is_err());

    let unsupported_child = format!(
        r#"<style:handout-master xmlns:style="{STYLE}" style:page-layout-name="physical"><style:master-page/></style:handout-master>"#
    );
    assert!(Master::from_xml_fragment(&unsupported_child).is_err());

    let mut presentation = Presentation::from_bytes(package(&styles_xml(false, ""))).unwrap();
    let mut master = Master::new("physical").unwrap();
    master.presentation_page_layout_name = Some("missing-layout".to_string());
    let before = presentation.styles_xml().unwrap().to_string();

    assert!(presentation.set_handout_master(&master).is_err());
    assert_eq!(presentation.styles_xml().unwrap(), before);
    assert!(presentation.handout_master().unwrap().is_none());
}

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/odf/odp")
        .join(name)
}

#[test]
fn reads_handout_master_from_real_presentations() {
    for name in [
        "cellspan.odp",
        "tdf102223.odp",
        "tdf105502.odp",
        "tdf169979.odp",
        "text-in-image.odp",
    ] {
        let presentation = Presentation::open(fixture(name)).unwrap();
        let master = presentation
            .handout_master()
            .unwrap_or_else(|error| panic!("{name}: {error}"))
            .unwrap_or_else(|| panic!("{name} has no handout master"));
        assert!(!master.page_layout_name.is_empty(), "{name}");
        let fragment = master.to_xml_fragment().unwrap();
        assert!(fragment.starts_with("<style:handout-master"), "{name}");
    }
}

#[test]
fn bom_prefixed_styles_xml_keeps_fragments_exact() {
    let presentation = Presentation::open(fixture("tdf169979.odp")).unwrap();
    for page in presentation.master_pages().unwrap() {
        assert!(
            page.master_page.xml.starts_with("<style:master-page"),
            "BOM shifted master-page fragment"
        );
    }
    let fragment = presentation
        .handout_master()
        .unwrap()
        .unwrap()
        .to_xml_fragment()
        .unwrap();
    assert!(fragment.starts_with("<style:handout-master"));
}
