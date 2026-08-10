//! Focused regression tests for the layered header/footer owner.

use super::properties::{Color, Length, Region, StyleProperties};

#[test]
fn property_model_round_trips_through_the_canonical_owner() {
    let properties = StyleProperties {
        height: Some(Length::new("1.25cm").unwrap()),
        background_color: Some(Color::Rgb(0x10, 0x20, 0x30)),
        dynamic_spacing: Some(true),
        ..Default::default()
    };
    let fragment = properties.to_region_fragment(Region::Header).unwrap();
    let xml = format!(
        r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:automatic-styles><style:page-layout style:name="layout">{fragment}</style:page-layout></office:automatic-styles></office:document-styles>"#
    );
    let entry = super::parse_page_layout_header_footer_properties(&xml)
        .unwrap()
        .pop()
        .unwrap();
    assert_eq!(entry.region, Region::Header);
    assert_eq!(entry.properties, properties);
}

const SOURCE_LAYOUT: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles><style:page-layout style:name="pmA"><style:page-layout-properties style:print-orientation="portrait"/></style:page-layout><style:page-layout style:name="pmB"><style:page-layout-properties style:print-orientation="landscape"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="A" style:page-layout-name="pmA"><style:header><text:p>alpha</text:p></style:header></style:master-page><style:master-page style:name="B" style:page-layout-name="pmB"><style:footer><text:p>beta</text:p></style:footer></style:master-page></office:master-styles></office:document-styles>"#;

const DESTINATION_LAYOUT: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles><style:page-layout style:name="pmC"/></office:automatic-styles><office:master-styles><style:master-page style:name="C" style:page-layout-name="pmC"/></office:master-styles></office:document-styles>"#;

#[test]
fn advanced_layout_patch_is_durable_reversible_and_source_checked() {
    let source = super::Snapshot::parse(SOURCE_LAYOUT).expect("layout source parses");
    let mut edit = source.edit();
    edit.set_region_text("A", super::Kind::Header, "changed")
        .expect("header stages");
    let commit = edit.commit().expect("layout commits");
    let durable = commit.patch().durable().expect("durable layout patch");
    let wire = durable.to_deterministic_json().expect("deterministic JSON");
    let reopened =
        super::DurablePatch::from_deterministic_json(&wire).expect("durable layout patch reopens");
    let replayed = reopened.apply(&source).expect("layout patch replays");
    assert_eq!(
        replayed.snapshot().master_pages()[0]
            .region(super::Kind::Header)
            .expect("header remains")
            .text,
        "changed"
    );
    let restored = reopened
        .inverse()
        .apply(replayed.snapshot())
        .expect("inverse applies");
    assert_eq!(restored.snapshot().source_xml(), SOURCE_LAYOUT);
    assert!(
        reopened
            .apply(&super::Snapshot::parse(DESTINATION_LAYOUT).expect("destination parses"))
            .is_err()
    );
}

#[test]
fn master_transfer_carries_its_page_layout_dependency() {
    let source = super::Snapshot::parse(SOURCE_LAYOUT).expect("layout source parses");
    let transfer = source
        .prepare_master_page_transfer("A")
        .expect("master transfer prepares");
    assert_eq!(transfer.page_layout_name(), "pmA");
    assert!(transfer.dependencies().is_empty());

    let destination = super::Snapshot::parse(DESTINATION_LAYOUT).expect("destination parses");
    let mut edit = destination.edit();
    edit.insert_transfer(&transfer).expect("transfer stages");
    let transferred = edit.commit().expect("transfer commits");
    assert!(
        transferred
            .snapshot()
            .master_pages()
            .iter()
            .any(|master| master.name == "A")
    );
    assert!(
        transferred
            .snapshot()
            .page_layouts()
            .iter()
            .any(|layout| layout.name == "pmA")
    );
}

#[test]
fn layout_merge_composes_disjoint_owners_and_reports_overlap() {
    let source = super::Snapshot::parse(SOURCE_LAYOUT).expect("layout source parses");
    let mut left = source.edit();
    left.set_region_text("A", super::Kind::Header, "left")
        .expect("left stages");
    let left = left.commit().expect("left commits");

    let mut right = source.edit();
    right
        .replace_page_layout(
            "pmB",
            r#"<style:page-layout style:name="pmB"><style:page-layout-properties style:print-orientation="portrait"/></style:page-layout>"#,
        )
        .expect("right stages");
    let right = right.commit().expect("right commits");
    let merged = super::Patch::merge(left.patch(), right.patch())
        .expect("merge plans")
        .finish()
        .expect("disjoint merge finishes")
        .apply(&source)
        .expect("merged patch applies");
    assert_eq!(
        merged.snapshot().master_pages()[0]
            .region(super::Kind::Header)
            .expect("header remains")
            .text,
        "left"
    );
    assert_eq!(
        merged.snapshot().page_layouts()[1].page_usage,
        crate::page_layout::PageUsage::All
    );

    let mut competing = source.edit();
    competing
        .set_region_text("A", super::Kind::Header, "right")
        .expect("competing edit stages");
    let competing = competing.commit().expect("competing edit commits");
    let mut conflict = super::Patch::merge(left.patch(), competing.patch()).expect("merge plans");
    assert_eq!(
        conflict.conflicts().cloned().collect::<Vec<_>>(),
        vec![super::Target::MasterPage("A".to_string())]
    );
    assert!(conflict.clone().finish().is_err());
    conflict
        .resolve(
            &super::Target::MasterPage("A".to_string()),
            super::Resolution::Right,
        )
        .expect("conflict resolves");
    assert!(conflict.finish().is_ok());
}
