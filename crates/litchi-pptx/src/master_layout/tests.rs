#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::model::*;
use crate::Package;
use litchi_opc::PackageWriter;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use std::io::Cursor;

fn roundtrip(package: &Package) -> Package {
    let bytes = PackageWriter::to_bytes(package.opc().unwrap()).unwrap();
    Package::from_reader(Cursor::new(bytes)).unwrap()
}

fn uri(value: &str) -> PackURI {
    PackURI::new(value).unwrap()
}

#[test]
fn authored_master_and_layouts_roundtrip_through_read_side() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    assert_eq!(master.master_id, MIN_MASTER_OR_LAYOUT_ID + 1);

    let title_layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Title,
            "Custom Title",
            &[
                PlaceholderSpec::new(PlaceholderKind::CenteredTitle)
                    .with_text("Click to edit the custom title"),
                PlaceholderSpec::new(PlaceholderKind::Subtitle)
                    .with_index(1)
                    .with_text("Custom subtitle"),
            ],
        )
        .unwrap();
    let blank_layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "Custom Blank",
            &[],
        )
        .unwrap();
    assert!(title_layout.layout_id >= MIN_MASTER_OR_LAYOUT_ID);
    assert!(blank_layout.layout_id >= MIN_MASTER_OR_LAYOUT_ID);
    assert_ne!(title_layout.layout_id, blank_layout.layout_id);

    // Author placeholders on the master itself, then replace one.
    package
        .store_placeholder_shape(
            &master.part_name,
            &PlaceholderSpec::new(PlaceholderKind::Title).with_text("Master title"),
        )
        .unwrap();
    package
        .store_placeholder_shape(
            &master.part_name,
            &PlaceholderSpec::new(PlaceholderKind::DateTime).with_index(10),
        )
        .unwrap();
    package
        .store_placeholder_shape(
            &master.part_name,
            &PlaceholderSpec::new(PlaceholderKind::Title).with_text("Master title v2"),
        )
        .unwrap();
    package.validate_master_layout_graph().unwrap();

    let reopened = roundtrip(&package);
    reopened.validate_master_layout_graph().unwrap();
    let presentation = reopened.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();
    assert_eq!(masters.len(), 2, "default master plus authored master");

    let authored = masters
        .iter()
        .find(|candidate| candidate.part().part().partname().as_str() == master.part_name.as_str())
        .expect("authored master must resolve through the presentation");

    // Default text styles: title/body/other with nine levels each.
    assert_eq!(
        authored
            .part()
            .part()
            .blob()
            .windows(b"<a:lvl1pPr".len())
            .filter(|window| *window == b"<a:lvl1pPr")
            .count(),
        3
    );

    // Master placeholder inventory, including the replaced title text.
    let master_shapes = authored.shapes().unwrap();
    let titles = master_shapes
        .placeholders()
        .filter(|shape| {
            shape
                .placeholder()
                .is_some_and(|value| value.kind() == Some("title"))
        })
        .count();
    assert_eq!(titles, 1, "replaced title placeholder must not duplicate");
    let title = master_shapes
        .placeholders()
        .find(|shape| {
            shape
                .placeholder()
                .is_some_and(|value| value.kind() == Some("title"))
        })
        .unwrap();
    assert_eq!(title.text(), Some("Master title v2"));
    assert!(master_shapes.placeholders().any(|shape| {
        shape
            .placeholder()
            .is_some_and(|value| value.kind() == Some("dt") && value.index() == 10)
    }));

    // Layout inventory: kinds, names, placeholders, and back-references.
    let layouts = authored.layouts().unwrap();
    assert_eq!(layouts.len(), 2);
    let title_layout_read = &layouts[0];
    assert_eq!(title_layout_read.kind().unwrap().as_deref(), Some("title"));
    assert_eq!(title_layout_read.name().unwrap(), "Custom Title");
    assert_eq!(
        title_layout_read
            .master()
            .unwrap()
            .part()
            .part()
            .partname()
            .as_str(),
        master.part_name.as_str()
    );
    let layout_shapes = title_layout_read.shapes().unwrap();
    assert_eq!(layout_shapes.placeholders().count(), 2);
    let centered = layout_shapes
        .placeholders()
        .find(|shape| {
            shape
                .placeholder()
                .is_some_and(|value| value.kind() == Some("ctrTitle"))
        })
        .unwrap();
    assert_eq!(centered.text(), Some("Click to edit the custom title"));
    assert!(layout_shapes.placeholders().any(|shape| {
        shape
            .placeholder()
            .is_some_and(|value| value.kind() == Some("subTitle") && value.index() == 1)
    }));
    assert_eq!(layouts[1].kind().unwrap().as_deref(), Some("blank"));
    assert!(layouts[1].shapes().unwrap().placeholders().next().is_none());

    // The authored master inherits a working theme relationship.
    assert!(
        authored
            .part()
            .part()
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::THEME)
    );

    // The default master and its eleven layouts are untouched.
    let default_master = masters
        .iter()
        .find(|candidate| {
            candidate.part().part().partname().as_str() == "/ppt/slideMasters/slideMaster1.xml"
        })
        .unwrap();
    assert_eq!(default_master.layouts().unwrap().len(), 11);
}

#[test]
fn master_ids_are_unique_across_multiple_adds() {
    let mut package = Package::new().unwrap();
    let first = package.add_slide_master().unwrap();
    let second = package.add_slide_master().unwrap();
    let third = package.add_slide_master().unwrap();
    assert_eq!(first.master_id, MIN_MASTER_OR_LAYOUT_ID + 1);
    assert_eq!(second.master_id, MIN_MASTER_OR_LAYOUT_ID + 2);
    assert_eq!(third.master_id, MIN_MASTER_OR_LAYOUT_ID + 3);
    package.validate_master_layout_graph().unwrap();

    let reopened = roundtrip(&package);
    assert_eq!(
        reopened
            .presentation()
            .unwrap()
            .slide_masters()
            .unwrap()
            .len(),
        4
    );
    reopened.validate_master_layout_graph().unwrap();
}

#[test]
fn authored_layout_attaches_to_default_master() {
    let mut package = Package::new().unwrap();
    let layout = package
        .add_slide_layout(
            &uri("/ppt/slideMasters/slideMaster1.xml"),
            SlideLayoutKind::TwoObjects,
            "Two Objects Extra",
            &[PlaceholderSpec::new(PlaceholderKind::Object).with_index(7)],
        )
        .unwrap();
    assert!(layout.layout_id > MIN_MASTER_OR_LAYOUT_ID + 11);

    let reopened = roundtrip(&package);
    let presentation = reopened.presentation().unwrap();
    let default_master = &presentation.slide_masters().unwrap()[0];
    let layouts = default_master.layouts().unwrap();
    assert_eq!(layouts.len(), 12);
    let added = layouts
        .iter()
        .find(|candidate| candidate.name().unwrap() == "Two Objects Extra")
        .unwrap();
    assert_eq!(added.kind().unwrap().as_deref(), Some("twoObj"));
    let added_shapes = added.shapes().unwrap();
    let mut placeholders = added_shapes.placeholders();
    let placeholder = placeholders.next().unwrap().placeholder().unwrap();
    assert_eq!(placeholder.index(), 7);
    assert!(placeholders.next().is_none());
}

#[test]
fn invalid_references_are_rejected() {
    let mut package = Package::new().unwrap();

    // Unknown master part.
    assert!(
        package
            .add_slide_layout(
                &uri("/ppt/slideMasters/slideMaster99.xml"),
                SlideLayoutKind::Blank,
                "Nope",
                &[],
            )
            .is_err()
    );
    // Master part name pointing at a non-master part.
    assert!(
        package
            .add_slide_layout(
                &uri("/ppt/presentation.xml"),
                SlideLayoutKind::Blank,
                "Nope",
                &[],
            )
            .is_err()
    );
    // Placeholder authoring on a part that is not a master or layout.
    assert!(
        package
            .store_placeholder_shape(
                &uri("/ppt/presentation.xml"),
                &PlaceholderSpec::new(PlaceholderKind::Title),
            )
            .is_err()
    );
    // Empty layout names are rejected.
    let master = package.add_slide_master().unwrap();
    assert!(
        package
            .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "", &[])
            .is_err()
    );
    // Duplicate placeholder identities are rejected.
    assert!(
        package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::Blank,
                "Dup",
                &[
                    PlaceholderSpec::new(PlaceholderKind::Body).with_index(1),
                    PlaceholderSpec::new(PlaceholderKind::Body).with_index(1),
                ],
            )
            .is_err()
    );
    // Removing unknown layouts is rejected.
    assert!(
        package
            .remove_slide_layout(&uri("/ppt/slideLayouts/slideLayout99.xml"))
            .is_err()
    );
    package.validate_master_layout_graph().unwrap();
}

#[test]
fn remove_layout_rejects_slide_references() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "In Use", &[])
        .unwrap();

    // Attach a slide part that references the layout.
    package
            .edit_opc(|opc| {
            let slide_uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
            let mut slide = BlobPart::new(
                slide_uri,
                ct::PML_SLIDE.to_string(),
                b"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>".to_vec(),
            );
            slide.relate_to(
                &format!("../{}", layout.part_name.as_str().trim_start_matches("/ppt/")),
                rt::SLIDE_LAYOUT,
            );
            opc.add_part(Box::new(slide));
                Ok(())
            })
            .unwrap();

    assert!(package.remove_slide_layout(&layout.part_name).is_err());
    package.validate_master_layout_graph().unwrap();
}

#[test]
fn remove_empty_layout_keeps_graph_consistent() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "Temporary", &[])
        .unwrap();
    package
        .add_slide_layout(&master.part_name, SlideLayoutKind::TitleOnly, "Kept", &[])
        .unwrap();

    package.remove_slide_layout(&layout.part_name).unwrap();
    package.validate_master_layout_graph().unwrap();
    assert!(
        package.opc().unwrap().get_part(&layout.part_name).is_err(),
        "layout part must be gone"
    );

    let reopened = roundtrip(&package);
    reopened.validate_master_layout_graph().unwrap();
    let presentation = reopened.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();
    let authored = masters
        .iter()
        .find(|candidate| candidate.part().part().partname().as_str() == master.part_name.as_str())
        .unwrap();
    let layouts = authored.layouts().unwrap();
    assert_eq!(layouts.len(), 1);
    assert_eq!(layouts[0].name().unwrap(), "Kept");

    // Deleting it a second time is an error.
    assert!(package.remove_slide_layout(&layout.part_name).is_err());
}

#[test]
fn authored_parts_serialize_deterministically() {
    let build = || {
        let mut package = Package::new().unwrap();
        let master = package.add_slide_master().unwrap();
        package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::SectionHeader,
                "Deterministic",
                &[PlaceholderSpec::new(PlaceholderKind::Title).with_text("Same")],
            )
            .unwrap();
        package
    };
    let first = build();
    let second = build();
    for part_name in [
        "/ppt/slideMasters/slideMaster2.xml",
        "/ppt/slideLayouts/slideLayout12.xml",
    ] {
        let uri = PackURI::new(part_name).unwrap();
        assert_eq!(
            first.opc().unwrap().get_part(&uri).unwrap().blob(),
            second.opc().unwrap().get_part(&uri).unwrap().blob(),
            "part {part_name} must serialize deterministically"
        );
    }
}
