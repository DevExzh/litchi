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

fn corrupt_zip_member(mut bytes: Vec<u8>, member: &str) -> Vec<u8> {
    let member = member.as_bytes();
    let mut local_found = false;
    let mut cursor = 0_usize;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let header = cursor + relative;
        let name_length = u16::from_le_bytes([bytes[header + 26], bytes[header + 27]]) as usize;
        let extra_length = u16::from_le_bytes([bytes[header + 28], bytes[header + 29]]) as usize;
        let name_start = header + 30;
        let data_start = name_start + name_length + extra_length;
        if &bytes[name_start..name_start + name_length] == member {
            assert!(data_start < bytes.len(), "ZIP member has no payload");
            local_found = true;
            break;
        }
        cursor = header + 4;
    }
    assert!(local_found, "ZIP member {member:?} was not found");

    // Keep the archive structurally admissible while making the deferred read
    // fail deterministically at CRC verification time. Mutating a compressed
    // bit can still produce an equivalent stream, so a central-directory CRC
    // mismatch is the stable malformed-source fixture.
    let mut cursor = 0_usize;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let header = cursor + relative;
        if header + 46 > bytes.len() {
            break;
        }
        let name_length = u16::from_le_bytes([bytes[header + 28], bytes[header + 29]]) as usize;
        let extra_length = u16::from_le_bytes([bytes[header + 30], bytes[header + 31]]) as usize;
        let comment_length = u16::from_le_bytes([bytes[header + 32], bytes[header + 33]]) as usize;
        let name_start = header + 46;
        let name_end = name_start + name_length;
        let record_end = name_end + extra_length + comment_length;
        if record_end > bytes.len() {
            break;
        }
        if &bytes[name_start..name_end] == member {
            bytes[header + 16] ^= 1;
            return bytes;
        }
        cursor = header + 4;
    }
    panic!("central record for ZIP member {member:?} was not found");
}

fn malformed_layout_source() -> Vec<u8> {
    let mut package = Package::new().unwrap();
    let master = uri("/ppt/slideMasters/slideMaster1.xml");
    let original_layout = uri("/ppt/slideLayouts/slideLayout1.xml");
    let corrupt_layout = uri("/ppt/slideLayouts/corrupt.xml");
    package
        .edit_opc(|opc| {
            let relationship_id = opc
                .get_part(&master)
                .unwrap()
                .rels()
                .iter()
                .find(|relationship| relationship.reltype() == rt::SLIDE_LAYOUT)
                .unwrap()
                .r_id()
                .to_owned();
            opc.get_part_mut(&master)
                .unwrap()
                .rels_mut()
                .retarget(&relationship_id, "../slideLayouts/corrupt.xml".to_owned())
                .unwrap();
            assert!(opc.remove_part(&original_layout));
            opc.add_part(Box::new(BlobPart::new(
                corrupt_layout,
                ct::PML_SLIDE_LAYOUT.to_owned(),
                b"<corrupt/>".to_vec(),
            )));
            Ok(())
        })
        .unwrap();
    let bytes = PackageWriter::to_bytes(package.opc().unwrap()).unwrap();
    corrupt_zip_member(bytes, "ppt/slideLayouts/corrupt.xml")
}

fn malformed_orphan_theme_source() -> Vec<u8> {
    let mut package = Package::new().unwrap();
    let master = uri("/ppt/slideMasters/slideMaster1.xml");
    let notes_master = uri("/ppt/notesMasters/notesMaster1.xml");
    let theme = uri("/ppt/theme/theme1.xml");
    let notes_theme = uri("/ppt/theme/theme2.xml");
    let orphan_theme = uri("/ppt/theme/orphan.xml");
    package
        .edit_opc(|opc| {
            for part_name in [&master, &notes_master] {
                let relationship_ids: Vec<String> = opc
                    .get_part(part_name)
                    .unwrap()
                    .rels()
                    .iter()
                    .filter(|relationship| relationship.reltype() == rt::THEME)
                    .map(|relationship| relationship.r_id().to_owned())
                    .collect();
                let part = opc.get_part_mut(part_name).unwrap();
                for relationship_id in relationship_ids {
                    part.rels_mut().remove(&relationship_id);
                }
            }
            assert!(opc.remove_part(&theme));
            assert!(opc.remove_part(&notes_theme));
            opc.add_part(Box::new(BlobPart::new(
                orphan_theme,
                ct::OFC_THEME.to_owned(),
                b"<orphan-theme/>".to_vec(),
            )));
            Ok(())
        })
        .unwrap();
    let bytes = PackageWriter::to_bytes(package.opc().unwrap()).unwrap();
    corrupt_zip_member(bytes, "ppt/theme/orphan.xml")
}

#[test]
fn deferred_theme_fallback_is_forced_before_master_authoring() {
    let source = malformed_orphan_theme_source();
    let mut package = Package::from_vec(source.clone()).unwrap();
    assert!(package.add_slide_master().is_err());
    assert_eq!(
        PackageWriter::to_bytes(package.opc().unwrap()).unwrap(),
        source
    );
}

#[test]
fn deferred_graph_failure_rolls_back_master_and_layout_authoring() {
    let source = malformed_layout_source();
    let mut master_package = Package::from_vec(source.clone()).unwrap();
    assert!(master_package.add_slide_master().is_err());
    assert_eq!(
        PackageWriter::to_bytes(master_package.opc().unwrap()).unwrap(),
        source
    );

    let mut layout_package = Package::from_vec(source.clone()).unwrap();
    assert!(
        layout_package
            .add_slide_layout(
                &uri("/ppt/slideMasters/slideMaster1.xml"),
                SlideLayoutKind::Blank,
                "Deferred failure",
                &[],
            )
            .is_err()
    );
    assert_eq!(
        PackageWriter::to_bytes(layout_package.opc().unwrap()).unwrap(),
        source
    );
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
