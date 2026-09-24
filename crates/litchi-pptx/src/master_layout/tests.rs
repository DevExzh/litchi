#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::model::*;
use super::{PLACEHOLDER_TYPE_EXTENSION_URI, PlaceholderTypeExtension};
use super::{PlaceholderSignaturePolicy, PlaceholderTypeExtensionLimits};
use crate::Package;
use litchi_opc::PackageWriter;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use std::io::Cursor;
use std::sync::Arc;

fn roundtrip(package: &Package) -> Package {
    let bytes = PackageWriter::to_bytes(package.opc().unwrap()).unwrap();
    Package::from_reader(Cursor::new(bytes)).unwrap()
}

fn uri(value: &str) -> PackURI {
    PackURI::new(value).unwrap()
}

fn zip_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}

fn corrupt_zip_member(mut bytes: Vec<u8>, member: &str) -> Vec<u8> {
    let member = member.as_bytes();
    let mut local = None;
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
            local = Some((header, data_start));
            break;
        }
        cursor = header + 4;
    }
    let Some((local_header, data_start)) = local else {
        panic!("ZIP member {member:?} was not found");
    };

    // Keep the archive structurally admissible while making the deferred read
    // fail deterministically at CRC verification time. Mutating a compressed
    // bit can still produce an equivalent stream, so a declared-CRC mismatch
    // is the stable malformed-source fixture. Admission cross-checks the
    // central record against the local header or the data descriptor, so the
    // same CRC bit is flipped in every header that records it: the headers
    // stay consistent with one another and only the payload disagrees.
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
            let crc = zip_u32(&bytes, header + 16);
            let compressed_size = zip_u32(&bytes, header + 20);
            assert_ne!(
                compressed_size,
                u32::MAX,
                "fixture member must not be ZIP64"
            );
            bytes[header + 16] ^= 1;
            if zip_u32(&bytes, local_header + 14) == crc {
                bytes[local_header + 14] ^= 1;
            }
            let flags = u16::from_le_bytes([bytes[local_header + 6], bytes[local_header + 7]]);
            if flags & 0x0008 != 0 {
                let mut descriptor = data_start + compressed_size as usize;
                if &bytes[descriptor..descriptor + 4] == b"PK\x07\x08" {
                    descriptor += 4;
                }
                assert_eq!(
                    zip_u32(&bytes, descriptor),
                    crc,
                    "data descriptor CRC for ZIP member {member:?}"
                );
                bytes[descriptor] ^= 1;
            }
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
                    .with_type_extension(PlaceholderTypeExtension::Cameo)
                    .with_text("Click to edit the custom title"),
                PlaceholderSpec::new(PlaceholderKind::Subtitle)
                    .with_index(1)
                    .with_type_extension(PlaceholderTypeExtension::Unknown)
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
    assert_eq!(
        centered
            .placeholder()
            .expect("centered placeholder")
            .type_extension(),
        Some(PlaceholderTypeExtension::Cameo)
    );
    assert!(layout_shapes.placeholders().any(|shape| {
        shape
            .placeholder()
            .is_some_and(|value| value.kind() == Some("subTitle") && value.index() == 1)
    }));
    let subtitle = layout_shapes
        .placeholders()
        .find(|shape| {
            shape
                .placeholder()
                .is_some_and(|value| value.kind() == Some("subTitle"))
        })
        .expect("subtitle placeholder");
    assert_eq!(
        subtitle.placeholder().expect("subtitle").type_extension(),
        Some(PlaceholderTypeExtension::Unknown)
    );
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
fn p232_scalar_patch_preserves_unknown_shape_bytes_and_inverse() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "P232",
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Cameo")
                .with_type_extension(PlaceholderTypeExtension::Cameo)],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let patched = source.replacen(
                "</p:sp>",
                r#"<q:opaque xmlns:q="urn:opaque">keep</q:opaque></p:sp>"#,
                1,
            );
            part.set_blob(patched.into_bytes());
            Ok(())
        })
        .unwrap();

    let source = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Cameo")
        .unwrap();
    let before = source.source_xml().to_vec();
    let mut edit = source.edit();
    edit.set(PlaceholderTypeExtension::Unknown);
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());
    assert!(!commit.is_noop());
    assert!(
        commit
            .snapshot()
            .source_xml()
            .windows(b"<q:opaque xmlns:q=\"urn:opaque\">keep</q:opaque>".len())
            .any(|window| window == b"<q:opaque xmlns:q=\"urn:opaque\">keep</q:opaque>")
    );

    package
        .apply_placeholder_type_extension_patch(commit.patch())
        .unwrap();
    let after = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Cameo")
        .unwrap();
    assert_eq!(after.value(), PlaceholderTypeExtension::Unknown);
    assert_ne!(after.source_xml(), before.as_slice());

    package
        .apply_placeholder_type_extension_patch(&commit.patch().inverse())
        .unwrap();
    let restored = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Cameo")
        .unwrap();
    assert_eq!(restored.value(), PlaceholderTypeExtension::Cameo);
    assert_eq!(restored.source_xml(), before.as_slice());
}

#[test]
fn p232_scalar_patch_preserves_an_inherited_alias_and_lexical_owner_bytes() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "P232",
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Alias")
                .with_type_extension(PlaceholderTypeExtension::Cameo)],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let source = source
                .replace("xmlns:p232=", "xmlns:x=")
                .replace("p232:", "x:")
                .replace(
                    "urn:litchi:pptx:x:phTypeExt",
                    "urn:litchi:pptx:p232:phTypeExt",
                );
            part.set_blob(source.into_bytes());
            Ok(())
        })
        .unwrap();

    let source = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Alias")
        .unwrap();
    let before = source.source_xml().to_vec();
    let mut edit = source.edit();
    edit.set(PlaceholderTypeExtension::Unknown);
    let commit = edit.commit().unwrap();
    package
        .apply_placeholder_type_extension_patch(commit.patch())
        .unwrap();
    let after = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Alias")
        .unwrap();
    assert_eq!(after.value(), PlaceholderTypeExtension::Unknown);
    assert!(
        after
            .source_xml()
            .windows(b"<x:unknown/>".len())
            .any(|window| window == b"<x:unknown/>")
    );
    assert!(
        !after
            .source_xml()
            .windows(b"<p232:unknown/>".len())
            .any(|window| window == b"<p232:unknown/>")
    );
    assert_eq!(
        after.source_xml().len(),
        before.len() + b"unknown".len() - b"cameo".len()
    );
}

#[test]
fn p232_scalar_patch_updates_explicit_end_tag_without_touching_opaque_siblings() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "Explicit",
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Explicit")
                .with_type_extension(PlaceholderTypeExtension::Cameo)],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let source = source
                .replace(
                    "<p232:cameo/>",
                    "<p232:cameo></p232:cameo>",
                )
                .replace(
                    "</p:extLst>",
                    "<!--before--><p:ext uri=\"urn:opaque\" xmlns:q=\"urn:q\"><q:future/></p:ext><!--after--></p:extLst>",
                );
            part.set_blob(source.into_bytes());
            Ok(())
        })
        .unwrap();

    let snapshot = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Explicit")
        .unwrap();
    let before = snapshot.source_xml().to_vec();
    let mut edit = snapshot.edit();
    edit.set(PlaceholderTypeExtension::Unknown);
    let commit = edit.commit().unwrap();
    package
        .apply_placeholder_type_extension_patch(commit.patch())
        .unwrap();
    let changed = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Explicit")
        .unwrap();
    assert!(
        std::str::from_utf8(changed.source_xml())
            .unwrap()
            .contains("<p232:unknown></p232:unknown>")
    );
    assert!(
        std::str::from_utf8(changed.source_xml())
            .unwrap()
            .contains("<q:future/>")
    );

    package
        .apply_placeholder_type_extension_patch(&commit.patch().inverse())
        .unwrap();
    assert_eq!(
        package
            .placeholder_type_extension_snapshot(&layout.part_name, "Explicit")
            .unwrap()
            .source_xml(),
        before.as_slice()
    );
}

#[test]
fn p232_noop_is_byte_stable_and_signature_policy_is_explicit() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "P232",
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Cameo")
                .with_type_extension(PlaceholderTypeExtension::Cameo)],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            opc.rels_mut().add_relationship(
                rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rIdSignature".to_owned(),
                false,
            );
            Ok(())
        })
        .unwrap();
    assert!(package.opc().unwrap().is_signed());

    let source = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Cameo")
        .unwrap();
    let mut noop_edit = source.clone().edit();
    noop_edit.set(PlaceholderTypeExtension::Cameo);
    let noop = noop_edit.commit().unwrap();
    assert!(noop.is_noop());
    package
        .apply_placeholder_type_extension_patch(noop.patch())
        .unwrap();
    assert!(package.opc().unwrap().is_signed());

    let mut changed_edit = source.edit();
    changed_edit.set(PlaceholderTypeExtension::Unknown);
    let changed = changed_edit.commit().unwrap();
    let error = package
        .apply_placeholder_type_extension_patch(changed.patch())
        .unwrap_err();
    assert!(matches!(
        error,
        crate::Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(package.opc().unwrap().is_signed());
    package
        .apply_placeholder_type_extension_patch_with_policy(
            changed.patch(),
            PlaceholderSignaturePolicy::Invalidate,
        )
        .unwrap();
    assert!(!package.opc().unwrap().is_signed());
}

#[test]
fn p232_patch_rejects_stale_source_and_growth_before_staging() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "P232",
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Cameo")
                .with_type_extension(PlaceholderTypeExtension::Cameo)],
        )
        .unwrap();
    let source = package
        .placeholder_type_extension_snapshot(&layout.part_name, "Cameo")
        .unwrap();
    let mut edit = source.clone().edit();
    edit.set(PlaceholderTypeExtension::Unknown);
    let patch = edit.commit().unwrap().into_patch();

    let exact_limits =
        PlaceholderTypeExtensionLimits::new(source.source_xml().len() + 2, 2).unwrap();
    let exact_snapshot = super::load_placeholder_type_extension_snapshot_with_limits(
        package.opc().unwrap(),
        &layout.part_name,
        "Cameo",
        exact_limits,
    )
    .unwrap();
    let mut exact_edit = exact_snapshot.edit();
    exact_edit.set(PlaceholderTypeExtension::Unknown);
    assert!(exact_edit.commit().is_ok());

    let tight_limits =
        PlaceholderTypeExtensionLimits::new(source.source_xml().len() + 2, 1).unwrap();
    let tight_snapshot = super::load_placeholder_type_extension_snapshot_with_limits(
        package.opc().unwrap(),
        &layout.part_name,
        "Cameo",
        tight_limits,
    )
    .unwrap();
    let mut tight_edit = tight_snapshot.edit();
    tight_edit.set(PlaceholderTypeExtension::Unknown);
    assert!(tight_edit.commit().is_err());

    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let mut bytes = part.blob().to_vec();
            bytes.extend_from_slice(b" ");
            part.set_blob(bytes);
            Ok(())
        })
        .unwrap();
    let stale_before = package
        .opc()
        .unwrap()
        .get_part(&layout.part_name)
        .unwrap()
        .blob()
        .to_vec();
    assert!(matches!(
        package.apply_placeholder_type_extension_patch(&patch),
        Err(crate::Error::StaleSource)
    ));
    assert_eq!(
        package
            .opc()
            .unwrap()
            .get_part(&layout.part_name)
            .unwrap()
            .blob(),
        stale_before.as_slice()
    );

    let bounded = PlaceholderTypeExtensionLimits::new(source.source_xml().len(), 1).unwrap();
    let bounded_source = super::load_placeholder_type_extension_snapshot_with_limits(
        package.opc().unwrap(),
        &layout.part_name,
        "Cameo",
        bounded,
    );
    assert!(
        bounded_source.is_err(),
        "stale owner must fail before limits"
    );
}

#[test]
fn p232_slot_add_remove_preserves_shape_bytes_and_inverse() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "P232 slot",
            &[PlaceholderSpec::new(PlaceholderKind::Object).with_name("Slot")],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let patched = source.replacen(
                "</p:sp>",
                r#"<q:opaque xmlns:q="urn:opaque">keep</q:opaque></p:sp>"#,
                1,
            );
            part.set_blob(patched.into_bytes());
            Ok(())
        })
        .unwrap();

    let absent = package
        .placeholder_type_extension_slot_snapshot(&layout.part_name, "Slot")
        .unwrap();
    assert_eq!(absent.value(), None);
    let before = absent.source_xml().to_vec();
    let mut add = absent.edit();
    add.set(PlaceholderTypeExtension::Cameo);
    let added = add.commit().unwrap();
    assert!(added.is_changed());
    assert!(
        added
            .snapshot()
            .source_xml()
            .windows(b"keep".len())
            .any(|w| w == b"keep")
    );
    package
        .apply_placeholder_type_extension_slot_patch(added.patch())
        .unwrap();
    let present = package
        .placeholder_type_extension_slot_snapshot(&layout.part_name, "Slot")
        .unwrap();
    assert_eq!(present.value(), Some(PlaceholderTypeExtension::Cameo));

    let mut remove = present.clone().edit();
    remove.remove();
    let removed = remove.commit().unwrap();
    package
        .apply_placeholder_type_extension_slot_patch(removed.patch())
        .unwrap();
    let restored_absent = package
        .placeholder_type_extension_slot_snapshot(&layout.part_name, "Slot")
        .unwrap();
    assert_eq!(restored_absent.value(), None);

    package
        .apply_placeholder_type_extension_slot_patch(&removed.patch().inverse())
        .unwrap();
    let restored_present = package
        .placeholder_type_extension_slot_snapshot(&layout.part_name, "Slot")
        .unwrap();
    assert_eq!(
        restored_present.value(),
        Some(PlaceholderTypeExtension::Cameo)
    );
    package
        .apply_placeholder_type_extension_slot_patch(&added.patch().inverse())
        .unwrap();
    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&layout.part_name, "Slot")
            .unwrap()
            .source_xml(),
        before.as_slice()
    );
}

#[test]
fn p232_slot_remove_inverse_restores_lexical_owner_with_opaque_variant_sibling() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "Lexical slot",
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Lexical")
                .with_type_extension(PlaceholderTypeExtension::Cameo)],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let source = source.replace(
                "<p:extLst><p:ext uri=\"urn:litchi:pptx:p232:phTypeExt\">",
                "<p:extLst> <!--before--> <p:ext uri=\"urn:litchi:pptx:p232:phTypeExt\">",
            ).replace(
                "</p:extLst>",
                "<!--between--><p:ext uri=\"urn:opaque\" xmlns:q=\"urn:q\"><q:cameo/></p:ext> <!--after--></p:extLst>",
            );
            part.set_blob(source.into_bytes());
            Ok(())
        })
        .unwrap();

    let present = package
        .placeholder_type_extension_slot_snapshot(&layout.part_name, "Lexical")
        .unwrap();
    let before = present.source_xml().to_vec();
    let mut remove = present.edit();
    remove.remove();
    let commit = remove.commit().unwrap();
    package
        .apply_placeholder_type_extension_slot_patch(commit.patch())
        .unwrap();
    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&layout.part_name, "Lexical")
            .unwrap()
            .value(),
        None
    );
    package
        .apply_placeholder_type_extension_slot_patch(&commit.patch().inverse())
        .unwrap();
    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&layout.part_name, "Lexical")
            .unwrap()
            .source_xml(),
        before.as_slice()
    );
}

#[test]
fn p232_slot_add_targets_selected_shape_when_another_shape_has_ext_lst() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "P232 multiple slots",
            &[
                PlaceholderSpec::new(PlaceholderKind::Object)
                    .with_name("Existing")
                    .with_type_extension(PlaceholderTypeExtension::Cameo),
                PlaceholderSpec::new(PlaceholderKind::Object)
                    .with_index(1)
                    .with_name("Missing"),
            ],
        )
        .unwrap();
    let before = package
        .placeholder_type_extension_slot_snapshot(&layout.part_name, "Missing")
        .unwrap();
    assert_eq!(before.value(), None);
    let mut edit = before.edit();
    edit.add(PlaceholderTypeExtension::Unknown);
    let commit = edit.commit().unwrap();
    package
        .apply_placeholder_type_extension_slot_patch(commit.patch())
        .unwrap();
    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&layout.part_name, "Missing")
            .unwrap()
            .value(),
        Some(PlaceholderTypeExtension::Unknown)
    );
    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&layout.part_name, "Existing")
            .unwrap()
            .value(),
        Some(PlaceholderTypeExtension::Cameo)
    );
}

#[test]
fn p232_source_store_preserves_unknown_shape_bytes_on_slide_and_master() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let slide_uri = uri("/ppt/slides/slide-p232.xml");
    let slide_xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#;
    package
        .edit_opc(|opc| {
            opc.add_part(Box::new(BlobPart::new(
                slide_uri.clone(),
                ct::PML_SLIDE.to_string(),
                slide_xml.to_vec(),
            )));
            Ok(())
        })
        .unwrap();

    let slide_spec = PlaceholderSpec::new(PlaceholderKind::Object)
        .with_name("SlideSlot")
        .with_text("before")
        .with_type_extension(PlaceholderTypeExtension::Cameo);
    package
        .store_placeholder_shape(&slide_uri, &slide_spec)
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&slide_uri)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let patched = source.replacen(
                "</p:sp>",
                r#"<q:opaque xmlns:q="urn:opaque">slide-keep</q:opaque></p:sp>"#,
                1,
            );
            part.set_blob(patched.into_bytes());
            Ok(())
        })
        .unwrap();
    package
        .store_placeholder_shape(
            &slide_uri,
            &PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("SlideSlot")
                .with_text("after")
                .with_type_extension(PlaceholderTypeExtension::Unknown),
        )
        .unwrap();
    let slide_part = package.opc().unwrap().get_part(&slide_uri).unwrap();
    assert!(
        slide_part
            .blob()
            .windows(b"slide-keep".len())
            .any(|w| w == b"slide-keep")
    );
    let slide = package
        .placeholder_type_extension_slot_snapshot(&slide_uri, "SlideSlot")
        .unwrap();
    assert_eq!(slide.value(), Some(PlaceholderTypeExtension::Unknown));
    let mut remove = slide.clone().edit();
    remove.remove();
    let removed = remove.commit().unwrap();
    package
        .apply_placeholder_type_extension_slot_patch(removed.patch())
        .unwrap();
    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&slide_uri, "SlideSlot")
            .unwrap()
            .value(),
        None
    );
    package
        .apply_placeholder_type_extension_slot_patch(&removed.patch().inverse())
        .unwrap();
    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&slide_uri, "SlideSlot")
            .unwrap()
            .value(),
        Some(PlaceholderTypeExtension::Unknown)
    );

    package
        .store_placeholder_shape(
            &master.part_name,
            &PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("MasterSlot")
                .with_type_extension(PlaceholderTypeExtension::Cameo),
        )
        .unwrap();
    let master_source = package
        .placeholder_type_extension_slot_snapshot(&master.part_name, "MasterSlot")
        .unwrap();
    assert_eq!(master_source.value(), Some(PlaceholderTypeExtension::Cameo));
}

#[test]
fn p232_legacy_facade_accepts_token_uri_whitespace_without_rewriting_it() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "Token URI",
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("TokenSlot")
                .with_type_extension(PlaceholderTypeExtension::Cameo)],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let canonical = format!("uri=\"{PLACEHOLDER_TYPE_EXTENSION_URI}\"");
            let spaced = format!("uri=\" \t{PLACEHOLDER_TYPE_EXTENSION_URI}\n \"");
            let patched = source.replacen(&canonical, &spaced, 1);
            assert_ne!(patched, source);
            part.set_blob(patched.into_bytes());
            Ok(())
        })
        .unwrap();

    assert_eq!(
        package
            .placeholder_type_extension_slot_snapshot(&layout.part_name, "TokenSlot")
            .unwrap()
            .value(),
        Some(PlaceholderTypeExtension::Cameo)
    );
    package
        .store_placeholder_shape(
            &layout.part_name,
            &PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("TokenSlot")
                .with_type_extension(PlaceholderTypeExtension::Unknown),
        )
        .unwrap();
    let raw = std::str::from_utf8(
        package
            .opc()
            .unwrap()
            .get_part(&layout.part_name)
            .unwrap()
            .blob(),
    )
    .unwrap();
    assert!(raw.contains(&format!("uri=\" \t{PLACEHOLDER_TYPE_EXTENSION_URI}\n \"")));
    assert!(raw.contains("<p232:unknown/>") || raw.contains("<p232:unknown />"));
}

#[test]
fn p232_legacy_facade_preflights_validation_and_preserves_noop_blob_sharing() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let spec = PlaceholderSpec::new(PlaceholderKind::Title)
        .with_name("Shared")
        .with_text("same")
        .with_type_extension(PlaceholderTypeExtension::Cameo);
    package
        .store_placeholder_shape(&master.part_name, &spec)
        .unwrap();
    let before = package
        .opc()
        .unwrap()
        .get_part(&master.part_name)
        .unwrap()
        .blob_arc();
    package
        .store_placeholder_shape(&master.part_name, &spec)
        .unwrap();
    let after = package
        .opc()
        .unwrap()
        .get_part(&master.part_name)
        .unwrap()
        .blob_arc();
    assert!(
        Arc::ptr_eq(&before, &after),
        "legacy no-op must retain its blob"
    );

    let too_many: Vec<_> = (0..65)
        .map(|index| PlaceholderSpec::new(PlaceholderKind::Object).with_index(index))
        .collect();
    assert!(
        package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::Blank,
                "Too many",
                &too_many
            )
            .is_err()
    );
    assert!(
        package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::Blank,
                &"x".repeat(257),
                &[],
            )
            .is_err()
    );
    assert!(
        package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::Blank,
                "Control",
                &[PlaceholderSpec::new(PlaceholderKind::Body).with_name("bad\u{1}")],
            )
            .is_err()
    );
    assert!(
        package
            .store_placeholder_shape(
                &master.part_name,
                &PlaceholderSpec::new(PlaceholderKind::Title).with_text("bad\u{1}"),
            )
            .is_err()
    );
    assert!(
        package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::Blank,
                "Duplicate",
                &[
                    PlaceholderSpec::new(PlaceholderKind::Body).with_index(4),
                    PlaceholderSpec::new(PlaceholderKind::Body).with_index(4),
                ],
            )
            .is_err()
    );
}

#[test]
fn p232_legacy_facade_flushes_new_slides_before_preflight_and_retires_writer() {
    let mut package = Package::new().unwrap();
    {
        let presentation = package.presentation_mut().unwrap();
        presentation.add_slide().unwrap();
    }
    let slide_uri = uri("/ppt/slides/slide1.xml");
    package
        .store_placeholder_shape(
            &slide_uri,
            &PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Pending slide")
                .with_type_extension(PlaceholderTypeExtension::Cameo),
        )
        .unwrap();
    assert!(
        package.presentation_mut().is_err(),
        "raw placeholder mutation must retire the stale mutable writer"
    );

    let bytes = package.to_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened
            .placeholder_type_extension_slot_snapshot(&slide_uri, "Pending slide")
            .unwrap()
            .value(),
        Some(PlaceholderTypeExtension::Cameo)
    );
}

#[test]
fn authored_crlf_names_and_text_round_trip_through_numeric_references() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout_name = "Layout\r\nName";
    let placeholder_name = "Placeholder\rName\n";
    let placeholder_text = "Text\r\nLine\nTail";
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            layout_name,
            &[PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name(placeholder_name)
                .with_text(placeholder_text)],
        )
        .unwrap();
    let raw = package
        .opc()
        .unwrap()
        .get_part(&layout.part_name)
        .unwrap()
        .blob();
    assert!(raw.windows(b"&#xD;".len()).any(|window| window == b"&#xD;"));
    assert!(raw.windows(b"&#xA;".len()).any(|window| window == b"&#xA;"));
    let bytes = package.to_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    let layout_view = reopened
        .presentation()
        .unwrap()
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|candidate| candidate.part().part().partname() == &layout.part_name)
        .unwrap();
    assert_eq!(layout_view.name().unwrap(), layout_name);
    let shapes = layout_view.shapes().unwrap();
    let shape = shapes.get(placeholder_name).unwrap().unwrap();
    assert_eq!(shape.text(), Some(placeholder_text));
}

#[test]
fn authored_layout_accepts_exact_part_limit_and_rejects_one_byte_over() {
    let empty = PlaceholderSpec::new(PlaceholderKind::Object).with_text("");
    let base = super::codec::layout_xml(SlideLayoutKind::Blank, "Exact", &[empty])
        .unwrap()
        .len();
    let exact_text = "x".repeat(super::codec::MAX_PART_XML_BYTES - base);
    let exact = PlaceholderSpec::new(PlaceholderKind::Object).with_text(exact_text);
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "Exact", &[exact])
        .unwrap();
    assert_eq!(
        package
            .opc()
            .unwrap()
            .get_part(&layout.part_name)
            .unwrap()
            .blob()
            .len(),
        super::codec::MAX_PART_XML_BYTES
    );

    let over_text = "x".repeat(super::codec::MAX_PART_XML_BYTES - base + 1);
    let over = PlaceholderSpec::new(PlaceholderKind::Object).with_text(over_text);
    assert!(
        package
            .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "Exact", &[over])
            .is_err()
    );
}

#[test]
fn p232_source_store_replaces_all_drawingml_text_runs() {
    let mut package = Package::new().unwrap();
    let master = package.add_slide_master().unwrap();
    let layout = package
        .add_slide_layout(
            &master.part_name,
            SlideLayoutKind::Blank,
            "Runs",
            &[PlaceholderSpec::new(PlaceholderKind::Object).with_name("Runs")],
        )
        .unwrap();
    package
        .edit_opc(|opc| {
            let part = opc.get_part_mut(&layout.part_name)?;
            let source = std::str::from_utf8(part.blob())
                .map_err(|error| crate::Error::Invalid(error.to_string()))?;
            let patched = source.replacen(
                "<a:p><a:endParaRPr",
                "<a:p><a:r><a:t>old one</a:t></a:r><a:r><a:t>old two</a:t></a:r><a:endParaRPr",
                1,
            );
            part.set_blob(patched.into_bytes());
            Ok(())
        })
        .unwrap();

    package
        .store_placeholder_shape(
            &layout.part_name,
            &PlaceholderSpec::new(PlaceholderKind::Object)
                .with_name("Runs")
                .with_text("updated & escaped"),
        )
        .unwrap();
    let part = package.opc().unwrap().get_part(&layout.part_name).unwrap();
    let scene = crate::shape::Scene::read(part.blob()).unwrap();
    let shape = scene.get("Runs").unwrap().unwrap();
    assert_eq!(shape.text(), Some("updated & escaped"));
    assert!(
        !std::str::from_utf8(part.blob())
            .unwrap()
            .contains("old one")
    );
    assert!(
        !std::str::from_utf8(part.blob())
            .unwrap()
            .contains("old two")
    );
}

#[test]
fn p232_slot_rejects_mce_rewritten_owner_before_mutation() {
    let mut package = Package::new().unwrap();
    let slide_uri = uri("/ppt/slides/mce-p232.xml");
    let xml = r#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="http://schemas.microsoft.com/office/powerpoint/2023/02/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:nvGrpSpPr/><p:grpSpPr/><mc:AlternateContent><mc:Choice Requires="x"><p:sp><p:nvSpPr><p:cNvPr id="2" name="Cameo"/><p:nvPr><p:ph type="obj"><p:extLst><p:ext uri="urn:litchi:pptx:p232:phTypeExt"><x:phTypeExt><x:type><x:cameo/></x:type></x:phTypeExt></p:ext></p:extLst></p:ph></p:nvPr></p:nvSpPr></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="3" name="Fallback"/><p:nvPr><p:ph type="obj"/></p:nvPr></p:nvSpPr></p:sp></mc:Fallback></mc:AlternateContent></p:spTree>"#;
    package
        .edit_opc(|opc| {
            opc.add_part(Box::new(BlobPart::new(
                slide_uri.clone(),
                ct::PML_SLIDE.to_string(),
                xml.as_bytes().to_vec(),
            )));
            Ok(())
        })
        .unwrap();
    let before = package
        .opc()
        .unwrap()
        .get_part(&slide_uri)
        .unwrap()
        .blob()
        .to_vec();
    let error = package
        .placeholder_type_extension_slot_snapshot(&slide_uri, "Cameo")
        .unwrap_err();
    assert!(matches!(error, crate::Error::UnsafeEdit { .. }));
    assert_eq!(
        package.opc().unwrap().get_part(&slide_uri).unwrap().blob(),
        before.as_slice()
    );
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
