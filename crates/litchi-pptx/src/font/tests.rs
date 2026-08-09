#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Focused semantic and package regression tests for embedded fonts.

use super::codec::*;
use super::model::*;
use super::package::*;
use super::*;
use crate::error::Error;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::sync::Arc;

fn root() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}
fn package(conformance: Conformance) -> OpcPackage {
    let mut package = OpcPackage::new();
    let uri = PackURI::new("/ppt/presentation.xml").unwrap();
    let xml = format!(
        "<p:presentation xmlns:p=\"{}\"><p:sldMasterIdLst/><p:defaultTextStyle/></p:presentation>",
        conformance.pml()
    );
    package.add_part(Box::new(BlobPart::new(
        uri,
        PRESENTATION_CT.into(),
        xml.into_bytes(),
    )));
    let office_rel = match conformance {
        Conformance::Transitional => {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        },
        Conformance::Strict => {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
        },
    };
    package.rels_mut().add_relationship(
        office_rel.into(),
        "ppt/presentation.xml".into(),
        "rId1".into(),
        false,
    );
    package
}
fn eot(marker: u8) -> Vec<u8> {
    let mut value = vec![0; 96];
    value[0..4].copy_from_slice(&108u32.to_le_bytes());
    value[4..8].copy_from_slice(&12u32.to_le_bytes());
    value[8..12].copy_from_slice(&0x0001_0000u32.to_le_bytes());
    value[16] = marker;
    value[34..36].copy_from_slice(&0x504Cu16.to_le_bytes());
    value.extend_from_slice(&[0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    value
}
fn value() -> RawFonts {
    RawFonts {
        fonts: vec![RawFont {
            has_descriptor: true,
            typeface: "A&B".into(),
            panose: Some(Panose::new([2, 11, 6, 4, 2, 2, 2, 2, 2, 4])),
            pitch_family: Some(PitchFamily::new(Pitch::Variable, Family::Swiss)),
            charset: Some(Charset::ANSI),
            faces: vec![
                RawFace {
                    style: Style::Regular,
                    relationship_id: "rIdFont1".into(),
                    resource: Some(RawResource {
                        part_name: "/ppt/fonts/font1.fntdata".into(),
                        content_type: FONT_DATA_CT.into(),
                        data: Arc::new(vec![0, 1, 2, 3]),
                    }),
                },
                RawFace {
                    style: Style::BoldItalic,
                    relationship_id: "rIdFont2".into(),
                    resource: Some(RawResource {
                        part_name: "/ppt/fonts/font2.fntdata".into(),
                        content_type: FONT_DATA_CT.into(),
                        data: Arc::new(vec![4, 5, 6]),
                    }),
                },
            ],
        }],
    }
}

#[test]
fn strict_xml_round_trip_and_mce_fallback() {
    let expected = value();
    let fragment = write_raw(&expected, Conformance::Strict).unwrap();
    let xml = [
        format!("<p:presentation xmlns:p=\"{STRICT_PML}\">").as_bytes(),
        fragment.as_slice(),
        b"</p:presentation>",
    ]
    .concat();
    let parsed = parse_raw(&xml).unwrap().unwrap();
    assert_eq!(parsed.fonts[0].typeface, "A&B");
    assert!(
        parsed.fonts[0]
            .faces
            .iter()
            .all(|face| face.resource.is_none())
    );
    let mce = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback><p:embeddedFontLst><p:embeddedFont><p:font typeface="Fallback"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></mc:Fallback></mc:AlternateContent></p:presentation>"#
    );
    assert_eq!(
        parse_raw(mce.as_bytes()).unwrap().unwrap().fonts[0].typeface,
        "Fallback"
    );
}

#[test]
fn loads_libreoffice_and_poi_reference_packages() {
    let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
        root().join("test-data/libreoffice-core/sd/qa/unit/data/BoldonseFontEmbedded.pptx"),
    )
    .unwrap();
    let mut libreoffice = package(Conformance::Transitional);
    let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
    libreoffice
        .get_part_mut(&presentation_uri)
        .unwrap()
        .set_blob(physical.blob_for(&presentation_uri).unwrap());
    let font_uri = PackURI::new("/ppt/fonts/font1.fntdata").unwrap();
    let font_data = physical.blob_for(&font_uri).unwrap();
    assert!(Data::powerpoint(font_data.clone()).is_ok());
    libreoffice.add_part(Box::new(BlobPart::new(
        font_uri.clone(),
        FONT_DATA_CT.into(),
        font_data,
    )));
    libreoffice
        .get_part_mut(&presentation_uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            FONT_REL.into(),
            "fonts/font1.fntdata".into(),
            "rId3".into(),
            false,
        );
    let fonts = load_raw(&libreoffice).unwrap().unwrap();
    assert_eq!(fonts.fonts[0].typeface, "Boldonse");
    assert_eq!(
        fonts.fonts[0].faces[0]
            .resource
            .as_ref()
            .unwrap()
            .data
            .len(),
        36_187
    );
    let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
        root().join("test-data/poi/test-data/slideshow/placeholder-layout-color.pptx"),
    )
    .unwrap();
    let mut poi = package(Conformance::Transitional);
    let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
    poi.get_part_mut(&presentation_uri)
        .unwrap()
        .set_blob(physical.blob_for(&presentation_uri).unwrap());
    for (index, relationship_id) in (1..=6).zip(["rId4", "rId5", "rId6", "rId7", "rId8", "rId9"]) {
        let uri = PackURI::new(format!("/ppt/fonts/font{index}.fntdata")).unwrap();
        let data = physical.blob_for(&uri).unwrap();
        poi.add_part(Box::new(BlobPart::new(uri, FONT_DATA_CT.into(), data)));
        poi.get_part_mut(&presentation_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                FONT_REL.into(),
                format!("fonts/font{index}.fntdata"),
                relationship_id.into(),
                false,
            );
    }
    let fonts = load_raw(&poi).unwrap().unwrap();
    assert_eq!(fonts.fonts.len(), 3);
    let roboto = fonts
        .fonts
        .iter()
        .find(|font| font.typeface == "Roboto")
        .unwrap();
    assert_eq!(roboto.faces.len(), 4);
    assert!(roboto.faces.iter().all(|face| {
        face.resource
            .as_ref()
            .is_some_and(|resource| !resource.data.is_empty())
    }));
}

#[test]
fn package_writer_round_trips_strict_graph_and_schema_position() {
    let mut package = package(Conformance::Strict);
    let expected = value();
    put_raw(&mut package, &expected, Conformance::Strict).unwrap();
    assert_eq!(load_raw(&package).unwrap().unwrap(), expected);
    let xml = package.main_document_part().unwrap().blob();
    let list = memchr::memmem::find(xml, b"<p:embeddedFontLst").unwrap();
    let defaults = memchr::memmem::find(xml, b"<p:defaultTextStyle").unwrap();
    assert!(list < defaults);
}

#[test]
fn rejects_malformed_xml_duplicates_and_caps() {
    for xml in [
            format!(r#"<p:presentation xmlns:p="{PML}"/>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" panose="12"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:bold xmlns:r="{REL_NS}" r:id="rId1"/><p:regular r:id="rId2"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<!DOCTYPE x><p:presentation xmlns:p="{PML}"/>"#),
        ].into_iter().skip(1) { assert!(parse_raw(xml.as_bytes()).is_err(), "{xml}"); }
    assert!(parse_raw(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
    let face = Face::new(Style::Regular, Data::powerpoint(eot(1)).unwrap());
    let font = Font::from_face("Duplicate", face).unwrap();
    let mut duplicate = Fonts::new();
    duplicate.add(font.clone()).unwrap();
    assert!(duplicate.add(font).is_err());
}

#[test]
fn rejects_external_orphan_and_outbound_graphs() {
    let mut external = package(Conformance::Transitional);
    let xml = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL_NS}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:regular r:id="rIdFont1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
    );
    external
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .set_blob(xml.into_bytes());
    external
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            FONT_REL.into(),
            "https://invalid.example/font".into(),
            "rIdFont1".into(),
            true,
        );
    assert!(load_raw(&external).is_err());

    let mut orphan = package(Conformance::Transitional);
    orphan.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/fonts/orphan.fntdata").unwrap(),
        FONT_DATA_CT.into(),
        vec![1],
    )));
    assert!(load_raw(&orphan).is_err());

    let mut outbound = package(Conformance::Transitional);
    put_raw(&mut outbound, &value(), Conformance::Transitional).unwrap();
    outbound
        .get_part_mut(&PackURI::new("/ppt/fonts/font1.fntdata").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:forbidden".into(),
            "other.bin".into(),
            "rId1".into(),
            false,
        );
    assert!(load_raw(&outbound).is_err());

    let mut root_owned = package(Conformance::Transitional);
    put_raw(&mut root_owned, &value(), Conformance::Transitional).unwrap();
    root_owned.rels_mut().add_relationship(
        "urn:not-a-font-owner".into(),
        "ppt/fonts/font1.fntdata".into(),
        "rIdOther".into(),
        false,
    );
    assert!(load_raw(&root_owned).is_err());
}

#[test]
fn fs_type_is_compact_and_validated() {
    assert!(License::from_fs_type(0).unwrap().installable());
    assert!(
        License::from_fs_type(0x0008)
            .unwrap()
            .restrictions()
            .is_empty()
    );
    let editable = License::from_fs_type(0x0108).unwrap();
    assert_eq!(editable.permission(), Permission::Editable);
    assert!(
        editable
            .restrictions()
            .contains(Restrictions::NO_SUBSETTING)
    );
    assert!(!editable.installable());
    assert!(License::from_fs_type(0x0006).is_err());
    assert!(License::from_fs_type(0x8000).is_err());
}

#[test]
fn typed_metadata_can_be_cleared_without_rebuilding_the_font() {
    let panose = Panose::new([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    let pitch = PitchFamily::new(Pitch::Fixed, Family::Roman);
    let mut font = Font::new("Metadata")
        .unwrap()
        .with_panose(panose)
        .with_pitch_family(pitch)
        .with_charset(Charset::SHIFT_JIS);
    assert_eq!(font.set_panose(None), Some(panose));
    assert_eq!(font.set_pitch_family(None), Some(pitch));
    assert_eq!(font.set_charset(None), Some(Charset::SHIFT_JIS));
    assert_eq!(font.panose(), None);
    assert_eq!(font.pitch_family(), None);
    assert_eq!(font.charset(), None);
}

#[test]
fn fresh_font_containers_are_structurally_checked() {
    assert!(Data::powerpoint(eot(3)).is_ok());
    let mut wrong_size = eot(3);
    wrong_size[0..4].copy_from_slice(&1u32.to_le_bytes());
    assert!(Data::powerpoint(wrong_size).is_err());
    let mut reserved = eot(3);
    reserved[64] = 1;
    assert!(Data::powerpoint(reserved).is_err());
    let sfnt = vec![0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    assert!(Data::standard(sfnt).is_ok());
    assert!(Data::standard(b"not-a-font".to_vec()).is_err());
}

#[test]
fn present_empty_and_loaded_false_flag_round_trip_exactly() {
    let uri = PackURI::new("/ppt/presentation.xml").unwrap();
    let mut empty = package(Conformance::Transitional);
    let xml = format!(
        r#"<p:presentation xmlns:p="{PML}"><p:sldMasterIdLst/><p:embeddedFontLst/><p:defaultTextStyle/></p:presentation>"#
    );
    empty.get_part_mut(&uri).unwrap().set_blob(xml.into_bytes());
    let loaded = load(&empty).unwrap().unwrap();
    assert!(loaded.is_empty());
    let before = empty.get_part(&uri).unwrap().blob().to_vec();
    empty.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(!put(&mut empty, loaded).unwrap());
    assert!(empty.is_signed());
    assert_eq!(empty.get_part(&uri).unwrap().blob(), before);
    assert!(remove(&mut empty).unwrap().is_some());
    assert!(load(&empty).unwrap().is_none());
    assert!(
        memchr::memmem::find(empty.get_part(&uri).unwrap().blob(), b"embeddedFontLst").is_none()
    );

    let mut disabled = package(Conformance::Transitional);
    put_raw(&mut disabled, &value(), Conformance::Transitional).unwrap();
    let xml = patch_embedding_flag(disabled.get_part(&uri).unwrap().blob(), false).unwrap();
    disabled.get_part_mut(&uri).unwrap().set_blob(xml);
    let loaded = load(&disabled).unwrap().unwrap();
    let before = disabled.get_part(&uri).unwrap().blob().to_vec();
    disabled.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(!put(&mut disabled, loaded).unwrap());
    assert!(disabled.is_signed());
    assert_eq!(disabled.get_part(&uri).unwrap().blob(), before);
}

#[test]
fn font_list_crud_edits_only_the_active_direct_mce_branch() {
    let mut package = package(Conformance::Transitional);
    let uri = PackURI::new("/ppt/presentation.xml").unwrap();
    let xml = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:mc="{MCE_NS}" xmlns:x="urn:future"><p:sldMasterIdLst/><mc:AlternateContent><mc:Choice Requires="x"><p:embeddedFontLst><p:embeddedFont><p:font typeface="Inactive"/></p:embeddedFont></p:embeddedFontLst></mc:Choice><mc:Fallback><p:defaultTextStyle/></mc:Fallback></mc:AlternateContent></p:presentation>"#
    );
    package
        .get_part_mut(&uri)
        .unwrap()
        .set_blob(xml.into_bytes());
    let mut fonts = Fonts::new();
    fonts
        .add(
            Font::from_face(
                "Active",
                Face::new(Style::Regular, Data::powerpoint(eot(9)).unwrap()),
            )
            .unwrap(),
        )
        .unwrap();
    assert!(put(&mut package, fonts).unwrap());
    let xml = package.get_part(&uri).unwrap().blob();
    assert!(memchr::memmem::find(xml, b"typeface=\"Inactive\"").is_some());
    assert!(memchr::memmem::find(xml, b"typeface=\"Active\"").is_some());
    let active = memchr::memmem::find(xml, b"typeface=\"Active\"").unwrap();
    let defaults = memchr::memmem::find(xml, b"<p:defaultTextStyle").unwrap();
    assert!(active < defaults);
    assert_eq!(
        load(&package)
            .unwrap()
            .unwrap()
            .get("Active")
            .unwrap()
            .name(),
        "Active"
    );
    assert!(remove(&mut package).unwrap().is_some());
    let xml = package.get_part(&uri).unwrap().blob();
    assert!(memchr::memmem::find(xml, b"typeface=\"Inactive\"").is_some());
    assert!(memchr::memmem::find(xml, b"typeface=\"Active\"").is_none());
    assert!(load(&package).unwrap().is_none());
}

#[test]
fn descriptor_root_order_and_unicode_relationship_ids_are_checked() {
    let face_first = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL_NS}"><p:embeddedFontLst><p:embeddedFont><p:regular r:id="rId1"/><p:font typeface="A"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
    );
    assert!(parse_raw(face_first.as_bytes()).is_err());
    let root_out_of_order = format!(
        r#"<p:presentation xmlns:p="{PML}"><p:defaultTextStyle/><p:embeddedFontLst/></p:presentation>"#
    );
    assert!(parse_raw(root_out_of_order.as_bytes()).is_err());
    let unicode = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL_NS}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:regular r:id="字体"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
    );
    assert_eq!(
        parse_raw(unicode.as_bytes()).unwrap().unwrap().fonts[0].faces[0].relationship_id,
        "字体"
    );
}

#[test]
fn rejected_collection_edits_leave_indexes_and_order_unchanged() {
    let mut fonts = Fonts::new();
    fonts.add(Font::new("First").unwrap()).unwrap();
    fonts.add(Font::new("Second").unwrap()).unwrap();
    let before = fonts.clone();
    assert!(fonts.reorder(&["First", "First"]).is_err());
    assert_eq!(fonts, before);
    assert!(
        fonts
            .replace("First", Font::new("SECOND").unwrap())
            .is_err()
    );
    assert_eq!(fonts, before);
    assert!(fonts.remove(9_usize).is_err());
    assert_eq!(fonts, before);
}

#[test]
fn generated_crud_allocates_collisions_and_preserves_unknown_xml_atomically() {
    let mut package = package(Conformance::Transitional);
    let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
    let original = package.get_part(&presentation_uri).unwrap().blob();
    let marker = memchr::memmem::find(original, b"<p:defaultTextStyle").unwrap();
    let mut xml = original.to_vec();
    xml.splice(marker..marker, b"<!--font-preserve-->".iter().copied());
    package
        .get_part_mut(&presentation_uri)
        .unwrap()
        .set_blob(xml);

    let generated = Font::from_face(
        "Generated",
        Face::new(Style::Regular, Data::powerpoint(eot(7)).unwrap()),
    )
    .unwrap()
    .with_panose([2, 11, 6, 4, 2, 2, 2, 2, 2, 4])
    .with_pitch_family(PitchFamily::new(Pitch::Variable, Family::Swiss))
    .with_charset(Charset::ANSI);
    let mut fonts = Fonts::new();
    fonts.add(generated).unwrap();
    assert!(put(&mut package, fonts).unwrap());
    let loaded = load(&package).unwrap().unwrap();
    let found = loaded.get("generated").unwrap();
    assert_eq!(found.faces()[0].data().bytes(), eot(7));
    assert!(package.contains_part(&PackURI::new("/ppt/fonts/font1.fntdata").unwrap()));
    assert!(
        package
            .get_part(&presentation_uri)
            .unwrap()
            .blob()
            .windows(b"<!--font-preserve-->".len())
            .any(|window| window == b"<!--font-preserve-->")
    );
    assert!(embedding_enabled(package.get_part(&presentation_uri).unwrap().blob()).unwrap());

    let before = package.get_part(&presentation_uri).unwrap().blob().to_vec();
    let parts = package.part_count();
    let mut duplicate = loaded.clone();
    assert!(duplicate.add(found.clone()).is_err());
    assert_eq!(package.get_part(&presentation_uri).unwrap().blob(), before);
    assert_eq!(package.part_count(), parts);
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());
    assert!(!put(&mut package, loaded.clone()).unwrap());
    assert!(package.is_signed());
    assert_eq!(package.get_part(&presentation_uri).unwrap().blob(), before);

    let mut changed = loaded;
    let mut replacement = changed
        .remove("Generated")
        .unwrap()
        .with_charset(Charset::DEFAULT);
    replacement.rename("Renamed").unwrap();
    changed.add(replacement).unwrap();
    assert!(put(&mut package, changed).unwrap());
    assert!(!package.is_signed());
    assert_eq!(
        load(&package)
            .unwrap()
            .unwrap()
            .get("renamed")
            .unwrap()
            .charset(),
        Some(Charset::DEFAULT)
    );
    assert!(remove(&mut package).unwrap().is_some());
    assert!(load(&package).unwrap().is_none());
    assert!(!embedding_enabled(package.get_part(&presentation_uri).unwrap().blob()).unwrap());
}

#[test]
fn pitch_charset_zero_face_and_mixed_dialects_are_checked() {
    for value in [
        0_u8, 1, 2, 16, 17, 18, 32, 33, 34, 48, 49, 50, 64, 65, 66, 80, 81, 82,
    ] {
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" pitchFamily="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        assert!(parse_raw(xml.as_bytes()).is_ok(), "pitchFamily={value}");
    }
    for value in [3_u8, 15, 19, 31, 35, 255] {
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" pitchFamily="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        assert!(parse_raw(xml.as_bytes()).is_err(), "pitchFamily={value}");
    }
    for value in ["-128", "127"] {
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" charset="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        assert!(parse_raw(xml.as_bytes()).is_ok(), "charset={value}");
    }
    for value in ["-129", "128"] {
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" charset="{value}"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        assert!(parse_raw(xml.as_bytes()).is_err(), "charset={value}");
    }
    let mixed = format!(
        r#"<p:presentation xmlns:p="{STRICT_PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
    );
    assert!(parse_raw(mixed.as_bytes()).is_err());
}

#[test]
fn malformed_unicode_duplicates_remain_numerically_repairable() {
    let mut package = package(Conformance::Transitional);
    let uri = PackURI::new("/ppt/presentation.xml").unwrap();
    let xml = format!(
        r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="Straße"/></p:embeddedFont><p:embeddedFont><p:font typeface="STRASSE"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
    );
    package
        .get_part_mut(&uri)
        .unwrap()
        .set_blob(xml.into_bytes());
    let mut fonts = load(&package).unwrap().unwrap();
    assert!(matches!(
        fonts.get("strasse"),
        Err(Error::AmbiguousFontName { matches: 2, .. })
    ));
    assert_eq!(fonts.get(0_usize).unwrap().name(), "Straße");
    fonts.remove(1_usize).unwrap();
    assert_eq!(fonts.get("STRASSE").unwrap().name(), "Straße");

    let mut authored = Fonts::new();
    authored.add(Font::new("é").unwrap()).unwrap();
    assert!(authored.add(Font::new("e\u{301}").unwrap()).is_err());
}

#[test]
fn noncanonical_targets_and_every_main_profile_round_trip() {
    let mut noncanonical = package(Conformance::Transitional);
    let mut value = value();
    for face in &mut value.fonts[0].faces {
        if let Some(resource) = &mut face.resource {
            resource.part_name = format!("/custom/{}.bin", face.style.element());
        }
    }
    put_raw(&mut noncanonical, &value, Conformance::Transitional).unwrap();
    assert_eq!(load_raw(&noncanonical).unwrap().unwrap(), value);

    for content_type in [
        ct::PML_PRESENTATION_MAIN,
        ct::PML_SLIDESHOW_MAIN,
        ct::PML_TEMPLATE_MAIN,
        ct::PML_PRES_MACRO_MAIN,
        ct::PML_SLIDESHOW_MACRO_MAIN,
        ct::PML_TEMPLATE_MACRO_MAIN,
    ] {
        let mut package = package(Conformance::Transitional);
        let uri = PackURI::new("/ppt/presentation.xml").unwrap();
        package
            .get_part_mut(&uri)
            .unwrap()
            .set_content_type(content_type.into())
            .unwrap();
        let mut fonts = Fonts::new();
        fonts
            .add(
                Font::from_face(
                    "Profile",
                    Face::new(Style::Regular, Data::powerpoint(eot(3)).unwrap()),
                )
                .unwrap(),
            )
            .unwrap();
        assert!(put(&mut package, fonts).unwrap(), "{content_type}");
        assert_eq!(load(&package).unwrap().unwrap().len(), 1, "{content_type}");
    }
}

#[test]
fn shared_font_parts_survive_face_removal_and_reject_other_owners() {
    let mut package = package(Conformance::Transitional);
    let shared = RawResource {
        part_name: "/ppt/fonts/shared.fntdata".into(),
        content_type: FONT_DATA_CT.into(),
        data: Arc::new(vec![3; 64]),
    };
    let graph = RawFonts {
        fonts: vec![
            RawFont {
                has_descriptor: true,
                typeface: "First".into(),
                panose: None,
                pitch_family: None,
                charset: None,
                faces: vec![RawFace {
                    style: Style::Regular,
                    relationship_id: "rIdFontA".into(),
                    resource: Some(shared.clone()),
                }],
            },
            RawFont {
                has_descriptor: true,
                typeface: "Second".into(),
                panose: None,
                pitch_family: None,
                charset: None,
                faces: vec![RawFace {
                    style: Style::Regular,
                    relationship_id: "rIdFontA".into(),
                    resource: Some(shared),
                }],
            },
        ],
    };
    put_raw(&mut package, &graph, Conformance::Transitional).unwrap();
    assert_eq!(load_raw(&package).unwrap().unwrap().fonts.len(), 2);
    let mut fonts = load(&package).unwrap().unwrap();
    fonts.remove("First").unwrap();
    put(&mut package, fonts).unwrap();
    let font_uri = PackURI::new("/ppt/fonts/shared.fntdata").unwrap();
    assert!(package.contains_part(&font_uri));

    let owner_uri = PackURI::new("/ppt/unknown-owner.bin").unwrap();
    let mut owner = BlobPart::new(
        owner_uri.clone(),
        "application/octet-stream".into(),
        vec![1],
    );
    owner.rels_mut().add_relationship(
        "urn:shared-resource".into(),
        "fonts/shared.fntdata".into(),
        "rIdShared".into(),
        false,
    );
    package.add_part(Box::new(owner));
    assert!(matches!(load(&package), Err(Error::Invalid(_))));
    assert!(package.contains_part(&font_uri));
}
