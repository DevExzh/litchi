use super::*;
use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::package_metadata::{PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE};
use crate::protobuf::tsd;
use crate::protobuf::tsp::{ComponentInfo, ObjectUuidMapEntry, PackageMetadata, Reference, Uuid};
use crate::protobuf::tswp::{
    ObjectAttributeTable, StorageArchive, object_attribute_table::ObjectAttribute,
};
use crate::shapes::{DrawablePoint, DrawableSize};

#[test]
fn pages_native_discriminants_are_typed_and_lossless() {
    for (raw, value) in [
        (0, PagesSectionStart::NextPage),
        (1, PagesSectionStart::RightPage),
        (2, PagesSectionStart::LeftPage),
        (7, PagesSectionStart::Unknown(7)),
    ] {
        assert_eq!(PagesSectionStart::from_raw(raw), value);
        assert_eq!(value.as_raw(), raw);
    }
    for (raw, value) in [
        (0, PagesSectionPageNumbering::ContinueFromPrevious),
        (1, PagesSectionPageNumbering::Restart),
        (3, PagesSectionPageNumbering::Unknown(3)),
    ] {
        assert_eq!(PagesSectionPageNumbering::from_raw(raw), value);
        assert_eq!(value.as_raw(), raw);
    }
    for (raw, value) in [
        (0, PagesPageOrientation::Portrait),
        (1, PagesPageOrientation::Landscape),
        (9, PagesPageOrientation::Unknown(9)),
    ] {
        assert_eq!(PagesPageOrientation::from_raw(raw), value);
        assert_eq!(value.as_raw(), raw);
    }
    assert!(PagesPageNumber::new(0).is_err());
    assert_eq!(PagesPageNumber::new(42).unwrap().get(), 42);
}

#[test]
fn semantic_body_update_and_clear_are_transactional() {
    let mut editor = PagesEditor::from_package(test_package("A🚀B")).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.replace_body_text(2..3, "x").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    editor.replace_body_text(1..3, "東京").unwrap();
    assert_eq!(editor.body_text().unwrap(), "A東京B");
    editor.clear_body().unwrap();
    assert_eq!(editor.body_text().unwrap(), "");
}

#[test]
fn page_layout_crud_preserves_unknown_wire_and_restores_exact_bytes() {
    let mut package = test_package("Body");
    let unknown = {
        let mut field = crate::varint::encode_varint(99 << 3);
        field.extend(crate::varint::encode_varint(999));
        field
    };
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(1).unwrap();
            let mut message = object.messages[0].clone();
            message.data.extend_from_slice(&unknown);
            object.replace_message(0, message).map(|_| ())
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let original_payload = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .objects[0]
        .messages[0]
        .data
        .clone();

    let layout = PagesPageLayout {
        page_width: Some(612.0),
        page_height: Some(792.0),
        left_margin: Some(72.0),
        right_margin: Some(72.0),
        top_margin: Some(54.0),
        bottom_margin: Some(54.0),
        header_margin: Some(24.0),
        footer_margin: Some(24.0),
        page_scale: Some(1.0),
        orientation: Some(PagesPageOrientation::Landscape),
        lays_out_body_vertically: Some(false),
    };
    editor.set_page_layout(layout.clone()).unwrap();
    assert_eq!(editor.page_layout().unwrap(), layout);
    let updated = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .objects[0]
        .messages[0]
        .data
        .clone();
    assert!(updated.starts_with(&original_payload));
    assert_eq!(
        updated
            .windows(unknown.len())
            .filter(|window| *window == unknown)
            .count(),
        1
    );

    let mut unknown_orientation = layout.clone();
    unknown_orientation.orientation = Some(PagesPageOrientation::Unknown(9));
    editor.set_page_layout(unknown_orientation.clone()).unwrap();
    assert_eq!(editor.page_layout().unwrap(), unknown_orientation);
    let before_invalid = editor.to_bytes().unwrap();
    unknown_orientation.orientation = Some(PagesPageOrientation::Unknown(0));
    assert!(editor.set_page_layout(unknown_orientation).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    editor
        .set_page_layout(PagesPageLayout::from(
            &DocumentArchive::decode(original_payload.as_slice()).unwrap(),
        ))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn page_layout_rejects_duplicate_scalar_fields_transactionally() {
    let mut package = test_package("Body");
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(1).unwrap();
            let mut message = object.messages[0].clone();
            for width in [612.0_f32, 640.0] {
                message
                    .data
                    .extend(crate::varint::encode_varint((30 << 3) | 5));
                message.data.extend(width.to_bits().to_le_bytes());
            }
            object.replace_message(0, message).map(|_| ())
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    let mut layout = editor.page_layout().unwrap();
    layout.page_width = Some(700.0);
    assert!(editor.set_page_layout(layout).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn document_options_crud_is_lossless_transactional_and_wire_exact() {
    let baseline_settings = tp::SettingsArchive {
        language: Some("en".to_owned()),
        hyphenation_language: Some(String::new()),
        footnote_kind: Some(0),
        footnote_format: Some(0),
        footnote_numbering: Some(0),
        footnote_gap: Some(10),
        ..Default::default()
    };
    let mut baseline_data = baseline_settings.encode_to_vec();
    let unknown = append_unknown_varint(&mut baseline_data, 99, 990);
    let mut editor = PagesEditor::from_package(test_package_with_settings(baseline_data)).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert_eq!(
        editor.document_options().unwrap(),
        PagesDocumentOptions::default()
    );

    let options = PagesDocumentOptions {
        body_enabled: Some(true),
        headers_enabled: Some(false),
        footers_enabled: Some(true),
        facing_pages: Some(true),
        automatic_hyphenation: Some(true),
        ligatures_enabled: Some(false),
    };
    editor.set_document_options(options).unwrap();
    assert_eq!(editor.document_options().unwrap(), options);
    assert!(options.body_is_enabled());
    assert!(!options.headers_are_enabled());
    assert!(options.footers_are_enabled());
    assert!(options.uses_facing_pages());
    assert!(options.uses_automatic_hyphenation());
    assert!(!options.uses_ligatures());

    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let payload = &archive.object(43).unwrap().messages[0].data;
    let settings = tp::SettingsArchive::decode(payload.as_slice()).unwrap();
    assert_eq!(settings.language.as_deref(), Some("en"));
    assert_eq!(settings.footnote_gap, Some(10));
    assert_eq!(
        payload
            .windows(unknown.len())
            .filter(|window| *window == unknown)
            .count(),
        1
    );
    let changed = editor.to_bytes().unwrap();
    let reparsed = PagesEditor::from_bytes(&changed).unwrap();
    assert_eq!(reparsed.document_options().unwrap(), options);
    editor.set_document_options(options).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), changed);

    editor
        .set_document_options(PagesDocumentOptions::default())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn document_options_update_and_restore_native_presence_exactly() {
    let native = PagesDocumentOptions {
        body_enabled: Some(true),
        headers_enabled: Some(true),
        footers_enabled: Some(true),
        facing_pages: Some(false),
        automatic_hyphenation: Some(false),
        ligatures_enabled: Some(true),
    };
    let mut data = tp::SettingsArchive {
        body: native.body_enabled,
        headers: native.headers_enabled,
        footers: native.footers_enabled,
        facing_pages: native.facing_pages,
        hyphenation: native.automatic_hyphenation,
        use_ligatures: native.ligatures_enabled,
        language: Some("en".to_owned()),
        ..Default::default()
    }
    .encode_to_vec();
    append_unknown_varint(&mut data, 99, 990);
    let mut editor = PagesEditor::from_package(test_package_with_settings(data)).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert_eq!(editor.document_options().unwrap(), native);

    let changed = PagesDocumentOptions {
        headers_enabled: Some(false),
        footers_enabled: Some(false),
        facing_pages: Some(true),
        automatic_hyphenation: Some(true),
        ligatures_enabled: Some(false),
        ..native
    };
    editor.set_document_options(changed).unwrap();
    assert_eq!(editor.document_options().unwrap(), changed);
    editor.set_document_options(native).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn document_options_reject_malformed_wire_transactionally() {
    for malformed in [vec![(1, 0), (1, 1)], vec![(9, 2)], vec![(34, 0), (34, 1)]] {
        let mut data = tp::SettingsArchive::default().encode_to_vec();
        for (field, value) in malformed {
            append_unknown_varint(&mut data, field, value);
        }
        let mut editor = PagesEditor::from_package(test_package_with_settings(data)).unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(editor.document_options().is_err());
        assert!(
            editor
                .set_document_options(PagesDocumentOptions::default())
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    let mut invalid_payload =
        PagesEditor::from_package(test_package_with_settings(vec![0xff])).unwrap();
    let before = invalid_payload.to_bytes().unwrap();
    assert!(invalid_payload.document_options().is_err());
    assert!(
        invalid_payload
            .set_document_options(PagesDocumentOptions::default())
            .is_err()
    );
    assert_eq!(invalid_payload.to_bytes().unwrap(), before);

    let mut duplicate_reference_package =
        test_package_with_settings(tp::SettingsArchive::default().encode_to_vec());
    duplicate_reference_package
        .update_archive(DOCUMENT_ARCHIVE_NAME, |archive| {
            let object = archive.object_mut(DOCUMENT_OBJECT_ID).unwrap();
            let message = object.messages[0].clone();
            let data = append_repeated_length_delimited_field(
                &message.data,
                7,
                &reference(43).encode_to_vec(),
            )?;
            object.replace_message(
                0,
                RawMessage {
                    type_: message.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut duplicate_reference = PagesEditor::from_package(duplicate_reference_package).unwrap();
    let before = duplicate_reference.to_bytes().unwrap();
    assert!(duplicate_reference.document_options().is_err());
    assert!(
        duplicate_reference
            .set_document_options(PagesDocumentOptions::default())
            .is_err()
    );
    assert_eq!(duplicate_reference.to_bytes().unwrap(), before);

    let mut editor = PagesEditor::from_package(test_package("Body")).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.document_options().is_err());
    assert!(
        editor
            .set_document_options(PagesDocumentOptions::default())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn section_settings_crud_is_lossless_validated_and_transactional() {
    let body_id = 42;
    let section_id = 43;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        section: Some(reference(section_id)),
        ..Default::default()
    };
    let body = StorageArchive {
        text: vec!["Body".to_owned()],
        ..Default::default()
    };
    let mut fill_payload = crate::varint::encode_varint(99 << 3);
    fill_payload.extend(crate::varint::encode_varint(7));
    let mut section_data = SectionArchive {
        inherit_previous_header_footer: Some(true),
        section_template_first_page_different: Some(false),
        section_template_even_odd_pages_different: Some(false),
        section_start_kind: Some(PagesSectionStart::NextPage.as_raw()),
        section_page_number_kind: Some(PagesSectionPageNumbering::ContinueFromPrevious.as_raw()),
        section_page_number_start: Some(PagesPageNumber::new(1).unwrap().get()),
        name: Some("Blank".to_owned()),
        section_template_first_page_hides_header_footer: Some(false),
        background_fill: Some(tsd::FillArchive::default()),
        ..Default::default()
    }
    .encode_to_vec();
    section_data =
        patch_length_delimited_field(&section_data, 30, true, Some(&fill_payload)).unwrap();
    let mut unknown_section_field = crate::varint::encode_varint(101 << 3);
    unknown_section_field.extend(crate::varint::encode_varint(999));
    section_data.extend_from_slice(&unknown_section_field);

    let objects = vec![
        object(1, 10000, root.encode_to_vec()),
        object(body_id, 2001, body.encode_to_vec()),
        object(section_id, SECTION_MESSAGE_TYPE, section_data),
    ];
    let mut package = IWorkPackage::new();
    package
        .replace_archive("Index/Document.iwa", &Archive { objects })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let original = PagesSectionSettings {
        name: Some("Blank".to_owned()),
        inherit_previous_header_footer: Some(true),
        first_page_different: Some(false),
        even_odd_pages_different: Some(false),
        start: Some(PagesSectionStart::NextPage),
        page_numbering: Some(PagesSectionPageNumbering::ContinueFromPrevious),
        starting_page_number: Some(PagesPageNumber::new(1).unwrap()),
        first_page_hides_header_footer: Some(false),
        background_fill_payload: Some(fill_payload),
    };
    assert_eq!(editor.section_settings(section_id).unwrap(), original);
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_section_settings(section_id, original.clone())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let mut updated = original.clone();
    updated.name = Some("Chapter Two".to_owned());
    updated.inherit_previous_header_footer = Some(false);
    updated.first_page_different = Some(true);
    updated.even_odd_pages_different = Some(true);
    updated.start = Some(PagesSectionStart::LeftPage);
    updated.page_numbering = Some(PagesSectionPageNumbering::Restart);
    updated.starting_page_number = Some(PagesPageNumber::new(42).unwrap());
    updated.first_page_hides_header_footer = Some(true);
    editor
        .set_section_settings(section_id, updated.clone())
        .unwrap();
    assert_eq!(editor.section_settings(section_id).unwrap(), updated);
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let section_payload = &archive.object(section_id).unwrap().messages[0].data;
    let section = SectionArchive::decode(section_payload.as_slice()).unwrap();
    assert_eq!(section.section_start_kind, Some(2));
    assert_eq!(section.section_page_number_kind, Some(1));
    assert_eq!(section.section_page_number_start, Some(42));
    assert!(
        section_payload
            .windows(unknown_section_field.len())
            .any(|window| window == unknown_section_field)
    );
    let reparsed = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reparsed.section_settings(section_id).unwrap(), updated);

    updated.start = Some(PagesSectionStart::Unknown(7));
    updated.page_numbering = Some(PagesSectionPageNumbering::Unknown(3));
    editor
        .set_section_settings(section_id, updated.clone())
        .unwrap();
    assert_eq!(editor.section_settings(section_id).unwrap(), updated);
    let section = SectionArchive::decode(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .object(section_id)
            .unwrap()
            .messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    assert_eq!(section.section_start_kind, Some(7));
    assert_eq!(section.section_page_number_kind, Some(3));

    editor
        .set_section_settings(section_id, original.clone())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    for invalid in [
        PagesSectionSettings {
            name: Some("bad\0name".to_owned()),
            ..original.clone()
        },
        PagesSectionSettings {
            background_fill_payload: Some(vec![0xff]),
            ..original.clone()
        },
        PagesSectionSettings {
            start: Some(PagesSectionStart::Unknown(0)),
            ..original.clone()
        },
        PagesSectionSettings {
            page_numbering: Some(PagesSectionPageNumbering::Unknown(1)),
            ..original.clone()
        },
    ] {
        assert!(editor.set_section_settings(section_id, invalid).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }
    assert!(editor.section_settings(999).is_err());
    assert!(editor.set_section_settings(999, original.clone()).is_err());
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let mut malformed = editor.package().clone();
    malformed
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(section_id).unwrap();
            let mut message = object.messages[0].clone();
            message.data.extend(crate::varint::encode_varint(17 << 3));
            message.data.push(0);
            object.replace_message(0, message).map(|_| ())
        })
        .unwrap();
    let mut malformed = PagesEditor::from_package(malformed).unwrap();
    let malformed_baseline = malformed.to_bytes().unwrap();
    assert!(malformed.section_settings(section_id).is_err());
    assert!(
        malformed
            .set_section_settings(section_id, original)
            .is_err()
    );
    assert_eq!(malformed.to_bytes().unwrap(), malformed_baseline);
}

#[test]
fn section_settings_reject_zero_starting_page_number_transactionally() {
    let body_id = 42;
    let section_id = 43;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        section: Some(reference(section_id)),
        ..Default::default()
    };
    let section = SectionArchive {
        section_start_kind: Some(PagesSectionStart::NextPage.as_raw()),
        section_page_number_kind: Some(PagesSectionPageNumbering::Restart.as_raw()),
        section_page_number_start: Some(0),
        ..Default::default()
    };
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    object(1, 10000, root.encode_to_vec()),
                    object(
                        body_id,
                        2001,
                        StorageArchive {
                            text: vec!["Body".to_owned()],
                            ..Default::default()
                        }
                        .encode_to_vec(),
                    ),
                    object(section_id, SECTION_MESSAGE_TYPE, section.encode_to_vec()),
                ],
            },
        )
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.section_settings(section_id).is_err());
    assert!(
        editor
            .set_section_settings(section_id, PagesSectionSettings::default())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn solid_section_background_crud_preserves_nested_unknown_wire() {
    let body_id = 42;
    let section_id = 43;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        section: Some(reference(section_id)),
        ..Default::default()
    };
    let body = StorageArchive {
        text: vec!["Body".to_owned()],
        ..Default::default()
    };
    let original_color = PagesRgbaColor {
        red: 1.0,
        green: 0.588_738_74,
        blue: 0.552_926_2,
        alpha: 1.0,
        color_space: PagesRgbColorSpace::Srgb,
    };
    let mut color_payload = tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(original_color.red),
        g: Some(original_color.green),
        b: Some(original_color.blue),
        rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
        a: Some(original_color.alpha),
        ..Default::default()
    }
    .encode_to_vec();
    let mut unknown_color_field = crate::varint::encode_varint(99 << 3);
    unknown_color_field.extend(crate::varint::encode_varint(123));
    color_payload.extend_from_slice(&unknown_color_field);
    let mut fill_payload = tsd::FillArchive {
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            ..Default::default()
        }),
        ..Default::default()
    }
    .encode_to_vec();
    fill_payload =
        patch_length_delimited_field(&fill_payload, 1, true, Some(&color_payload)).unwrap();
    let mut unknown_fill_field = crate::varint::encode_varint(100 << 3);
    unknown_fill_field.extend(crate::varint::encode_varint(456));
    fill_payload.extend_from_slice(&unknown_fill_field);
    let mut section_data = SectionArchive {
        name: Some("Blank".to_owned()),
        background_fill: Some(tsd::FillArchive::default()),
        ..Default::default()
    }
    .encode_to_vec();
    section_data =
        patch_length_delimited_field(&section_data, 30, true, Some(&fill_payload)).unwrap();
    let objects = vec![
        object(1, 10000, root.encode_to_vec()),
        object(body_id, 2001, body.encode_to_vec()),
        object(section_id, SECTION_MESSAGE_TYPE, section_data),
    ];
    let mut package = IWorkPackage::new();
    package
        .replace_archive("Index/Document.iwa", &Archive { objects })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    assert_eq!(
        editor.section_background(section_id).unwrap(),
        PagesSectionBackground::Solid(original_color)
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_section_background(section_id, PagesSectionBackground::Solid(original_color))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let updated = PagesRgbaColor {
        red: 0.125,
        green: original_color.green,
        blue: 0.75,
        alpha: 0.5,
        color_space: PagesRgbColorSpace::DisplayP3,
    };
    editor
        .set_section_background(section_id, PagesSectionBackground::Solid(updated))
        .unwrap();
    assert_eq!(
        editor.section_background(section_id).unwrap(),
        PagesSectionBackground::Solid(updated)
    );
    let payload = editor
        .section_settings(section_id)
        .unwrap()
        .background_fill_payload
        .unwrap();
    for unknown in [&unknown_color_field, &unknown_fill_field] {
        assert!(
            payload
                .windows(unknown.len())
                .any(|window| window == unknown.as_slice())
        );
    }
    editor
        .set_section_background(section_id, PagesSectionBackground::Solid(original_color))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for invalid in [
        PagesRgbaColor {
            red: f32::NAN,
            ..original_color
        },
        PagesRgbaColor {
            alpha: 1.01,
            ..original_color
        },
    ] {
        assert!(
            editor
                .set_section_background(section_id, PagesSectionBackground::Solid(invalid))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }
    assert!(
        editor
            .set_section_background(section_id, PagesSectionBackground::Opaque(vec![0xff]),)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_section_background(section_id, PagesSectionBackground::None)
        .unwrap();
    assert_eq!(
        editor.section_background(section_id).unwrap(),
        PagesSectionBackground::None
    );
    let opaque = tsd::FillArchive {
        gradient: Some(tsd::GradientArchive::default()),
        ..Default::default()
    }
    .encode_to_vec();
    editor
        .set_section_background(section_id, PagesSectionBackground::Opaque(opaque.clone()))
        .unwrap();
    assert_eq!(
        editor.section_background(section_id).unwrap(),
        PagesSectionBackground::Opaque(opaque)
    );
}

#[test]
fn reachable_header_footer_crud_is_typed_and_transactional() {
    let body_id = 42;
    let section_id = 43;
    let template_id = 44;
    let header_id = 45;
    let footer_id = 46;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        ..Default::default()
    };
    let body = StorageArchive {
        text: vec!["Body".to_owned()],
        table_section: Some(ObjectAttributeTable {
            entries: vec![ObjectAttribute {
                character_index: 0,
                object: Some(reference(section_id)),
            }],
        }),
        ..Default::default()
    };
    let section = SectionArchive {
        name: Some("Chapter".to_owned()),
        odd_section_template_page: Some(reference(template_id)),
        ..Default::default()
    };
    let template = SectionTemplateArchive {
        headers: vec![reference(header_id)],
        footers: vec![reference(footer_id)],
        ..Default::default()
    };
    let storage = |text: &str| StorageArchive {
        kind: Some(1),
        text: vec![text.to_owned()],
        ..Default::default()
    };
    let objects = vec![
        object(1, 10000, root.encode_to_vec()),
        object(body_id, 2001, body.encode_to_vec()),
        object(section_id, SECTION_MESSAGE_TYPE, section.encode_to_vec()),
        object(
            template_id,
            SECTION_TEMPLATE_MESSAGE_TYPE,
            template.encode_to_vec(),
        ),
        object(header_id, 2001, storage("A🚀B").encode_to_vec()),
        object(footer_id, 2001, storage("Footer").encode_to_vec()),
    ];
    let mut package = IWorkPackage::new();
    package
        .replace_archive("Index/Document.iwa", &Archive { objects })
        .unwrap();

    let mut editor = PagesEditor::from_package(package).unwrap();
    assert_eq!(editor.sections().len(), 1);
    assert_eq!(editor.sections()[0].name.as_deref(), Some("Chapter"));
    let regions = editor.header_footers().unwrap();
    assert_eq!(regions.len(), 2);
    assert_eq!(regions[0].section_name.as_deref(), Some("Chapter"));
    assert_eq!(regions[0].template, PagesTemplateKind::Odd);
    assert_eq!(regions[0].kind, PagesHeaderFooterKind::Header);
    assert_eq!(regions[0].storage.text, "A🚀B");
    assert_eq!(regions[1].kind, PagesHeaderFooterKind::Footer);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .replace_header_footer_text(header_id, 2..3, "x")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
    editor
        .replace_header_footer_text(header_id, 1..3, "東京")
        .unwrap();
    editor.clear_header_footer(footer_id).unwrap();
    let regions = editor.header_footers().unwrap();
    assert_eq!(regions[0].storage.text, "A東京B");
    assert_eq!(regions[1].storage.text, "");
    assert!(editor.set_header_footer_text(body_id, "no").is_err());
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_section_name(999, Some("no")).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    editor
        .set_section_name(section_id, Some("Renamed"))
        .unwrap();
    assert_eq!(editor.sections()[0].name.as_deref(), Some("Renamed"));
    let mut layout = editor.page_layout().unwrap();
    layout.page_width = Some(612.0);
    layout.page_height = Some(792.0);
    layout.left_margin = Some(72.0);
    layout.right_margin = Some(72.0);
    layout.top_margin = Some(54.0);
    layout.bottom_margin = Some(54.0);
    layout.header_margin = Some(24.0);
    layout.footer_margin = Some(24.0);
    layout.page_scale = Some(1.0);
    let before = editor.to_bytes().unwrap();
    let mut invalid = layout.clone();
    invalid.page_width = Some(f32::NAN);
    assert!(editor.set_page_layout(invalid).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    editor.set_page_layout(layout.clone()).unwrap();
    assert_eq!(editor.page_layout().unwrap(), layout);

    let reparsed = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let regions = reparsed.header_footers().unwrap();
    assert_eq!(regions[0].storage.text, "A東京B");
    assert_eq!(regions[1].storage.text, "");
    assert_eq!(reparsed.sections()[0].name.as_deref(), Some("Renamed"));
    assert_eq!(reparsed.page_layout().unwrap(), layout);
}

#[test]
fn section_append_remove_is_wire_preserving_and_transactional() {
    let body_id = 42;
    let section_id = 43;
    let template_id = 44;
    let header_id = 45;
    let footer_id = 46;
    let guide_map_id = 47;
    let guide_storage_id = 48;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        ..Default::default()
    };
    let body = StorageArchive {
        text: vec!["Body".to_owned()],
        table_section: Some(crate::protobuf::tswp::ObjectAttributeTable {
            entries: vec![ObjectAttribute {
                character_index: 0,
                object: Some(reference(section_id)),
            }],
        }),
        ..Default::default()
    };
    let section = SectionArchive {
        name: Some("Original".to_owned()),
        odd_section_template_page: Some(reference(template_id)),
        user_defined_guide_storage: Some(reference(guide_map_id)),
        ..Default::default()
    };
    let template = SectionTemplateArchive {
        headers: vec![reference(header_id)],
        footers: vec![reference(footer_id)],
        ..Default::default()
    };
    let mut body_data = body.encode_to_vec();
    let body_unknown = append_unknown_varint(&mut body_data, 98, 980);
    let mut section_data = section.encode_to_vec();
    let section_unknown = append_unknown_varint(&mut section_data, 99, 990);
    let mut body_object = object(body_id, 2001, body_data);
    body_object.archive_info.message_infos[0]
        .object_references
        .push(section_id);
    let mut section_object = object(section_id, SECTION_MESSAGE_TYPE, section_data);
    section_object.archive_info.message_infos[0]
        .object_references
        .extend([template_id, guide_map_id]);
    let mut template_object = object(
        template_id,
        SECTION_TEMPLATE_MESSAGE_TYPE,
        template.encode_to_vec(),
    );
    template_object.archive_info.message_infos[0]
        .object_references
        .extend([header_id, footer_id]);
    let storage = |text: &str| StorageArchive {
        kind: Some(1),
        text: vec![text.to_owned()],
        ..Default::default()
    };
    let guide_map = tp::UserDefinedGuideMapArchive {
        user_defined_guide_storages: vec![tp::user_defined_guide_map_archive::UserDefinedGuide {
            page_index: 7,
            guide_storage: reference(guide_storage_id),
        }],
    };
    let guide_storage = crate::protobuf::tsd::GuideStorageArchive {
        user_defined_guides: vec![Default::default()],
    };
    let mut guide_map_object = object(
        guide_map_id,
        USER_DEFINED_GUIDE_MAP_MESSAGE_TYPE,
        guide_map.encode_to_vec(),
    );
    guide_map_object.archive_info.message_infos[0]
        .object_references
        .push(guide_storage_id);
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    object(1, 10000, root.encode_to_vec()),
                    body_object,
                    section_object,
                    template_object,
                    object(header_id, 2001, storage("Header").encode_to_vec()),
                    object(footer_id, 2001, storage("Footer").encode_to_vec()),
                    guide_map_object,
                    object(
                        guide_storage_id,
                        GUIDE_STORAGE_MESSAGE_TYPE,
                        guide_storage.encode_to_vec(),
                    ),
                ],
            },
        )
        .unwrap();
    let before = package.to_bytes().unwrap();

    let mut editor = PagesEditor::from_package(package).unwrap();
    let created = editor.append_section(section_id, "Appended").unwrap();
    assert_eq!(created.character_index, 5);
    assert_eq!(created.name.as_deref(), Some("Appended"));
    assert_ne!(created.odd_template_id, Some(template_id));
    assert_eq!(editor.body_text().unwrap(), "Body\u{4}");
    assert_eq!(editor.sections().len(), 2);
    assert_eq!(editor.header_footers().unwrap().len(), 4);
    let archive = editor.package().archive("Index/Document.iwa").unwrap();
    let created_section = SectionArchive::decode(
        archive.object(created.object_id).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    let created_template = SectionTemplateArchive::decode(
        archive
            .object(created.odd_template_id.unwrap())
            .unwrap()
            .messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    assert_ne!(created_template.headers[0].identifier, header_id);
    assert_ne!(created_template.footers[0].identifier, footer_id);
    let created_guide_map_id = created_section
        .user_defined_guide_storage
        .unwrap()
        .identifier;
    assert_ne!(created_guide_map_id, guide_map_id);
    let created_guide_map = tp::UserDefinedGuideMapArchive::decode(
        archive.object(created_guide_map_id).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    assert_eq!(created_guide_map.user_defined_guide_storages.len(), 1);
    assert_eq!(
        created_guide_map.user_defined_guide_storages[0].page_index,
        0
    );
    let created_guide_storage_id = created_guide_map.user_defined_guide_storages[0]
        .guide_storage
        .identifier;
    assert_ne!(created_guide_storage_id, guide_storage_id);
    let created_guide_storage = crate::protobuf::tsd::GuideStorageArchive::decode(
        archive.object(created_guide_storage_id).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    assert!(created_guide_storage.user_defined_guides.is_empty());
    assert!(
        archive.object(created.object_id).unwrap().messages[0]
            .data
            .ends_with(&section_unknown)
    );
    assert!(
        archive.object(body_id).unwrap().messages[0]
            .data
            .ends_with(&body_unknown)
    );

    let before_middle = editor.to_bytes().unwrap();
    let inserted = editor.insert_section(section_id, 2, "Middle").unwrap();
    assert_eq!(inserted.character_index, 3);
    assert_eq!(editor.body_text().unwrap(), "Bo\u{4}dy\u{4}");
    assert_eq!(
        editor
            .sections()
            .iter()
            .map(|section| section.character_index)
            .collect::<Vec<_>>(),
        [0, 3, 6]
    );
    editor.remove_section(inserted.object_id).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before_middle);

    assert_eq!(editor.section_text(section_id).unwrap(), "Body");
    assert_eq!(editor.section_text(created.object_id).unwrap(), "");
    editor
        .set_section_text(created.object_id, "Appended 🚀")
        .unwrap();
    assert_eq!(editor.body_text().unwrap(), "Body\u{4}Appended 🚀");
    let before_surrogate = editor.to_bytes().unwrap();
    assert!(
        editor
            .replace_section_text(created.object_id, 10..10, "x")
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_surrogate);
    editor
        .replace_section_text(created.object_id, 9..11, "東京")
        .unwrap();
    assert_eq!(
        editor.section_text(created.object_id).unwrap(),
        "Appended 東京"
    );
    editor.clear_section_text(section_id).unwrap();
    assert_eq!(editor.body_text().unwrap(), "\u{4}Appended 東京");
    assert_eq!(editor.sections()[1].character_index, 1);
    editor.set_section_text(section_id, "First").unwrap();
    assert_eq!(editor.body_text().unwrap(), "First\u{4}Appended 東京");
    assert_eq!(editor.sections()[1].character_index, 6);

    let before_rejected_body = editor.to_bytes().unwrap();
    assert!(editor.replace_body_text(0..6, "crossed").is_err());
    assert!(editor.replace_body_text(0..0, "bad\u{4}break").is_err());
    assert!(editor.set_body_text("flattened").is_err());
    assert!(editor.section_text(999).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_rejected_body);

    editor.set_section_text(section_id, "Body").unwrap();
    editor.clear_section_text(created.object_id).unwrap();
    assert_eq!(editor.body_text().unwrap(), "Body\u{4}");

    let before_rejected = editor.to_bytes().unwrap();
    assert!(editor.insert_section(section_id, 0, "Invalid").is_err());
    assert!(editor.insert_section(section_id, 99, "Invalid").is_err());
    assert!(editor.insert_section(section_id, 5, "Invalid").is_err());
    assert!(editor.insert_section(section_id, 4, "Invalid\0").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_rejected);

    let removed = editor.remove_section(created.object_id).unwrap();
    assert_eq!(removed, created);
    assert_eq!(editor.sections().len(), 1);
    assert_eq!(editor.body_text().unwrap(), "Body");
    assert_eq!(editor.to_bytes().unwrap(), before);
    let before_rejected = editor.to_bytes().unwrap();
    assert!(editor.remove_section(section_id).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before_rejected);
}

#[test]
fn reachable_floating_drawable_comment_crud_is_guarded() {
    let body_id = 42;
    let floating_id = 50;
    let drawable_id = 60;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        floating_drawables: Some(reference(floating_id)),
        ..Default::default()
    };
    let placeholder = || tp::PlaceholderArchive {
        super_: crate::protobuf::tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive::default(),
                ..Default::default()
            },
            ..Default::default()
        },
    };
    let floating = tp::FloatingDrawablesArchive {
        page_groups: vec![tp::floating_drawables_archive::PageGroup {
            page_index: 0,
            drawables: vec![tp::floating_drawables_archive::DrawableEntry {
                drawable: Some(reference(drawable_id)),
            }],
            ..Default::default()
        }],
        ..Default::default()
    };
    let mut root_object = object(1, 10000, root.encode_to_vec());
    root_object.archive_info.message_infos[0]
        .object_references
        .extend([body_id, floating_id]);
    let mut floating_object = object(floating_id, 10010, floating.encode_to_vec());
    floating_object.archive_info.message_infos[0]
        .object_references
        .push(drawable_id);
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    root_object,
                    object(
                        body_id,
                        2001,
                        StorageArchive {
                            text: vec!["Body".to_owned()],
                            ..Default::default()
                        }
                        .encode_to_vec(),
                    ),
                    floating_object,
                    object(drawable_id, 7, placeholder().encode_to_vec()),
                    object(61, 7, placeholder().encode_to_vec()),
                ],
            },
        )
        .unwrap();

    let mut editor = PagesEditor::from_package(package).unwrap();
    assert_eq!(
        editor
            .drawables()
            .unwrap()
            .into_iter()
            .map(|drawable| drawable.object_id)
            .collect::<Vec<_>>(),
        vec![drawable_id]
    );
    assert!(editor.set_drawable_comment(61, "Unreachable").is_err());
    editor
        .set_drawable_comment(drawable_id, "Page annotation")
        .unwrap();
    assert_eq!(
        editor
            .drawable_comment(drawable_id)
            .unwrap()
            .unwrap()
            .comment
            .text,
        "Page annotation"
    );
    assert_eq!(editor.body_text().unwrap(), "Body");
    let bytes = editor.to_bytes().unwrap();
    editor
        .set_drawable_comment(drawable_id, "Page annotation")
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), bytes);

    let mut reparsed = PagesEditor::from_bytes(&bytes).unwrap();
    reparsed.clear_drawable_comment(drawable_id).unwrap();
    assert!(reparsed.drawable_comment(drawable_id).unwrap().is_none());
    assert_eq!(reparsed.body_text().unwrap(), "Body");
}

#[test]
fn reachable_drawable_text_crud_covers_placeholders_and_text_boxes() {
    let mut editor = PagesEditor::from_package(floating_text_package()).unwrap();
    let text = editor.drawable_text_storages().unwrap();
    assert_eq!(text.len(), 2);
    assert_eq!(text[0].drawable_object_id, 60);
    assert_eq!(text[0].storage.object_id, 62);
    assert_eq!(text[0].storage.text, "A🚀B");
    assert_eq!(text[1].drawable_object_id, 64);
    assert_eq!(text[1].storage.object_id, 65);
    assert_eq!(text[1].storage.text, "Independent text box");

    let before = editor.to_bytes().unwrap();
    assert!(editor.replace_drawable_text(60, 2..3, "x").is_err());
    assert!(editor.set_drawable_text(61, "detached").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);

    editor.replace_drawable_text(60, 1..3, "東京").unwrap();
    assert_eq!(
        editor.drawable_text_storages().unwrap()[0].storage.text,
        "A東京B"
    );
    editor
        .set_drawable_text(64, "Replacement shape 🚀")
        .unwrap();
    assert_eq!(
        editor.drawable_text_storages().unwrap()[1].storage.text,
        "Replacement shape 🚀"
    );
    editor.clear_drawable_text(64).unwrap();
    assert_eq!(editor.drawable_text_storages().unwrap()[1].storage.text, "");
    assert_eq!(editor.body_text().unwrap(), "Body");
}

#[test]
fn drawable_text_updates_preserve_unknown_wire_and_restore_exactly() {
    let mut package = floating_text_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(65).unwrap();
            let mut message = object.messages[0].clone();
            append_unknown_varint(&mut message.data, 99, 990);
            object.replace_message(0, message).map(|_| ())
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let baseline = editor
        .package()
        .archive("Index/Document.iwa")
        .unwrap()
        .to_bytes()
        .unwrap();
    editor.set_drawable_text(64, "Temporary 東京").unwrap();
    editor
        .set_drawable_text(64, "Independent text box")
        .unwrap();
    assert_eq!(
        editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap(),
        baseline
    );
}

#[test]
fn ambiguous_drawable_text_ownership_fails_transactionally() {
    let mut package = floating_text_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(64).unwrap();
            object.replace_message(
                0,
                RawMessage {
                    type_: SHAPE_INFO_MESSAGE_TYPE,
                    data: crate::protobuf::tswp::ShapeInfoArchive {
                        owned_storage: Some(reference(62)),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.drawable_text_storages().is_err());
    assert!(editor.set_drawable_text(60, "rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn text_box_geometry_updates_are_guarded_and_byte_exact() {
    let mut editor = PagesEditor::from_package(anchored_text_box_package()).unwrap();
    let original = editor.text_box_geometry(64).unwrap();
    assert_eq!(
        original,
        DrawableGeometry {
            position: Some(DrawablePoint { x: 100.0, y: 80.0 }),
            size: Some(DrawableSize {
                width: 200.0,
                height: 60.0,
            }),
            flags: Some(0),
            angle: Some(0.0),
        }
    );
    let baseline = editor.to_bytes().unwrap();
    let changed = DrawableGeometry {
        position: Some(DrawablePoint { x: 125.5, y: 91.25 }),
        size: Some(DrawableSize {
            width: 320.0,
            height: 72.0,
        }),
        flags: Some(3),
        angle: Some(0.25),
    };
    editor.set_text_box_geometry(64, changed).unwrap();
    assert_eq!(editor.text_box_geometry(64).unwrap(), changed);
    assert_eq!(
        editor.drawable_text_storages().unwrap()[0].storage.text,
        "Source"
    );
    editor.set_text_box_geometry(64, original).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    let mut invalid = original;
    invalid.position.as_mut().unwrap().x = f32::NAN;
    assert!(editor.set_text_box_geometry(64, invalid).is_err());
    assert!(editor.text_box_geometry(60).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn text_box_properties_updates_are_guarded_and_byte_exact() {
    let mut editor = PagesEditor::from_package(anchored_text_box_package()).unwrap();
    let original = editor.text_box_properties(64).unwrap();
    assert_eq!(original, DrawableProperties::default());
    let baseline = editor.to_bytes().unwrap();
    let changed = DrawableProperties {
        hyperlink_url: Some("https://example.test/pages-text-box".to_owned()),
        locked: Some(true),
        aspect_ratio_locked: Some(true),
        accessibility_description: Some("Accessible Pages text box ✨".to_owned()),
    };

    editor.set_text_box_properties(64, changed.clone()).unwrap();
    assert_eq!(editor.text_box_properties(64).unwrap(), changed);
    assert_eq!(
        editor.drawable_text_storages().unwrap()[0].storage.text,
        "Source"
    );
    editor.set_text_box_properties(64, original).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_text_box_properties(60, DrawableProperties::default())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn body_anchored_text_box_duplicate_delete_is_independent_and_exact() {
    let mut editor = PagesEditor::from_package(anchored_text_box_package()).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let source = editor.drawable_text_storages().unwrap()[0].clone();
    assert_eq!(source.drawable_object_id, 64);
    assert_eq!(editor.text_box_graph(64).unwrap().anchor_character_index, 6);

    let created = editor.duplicate_text_box(64, 0, "Clone 🚀").unwrap();
    assert_ne!(created.drawable_object_id, source.drawable_object_id);
    assert_ne!(created.storage.object_id, source.storage.object_id);
    assert_eq!(created.storage.text, "Clone 🚀");
    assert_eq!(editor.body_text().unwrap(), "￼Before￼After");
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id,
        2001,
        "TSWP.StorageArchive",
    )
    .unwrap();
    assert_eq!(
        body.table_attachment
            .unwrap()
            .entries
            .into_iter()
            .map(|entry| entry.character_index)
            .collect::<Vec<_>>(),
        [0, 7]
    );
    assert_eq!(
        editor
            .text_box_graph(created.drawable_object_id)
            .unwrap()
            .anchor_character_index,
        0
    );
    assert_eq!(editor.text_box_graph(64).unwrap().anchor_character_index, 7);
    editor
        .set_drawable_text(created.drawable_object_id, "Changed")
        .unwrap();
    assert_eq!(
        editor
            .drawable_text_storages()
            .unwrap()
            .into_iter()
            .find(|item| item.drawable_object_id == 64)
            .unwrap()
            .storage
            .text,
        "Source"
    );
    let removed = editor.remove_text_box(created.drawable_object_id).unwrap();
    assert_eq!(removed.anchor_character_index, 0);
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let removed = editor.remove_text_box(64).unwrap();
    assert_eq!(removed.anchor_character_index, 6);
    assert_eq!(removed.text.storage.text, "Source");
    assert_eq!(editor.body_text().unwrap(), "BeforeAfter");
    assert!(editor.drawable_text_storages().unwrap().is_empty());
}

#[test]
fn text_box_duplicate_delete_tracks_package_highwater_and_document_uuids() {
    let mut package = anchored_text_box_package();
    let uuid_entry = |identifier| ObjectUuidMapEntry {
        identifier,
        uuid: Uuid {
            lower: identifier,
            upper: identifier + 1_000,
        },
    };
    let metadata = PackageMetadata {
        last_object_identifier: 69,
        components: vec![ComponentInfo {
            identifier: 1,
            preferred_locator: "Document".to_owned(),
            object_uuid_map_entries: [64, 65, 68, 69].into_iter().map(uuid_entry).collect(),
            ..Default::default()
        }],
        ..Default::default()
    };
    package
        .replace_archive(
            PACKAGE_METADATA_ENTRY,
            &Archive {
                objects: vec![object(
                    2,
                    PACKAGE_METADATA_MESSAGE_TYPE,
                    metadata.encode_to_vec(),
                )],
            },
        )
        .unwrap();

    let mut editor = PagesEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let created = editor.duplicate_text_box(64, 0, "Metadata clone").unwrap();
    let graph = editor.text_box_graph(created.drawable_object_id).unwrap();
    assert_eq!(graph.object_ids, [70, 71, 72, 73, 74]);
    assert_eq!(graph.uuid_object_ids, [70, 71, 73, 74]);

    let archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
    let metadata =
        PackageMetadata::decode(archive.object(2).unwrap().messages[0].data.as_slice()).unwrap();
    assert_eq!(metadata.last_object_identifier, 74);
    let entries = &metadata.components[0].object_uuid_map_entries;
    assert_eq!(entries.len(), 8);
    assert_eq!(
        entries
            .iter()
            .map(|entry| entry.identifier)
            .collect::<HashSet<_>>(),
        [64, 65, 68, 69, 70, 71, 73, 74].into_iter().collect()
    );
    assert_eq!(
        entries
            .iter()
            .map(|entry| (entry.uuid.lower, entry.uuid.upper))
            .collect::<HashSet<_>>()
            .len(),
        entries.len()
    );

    editor.remove_text_box(created.drawable_object_id).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn text_box_graph_crud_preserves_unknown_fields_and_rejects_external_owners() {
    let mut package = anchored_text_box_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            for identifier in [64, 65, 67, 68, 69] {
                let object = archive.object_mut(identifier).unwrap();
                let mut message = object.messages[0].clone();
                append_unknown_varint(&mut message.data, 99, identifier);
                object.replace_message(0, message)?;
            }
            Ok(())
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();
    let created = editor.duplicate_text_box(64, 11, "Unknowns").unwrap();
    let graph = editor.text_box_graph(created.drawable_object_id).unwrap();
    for identifier in graph.object_ids {
        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let payload = &archive.object(identifier).unwrap().messages[0].data;
        assert!(payload.windows(2).any(|window| window == [0x98, 0x06]));
    }
    editor.remove_text_box(created.drawable_object_id).unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let mut package = anchored_text_box_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let mut external = object(70, 3097, Vec::new());
            external.archive_info.message_infos[0]
                .object_references
                .push(65);
            archive.insert_object(external)
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_text_box(64).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn ambiguous_text_box_anchor_and_zorder_fail_transactionally() {
    let mut package = anchored_text_box_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let body = archive.object_mut(42).unwrap();
            let mut storage = StorageArchive::decode(body.messages[0].data.as_slice()).unwrap();
            storage
                .table_attachment
                .as_mut()
                .unwrap()
                .entries
                .push(ObjectAttribute {
                    character_index: 6,
                    object: Some(reference(67)),
                });
            body.replace_message(
                0,
                RawMessage {
                    type_: 2001,
                    data: storage.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.duplicate_text_box(64, 0, "rejected").is_err());
    assert!(editor.remove_text_box(64).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);

    let mut package = anchored_text_box_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let zorder = archive.object_mut(66).unwrap();
            let mut value =
                tp::DrawablesZOrderArchive::decode(zorder.messages[0].data.as_slice()).unwrap();
            value.drawables.push(reference(64));
            zorder.replace_message(
                0,
                RawMessage {
                    type_: 10015,
                    data: value.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_text_box(64).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[allow(deprecated)]
fn anchored_text_box_package() -> IWorkPackage {
    let body_id = 42;
    let drawable_id = 64;
    let storage_id = 65;
    let zorder_id = 66;
    let attachment_id = 67;
    let title_id = 68;
    let caption_id = 69;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        drawables_zorder: Some(reference(zorder_id)),
        ..Default::default()
    };
    let body = StorageArchive {
        text: vec!["Before￼After".to_owned()],
        table_attachment: Some(crate::protobuf::tswp::ObjectAttributeTable {
            entries: vec![ObjectAttribute {
                character_index: 6,
                object: Some(reference(attachment_id)),
            }],
        }),
        ..Default::default()
    };
    let shape = crate::protobuf::tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
            super_: tsd::DrawableArchive {
                geometry: Some(tsd::GeometryArchive {
                    position: Some(crate::protobuf::tsp::Point { x: 100.0, y: 80.0 }),
                    size: Some(crate::protobuf::tsp::Size {
                        width: 200.0,
                        height: 60.0,
                    }),
                    flags: Some(0),
                    angle: Some(0.0),
                }),
                parent: Some(reference(body_id)),
                title: Some(reference(title_id)),
                caption: Some(reference(caption_id)),
                ..Default::default()
            },
            ..Default::default()
        },
        deprecated_storage: Some(reference(storage_id)),
        owned_storage: Some(reference(storage_id)),
        is_text_box: Some(true),
        ..Default::default()
    };
    let storage = StorageArchive {
        text: vec!["Source".to_owned()],
        ..Default::default()
    };
    let zorder = tp::DrawablesZOrderArchive {
        drawables: vec![reference(body_id), reference(drawable_id)],
    };
    let attachment = DrawableAttachmentArchive {
        drawable: Some(reference(drawable_id)),
        h_offset_type: Some(0),
        h_offset: Some(109.0),
        v_offset_type: Some(1),
        v_offset: Some(84.0),
    };
    let mut root_object = object(1, 10000, root.encode_to_vec());
    root_object.archive_info.message_infos[0]
        .object_references
        .extend([body_id, zorder_id]);
    let mut body_object = object(body_id, 2001, body.encode_to_vec());
    body_object.archive_info.message_infos[0]
        .object_references
        .push(attachment_id);
    let mut shape_object = object(drawable_id, 2011, shape.encode_to_vec());
    shape_object.archive_info.message_infos[0]
        .object_references
        .extend([body_id, title_id, caption_id, storage_id]);
    let mut zorder_object = object(zorder_id, 10015, zorder.encode_to_vec());
    zorder_object.archive_info.message_infos[0]
        .object_references
        .extend([body_id, drawable_id]);
    let mut attachment_object = object(attachment_id, 2003, attachment.encode_to_vec());
    attachment_object.archive_info.message_infos[0]
        .object_references
        .push(drawable_id);
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    root_object,
                    body_object,
                    shape_object,
                    object(storage_id, 2001, storage.encode_to_vec()),
                    zorder_object,
                    attachment_object,
                    object(title_id, 3097, Vec::new()),
                    object(caption_id, 3097, Vec::new()),
                ],
            },
        )
        .unwrap();
    package
}

fn floating_text_package() -> IWorkPackage {
    let body_id = 42;
    let floating_id = 50;
    let placeholder_id = 60;
    let detached_id = 61;
    let placeholder_storage_id = 62;
    let detached_storage_id = 63;
    let text_box_id = 64;
    let text_box_storage_id = 65;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        floating_drawables: Some(reference(floating_id)),
        ..Default::default()
    };
    let floating = tp::FloatingDrawablesArchive {
        page_groups: vec![tp::floating_drawables_archive::PageGroup {
            page_index: 0,
            drawables: [placeholder_id, text_box_id]
                .into_iter()
                .map(|identifier| tp::floating_drawables_archive::DrawableEntry {
                    drawable: Some(reference(identifier)),
                })
                .collect(),
            ..Default::default()
        }],
        ..Default::default()
    };
    let placeholder = |storage_id| tp::PlaceholderArchive {
        super_: crate::protobuf::tswp::ShapeInfoArchive {
            owned_storage: Some(reference(storage_id)),
            ..Default::default()
        },
    };
    let shape = |storage_id| crate::protobuf::tswp::ShapeInfoArchive {
        owned_storage: Some(reference(storage_id)),
        ..Default::default()
    };
    let storage = |text: &str| StorageArchive {
        text: vec![text.to_owned()],
        ..Default::default()
    };

    let mut root_object = object(1, 10000, root.encode_to_vec());
    root_object.archive_info.message_infos[0]
        .object_references
        .extend([body_id, floating_id]);
    let mut floating_object = object(floating_id, 10010, floating.encode_to_vec());
    floating_object.archive_info.message_infos[0]
        .object_references
        .extend([placeholder_id, text_box_id]);
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    root_object,
                    object(body_id, 2001, storage("Body").encode_to_vec()),
                    floating_object,
                    object(
                        placeholder_id,
                        PLACEHOLDER_MESSAGE_TYPE,
                        placeholder(placeholder_storage_id).encode_to_vec(),
                    ),
                    object(
                        placeholder_storage_id,
                        2001,
                        storage("A🚀B").encode_to_vec(),
                    ),
                    object(
                        text_box_id,
                        SHAPE_INFO_MESSAGE_TYPE,
                        shape(text_box_storage_id).encode_to_vec(),
                    ),
                    object(
                        text_box_storage_id,
                        2001,
                        storage("Independent text box").encode_to_vec(),
                    ),
                    object(
                        detached_id,
                        PLACEHOLDER_MESSAGE_TYPE,
                        placeholder(detached_storage_id).encode_to_vec(),
                    ),
                    object(
                        detached_storage_id,
                        2001,
                        storage("Detached").encode_to_vec(),
                    ),
                ],
            },
        )
        .unwrap();
    package
}

fn test_package(text: &str) -> IWorkPackage {
    let body_id = 42;
    let root = DocumentArchive {
        body_storage: Some(Reference {
            identifier: body_id,
            ..Default::default()
        }),
        ..Default::default()
    };
    let storage = StorageArchive {
        text: vec![text.to_owned()],
        ..Default::default()
    };
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    ArchiveObject::new(
                        1,
                        vec![RawMessage {
                            type_: 10000,
                            data: root.encode_to_vec(),
                        }],
                    )
                    .unwrap(),
                    ArchiveObject::new(
                        body_id,
                        vec![RawMessage {
                            type_: 2001,
                            data: storage.encode_to_vec(),
                        }],
                    )
                    .unwrap(),
                ],
            },
        )
        .unwrap();
    package
}

fn test_package_with_settings(settings_data: Vec<u8>) -> IWorkPackage {
    let body_id = 42;
    let settings_id = 43;
    let root = DocumentArchive {
        body_storage: Some(reference(body_id)),
        settings: Some(reference(settings_id)),
        ..Default::default()
    };
    let storage = StorageArchive {
        text: vec!["Body".to_owned()],
        ..Default::default()
    };
    let mut root_object = object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec());
    root_object.archive_info.message_infos[0].object_references = vec![body_id, settings_id];
    let archive = Archive {
        objects: vec![
            root_object,
            object(body_id, 2_001, storage.encode_to_vec()),
            object(settings_id, 10_012, settings_data),
        ],
    };
    let mut package = IWorkPackage::new();
    package
        .replace_archive(DOCUMENT_ARCHIVE_NAME, &archive)
        .unwrap();
    package
}

fn reference(identifier: u64) -> Reference {
    Reference {
        identifier,
        ..Default::default()
    }
}

fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) -> Vec<u8> {
    let mut field = crate::varint::encode_varint(u64::from(field_number) << 3);
    field.extend(crate::varint::encode_varint(value));
    data.extend_from_slice(&field);
    field
}

fn object(identifier: u64, message_type: u32, data: Vec<u8>) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )
    .unwrap()
}
