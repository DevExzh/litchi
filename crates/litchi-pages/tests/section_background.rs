use std::io;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{
    WireLimits, encode_varint_into,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsd, tsp, tswp};
use litchi_pages::{
    Limits, Package, Position, SectionSelector,
    section::{
        Background,
        background::{Error, LimitKind},
    },
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const BODY_IDENTIFIER: u64 = 42;
const FIRST_SECTION_IDENTIFIER: u64 = 43;
const SECOND_SECTION_IDENTIFIER: u64 = 44;
const SECTION_MESSAGE_TYPE: u32 = 10_011;
const PRIVATE_MARKER: &str = "private-pages-background-marker-998244353";

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn assert_send_sync<T: Send + Sync>(_: &T) {}
fn assert_type_send_sync<T: Send + Sync>() {}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn solid(red: f32, green: f32, blue: f32, alpha: f32) -> Background {
    Background::Solid(
        litchi_iwa_common::color::Rgba::new(
            red,
            green,
            blue,
            alpha,
            litchi_iwa_common::color::RgbColorSpace::Srgb,
        )
        .expect("test color is valid"),
    )
}

fn solid_payload(red: f32, green: f32, blue: f32, alpha: f32) -> Vec<u8> {
    tsd::FillArchive {
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(red),
            g: Some(green),
            b: Some(blue),
            a: Some(alpha),
            rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
            ..tsp::Color::default()
        }),
        ..tsd::FillArchive::default()
    }
    .encode_to_vec()
}

fn unsupported_payload() -> Vec<u8> {
    tsd::FillArchive {
        gradient: Some(tsd::GradientArchive::default()),
        ..tsd::FillArchive::default()
    }
    .encode_to_vec()
}

fn solid_payload_with_nested_unknown() -> TestResult<Vec<u8>> {
    let fill = tsd::FillArchive::decode(solid_payload(0.1, 0.2, 0.3, 0.4).as_slice())?;
    let mut color = fill
        .color
        .ok_or_else(|| io::Error::other("solid color missing"))?
        .encode_to_vec();
    append_varint_field(&mut color, 99, 9_999)?;
    let mut output = Vec::new();
    append_length_delimited_field(&mut output, 1, &color)?;
    append_varint_field(&mut output, 98, 8_888)?;
    Ok(output)
}

fn section_message(name: &str, background: Option<&[u8]>, unknown: u64) -> TestResult<RawMessage> {
    let mut data = tp::SectionArchive {
        name: Some(name.to_owned()),
        ..tp::SectionArchive::default()
    }
    .encode_to_vec();
    if let Some(payload) = background {
        append_length_delimited_field(&mut data, 30, payload)?;
    }
    append_varint_field(&mut data, 99, unknown)?;
    Ok(RawMessage {
        type_: SECTION_MESSAGE_TYPE,
        data,
    })
}

fn package_with_first_message(first: RawMessage, second: RawMessage) -> TestResult<Vec<u8>> {
    package_with_first_message_metadata(first, second, false, false, false, false)
}

fn package_with_first_message_metadata(
    mut first: RawMessage,
    second: RawMessage,
    field_object_reference: bool,
    aggregate_object_reference: bool,
    aggregate_data_reference: bool,
    aliased_preserved_references: bool,
) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let body = tswp::StorageArchive {
        text: vec!["Alpha\u{0004}Beta".to_owned()],
        table_section: Some(tswp::ObjectAttributeTable {
            entries: vec![
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(reference(FIRST_SECTION_IDENTIFIER)),
                },
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 6,
                    object: Some(reference(SECOND_SECTION_IDENTIFIER)),
                },
            ],
        }),
        ..tswp::StorageArchive::default()
    };
    if aliased_preserved_references {
        let encoded = reference(60).encode_to_vec();
        append_length_delimited_field(&mut first.data, 23, &encoded)?;
        append_length_delimited_field(&mut first.data, 24, &encoded)?;
    }
    let mut first_section = ArchiveObject::new(
        FIRST_SECTION_IDENTIFIER,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-section".to_vec(),
            },
            first,
            RawMessage {
                type_: 778,
                data: b"after-section".to_vec(),
            },
        ],
    )?;
    if field_object_reference {
        let mut field = FieldInfo::new(vec![30]);
        field.object_references = vec![60];
        first_section.archive_info.message_infos[1]
            .field_infos
            .push(field);
    }
    if aggregate_object_reference || aliased_preserved_references {
        first_section.archive_info.message_infos[1].object_references = vec![60];
    }
    if aggregate_data_reference {
        first_section.archive_info.message_infos[1].data_references = vec![61];
    }
    let mut objects = vec![
        object(1, 10_000, root.encode_to_vec())?,
        object(BODY_IDENTIFIER, 2_001, body.encode_to_vec())?,
        first_section,
        ArchiveObject::new(SECOND_SECTION_IDENTIFIER, vec![second])?,
    ];
    if aliased_preserved_references {
        objects.push(object(60, 888, b"shared template".to_vec())?);
    }
    let document = component(objects)?;
    let unrelated = component(vec![object(99, 777, PRIVATE_MARKER.as_bytes().to_vec())?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document.as_slice()),
            (UNRELATED_MEMBER, unrelated.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn package(first: Option<&[u8]>, second: Option<&[u8]>) -> TestResult<Vec<u8>> {
    package_with_first_message(
        section_message("Alpha", first, 7_777)?,
        section_message("Beta", second, 8_888)?,
    )
}

fn legacy_package_bytes(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| {
            std::path::Path::new(entry.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("iwa"))
        })
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner = litchi_iwa_archive::package::to_bytes(inner_entries, Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.pages/Index.zip", inner.as_slice()),
            (
                "legacy.pages/Data/sentinel.bin",
                b"legacy outer sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("document member missing"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn message_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("section missing"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == SECTION_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("section payload missing").into())
}

fn background_payload(package: &[u8], identifier: u64) -> TestResult<Option<Vec<u8>>> {
    let payload = message_payload(package, identifier)?;
    let view = WireView::parse_with_limits(&payload, WireLimits::default())?;
    let mut values = view
        .fields()
        .filter(|field| field.number() == 30)
        .map(|field| field.payload().to_vec());
    let value = values.next();
    assert!(
        values.next().is_none(),
        "synthetic payload has duplicate background fields"
    );
    Ok(value)
}

fn fields_except_background(payload: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse_with_limits(payload, WireLimits::default())?
        .fields()
        .filter(|field| field.number() != 30)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn assert_untouched_zip_members(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let before_entries = before_catalog.iter().collect::<Vec<_>>();
    let after_entries = after_catalog.iter().collect::<Vec<_>>();
    assert_eq!(before_entries.len(), after_entries.len());
    let mut changed = 0_usize;
    for (before, after) in before_entries.into_iter().zip(after_entries) {
        assert_eq!(before.name(), after.name());
        if before.data() == after.data() {
            assert_eq!(before.metadata(), after.metadata());
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record()
            );
        } else {
            changed += 1;
            assert_eq!(before.name(), DOCUMENT_MEMBER);
        }
    }
    assert_eq!(changed, 1);
    Ok(())
}

#[test]
fn reads_absent_solid_and_unsupported_losslessly() -> TestResult<()> {
    let absent = Package::from_bytes(&package(None, None)?)?;
    assert_eq!(
        absent.section_background(SectionSelector::index(0))?,
        Background::None
    );

    let known = Package::from_bytes(&package(Some(&solid_payload(0.1, 0.2, 0.3, 0.4)), None)?)?;
    assert_eq!(
        known.section_background(SectionSelector::name("Alpha"))?,
        solid(0.1, 0.2, 0.3, 0.4)
    );

    let unsupported = Package::from_bytes(&package(Some(&unsupported_payload()), None)?)?;
    assert_eq!(
        unsupported.section_background(SectionSelector::index(0))?,
        Background::Unsupported
    );
    let source_pointer = unsupported.source_bytes().as_ptr();
    let noop = unsupported
        .edit_section_background(SectionSelector::index(0))?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes(), unsupported.source_bytes());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);
    Ok(())
}

#[test]
fn selector_and_malformed_background_fields_fail_closed_without_mutation() -> TestResult<()> {
    let same = Package::from_bytes(&package_with_first_message(
        section_message("Same", None, 1)?,
        section_message("Same", None, 2)?,
    )?)?;
    assert!(matches!(
        same.edit_section_background(SectionSelector::name("Missing")),
        Err(Error::NameNotFound)
    ));
    assert!(
        matches!(same.edit_section_background(SectionSelector::index(2)), Err(Error::PositionNotFound { position }) if position == Position::new(2))
    );
    assert!(
        matches!(same.edit_section_background(SectionSelector::name("Same")), Err(Error::AmbiguousSelector { first, duplicate }) if first == Position::new(0) && duplicate == Position::new(1))
    );

    let canonical = section_message("Alpha", None, 7_777)?;
    let mut malformed = Vec::new();
    let mut duplicate = canonical.data.clone();
    append_length_delimited_field(&mut duplicate, 30, &solid_payload(0.1, 0.2, 0.3, 0.4))?;
    append_length_delimited_field(&mut duplicate, 30, &solid_payload(0.4, 0.3, 0.2, 0.1))?;
    malformed.push(duplicate);
    let mut wrong_wire = canonical.data.clone();
    append_varint_field(&mut wrong_wire, 30, 1)?;
    malformed.push(wrong_wire);
    let mut noncanonical = canonical.data;
    encode_varint_into(&mut noncanonical, (u64::from(30_u32) << 3) | 2);
    noncanonical.extend_from_slice(&[0x80, 0x00]);
    malformed.push(noncanonical);

    let mut nested_wrong_wire = section_message("Alpha", None, 7_777)?.data;
    let mut fill = solid_payload(0.1, 0.2, 0.3, 1.0);
    append_varint_field(&mut fill, 1, 1)?;
    append_length_delimited_field(&mut nested_wrong_wire, 30, &fill)?;
    malformed.push(nested_wrong_wire);

    let mut nested_noncanonical = section_message("Alpha", None, 7_777)?.data;
    let mut fill = Vec::new();
    encode_varint_into(&mut fill, (u64::from(1_u32) << 3) | 2);
    fill.extend_from_slice(&[0x80, 0x00]);
    append_length_delimited_field(&mut nested_noncanonical, 30, &fill)?;
    malformed.push(nested_noncanonical);

    let mut duplicate_color = tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.1),
        g: Some(0.2),
        b: Some(0.3),
        a: Some(1.0),
        ..tsp::Color::default()
    }
    .encode_to_vec();
    encode_varint_into(&mut duplicate_color, (u64::from(3_u32) << 3) | 5);
    duplicate_color.extend_from_slice(&0.4_f32.to_le_bytes());
    let mut fill = Vec::new();
    append_length_delimited_field(&mut fill, 1, &duplicate_color)?;
    let mut nested_duplicate = section_message("Alpha", None, 7_777)?.data;
    append_length_delimited_field(&mut nested_duplicate, 30, &fill)?;
    malformed.push(nested_duplicate);

    let mut wrong_color_wire = tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.1),
        g: Some(0.2),
        b: Some(0.3),
        a: Some(1.0),
        ..tsp::Color::default()
    }
    .encode_to_vec();
    append_varint_field(&mut wrong_color_wire, 3, 1)?;
    let mut fill = Vec::new();
    append_length_delimited_field(&mut fill, 1, &wrong_color_wire)?;
    let mut nested_wrong_color_wire = section_message("Alpha", None, 7_777)?.data;
    append_length_delimited_field(&mut nested_wrong_color_wire, 30, &fill)?;
    malformed.push(nested_wrong_color_wire);

    let mut noncanonical_color_key = tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.1),
        g: Some(0.2),
        b: Some(0.3),
        a: Some(1.0),
        ..tsp::Color::default()
    }
    .encode_to_vec();
    noncanonical_color_key.extend_from_slice(&[0x9d, 0x80, 0x00, 0, 0, 0, 0]);
    let mut fill = Vec::new();
    append_length_delimited_field(&mut fill, 1, &noncanonical_color_key)?;
    let mut nested_noncanonical_color = section_message("Alpha", None, 7_777)?.data;
    append_length_delimited_field(&mut nested_noncanonical_color, 30, &fill)?;
    malformed.push(nested_noncanonical_color);

    let mut malformed_gradient = Vec::new();
    append_length_delimited_field(&mut malformed_gradient, 2, &[0x80])?;
    let mut malformed_gradient_section = section_message("Alpha", None, 7_777)?.data;
    append_length_delimited_field(&mut malformed_gradient_section, 30, &malformed_gradient)?;
    malformed.push(malformed_gradient_section);

    for payload in malformed {
        let bytes = package_with_first_message(
            RawMessage {
                type_: SECTION_MESSAGE_TYPE,
                data: payload,
            },
            section_message("Beta", None, 2)?,
        )?;
        let source = Package::from_bytes(&bytes)?;
        assert!(matches!(
            source.section_background(SectionSelector::index(0)),
            Err(Error::InvalidSource { .. })
        ));
        assert!(matches!(
            source.edit_section_background(SectionSelector::index(0)),
            Err(Error::InvalidSource { .. })
        ));
        assert_eq!(source.source_bytes(), bytes);
    }
    Ok(())
}

#[test]
fn color_projection_is_strict_but_preserves_supported_display_p3() -> TestResult<()> {
    let p3_fill = tsd::FillArchive {
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(0.1),
            g: Some(0.2),
            b: Some(0.3),
            a: Some(0.4),
            rgbspace: Some(tsp::color::RgbColorSpace::P3 as i32),
            ..tsp::Color::default()
        }),
        ..tsd::FillArchive::default()
    }
    .encode_to_vec();
    let source = Package::from_bytes(&package(Some(&p3_fill), None)?)?;
    assert_eq!(
        source.section_background(SectionSelector::index(0))?,
        Background::Solid(litchi_iwa_common::color::Rgba::new(
            0.1,
            0.2,
            0.3,
            0.4,
            litchi_iwa_common::color::RgbColorSpace::DisplayP3
        )?,)
    );

    for color in [
        tsp::Color {
            model: tsp::color::ColorModel::Cmyk as i32,
            r: Some(0.1),
            g: Some(0.2),
            b: Some(0.3),
            a: Some(1.0),
            ..tsp::Color::default()
        },
        tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(f32::NAN),
            g: Some(0.2),
            b: Some(0.3),
            a: Some(1.0),
            ..tsp::Color::default()
        },
        tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(1.1),
            g: Some(0.2),
            b: Some(0.3),
            a: Some(1.0),
            ..tsp::Color::default()
        },
        tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(0.1),
            g: Some(0.2),
            b: Some(0.3),
            a: Some(f32::INFINITY),
            ..tsp::Color::default()
        },
        tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(0.1),
            g: Some(0.2),
            b: Some(0.3),
            a: Some(1.0),
            rgbspace: Some(99),
            ..tsp::Color::default()
        },
    ] {
        let fill = tsd::FillArchive {
            color: Some(color),
            ..tsd::FillArchive::default()
        }
        .encode_to_vec();
        let source = Package::from_bytes(&package(Some(&fill), None)?)?;
        match source.section_background(SectionSelector::index(0)) {
            Ok(Background::Unsupported) => {
                let mut edit = source.edit_section_background(SectionSelector::index(0))?;
                edit.clear();
                assert!(matches!(
                    edit.commit(),
                    Err(Error::UnsupportedSource { .. })
                ));
            },
            Err(Error::InvalidSource { .. }) => {},
            other => panic!("invalid color had unexpected classification: {other:?}"),
        }
    }
    Ok(())
}

#[test]
fn set_clear_noop_and_patch_lifecycle_preserve_exact_source_rules() -> TestResult<()> {
    let bytes = package(None, Some(&solid_payload(0.7, 0.6, 0.5, 1.0)))?;
    let source = Package::from_bytes(&bytes)?;
    let source_ptr = source.source_bytes().as_ptr();
    let mut noop = source.edit_section_background(SectionSelector::index(0))?;
    noop.clear();
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_ptr);
    assert!(!noop.diagnostics().changed());

    let target_background = solid(0.1, 0.2, 0.3, 0.4);
    let mut edit = source.edit_section_background(SectionSelector::name("Alpha"))?;
    edit.set_solid(match target_background {
        Background::Solid(value) => value,
        _ => unreachable!(),
    })?;
    let commit = edit.commit()?;
    assert_eq!(source.source_bytes(), bytes);
    assert_eq!(
        commit
            .package()
            .section_background(SectionSelector::index(0))?,
        target_background
    );
    assert_eq!(
        commit
            .package()
            .section_background(SectionSelector::index(1))?,
        solid(0.7, 0.6, 0.5, 1.0)
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    let applied = source.apply_section_background(commit.patch())?;
    assert_eq!(
        applied.package().source_bytes(),
        commit.package().source_bytes()
    );
    let restored = commit
        .package()
        .apply_section_background(&commit.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);

    let unrelated = Package::from_bytes(&package(Some(&solid_payload(0.9, 0.8, 0.7, 1.0)), None)?)?;
    assert!(matches!(
        unrelated.apply_section_background(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert!(matches!(
        commit.package().apply_section_background(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains(PRIVATE_MARKER));
    assert!(!debug.contains("Index/"));
    assert_send_sync(&commit);
    assert_send_sync(commit.patch());
    assert_send_sync(commit.diagnostics());
    assert_type_send_sync::<Error>();
    Ok(())
}

#[test]
fn clear_and_changed_legacy_sources_obey_exact_source_rules() -> TestResult<()> {
    let original = solid_payload(0.2, 0.3, 0.4, 1.0);
    let bytes = package(Some(&original), None)?;
    let source = Package::from_bytes(&bytes)?;
    let source_payload = message_payload(&bytes, FIRST_SECTION_IDENTIFIER)?;
    let mut clear = source.edit_section_background(SectionSelector::index(0))?;
    clear.clear();
    let commit = clear.commit()?;
    assert_eq!(
        commit
            .package()
            .section_background(SectionSelector::index(0))?,
        Background::None
    );
    let target_payload =
        message_payload(commit.package().source_bytes(), FIRST_SECTION_IDENTIFIER)?;
    assert_eq!(
        fields_except_background(&target_payload)?,
        fields_except_background(&source_payload)?
    );
    assert_eq!(
        background_payload(commit.package().source_bytes(), FIRST_SECTION_IDENTIFIER)?,
        None
    );
    assert_untouched_zip_members(&bytes, commit.package().source_bytes())?;

    let legacy = Package::from_bytes(&legacy_package_bytes(&bytes)?)?;
    let mut noop = legacy.edit_section_background(SectionSelector::index(0))?;
    noop.set_solid(match solid(0.2, 0.3, 0.4, 1.0) {
        Background::Solid(value) => value,
        _ => unreachable!(),
    })?;
    assert!(noop.commit()?.patch().is_noop());
    let mut changed = legacy.edit_section_background(SectionSelector::index(0))?;
    changed.clear();
    assert!(matches!(
        changed.commit(),
        Err(Error::UnsupportedSource { .. })
    ));
    Ok(())
}

#[test]
fn changed_background_preserves_nested_unknowns_and_package_locality() -> TestResult<()> {
    let original_fill = solid_payload_with_nested_unknown()?;
    let bytes = package(Some(&original_fill), Some(&unsupported_payload()))?;
    let source = Package::from_bytes(&bytes)?;
    let source_payload = message_payload(&bytes, FIRST_SECTION_IDENTIFIER)?;
    let source_second = message_payload(&bytes, SECOND_SECTION_IDENTIFIER)?;
    let mut edit = source.edit_section_background(SectionSelector::index(0))?;
    edit.set_solid(match solid(0.4, 0.3, 0.2, 1.0) {
        Background::Solid(value) => value,
        _ => unreachable!(),
    })?;
    let commit = edit.commit()?;
    let target = commit.package().source_bytes();
    let target_payload = message_payload(target, FIRST_SECTION_IDENTIFIER)?;
    assert_eq!(
        fields_except_background(&source_payload)?,
        fields_except_background(&target_payload)?
    );
    let edited_fill = background_payload(target, FIRST_SECTION_IDENTIFIER)?
        .ok_or_else(|| io::Error::other("edited fill missing"))?;
    let fill_view = WireView::parse(&edited_fill)?;
    assert!(fill_view.fields().any(|field| field.number() == 98));
    let color = fill_view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("edited color missing"))?;
    assert!(
        WireView::parse(color.payload())?
            .fields()
            .any(|field| field.number() == 99)
    );
    assert_eq!(
        message_payload(target, SECOND_SECTION_IDENTIFIER)?,
        source_second
    );
    assert_untouched_zip_members(&bytes, target)?;
    Ok(())
}

#[test]
fn unsupported_or_reference_owned_backgrounds_refuse_changed_edits_atomically() -> TestResult<()> {
    let unsupported_bytes = package(Some(&unsupported_payload()), None)?;
    let unsupported = Package::from_bytes(&unsupported_bytes)?;
    let mut clear = unsupported.edit_section_background(SectionSelector::index(0))?;
    clear.clear();
    assert!(matches!(
        clear.commit(),
        Err(Error::UnsupportedSource { .. })
    ));
    assert_eq!(unsupported.source_bytes(), unsupported_bytes);
    let mut replace = unsupported.edit_section_background(SectionSelector::index(0))?;
    replace.set_solid(match solid(0.1, 0.2, 0.3, 1.0) {
        Background::Solid(value) => value,
        _ => unreachable!(),
    })?;
    assert!(matches!(
        replace.commit(),
        Err(Error::UnsupportedSource { .. })
    ));
    assert_eq!(unsupported.source_bytes(), unsupported_bytes);

    for bytes in [
        package_with_first_message_metadata(
            section_message("Alpha", Some(&solid_payload(0.1, 0.2, 0.3, 1.0)), 7_777)?,
            section_message("Beta", None, 8_888)?,
            true,
            false,
            false,
            false,
        )?,
        package_with_first_message_metadata(
            section_message("Alpha", None, 7_777)?,
            section_message("Beta", None, 8_888)?,
            false,
            false,
            true,
            false,
        )?,
        package_with_first_message_metadata(
            section_message("Alpha", Some(&solid_payload(0.1, 0.2, 0.3, 1.0)), 7_777)?,
            section_message("Beta", None, 8_888)?,
            false,
            true,
            false,
            false,
        )?,
    ] {
        let source = Package::from_bytes(&bytes)?;
        let mut edit = source.edit_section_background(SectionSelector::index(0))?;
        edit.set_solid(match solid(0.4, 0.3, 0.2, 1.0) {
            Background::Solid(value) => value,
            _ => unreachable!(),
        })?;
        assert!(matches!(
            edit.commit(),
            Err(Error::InvalidSource { .. } | Error::UnsupportedSource { .. })
        ));
        assert_eq!(source.source_bytes(), bytes);
    }

    let aliased = package_with_first_message_metadata(
        section_message("Alpha", Some(&solid_payload(0.1, 0.2, 0.3, 1.0)), 7_777)?,
        section_message("Beta", None, 8_888)?,
        false,
        false,
        false,
        true,
    )?;
    let source = Package::from_bytes(&aliased)?;
    let mut edit = source.edit_section_background(SectionSelector::index(0))?;
    edit.clear();
    let commit = edit.commit()?;
    assert_eq!(
        commit
            .package()
            .section_background(SectionSelector::index(0))?,
        Background::None
    );
    Ok(())
}

#[test]
fn output_limits_are_atomic() -> TestResult<()> {
    let bytes = package(None, None)?;
    let source = Package::from_bytes(&bytes)?;
    let mut edit = source.edit_section_background(SectionSelector::index(0))?;
    edit.set_solid(match solid(0.1, 0.2, 0.3, 1.0) {
        Background::Solid(value) => value,
        _ => unreachable!(),
    })?;
    let target_length = edit.commit()?.package().source_bytes().len();
    let limits = Limits::new(
        u64::try_from(target_length - 1)?,
        8,
        1_024 * 1_024,
        1_024 * 1_024,
        1_024 * 1_024,
    )?;
    let limited = Package::from_bytes_with_limits(&bytes, limits)?;
    let mut limited_edit = limited.edit_section_background(SectionSelector::index(0))?;
    limited_edit.set_solid(match solid(0.1, 0.2, 0.3, 1.0) {
        Background::Solid(value) => value,
        _ => unreachable!(),
    })?;
    assert!(matches!(
        limited_edit.commit(),
        Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(limited.source_bytes(), bytes);
    Ok(())
}
