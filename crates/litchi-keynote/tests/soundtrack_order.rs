//! Exact-source semantic coverage for Keynote soundtrack playback order.

use litchi_iwa_archive::{Limits, package::Catalog, package::EntryEdit};
use litchi_iwa_common::{encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{keynote_soundtrack_settings_codec as soundtrack_codec, kn, tsa, tsk, tsp};
use litchi_keynote::{Package, Position, ReadOptions, SemanticLimits, soundtrack::order::Error};
use prost::Message;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 100;
const DOCUMENT_COMPONENT: u64 = 1;
const ROOT_OBJECT: u64 = 1;
const SHOW_OBJECT: u64 = 2;
const SOUNDTRACK_OBJECT: u64 = 3;
const THEME_OBJECT: u64 = 4;
const STYLESHEET_OBJECT: u64 = 5;
const AUDIO_DATA_IDS: [u64; 2] = [7_001, 7_002];
const AUDIO_FILENAMES: [&str; 2] = ["soundtrack-first.wav", "soundtrack-second.wav"];
const FIRST_AUDIO: &[u8] = b"RIFF\x24\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x00\x00\x00\x00";
const SECOND_AUDIO: &[u8] = b"RIFF\x26\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x02\x00\x00\x00\x00\x00";
const AUDIO_DIGESTS: [[u8; 20]; 2] = [
    [
        0xab, 0xa0, 0x27, 0x8c, 0x07, 0x53, 0x2e, 0xc5, 0x72, 0x88, 0x8f, 0x76, 0xaf, 0x7e, 0x9e,
        0xbd, 0x60, 0x71, 0x36, 0x70,
    ],
    [
        0x5f, 0x85, 0xe2, 0xb9, 0x1b, 0xcd, 0x2b, 0x20, 0xd0, 0xad, 0xa5, 0x0e, 0xc9, 0xa3, 0x45,
        0x67, 0xef, 0x8e, 0xbc, 0xa1,
    ],
];

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
) -> TestResult<litchi_iwa_core::ArchiveObject> {
    Ok(litchi_iwa_core::ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn component(objects: Vec<litchi_iwa_core::ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

/// Build a complete Keynote source with a native soundtrack graph and two
/// materialized WAV assets.  Every payload, archive index, metadata owner, and
/// ZIP member is created together so the strict mutation validator sees one
/// internally consistent source rather than grafted references into a native
/// fixture's unrelated image graph.
fn fixture() -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(SHOW_OBJECT),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: reference(THEME_OBJECT),
        slide_tree: kn::SlideTreeArchive::default(),
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(STYLESHEET_OBJECT),
        soundtrack: Some(reference(SOUNDTRACK_OBJECT)),
        ..Default::default()
    };
    let soundtrack = kn::Soundtrack {
        volume: Some(1.0),
        mode: Some(kn::soundtrack::SoundtrackMode::KKnSoundtrackModePlayOnce as i32),
        movie_media: AUDIO_DATA_IDS
            .iter()
            .copied()
            .map(|identifier| tsp::DataReference { identifier })
            .collect(),
    };

    let mut root = object(ROOT_OBJECT, 1, document.encode_to_vec())?;
    root.archive_info.message_infos[0].object_references = vec![SHOW_OBJECT];
    let mut root_show_field = FieldInfo::new(vec![2]);
    root_show_field.r#type = Some(FieldType::ObjectReference);
    root_show_field.object_references = vec![SHOW_OBJECT];
    root.archive_info.message_infos[0]
        .field_infos
        .push(root_show_field);

    let mut show_object = object(SHOW_OBJECT, 2, show.encode_to_vec())?;
    show_object.archive_info.message_infos[0].object_references =
        vec![THEME_OBJECT, STYLESHEET_OBJECT, SOUNDTRACK_OBJECT];
    let mut show_soundtrack_field = FieldInfo::new(vec![17]);
    show_soundtrack_field.r#type = Some(FieldType::ObjectReference);
    show_soundtrack_field.object_references = vec![SOUNDTRACK_OBJECT];
    show_object.archive_info.message_infos[0]
        .field_infos
        .push(show_soundtrack_field);

    let mut soundtrack_object = object(SOUNDTRACK_OBJECT, 21, soundtrack.encode_to_vec())?;
    soundtrack_object.archive_info.message_infos[0].data_references = AUDIO_DATA_IDS.to_vec();
    let mut soundtrack_media_field = FieldInfo::new(vec![3]);
    soundtrack_media_field.r#type = Some(FieldType::DataReference);
    soundtrack_media_field.data_references = AUDIO_DATA_IDS.to_vec();
    soundtrack_object.archive_info.message_infos[0]
        .field_infos
        .push(soundtrack_media_field);

    let document_component = component(vec![
        root,
        show_object,
        soundtrack_object,
        object(THEME_OBJECT, 10, Vec::new())?,
        object(STYLESHEET_OBJECT, 9_002, Vec::new())?,
    ])?;

    let metadata = tsp::PackageMetadata {
        last_object_identifier: METADATA_OBJECT,
        components: vec![tsp::ComponentInfo {
            identifier: DOCUMENT_COMPONENT,
            preferred_locator: "Document".to_owned(),
            locator: Some("Document".to_owned()),
            data_references: AUDIO_DATA_IDS
                .iter()
                .copied()
                .map(|identifier| tsp::ComponentDataReference {
                    data_identifier: identifier,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: SOUNDTRACK_OBJECT,
                        count: 1,
                    }],
                })
                .collect(),
            ..Default::default()
        }],
        datas: AUDIO_DATA_IDS
            .iter()
            .enumerate()
            .map(|(index, identifier)| {
                let data = [FIRST_AUDIO, SECOND_AUDIO][index];
                tsp::DataInfo {
                    identifier: *identifier,
                    digest: AUDIO_DIGESTS[index].to_vec(),
                    preferred_file_name: AUDIO_FILENAMES[index].to_owned(),
                    file_name: Some(AUDIO_FILENAMES[index].to_owned()),
                    materialized_length: Some(
                        u64::try_from(data.len()).expect("audio length fits"),
                    ),
                    ..Default::default()
                }
            })
            .collect(),
        ..Default::default()
    };
    let metadata_component = component(vec![object(
        METADATA_OBJECT,
        11_006,
        metadata.encode_to_vec(),
    )?])?;

    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/soundtrack-first.wav", FIRST_AUDIO),
            ("Data/soundtrack-second.wav", SECOND_AUDIO),
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (METADATA_MEMBER, metadata_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn append_root_reference_bytes(source: &[u8], nested: &[u8]) -> TestResult<Vec<u8>> {
    const ROOT_OBJECT: u64 = 1;
    const DOCUMENT_MESSAGE_TYPE: u32 = 1;
    const SHOW_FIELD: u32 = 2;
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        for object in &mut archive.objects {
            if object.archive_info.identifier != Some(ROOT_OBJECT) {
                continue;
            }
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
            else {
                continue;
            };
            let view = WireView::parse(&object.messages[index].data)?;
            let mut rewritten = Vec::new();
            let mut changed = false;
            for field in view.fields() {
                if field.number() != SHOW_FIELD {
                    rewritten.extend_from_slice(field.raw());
                    continue;
                }
                let mut payload = field.payload().to_vec();
                payload.extend_from_slice(nested);
                encode_varint_into(&mut rewritten, (u64::from(SHOW_FIELD) << 3) | 2);
                encode_varint_into(&mut rewritten, u64::try_from(payload.len())?);
                rewritten.extend_from_slice(&payload);
                changed = true;
            }
            if !changed {
                continue;
            }
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data: rewritten,
                },
            )?;
            let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
            return Ok(catalog.reassemble_to_bytes(
                &[EntryEdit::new(entry.name(), &compressed)],
                Limits::default(),
            )?);
        }
    }
    Err("root show reference was not found in the fixture".into())
}

fn append_unknown_soundtrack_field(source: &[u8]) -> TestResult<Vec<u8>> {
    const UNKNOWN: &[u8] = &[0xa0, 0x06, 0x07];
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        for object in &mut archive.objects {
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == 21)
            else {
                continue;
            };
            let mut data = object.messages[index].data.clone();
            data.extend_from_slice(UNKNOWN);
            object.replace_message_preserving_header(index, RawMessage { type_: 21, data })?;
            let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
            return Ok(catalog.reassemble_to_bytes(
                &[EntryEdit::new(entry.name(), &compressed)],
                Limits::default(),
            )?);
        }
    }
    Err("soundtrack payload was not found in the fixture".into())
}

fn append_external_soundtrack_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        for object in &mut archive.objects {
            let Some(index) = object.messages.iter().position(|message| {
                message.type_ == 21 && object.archive_info.identifier == Some(SOUNDTRACK_OBJECT)
            }) else {
                continue;
            };
            let view = WireView::parse(&object.messages[index].data)?;
            let mut rewritten = Vec::new();
            let mut changed = false;
            for field in view.fields() {
                if field.number() != 3 || changed {
                    rewritten.extend_from_slice(field.raw());
                    continue;
                }
                let mut reference = field.payload().to_vec();
                reference.extend_from_slice(&[0x18, 1]);
                encode_varint_into(&mut rewritten, 26);
                encode_varint_into(
                    &mut rewritten,
                    u64::try_from(reference.len()).expect("reference length fits u64"),
                );
                rewritten.extend_from_slice(&reference);
                changed = true;
            }
            if !changed {
                continue;
            }
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: 21,
                    data: rewritten,
                },
            )?;
            let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
            return Ok(catalog.reassemble_to_bytes(
                &[EntryEdit::new(entry.name(), &compressed)],
                Limits::default(),
            )?);
        }
    }
    Err("synthetic soundtrack payload was not found".into())
}

fn append_unknown_soundtrack_metadata(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        for object in &mut archive.objects {
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == 21)
            else {
                continue;
            };
            object.archive_info.message_infos[index]
                .field_infos
                .push(FieldInfo::new(vec![99, 7]));
            let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
            return Ok(catalog.reassemble_to_bytes(
                &[EntryEdit::new(entry.name(), &compressed)],
                Limits::default(),
            )?);
        }
    }
    Err("soundtrack metadata was not found in the fixture".into())
}

fn soundtrack_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            if let Some(message) = object.messages.iter().find(|message| message.type_ == 21) {
                return Ok(message.data.clone());
            }
        }
    }
    Err("soundtrack payload was not found in the fixture".into())
}

#[test]
fn fixture_order_transaction_is_exact_and_reversible() -> TestResult {
    let source = fixture()?;
    let package = Package::from_bytes(&source)?;
    let edit = package.edit_soundtrack_order();
    assert!(matches!(edit.commit(), Err(Error::NoStagedOperation)));

    let mut edit = package.edit_soundtrack_order();
    edit.move_item(Position::new(0), Position::new(0))?;
    let noop = edit.commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(bytes(noop.package())?, source);

    let mut changed = package.edit_soundtrack_order();
    changed.move_item(Position::new(0), Position::new(1))?;
    let commit = changed.commit()?;
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    let applied = package.apply_soundtrack_order(commit.patch())?;
    assert_eq!(bytes(applied.package())?, bytes(commit.package())?);
    assert!(matches!(
        commit.package().apply_soundtrack_order(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let restored = commit
        .package()
        .apply_soundtrack_order(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn changed_order_rejects_a_tight_output_limit_atomically() -> TestResult {
    let source = fixture()?;
    let input_bytes = u64::try_from(source.len())?;
    let limits = Limits::new(input_bytes, 128, 1024 * 1024, 1024 * 1024, 1024 * 1024)?;
    let package = Package::from_bytes_with_limits(&source, limits)?;
    let mut edit = package.edit_soundtrack_order();
    edit.move_item(Position::new(0), Position::new(1))?;
    assert!(matches!(
        edit.commit(),
        Err(Error::LimitExceeded {
            kind: litchi_keynote::soundtrack::order::LimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(bytes(&package)?, source);
    Ok(())
}

#[test]
fn changed_order_preserves_unknown_soundtrack_payload_fields() -> TestResult {
    let source = append_unknown_soundtrack_field(&fixture()?)?;
    let package = Package::from_bytes(&source)?;
    let before_catalog = Catalog::from_bytes(&source)?;
    let mut edit = package.edit_soundtrack_order();
    edit.move_item(Position::new(0), Position::new(1))?;
    let commit = edit.commit()?;
    let after_bytes = bytes(commit.package())?;
    let after_catalog = Catalog::from_bytes(&after_bytes)?;
    assert_untouched_zip_records(&before_catalog, &after_catalog);
    assert!(soundtrack_payload(&after_bytes)?.ends_with(&[0xa0, 0x06, 0x07]));
    Ok(())
}

fn assert_untouched_zip_records(before: &Catalog, after: &Catalog) {
    let before_entries: Vec<_> = before.iter().collect();
    let after_entries: Vec<_> = after.iter().collect();
    assert_eq!(before_entries.len(), after_entries.len());
    let mut changed_entries = 0;
    for (before, after) in before_entries.iter().zip(after_entries) {
        assert_eq!(before.name(), after.name());
        if before.data() != after.data() {
            changed_entries += 1;
            continue;
        }
        assert_eq!(before.raw_name(), after.raw_name());
        assert_eq!(
            before.raw_record().local_record(),
            after.raw_record().local_record()
        );
        let before_central = before.raw_record().central_directory_record();
        let after_central = after.raw_record().central_directory_record();
        assert_eq!(&before_central[..42], &after_central[..42]);
        assert_eq!(&before_central[46..], &after_central[46..]);
    }
    assert_eq!(changed_entries, 1);
}

#[test]
fn external_soundtrack_media_reference_fails_before_staging() -> TestResult {
    let source = fixture()?;
    let malformed = append_external_soundtrack_reference(&source)?;
    let package = Package::from_bytes(&malformed)?;
    let mut edit = package.edit_soundtrack_order();
    assert!(matches!(
        edit.move_item(Position::new(0), Position::new(1)),
        Err(Error::InvalidSource)
    ));
    Ok(())
}

#[test]
fn changed_order_preserves_unknown_soundtrack_metadata() -> TestResult {
    let source = append_unknown_soundtrack_metadata(&fixture()?)?;
    let package = Package::from_bytes(&source)?;
    let before_catalog = Catalog::from_bytes(&source)?;
    let before_metadata = soundtrack_metadata(&before_catalog)?;

    let mut edit = package.edit_soundtrack_order();
    edit.move_item(Position::new(0), Position::new(1))?;
    let commit = edit.commit()?;
    let after_catalog = Catalog::from_bytes(&bytes(commit.package())?)?;
    let after_metadata = soundtrack_metadata(&after_catalog)?;
    assert_eq!(
        after_metadata.unknown_fields,
        before_metadata.unknown_fields
    );
    assert_eq!(
        after_metadata.media_field_count,
        before_metadata.media_field_count
    );
    Ok(())
}

#[derive(Debug, PartialEq, Eq)]
struct SoundtrackMetadata {
    unknown_fields: Vec<FieldInfo>,
    media_field_count: usize,
}

fn soundtrack_metadata(catalog: &Catalog) -> TestResult<SoundtrackMetadata> {
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == 21)
            else {
                continue;
            };
            let info = &object.archive_info.message_infos[index];
            return Ok(SoundtrackMetadata {
                unknown_fields: info
                    .field_infos
                    .iter()
                    .filter(|field| field.path.as_slice() == [99, 7])
                    .cloned()
                    .collect(),
                media_field_count: info
                    .field_infos
                    .iter()
                    .filter(|field| field.path.as_slice() == [3])
                    .count(),
            });
        }
    }
    Err("soundtrack metadata was not found in the fixture".into())
}

#[test]
fn source_order_permutations_match_position_moves() -> TestResult {
    let source = fixture()?;
    let fixture = Package::from_bytes(&source)?;
    let before = soundtrack_ids(&fixture)?;
    assert_eq!(before.len(), AUDIO_DATA_IDS.len());
    for source_position in 0..before.len() {
        for destination_position in 0..before.len() {
            let mut expected = before.clone();
            let value = expected.remove(source_position);
            expected.insert(destination_position, value);

            let package = Package::from_bytes(&source)?;
            let mut edit = package.edit_soundtrack_order();
            edit.move_item(
                Position::new(source_position),
                Position::new(destination_position),
            )?;
            let commit = edit.commit()?;
            assert_eq!(soundtrack_ids(commit.package())?, expected);
        }
    }
    Ok(())
}

#[test]
fn malformed_root_reference_ownership_fails_before_staging() -> TestResult {
    let source = fixture()?;
    let external = [0x18, 1];
    let noncanonical = [0x10, 0x80, 0];
    for nested in [&external[..], &noncanonical[..]] {
        let malformed = append_root_reference_bytes(&source, nested)?;
        match Package::from_bytes(&malformed) {
            Ok(package) => {
                let mut edit = package.edit_soundtrack_order();
                assert!(matches!(
                    edit.move_item(Position::new(0), Position::new(0)),
                    Err(Error::InvalidSource)
                ));
            },
            // A non-canonical nested varint is rejected by the strict root
            // decoder at package ingress, before a transaction can stage.
            Err(_) if nested == &noncanonical[..] => {},
            Err(error) => return Err(error.into()),
        }
    }
    Ok(())
}

#[test]
fn soundtrack_order_reference_budget_is_aggregate() -> TestResult {
    let source = fixture()?;
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        SemanticLimits::MAX_SLIDES,
        1,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )?;
    let package =
        Package::from_bytes_with_options(&source, ReadOptions::new(Limits::default(), semantic))?;
    let mut edit = package.edit_soundtrack_order();
    assert!(matches!(
        edit.move_item(Position::new(0), Position::new(0)),
        Err(Error::LimitExceeded { .. })
    ));
    assert_eq!(bytes(&package)?, source);
    Ok(())
}

fn soundtrack_ids(package: &Package) -> TestResult<Vec<u64>> {
    let payload = soundtrack_payload(&bytes(package)?)?;
    let options = soundtrack_codec::DecodeOptions::new(payload.len(), 256, 16 * 1024, 64);
    let mut ids = Vec::new();
    soundtrack_codec::visit_soundtrack_media_identifiers(&payload, options, &mut |id| {
        ids.push(id);
        Ok(())
    })?;
    Ok(ids)
}
