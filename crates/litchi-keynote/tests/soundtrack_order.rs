//! Exact-source semantic coverage for Keynote soundtrack playback order.

use litchi_iwa_archive::{Limits, package::Catalog, package::EntryEdit};
use litchi_iwa_common::{encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_soundtrack_settings_codec as soundtrack_codec;
use litchi_keynote::{Package, Position, ReadOptions, SemanticLimits, soundtrack::order::Error};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SYNTHETIC_MEDIA: [u64; 2] = [9074, 9071];
const SYNTHETIC_SOUNDTRACK: u64 = 2651093;

fn fixture() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/basic.key")
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

/// Add two existing fixture media assets to the otherwise settings-only
/// soundtrack.  The fixture intentionally has no soundtrack records, so the
/// order transaction tests must build a real, bounded media closure instead
/// of silently returning from every changed-order assertion.
fn with_soundtrack_media(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacements = Vec::new();
    let mut changed_soundtrack = false;
    let mut changed_metadata = false;

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            let Some(index) = object.messages.iter().position(|message| {
                message.type_ == 21 && object.archive_info.identifier == Some(SYNTHETIC_SOUNDTRACK)
            }) else {
                continue;
            };
            let mut payload = object.messages[index].data.clone();
            for identifier in SYNTHETIC_MEDIA {
                let mut reference = Vec::new();
                encode_varint_into(&mut reference, 8);
                encode_varint_into(&mut reference, identifier);
                encode_varint_into(&mut payload, 26);
                encode_varint_into(
                    &mut payload,
                    u64::try_from(reference.len()).expect("reference length fits u64"),
                );
                payload.extend_from_slice(&reference);
            }
            let info = object
                .archive_info
                .message_infos
                .get_mut(index)
                .ok_or("soundtrack message metadata is missing")?;
            info.data_references = SYNTHETIC_MEDIA.to_vec();
            if let Some(media_field) = info
                .field_infos
                .iter_mut()
                .find(|field| field.path.as_slice() == [3])
            {
                media_field.r#type = Some(FieldType::DataReference);
                media_field.data_references = SYNTHETIC_MEDIA.to_vec();
            } else {
                let mut media_field = FieldInfo::new(vec![3]);
                media_field.r#type = Some(FieldType::DataReference);
                media_field.data_references = SYNTHETIC_MEDIA.to_vec();
                info.field_infos.push(media_field);
            }
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: 21,
                    data: payload,
                },
            )?;
            changed = true;
            changed_soundtrack = true;
        }
        if changed {
            replacements.push((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?,
            ));
        }
    }

    for entry in catalog
        .iter()
        .filter(|entry| entry.name() == "Index/Metadata.iwa")
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == 11_006)
            else {
                continue;
            };
            let payload = object.messages[index].data.clone();
            let view = WireView::parse(&payload)?;
            let mut rewritten = Vec::new();
            for field in view.fields() {
                if field.number() != 3 {
                    rewritten.extend_from_slice(field.raw());
                    continue;
                }
                let component = WireView::parse(field.payload())?;
                let preferred = component
                    .fields()
                    .find(|nested| nested.number() == 2)
                    .map(|nested| nested.payload());
                if preferred != Some(b"Document") {
                    rewritten.extend_from_slice(field.raw());
                    continue;
                }
                let mut component_bytes = Vec::new();
                for nested in component.fields() {
                    if nested.number() != 7 {
                        component_bytes.extend_from_slice(nested.raw());
                        continue;
                    }
                    let data_reference = WireView::parse(nested.payload())?;
                    let identifier = data_reference
                        .fields()
                        .find(|reference| reference.number() == 1)
                        .and_then(|reference| {
                            let bytes = reference.payload();
                            let (value, consumed) =
                                litchi_iwa_common::decode_varint_from_bytes(bytes).ok()?;
                            (consumed == bytes.len()).then_some(value)
                        });
                    if !SYNTHETIC_MEDIA.contains(&identifier.unwrap_or_default()) {
                        component_bytes.extend_from_slice(nested.raw());
                        continue;
                    }
                    let mut data_reference_bytes = Vec::new();
                    for reference in data_reference.fields() {
                        data_reference_bytes.extend_from_slice(reference.raw());
                    }
                    let mut owner = Vec::new();
                    encode_varint_into(&mut owner, 8);
                    encode_varint_into(&mut owner, SYNTHETIC_SOUNDTRACK);
                    encode_varint_into(&mut owner, 16);
                    encode_varint_into(&mut owner, 1);
                    encode_varint_into(&mut data_reference_bytes, 18);
                    encode_varint_into(
                        &mut data_reference_bytes,
                        u64::try_from(owner.len()).expect("owner length fits u64"),
                    );
                    data_reference_bytes.extend_from_slice(&owner);
                    encode_varint_into(&mut component_bytes, 58);
                    encode_varint_into(
                        &mut component_bytes,
                        u64::try_from(data_reference_bytes.len())
                            .expect("data reference length fits u64"),
                    );
                    component_bytes.extend_from_slice(&data_reference_bytes);
                    changed = true;
                    changed_metadata = true;
                }
                encode_varint_into(&mut rewritten, 26);
                encode_varint_into(
                    &mut rewritten,
                    u64::try_from(component_bytes.len()).expect("component length fits u64"),
                );
                rewritten.extend_from_slice(&component_bytes);
            }
            if changed {
                object.replace_message_preserving_header(
                    index,
                    RawMessage {
                        type_: 11_006,
                        data: rewritten,
                    },
                )?;
            }
        }
        if changed {
            replacements.push((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?,
            ));
        }
    }

    assert!(changed_soundtrack && changed_metadata);
    let edits: Vec<_> = replacements
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    Ok(catalog.reassemble_to_bytes(&edits, Limits::default())?)
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
                message.type_ == 21 && object.archive_info.identifier == Some(SYNTHETIC_SOUNDTRACK)
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
    let package = Package::open(fixture())?;
    let source = with_soundtrack_media(&bytes(&package)?)?;
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
    let native = Package::open(fixture())?;
    let source = with_soundtrack_media(&bytes(&native)?)?;
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
    let fixture = Package::open(fixture())?;
    let source = append_unknown_soundtrack_field(&with_soundtrack_media(&bytes(&fixture)?)?)?;
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
    let native = Package::open(fixture())?;
    let source = with_soundtrack_media(&bytes(&native)?)?;
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
    let fixture = Package::open(fixture())?;
    let source = append_unknown_soundtrack_metadata(&with_soundtrack_media(&bytes(&fixture)?)?)?;
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
    let native = Package::open(fixture())?;
    let source = with_soundtrack_media(&bytes(&native)?)?;
    let fixture = Package::from_bytes(&source)?;
    let before = soundtrack_ids(&fixture)?;
    assert_eq!(before.len(), SYNTHETIC_MEDIA.len());
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
    let native = Package::open(fixture())?;
    let source = with_soundtrack_media(&bytes(&native)?)?;
    let external = [0x18, 1];
    let noncanonical = [0x10, 0x80, 0];
    for nested in [&external[..], &noncanonical[..]] {
        let malformed = append_root_reference_bytes(&source, nested)?;
        let package = Package::from_bytes(&malformed)?;
        let mut edit = package.edit_soundtrack_order();
        assert!(matches!(
            edit.move_item(Position::new(0), Position::new(0)),
            Err(Error::InvalidSource)
        ));
    }
    Ok(())
}

#[test]
fn soundtrack_order_reference_budget_is_aggregate() -> TestResult {
    let native = Package::open(fixture())?;
    let source = with_soundtrack_media(&bytes(&native)?)?;
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
