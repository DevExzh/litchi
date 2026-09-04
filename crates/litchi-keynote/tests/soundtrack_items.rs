//! Semantic Keynote soundtrack-item lifecycle coverage.
//!
//! The fixture is intentionally small and self-contained.  It contains a
//! rooted `Document -> Show -> Soundtrack` graph, exact archive metadata, two
//! materialized audio resources, and one unrelated ZIP member.  Keeping the
//! graph in this test makes the lifecycle assertions independent from the
//! number of media records in the checked-in native presentation fixture.

use std::collections::HashMap;

use litchi_core::Position;
use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::encode_varint_into;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use litchi_keynote::{Package, soundtrack::items::AudioSource};
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

// These are deliberately tiny, but have a real RIFF/WAVE signature and are
// accepted by the shared media classifier.  The lifecycle owner treats the
// bytes as opaque and computes metadata from the exact supplied payload.
const FIRST_AUDIO: &[u8] =
    b"RIFF\x24\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x00\x00\x00\x00";
const SECOND_AUDIO: &[u8] =
    b"RIFF\x26\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x02\x00\x00\x00\x00\x00";
const THIRD_AUDIO: &[u8] =
    b"RIFF\x28\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x04\x00\x00\x00\x01\x02\x03\x04";

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

/// Build a complete, strict, two-item Keynote soundtrack package.
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
                    // These digests are checked by the strict metadata owner.
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
            ("Data/unrelated.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (METADATA_MEMBER, metadata_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn package() -> TestResult<Package> {
    Ok(Package::from_bytes(&fixture()?)?)
}

/// Rewrite one object in one flat IWA member while retaining the source ZIP
/// envelope.  The helper is used only to derive strict boundary fixtures;
/// production code must continue to use the focused package owner.
fn rewrite_iwa_object<F>(
    source: &[u8],
    object_identifier: u64,
    message_type: u32,
    mut rewrite: F,
) -> TestResult<Vec<u8>>
where
    F: FnMut(&mut litchi_iwa_core::ArchiveObject, usize) -> TestResult<()>,
{
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        for object in &mut archive.objects {
            if object.archive_info.identifier != Some(object_identifier) {
                continue;
            }
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == message_type)
            else {
                continue;
            };
            rewrite(object, index)?;
            let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
            return Ok(catalog.reassemble_to_bytes(
                &[litchi_iwa_archive::package::EntryEdit::new(
                    entry.name(),
                    &compressed,
                )],
                Limits::default(),
            )?);
        }
    }
    Err("requested IWA object was not found in fixture".into())
}

/// Derive a source without a rooted soundtrack.  Keeping the unreferenced
/// soundtrack object and media records in the package is intentional: the
/// read boundary must answer `None` solely from the root → show graph and
/// must not infer a soundtrack from orphaned physical records.
fn fixture_without_soundtrack() -> TestResult<Vec<u8>> {
    rewrite_iwa_object(&fixture()?, SHOW_OBJECT, 2, |object, index| {
        let mut show = kn::ShowArchive::decode(object.messages[index].data.as_slice())?;
        show.soundtrack = None;
        let info = &mut object.archive_info.message_infos[index];
        info.object_references
            .retain(|identifier| *identifier != SOUNDTRACK_OBJECT);
        info.field_infos
            .retain(|field| field.path.as_slice() != [17]);
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 2,
                data: show.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

/// Derive a source with an existing, but empty, soundtrack collection.
fn fixture_with_empty_soundtrack() -> TestResult<Vec<u8>> {
    rewrite_iwa_object(&fixture()?, SOUNDTRACK_OBJECT, 21, |object, index| {
        let mut soundtrack = kn::Soundtrack::decode(object.messages[index].data.as_slice())?;
        soundtrack.movie_media.clear();
        let info = &mut object.archive_info.message_infos[index];
        info.data_references.clear();
        for field in &mut info.field_infos {
            if field.path.as_slice() == [3] {
                field.data_references.clear();
            }
        }
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 21,
                data: soundtrack.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

/// Derive a source whose two semantic occurrences share the first physical
/// media resource.  The metadata owner count is raised to two so this remains
/// a valid closure and exercises occurrence-local removal/replacement rather
/// than accidentally testing global media garbage collection.
fn fixture_with_shared_first_media() -> TestResult<Vec<u8>> {
    let source = rewrite_iwa_object(&fixture()?, SOUNDTRACK_OBJECT, 21, |object, index| {
        let mut soundtrack = kn::Soundtrack::decode(object.messages[index].data.as_slice())?;
        soundtrack.movie_media = vec![
            tsp::DataReference {
                identifier: AUDIO_DATA_IDS[0],
            };
            2
        ];
        let info = &mut object.archive_info.message_infos[index];
        info.data_references = vec![AUDIO_DATA_IDS[0]; 2];
        for field in &mut info.field_infos {
            if field.path.as_slice() == [3] {
                field.data_references = vec![AUDIO_DATA_IDS[0]; 2];
            }
        }
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 21,
                data: soundtrack.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;

    rewrite_iwa_object(&source, METADATA_OBJECT, 11_006, |object, index| {
        let mut metadata = tsp::PackageMetadata::decode(object.messages[index].data.as_slice())?;
        let mut updated = false;
        for component in &mut metadata.components {
            for data_reference in &mut component.data_references {
                if data_reference.data_identifier != AUDIO_DATA_IDS[0] {
                    continue;
                }
                for owner in &mut data_reference.object_reference_list {
                    if owner.object_identifier == SOUNDTRACK_OBJECT {
                        owner.count = 2;
                        updated = true;
                    }
                }
            }
        }
        if !updated {
            return Err("shared fixture owner was not found".into());
        }
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 11_006,
                data: metadata.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn item_summary(package: &Package) -> TestResult<Vec<(Position, String, usize)>> {
    let items = package
        .soundtrack_items()?
        .ok_or("soundtrack unexpectedly absent")?;
    Ok(items
        .iter()
        .map(|item| {
            (
                item.position(),
                item.filename().to_owned(),
                item.byte_length(),
            )
        })
        .collect())
}

fn audio(filename: &str, data: &[u8]) -> TestResult<AudioSource> {
    Ok(AudioSource::new(filename, data.to_vec())?)
}

fn entries(source: &[u8]) -> TestResult<HashMap<String, Vec<u8>>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect())
}

/// Verify that every unselected member retains its local record and all
/// central-directory bytes except the relative local-header offset.  A
/// changed member may legitimately alter CRC/size and a new/deleted member
/// changes offsets for later records; those are the only ZIP bookkeeping
/// differences this focused transaction is allowed to introduce.
fn assert_untouched_zip_records(before: &Catalog, after: &Catalog, changed: &[&str]) {
    for before_entry in before.iter() {
        if changed.contains(&before_entry.name()) {
            continue;
        }
        let after_entry = after
            .iter()
            .find(|entry| entry.name() == before_entry.name())
            .unwrap_or_else(|| {
                panic!("unselected ZIP member disappeared: {}", before_entry.name())
            });
        assert_eq!(before_entry.raw_name(), after_entry.raw_name());
        assert_eq!(before_entry.data(), after_entry.data());
        assert_eq!(
            before_entry.raw_record().local_record(),
            after_entry.raw_record().local_record(),
            "unselected local record changed: {}",
            before_entry.name()
        );
        let before_central = before_entry.raw_record().central_directory_record();
        let after_central = after_entry.raw_record().central_directory_record();
        assert_eq!(
            &before_central[..42],
            &after_central[..42],
            "unselected central record changed before offset: {}",
            before_entry.name()
        );
        assert_eq!(
            &before_central[46..],
            &after_central[46..],
            "unselected central record changed after offset: {}",
            before_entry.name()
        );
    }
}

fn append_unknown_soundtrack_field(source: &[u8]) -> TestResult<Vec<u8>> {
    // Unknown fields are legal and must remain byte-exact through a changed
    // lifecycle transaction.  The selected payload is length-delimited, so
    // appending an unknown varint is enough to exercise the preservation path.
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
                &[litchi_iwa_archive::package::EntryEdit::new(
                    entry.name(),
                    &compressed,
                )],
                Limits::default(),
            )?);
        }
    }
    Err("soundtrack payload was not found in fixture".into())
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
    Err("soundtrack payload was not found in fixture".into())
}

fn append_external_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    // Add an explicit external marker to one selected DataReference.  The
    // strict owner must reject the source before it allocates a lifecycle
    // rewrite, and the immutable package bytes must stay unchanged.
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
                let mut payload = field.payload().to_vec();
                payload.extend_from_slice(&[0x18, 1]);
                encode_varint_into(&mut rewritten, 26);
                encode_varint_into(
                    &mut rewritten,
                    u64::try_from(payload.len()).expect("payload length fits"),
                );
                rewritten.extend_from_slice(&payload);
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
                &[litchi_iwa_archive::package::EntryEdit::new(
                    entry.name(),
                    &compressed,
                )],
                Limits::default(),
            )?);
        }
    }
    Err("soundtrack payload was not found in fixture".into())
}

fn without_soundtrack_field_info(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let Some(object) = archive.objects.iter_mut().find(|object| {
            object.archive_info.identifier == Some(SOUNDTRACK_OBJECT)
                && object.messages.iter().any(|message| message.type_ == 21)
        }) else {
            continue;
        };
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == 21)
            .ok_or("soundtrack message was absent")?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or("soundtrack message info was absent")?;
        info.field_infos
            .retain(|field| field.path.as_slice() != [3]);
        let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
        return Ok(catalog.reassemble_to_bytes(
            &[litchi_iwa_archive::package::EntryEdit::new(
                entry.name(),
                &compressed,
            )],
            Limits::default(),
        )?);
    }
    Err("soundtrack object was not found in fixture".into())
}

fn soundtrack_has_field_info(source: &[u8]) -> TestResult<bool> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            if object.archive_info.identifier != Some(SOUNDTRACK_OBJECT) {
                continue;
            }
            let Some(index) = object
                .messages
                .iter()
                .position(|message| message.type_ == 21)
            else {
                continue;
            };
            let info = object
                .archive_info
                .message_infos
                .get(index)
                .ok_or("soundtrack message info was absent")?;
            return Ok(info
                .field_infos
                .iter()
                .any(|field| field.path.as_slice() == [3]));
        }
    }
    Err("soundtrack object was not found in fixture".into())
}

#[test]
fn absent_and_existing_empty_soundtracks_are_distinct() -> TestResult {
    let absent = Package::from_bytes(&fixture_without_soundtrack()?)?;
    assert_eq!(absent.soundtrack_items()?, None);

    let empty = Package::from_bytes(&fixture_with_empty_soundtrack()?)?;
    let items = empty
        .soundtrack_items()?
        .ok_or("existing empty soundtrack was reported absent")?;
    assert!(items.is_empty());
    Ok(())
}

#[test]
fn duplicate_media_occurrences_keep_shared_data_and_inverse_exact() -> TestResult {
    let source = fixture_with_shared_first_media()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        item_summary(&package)?,
        vec![
            (
                Position::new(0),
                AUDIO_FILENAMES[0].to_owned(),
                FIRST_AUDIO.len()
            ),
            (
                Position::new(1),
                AUDIO_FILENAMES[0].to_owned(),
                FIRST_AUDIO.len()
            ),
        ]
    );

    let before_catalog = Catalog::from_bytes(&source)?;
    let mut edit = package.edit_soundtrack_items()?;
    edit.remove(Position::new(0))?;
    let commit = edit.commit()?;
    assert_eq!(
        item_summary(commit.package())?,
        vec![(
            Position::new(0),
            AUDIO_FILENAMES[0].to_owned(),
            FIRST_AUDIO.len()
        )]
    );

    let target = bytes(commit.package())?;
    let after_catalog = Catalog::from_bytes(&target)?;
    assert_eq!(
        after_catalog
            .iter()
            .find(|entry| entry.name() == "Data/soundtrack-first.wav")
            .map(|entry| entry.data()),
        Some(FIRST_AUDIO)
    );
    assert_untouched_zip_records(
        &before_catalog,
        &after_catalog,
        &[DOCUMENT_MEMBER, METADATA_MEMBER],
    );

    let restored = commit
        .package()
        .apply_soundtrack_items(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn read_add_insert_replace_remove_are_ordered_and_semantic() -> TestResult {
    let package = package()?;
    assert_eq!(
        item_summary(&package)?,
        vec![
            (
                Position::new(0),
                AUDIO_FILENAMES[0].to_owned(),
                FIRST_AUDIO.len()
            ),
            (
                Position::new(1),
                AUDIO_FILENAMES[1].to_owned(),
                SECOND_AUDIO.len()
            ),
        ]
    );

    let mut add = package.edit_soundtrack_items()?;
    add.add(audio("appended.wav", THIRD_AUDIO)?)?;
    let added = add.commit()?;
    assert!(added.diagnostics().changed());
    assert_eq!(item_summary(added.package())?.len(), 3);
    assert_eq!(item_summary(added.package())?[2].1, "appended.wav");

    let mut insert = added.package().edit_soundtrack_items()?;
    insert.insert(Position::new(0), audio("inserted.wav", THIRD_AUDIO)?)?;
    let inserted = insert.commit()?;
    let inserted_bytes = bytes(inserted.package())?;
    assert_eq!(
        Catalog::from_bytes(&inserted_bytes)?
            .iter()
            .filter(|entry| entry.name().starts_with("Data/"))
            .count(),
        4,
        "reusing identical media must not insert another physical member"
    );
    assert_eq!(
        item_summary(inserted.package())?
            .iter()
            .map(|(_, name, _)| name.as_str())
            .collect::<Vec<_>>(),
        [
            "appended.wav",
            "soundtrack-first.wav",
            "soundtrack-second.wav",
            "appended.wav"
        ]
    );

    let mut replace = inserted.package().edit_soundtrack_items()?;
    replace.replace(Position::new(1), audio("replacement.wav", THIRD_AUDIO)?)?;
    let replaced = replace.commit()?;
    assert_eq!(item_summary(replaced.package())?[1].1, "appended.wav");
    assert_eq!(item_summary(replaced.package())?[1].2, THIRD_AUDIO.len());

    let mut remove = replaced.package().edit_soundtrack_items()?;
    remove.remove(Position::new(1))?;
    let removed = remove.commit()?;
    assert_eq!(
        item_summary(removed.package())?
            .iter()
            .map(|(_, name, _)| name.as_str())
            .collect::<Vec<_>>(),
        ["appended.wav", "soundtrack-second.wav", "appended.wav"]
    );
    Ok(())
}

#[test]
fn lifecycle_patch_apply_inverse_and_foreign_source_conflict_are_exact() -> TestResult {
    let package = package()?;
    let source = bytes(&package)?;
    let mut edit = package.edit_soundtrack_items()?;
    edit.replace(Position::new(0), audio("replacement.wav", THIRD_AUDIO)?)?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().full_reparse_performed());
    assert!(!commit.patch().is_noop());
    assert_eq!(
        commit
            .patch()
            .before()
            .iter()
            .map(|item| item.filename())
            .collect::<Vec<_>>(),
        ["soundtrack-first.wav", "soundtrack-second.wav"]
    );
    assert_eq!(
        commit
            .patch()
            .after()
            .iter()
            .map(|item| item.filename())
            .collect::<Vec<_>>(),
        ["replacement.wav", "soundtrack-second.wav"]
    );

    let target = bytes(commit.package())?;
    let applied = package.apply_soundtrack_items(commit.patch())?;
    assert_eq!(bytes(applied.package())?, target);
    assert!(matches!(
        commit.package().apply_soundtrack_items(commit.patch()),
        Err(litchi_keynote::soundtrack::items::Error::PatchConflict)
    ));

    let restored = commit
        .package()
        .apply_soundtrack_items(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    assert_eq!(item_summary(restored.package())?, item_summary(&package)?);

    // Handles are opaque capabilities tied to the immutable source lineage;
    // an item read before publication cannot accidentally retarget the new
    // package snapshot.
    let original_item = package.soundtrack_items()?.unwrap()[0].clone();
    let mut foreign_edit = commit.package().edit_soundtrack_items()?;
    foreign_edit.remove(original_item.handle())?;
    assert!(matches!(
        foreign_edit.commit(),
        Err(litchi_keynote::soundtrack::items::Error::ItemHandleConflict)
    ));
    Ok(())
}

#[test]
fn no_op_patch_apply_preserves_the_exact_snapshot_without_reparse() -> TestResult {
    let package = package()?;
    let source = bytes(&package)?;
    let current = package.soundtrack_items()?.unwrap()[0].clone();
    let mut edit = package.edit_soundtrack_items()?;
    // A caller-provided alias still reuses the package's canonical media
    // record by digest, so the candidate becomes an exact no-op only after
    // physical rewrite planning.
    edit.replace(current.handle(), audio("caller-alias.wav", FIRST_AUDIO)?)?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());

    let applied = package.apply_soundtrack_items(commit.patch())?;
    assert_eq!(bytes(applied.package())?, source);
    assert!(applied.patch().is_noop());
    assert!(!applied.diagnostics().changed());
    assert_eq!(applied.diagnostics().touched_components(), 0);
    assert!(!applied.diagnostics().full_reparse_performed());
    assert_eq!(item_summary(applied.package())?, item_summary(&package)?);
    Ok(())
}

#[test]
fn empty_edit_and_invalid_positions_are_atomic() -> TestResult {
    let package = package()?;
    let source = bytes(&package)?;
    assert!(matches!(
        package.edit_soundtrack_items()?.commit(),
        Err(litchi_keynote::soundtrack::items::Error::NoStagedOperation)
    ));

    let mut already_staged = package.edit_soundtrack_items()?;
    already_staged.add(audio("first-new.wav", THIRD_AUDIO)?)?;
    assert!(matches!(
        already_staged.remove(Position::new(0)),
        Err(litchi_keynote::soundtrack::items::Error::OperationAlreadyStaged)
    ));

    for stage in [0_u8, 1, 2] {
        let mut edit = package.edit_soundtrack_items()?;
        let result = match stage {
            0 => edit.insert(Position::new(3), audio("bad.wav", THIRD_AUDIO)?),
            1 => edit.replace(Position::new(2), audio("bad.wav", THIRD_AUDIO)?),
            _ => edit.remove(Position::new(2)),
        };
        assert!(result.is_err(), "out-of-bounds operation was staged");
        assert_eq!(bytes(&package)?, source);
    }

    for (filename, data) in [("not-audio.png", THIRD_AUDIO), ("empty.wav", &[][..])] {
        assert!(AudioSource::new(filename, data.to_vec()).is_err());
        assert_eq!(bytes(&package)?, source);
    }
    Ok(())
}

#[test]
fn changed_item_transaction_preserves_unknown_fields_and_unselected_zip_records() -> TestResult {
    let source = append_unknown_soundtrack_field(&fixture()?)?;
    let package = Package::from_bytes(&source)?;
    let before_entries = entries(&source)?;
    let mut edit = package.edit_soundtrack_items()?;
    edit.replace(Position::new(0), audio("replacement.wav", THIRD_AUDIO)?)?;
    let commit = edit.commit()?;
    let target = bytes(commit.package())?;
    let after_entries = entries(&target)?;
    let after_catalog = Catalog::from_bytes(&target)?;
    assert_untouched_zip_records(
        &Catalog::from_bytes(&source)?,
        &after_catalog,
        &[
            DOCUMENT_MEMBER,
            METADATA_MEMBER,
            "Data/soundtrack-first.wav",
        ],
    );

    assert!(soundtrack_payload(&target)?.ends_with(&[0xa0, 0x06, 0x07]));
    for (name, data) in &before_entries {
        if name == DOCUMENT_MEMBER || name == METADATA_MEMBER || name.starts_with("Data/") {
            continue;
        }
        assert_eq!(after_entries.get(name), Some(data));
    }
    assert_eq!(
        after_entries.get("Data/soundtrack-first.wav"),
        before_entries.get("Data/soundtrack-first.wav")
    );
    assert_eq!(
        after_entries.get("Data/soundtrack-second.wav"),
        before_entries.get("Data/soundtrack-second.wav")
    );
    assert_eq!(
        after_entries.get("Data/unrelated.bin"),
        before_entries.get("Data/unrelated.bin")
    );
    Ok(())
}

#[test]
fn malformed_external_reference_fails_before_publication() -> TestResult {
    let malformed = append_external_reference(&fixture()?)?;
    let package = Package::from_bytes(&malformed)?;
    let before = bytes(&package)?;
    match package.edit_soundtrack_items() {
        Err(_) => {},
        Ok(mut edit) => {
            let result = edit.replace(Position::new(0), audio("replacement.wav", THIRD_AUDIO)?);
            if result.is_ok() {
                assert!(edit.commit().is_err());
            }
        },
    }
    assert_eq!(bytes(&package)?, before);
    Ok(())
}

#[test]
fn aggregate_only_native_metadata_is_admitted_and_preserved_on_write() -> TestResult {
    let source = without_soundtrack_field_info(&fixture()?)?;
    let package = Package::from_bytes(&source)?;
    let items = package
        .soundtrack_items()?
        .ok_or("aggregate-only soundtrack was reported absent")?;
    assert_eq!(items.len(), 2);

    let mut edit = package.edit_soundtrack_items()?;
    edit.replace(Position::new(0), audio("replacement.wav", THIRD_AUDIO)?)?;
    let commit = edit.commit()?;
    assert_eq!(commit.package().soundtrack_items()?.unwrap().len(), 2);
    assert!(!soundtrack_has_field_info(&bytes(commit.package())?)?);

    let restored = commit
        .package()
        .apply_soundtrack_items(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}
