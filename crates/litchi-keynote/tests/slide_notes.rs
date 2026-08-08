use std::io;

use litchi_iwa_archive::{Limits, package::Catalog, package::EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp, tswp};
use litchi_keynote::{
    Package, Position, SlideNotesCommit, SlideNotesDiagnostics, SlideNotesError,
    SlideNotesLimitKind, SlideNotesPatch, SlideSelector, TextPosition, TextSpan,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const FIRST_SLIDE: u64 = 4;
const SECOND_SLIDE: u64 = 31;
const NOTE_OBJECT: u64 = 8;
const NOTE_STORAGE: u64 = 13;
const NOTE_MESSAGE_TYPE: u32 = 15;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const PRIVATE_MARKER: &[u8] = b"private-keynote-notes-marker-2147483647";

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(identifier: u64, type_: u32, value: &impl prost::Message) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_,
            data: value.encode_to_vec(),
        }],
    )?)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn component_with_unknown_header(
    objects: Vec<ArchiveObject>,
    target_identifier: u64,
) -> TestResult<Vec<u8>> {
    let bytes = Archive { objects }.to_bytes()?;
    let parsed = Archive::parse(&bytes)?;
    let object = parsed
        .object(target_identifier)
        .ok_or_else(|| io::Error::other("synthetic target object is missing"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length_u64, prefix_length) = decode_varint_from_bytes(
        bytes
            .get(header_offset..)
            .ok_or_else(|| io::Error::other("synthetic header offset is invalid"))?,
    )?;
    let header_length = usize::try_from(header_length_u64)?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflow"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("synthetic header end overflow"))?;
    if header_end != data_offset {
        return Err(io::Error::other("synthetic object offsets disagree").into());
    }

    let mut header = bytes
        .get(header_start..header_end)
        .ok_or_else(|| io::Error::other("synthetic header range is invalid"))?
        .to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut header, 99, 9_999)?;
    let mut with_unknown = Vec::with_capacity(bytes.len().saturating_add(8));
    with_unknown.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut with_unknown, u64::try_from(header.len())?);
    with_unknown.extend_from_slice(&header);
    with_unknown.extend_from_slice(&bytes[data_offset..]);

    let reparsed = Archive::parse(&with_unknown)?;
    assert_eq!(reparsed.to_bytes()?, with_unknown);
    Ok(SnappyStream::compress(&with_unknown)?)
}

fn slide(identifier: u64, name: &str, note: Option<u64>) -> kn::SlideArchive {
    kn::SlideArchive {
        style: reference(identifier.saturating_add(1_000)),
        transition: kn::TransitionArchive::default(),
        name: Some(name.to_owned()),
        note: note.map(reference),
        in_document: true,
        ..kn::SlideArchive::default()
    }
}

fn synthetic_package(dependent_note_content: bool) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(3), reference(30)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema requires legacy cache fields")]
    let first_node = kn::SlideNodeArchive {
        slide: Some(reference(FIRST_SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };
    #[allow(deprecated, reason = "native schema requires legacy cache fields")]
    let second_node = kn::SlideNodeArchive {
        slide: Some(reference(SECOND_SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };

    let note_payload = kn::NoteArchive {
        contained_storage: reference(NOTE_STORAGE),
    }
    .encode_to_vec();

    let mut storage = tswp::StorageArchive {
        text: vec!["Speaker ".to_owned(), "🚀 notes".to_owned()],
        ..tswp::StorageArchive::default()
    };
    if dependent_note_content {
        storage.table_attachment = Some(tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index: 8,
                object: Some(reference(500)),
            }],
        });
    }
    let mut storage_payload = storage.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut storage_payload, 98, 98_998)?;

    let mut note_object = ArchiveObject::new(
        NOTE_OBJECT,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-note-message".to_vec(),
            },
            RawMessage {
                type_: NOTE_MESSAGE_TYPE,
                data: note_payload,
            },
            RawMessage {
                type_: 778,
                data: b"after-note-message".to_vec(),
            },
        ],
    )?;
    note_object.archive_info.message_infos[1].object_references = vec![NOTE_STORAGE];
    let storage_object = ArchiveObject::new(
        NOTE_STORAGE,
        vec![
            RawMessage {
                type_: 779,
                data: b"before-storage-message".to_vec(),
            },
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: storage_payload,
            },
            RawMessage {
                type_: 780,
                data: b"after-storage-message".to_vec(),
            },
        ],
    )?;

    let mut first_slide_object = object(
        FIRST_SLIDE,
        5,
        &slide(FIRST_SLIDE, "Agenda", Some(NOTE_OBJECT)),
    )?;
    first_slide_object.archive_info.message_infos[0].object_references = vec![NOTE_OBJECT];
    let document_component = component_with_unknown_header(
        vec![
            object(1, 1, &document)?,
            object(2, 2, &show)?,
            object(3, 4, &first_node)?,
            first_slide_object,
            note_object,
            storage_object,
            object(30, 4, &second_node)?,
            object(SECOND_SLIDE, 5, &slide(SECOND_SLIDE, "No Notes", None))?,
            ArchiveObject::new(
                500,
                vec![RawMessage {
                    type_: 999,
                    data: PRIVATE_MARKER.to_vec(),
                }],
            )?,
        ],
        NOTE_STORAGE,
    )?;
    let unrelated_component = component(vec![ArchiveObject::new(
        900,
        vec![RawMessage {
            type_: 999,
            data: b"unrelated-iwa-component".to_vec(),
        }],
    )?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            ("Index/Unrelated.iwa", unrelated_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn rewrite_document(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let mut archive = Archive::parse(&stream)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn with_overlong_object_length_prefix(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let mut stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote object"))?;
    let offset = usize::try_from(object.header_offset)?;
    let (_length, prefix_bytes) = decode_varint_from_bytes(&stream[offset..])?;
    if prefix_bytes != 1 {
        return Err(io::Error::other("synthetic prefix is not one byte").into());
    }
    stream[offset] |= 0x80;
    stream.insert(offset + 1, 0);
    Archive::parse(&stream)?;
    let compressed = SnappyStream::compress(&stream)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote object"))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote message"))?;
    Ok(message.data.clone())
}

fn object_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote object"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&stream[header_offset..])?;
    let start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflow"))?;
    let end = start
        .checked_add(usize::try_from(header_length)?)
        .ok_or_else(|| io::Error::other("synthetic header end overflow"))?;
    assert_eq!(end, data_offset);
    Ok(stream[start..end].to_vec())
}

fn assert_only_document_payload_changed(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let before_entries = before_catalog.iter().collect::<Vec<_>>();
    let after_entries = after_catalog.iter().collect::<Vec<_>>();
    assert_eq!(before_entries.len(), after_entries.len());
    let mut changed = 0usize;
    for (before_entry, after_entry) in before_entries.into_iter().zip(after_entries) {
        assert_eq!(before_entry.name(), after_entry.name());
        assert_eq!(before_entry.raw_name(), after_entry.raw_name());
        if before_entry.name() == DOCUMENT_MEMBER {
            assert_ne!(before_entry.data(), after_entry.data());
            changed += 1;
        } else {
            assert_eq!(before_entry.data(), after_entry.data());
            assert_eq!(before_entry.metadata(), after_entry.metadata());
            assert_eq!(
                before_entry.raw_record().local_record(),
                after_entry.raw_record().local_record()
            );
        }
    }
    assert_eq!(changed, 1);
    Ok(())
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn assert_send_sync<T: Send + Sync>(_: &T) {}
fn assert_type_send_sync<T: Send + Sync>() {}

#[test]
fn synthetic_notes_fixture_reads_semantically() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    let show = package.show()?;
    assert_eq!(show.slide_count(), 2);
    assert_eq!(show.slides()[0].name(), Some("Agenda"));
    assert_eq!(show.slides()[0].notes(), Some("Speaker 🚀 notes"));
    assert_eq!(show.slides()[1].name(), Some("No Notes"));
    assert_eq!(show.slides()[1].notes(), None);
    assert!(package.text()?.ends_with("Speaker 🚀 notes"));
    Ok(())
}

#[test]
fn selector_first_utf16_notes_operations_are_semantic_and_transactional() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.slide_notes("Agenda")?,
        Some("Speaker 🚀 notes".to_owned())
    );
    assert_eq!(
        package.slide_notes(SlideSelector::index(0))?,
        Some("Speaker 🚀 notes".to_owned())
    );
    assert_eq!(package.slide_notes("No Notes")?, None);

    let source_pointer = package.source_bytes().as_ptr();
    let untouched = package.show()?.slides()[1].clone();
    let span = TextSpan::from_utf16_indexes(8, 10)?;
    let mut edit = package.edit_slide_notes("Agenda")?;
    assert_eq!(edit.position(), Position::new(0));
    assert_eq!(edit.text(), "Speaker 🚀 notes");
    edit.replace(span, "東京😀")?;
    assert_eq!(edit.span(), Some(span));
    let commit = edit.commit()?;
    assert_eq!(commit.patch().position(), Position::new(0));
    assert_eq!(commit.patch().span(), span);
    assert_eq!(commit.patch().before(), "Speaker 🚀 notes");
    assert_eq!(commit.patch().after(), "Speaker 東京😀 notes");
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit.package().slide_notes(SlideSelector::index(0))?,
        Some("Speaker 東京😀 notes".to_owned())
    );
    assert_eq!(commit.package().show()?.slides()[1], untouched);
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);

    let mut insert = package.edit_slide_notes(SlideSelector::index(0))?;
    insert.insert(TextPosition::ZERO, "Intro: ")?;
    assert_eq!(
        insert.commit()?.package().slide_notes("Agenda")?,
        Some("Intro: Speaker 🚀 notes".to_owned())
    );

    let mut delete = package.edit_slide_notes("Agenda")?;
    delete.delete(TextSpan::from_utf16_indexes(8, 10)?)?;
    assert_eq!(
        delete.commit()?.package().slide_notes("Agenda")?,
        Some("Speaker  notes".to_owned())
    );

    let mut set = package.edit_slide_notes("Agenda")?;
    set.set("Entirely new notes")?;
    assert_eq!(
        set.commit()?.package().slide_notes("Agenda")?,
        Some("Entirely new notes".to_owned())
    );

    let mut clear = package.edit_slide_notes("Agenda")?;
    clear.clear()?;
    let cleared = clear.commit()?;
    assert_eq!(
        cleared.package().slide_notes("Agenda")?,
        Some(String::new())
    );
    assert_eq!(cleared.package().show()?.slides()[0].notes(), None);
    let mut refill = cleared.package().edit_slide_notes("Agenda")?;
    assert_eq!(refill.text(), "");
    refill.set("Restored from an existing empty graph")?;
    assert_eq!(
        refill.commit()?.package().slide_notes("Agenda")?,
        Some("Restored from an existing empty graph".to_owned())
    );
    Ok(())
}

#[test]
fn selectors_utf16_boundaries_and_staging_errors_leave_source_unchanged() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();

    assert!(matches!(
        package.edit_slide_notes("Missing"),
        Err(SlideNotesError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.edit_slide_notes(SlideSelector::index(2)),
        Err(SlideNotesError::SlidePositionNotFound { position })
            if position == Position::new(2)
    ));
    assert!(matches!(
        package.edit_slide_notes("No Notes"),
        Err(SlideNotesError::NotesStorageNotFound)
    ));

    let mut split_start = package.edit_slide_notes("Agenda")?;
    assert!(matches!(
        split_start.replace(TextSpan::from_utf16_indexes(9, 10)?, "x"),
        Err(SlideNotesError::SurrogateBoundary { position })
            if position == TextPosition::from_utf16_code_units(9)
    ));
    let mut split_end = package.edit_slide_notes("Agenda")?;
    assert!(matches!(
        split_end.delete(TextSpan::from_utf16_indexes(8, 9)?),
        Err(SlideNotesError::SurrogateBoundary { position })
            if position == TextPosition::from_utf16_code_units(9)
    ));
    let mut out_of_bounds = package.edit_slide_notes("Agenda")?;
    assert!(matches!(
        out_of_bounds.delete(TextSpan::from_utf16_indexes(0, 100)?),
        Err(SlideNotesError::SpanOutOfBounds { length, .. })
            if length == TextPosition::from_utf16_code_units(16)
    ));
    let mut marker = package.edit_slide_notes("Agenda")?;
    assert!(matches!(
        marker.insert(TextPosition::ZERO, "bad\u{fffc}marker"),
        Err(SlideNotesError::ObjectMarkerReplacement)
    ));
    let mut one_operation = package.edit_slide_notes("Agenda")?;
    one_operation.insert(TextPosition::ZERO, "first")?;
    assert!(matches!(
        one_operation.insert(TextPosition::ZERO, "second"),
        Err(SlideNotesError::OperationAlreadyStaged)
    ));

    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);
    Ok(())
}

#[test]
fn semantic_noops_share_the_exact_source_allocation() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();

    let mut edit = package.edit_slide_notes("Agenda")?;
    edit.set("Speaker 🚀 notes")?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.package().source_bytes(), bytes);
    assert_eq!(commit.package().source_bytes().as_ptr(), source_pointer);
    let applied = package.apply_slide_notes(commit.patch())?;
    assert!(applied.patch().is_noop());
    assert_eq!(applied.package().source_bytes().as_ptr(), source_pointer);

    let unstaged = package.edit_slide_notes("Agenda")?.commit()?;
    assert!(unstaged.patch().is_noop());
    assert_eq!(unstaged.package().source_bytes().as_ptr(), source_pointer);
    Ok(())
}

#[test]
fn changed_notes_preserve_unknowns_scope_and_exact_inverse_application() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let source_note = message_payload(&bytes, NOTE_OBJECT, NOTE_MESSAGE_TYPE)?;
    let source_storage = message_payload(&bytes, NOTE_STORAGE, STORAGE_MESSAGE_TYPE)?;
    let source_before_note = message_payload(&bytes, NOTE_OBJECT, 777)?;
    let source_after_note = message_payload(&bytes, NOTE_OBJECT, 778)?;
    let source_before_storage = message_payload(&bytes, NOTE_STORAGE, 779)?;
    let source_after_storage = message_payload(&bytes, NOTE_STORAGE, 780)?;
    let source_header = object_header(&bytes, NOTE_STORAGE)?;

    let mut edit = package.edit_slide_notes("Agenda")?;
    edit.replace(TextSpan::from_utf16_indexes(8, 10)?, "東京😀")?;
    let edit_debug = format!("{edit:?}");
    assert!(!edit_debug.contains("Speaker"));
    assert!(!edit_debug.contains("Agenda"));
    assert!(!edit_debug.contains("Index/"));
    let commit = edit.commit()?;
    let target = commit.package().source_bytes();
    assert_only_document_payload_changed(&bytes, target)?;
    assert_eq!(
        message_payload(target, NOTE_OBJECT, NOTE_MESSAGE_TYPE)?,
        source_note
    );
    assert_eq!(
        message_payload(target, NOTE_OBJECT, 777)?,
        source_before_note
    );
    assert_eq!(
        message_payload(target, NOTE_OBJECT, 778)?,
        source_after_note
    );
    assert_eq!(
        message_payload(target, NOTE_STORAGE, 779)?,
        source_before_storage
    );
    assert_eq!(
        message_payload(target, NOTE_STORAGE, 780)?,
        source_after_storage
    );
    let target_storage = message_payload(target, NOTE_STORAGE, STORAGE_MESSAGE_TYPE)?;
    assert_eq!(
        raw_fields(&target_storage, 98)?,
        raw_fields(&source_storage, 98)?
    );
    let target_header = object_header(target, NOTE_STORAGE)?;
    assert_eq!(
        raw_fields(&target_header, 99)?,
        raw_fields(&source_header, 99)?
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);

    let applied = package.apply_slide_notes(commit.patch())?;
    assert_eq!(applied.package().source_bytes(), target);
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), commit.patch().after());
    assert_eq!(inverse.after(), commit.patch().before());
    assert_eq!(inverse.inverse(), commit.patch().clone());
    let restored = commit.package().apply_slide_notes(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(
        restored.package().slide_notes("Agenda")?,
        Some("Speaker 🚀 notes".to_owned())
    );

    let catalog = Catalog::from_bytes(&bytes)?;
    let equivalent_bytes = catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            "Data/sentinel.bin",
            b"different unrelated ZIP sentinel",
        )],
        Limits::default(),
    )?;
    let equivalent = Package::from_bytes(&equivalent_bytes)?;
    assert_eq!(
        equivalent.slide_notes("Agenda")?,
        package.slide_notes("Agenda")?
    );
    assert!(matches!(
        equivalent.apply_slide_notes(commit.patch()),
        Err(SlideNotesError::PatchConflict)
    ));

    let patch_debug = format!("{:?}", commit.patch());
    assert!(!patch_debug.contains("Speaker"));
    assert!(!patch_debug.contains("Agenda"));
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains(std::str::from_utf8(PRIVATE_MARKER)?));
    assert_send_sync(&package);
    assert_send_sync(&commit);
    assert_send_sync(commit.patch());
    assert_send_sync(commit.diagnostics());
    assert_type_send_sync::<SlideNotesCommit>();
    assert_type_send_sync::<SlideNotesPatch>();
    assert_type_send_sync::<SlideNotesDiagnostics>();
    assert_type_send_sync::<SlideNotesError>();
    Ok(())
}

#[test]
fn duplicate_names_missing_graph_and_malformed_references_fail_closed() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let duplicate_name = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(SECOND_SLIDE)
            .ok_or_else(|| io::Error::other("missing second slide"))?;
        let mut value = kn::SlideArchive::decode(object.messages[0].data.as_slice())?;
        value.name = Some("Agenda".to_owned());
        object.replace_message(
            0,
            RawMessage {
                type_: 5,
                data: value.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let duplicate_name_package = Package::from_bytes(&duplicate_name)?;
    assert!(matches!(
        duplicate_name_package.edit_slide_notes("Agenda"),
        Err(SlideNotesError::AmbiguousSelector)
    ));

    let malformed = [
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(FIRST_SLIDE)
                .ok_or_else(|| io::Error::other("missing first slide"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == 5)
                .ok_or_else(|| io::Error::other("missing slide message"))?;
            let mut data = object.messages[index].data.clone();
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut data,
                27,
                &reference(NOTE_OBJECT).encode_to_vec(),
            )?;
            object.replace_message_preserving_header(index, RawMessage { type_: 5, data })?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            archive
                .objects
                .retain(|object| object.archive_info.identifier != Some(NOTE_OBJECT));
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(NOTE_OBJECT)
                .ok_or_else(|| io::Error::other("missing note object"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == NOTE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing note message"))?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: NOTE_MESSAGE_TYPE,
                    data: Vec::new(),
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(NOTE_OBJECT)
                .ok_or_else(|| io::Error::other("missing note object"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == NOTE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing note message"))?;
            let mut data = object.messages[index].data.clone();
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut data,
                1,
                &reference(NOTE_STORAGE).encode_to_vec(),
            )?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: NOTE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(NOTE_OBJECT)
                .ok_or_else(|| io::Error::other("missing note object"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == NOTE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing note message"))?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: NOTE_MESSAGE_TYPE,
                    data: vec![0x0a, 0x00],
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(NOTE_OBJECT)
                .ok_or_else(|| io::Error::other("missing note object"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == NOTE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing note message"))?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: NOTE_MESSAGE_TYPE,
                    data: kn::NoteArchive {
                        contained_storage: reference(999),
                    }
                    .encode_to_vec(),
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            archive
                .objects
                .retain(|object| object.archive_info.identifier != Some(NOTE_STORAGE));
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(NOTE_STORAGE)
                .ok_or_else(|| io::Error::other("missing storage object"))?;
            let duplicate = object
                .messages
                .iter()
                .find(|message| message.type_ == STORAGE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing storage message"))?
                .clone();
            object.push_message(duplicate)?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(NOTE_STORAGE)
                .ok_or_else(|| io::Error::other("missing storage object"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == STORAGE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing storage message"))?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: STORAGE_MESSAGE_TYPE,
                    data: vec![0x1a, 0x80],
                },
            )?;
            Ok(())
        })?,
    ];

    for malformed_bytes in malformed {
        let package = Package::from_bytes(&malformed_bytes)?;
        let result = package.edit_slide_notes(SlideSelector::index(0));
        assert!(
            matches!(result, Err(SlideNotesError::InvalidSource)),
            "unexpected malformed-notes result: {result:?}"
        );
        assert_eq!(package.source_bytes(), malformed_bytes);
    }
    Ok(())
}

#[test]
fn note_graph_metadata_ownership_is_proven_before_rewriting() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let missing_slide_ownership = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        object.archive_info.message_infos[0]
            .object_references
            .clear();
        Ok(())
    })?;
    let missing_note_ownership = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(NOTE_OBJECT)
            .ok_or_else(|| io::Error::other("missing note object"))?;
        let message_info = object
            .archive_info
            .message_infos
            .iter_mut()
            .find(|info| info.type_ == NOTE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing note message metadata"))?;
        message_info.object_references.clear();
        Ok(())
    })?;
    let extra_owner = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(500)
            .ok_or_else(|| io::Error::other("missing synthetic dependent object"))?;
        object.archive_info.message_infos[0]
            .object_references
            .push(NOTE_OBJECT);
        Ok(())
    })?;

    for malformed_bytes in [missing_slide_ownership, missing_note_ownership] {
        let package = Package::from_bytes(&malformed_bytes)?;
        let mut edit = package.edit_slide_notes("Agenda")?;
        edit.set("ownership must be proven")?;
        assert!(matches!(edit.commit(), Err(SlideNotesError::InvalidSource)));
        assert_eq!(package.source_bytes(), malformed_bytes);
    }

    let package = Package::from_bytes(&extra_owner)?;
    let mut edit = package.edit_slide_notes("Agenda")?;
    edit.set("exclusive ownership is required")?;
    assert!(matches!(
        edit.commit(),
        Err(SlideNotesError::DependentContent)
    ));
    assert_eq!(package.source_bytes(), extra_owner);
    Ok(())
}

#[test]
fn metadata_and_payload_ownership_aliases_fail_closed() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let duplicate_same_message_metadata = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        object.archive_info.message_infos[0]
            .object_references
            .push(NOTE_OBJECT);
        Ok(())
    })?;
    let metadata_payload_mismatch = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(NOTE_OBJECT)
            .ok_or_else(|| io::Error::other("missing note object"))?;
        object.archive_info.message_infos[1].object_references = vec![999];
        Ok(())
    })?;
    let other_slide_alias = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(SECOND_SLIDE)
            .ok_or_else(|| io::Error::other("missing second slide"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == 5)
            .ok_or_else(|| io::Error::other("missing second slide message"))?;
        let mut value = kn::SlideArchive::decode(object.messages[index].data.as_slice())?;
        value.note = Some(reference(NOTE_OBJECT));
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 5,
                data: value.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let other_note_alias = rewrite_document(&bytes, |archive| {
        archive.objects.push(object(
            501,
            NOTE_MESSAGE_TYPE,
            &kn::NoteArchive {
                contained_storage: reference(NOTE_STORAGE),
            },
        )?);
        Ok(())
    })?;
    let selected_note_unknown_field = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(NOTE_OBJECT)
            .ok_or_else(|| io::Error::other("missing note object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == NOTE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing note message"))?;
        let mut data = object.messages[index].data.clone();
        litchi_iwa_common::wire::append_varint_field(&mut data, 99, 99_999)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: NOTE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;

    for adversarial in [
        duplicate_same_message_metadata,
        metadata_payload_mismatch,
        other_slide_alias,
        other_note_alias,
        selected_note_unknown_field,
    ] {
        let package = Package::from_bytes(&adversarial)?;
        let mut edit = package.edit_slide_notes("Agenda")?;
        edit.set("ownership ambiguity must not publish")?;
        assert!(matches!(
            edit.commit(),
            Err(SlideNotesError::DependentContent | SlideNotesError::InvalidSource)
        ));
        assert_eq!(package.source_bytes(), adversarial);
    }
    Ok(())
}

#[test]
fn exact_noop_and_replay_do_not_require_mutation_ownership() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let malformed_metadata = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        object.archive_info.message_infos[0]
            .object_references
            .clear();
        Ok(())
    })?;
    let package = Package::from_bytes(&malformed_metadata)?;
    let mut edit = package.edit_slide_notes("Agenda")?;
    edit.set("Speaker 🚀 notes")?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_noop());
    let replay = package.apply_slide_notes(commit.patch())?;
    assert!(replay.patch().is_noop());
    assert_eq!(replay.package().source_bytes(), malformed_metadata);
    Ok(())
}

#[test]
fn unrelated_noncanonical_object_prefix_is_refused_before_rewrite() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let noncanonical = with_overlong_object_length_prefix(&bytes, 500)?;
    let package = Package::from_bytes(&noncanonical)?;
    let mut edit = package.edit_slide_notes("Agenda")?;
    edit.set("must not canonicalize an unrelated object frame")?;
    assert!(matches!(edit.commit(), Err(SlideNotesError::InvalidSource)));
    assert_eq!(package.source_bytes(), noncanonical);
    Ok(())
}

#[test]
fn dependent_note_content_is_refused_without_blocking_unrelated_text() -> TestResult<()> {
    let bytes = synthetic_package(true)?;
    let package = Package::from_bytes(&bytes)?;
    let mut dependent = package.edit_slide_notes("Agenda")?;
    dependent.delete(TextSpan::from_utf16_indexes(8, 10)?)?;
    assert!(matches!(
        dependent.commit(),
        Err(SlideNotesError::DependentContent)
    ));
    assert_eq!(package.source_bytes(), bytes);

    let mut unrelated = package.edit_slide_notes("Agenda")?;
    unrelated.replace(TextSpan::from_utf16_indexes(0, 1)?, "X")?;
    let commit = unrelated.commit()?;
    assert_eq!(
        commit.package().slide_notes("Agenda")?,
        Some("Xpeaker 🚀 notes".to_owned())
    );
    Ok(())
}

#[test]
fn changed_notes_respect_the_retained_output_limit() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let input_bytes = u64::try_from(bytes.len())?;
    let limits = Limits::new(input_bytes, 16, 1024 * 1024, 1024 * 1024, 1024 * 1024)?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    let mut edit = package.edit_slide_notes("Agenda")?;
    edit.insert(TextPosition::ZERO, &"expanded".repeat(1_024))?;
    assert!(matches!(
        edit.commit(),
        Err(SlideNotesError::LimitExceeded {
            kind: SlideNotesLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}
