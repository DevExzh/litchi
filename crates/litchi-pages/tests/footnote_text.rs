use litchi_iwa_archive::{Limits as ArchiveLimits, package::to_bytes};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsp, tswp};
use litchi_pages::footnote::body::{Position, Selector};
use litchi_pages::{FootnoteTextError, Package};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_OBJECT: u64 = 1;
const BODY_OBJECT: u64 = 42;
const FIRST_REFERENCE_OBJECT: u64 = 100;
const FIRST_STORAGE_OBJECT: u64 = 101;
const FIRST_MARKER_OBJECT: u64 = 102;
const SECOND_REFERENCE_OBJECT: u64 = 110;
const SECOND_STORAGE_OBJECT: u64 = 111;
const SECOND_MARKER_OBJECT: u64 = 112;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;
const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;
const FOOTNOTE_CONTENT_PREFIX: &str = "\u{fffc} ";

struct Fixture {
    bytes: Vec<u8>,
    decompressed_iwa_bytes: usize,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(identifier: u64, message_type: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )?)
}

fn build_fixture(malformed_reference: bool) -> TestResult<Fixture> {
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["A😀\u{e}B\u{e}C".to_owned()],
        table_footnote: Some(tswp::ObjectAttributeTable {
            entries: vec![
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 3,
                    object: Some(reference(FIRST_REFERENCE_OBJECT)),
                },
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 5,
                    object: Some(reference(SECOND_REFERENCE_OBJECT)),
                },
            ],
        }),
        ..tswp::StorageArchive::default()
    };

    let note_objects = |reference_id: u64,
                        storage_id: u64,
                        marker_id: u64,
                        text: &str,
                        custom_mark: Option<&str>,
                        malformed: bool|
     -> TestResult<Vec<ArchiveObject>> {
        let mut reference_payload = tswp::FootnoteReferenceAttachmentArchive {
            super_: Some(tswp::TextualAttachmentArchive {
                string_equivalent: Some("*".to_owned()),
                kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
            }),
            contained_storage: Some(reference(storage_id)),
            custom_mark_string: custom_mark.map(str::to_owned),
        }
        .encode_to_vec();
        if malformed {
            // The package reader does not need to project the note graph to
            // open a source, while the focused selector must reject this
            // truncated length-delimited field before any edit is published.
            reference_payload.extend_from_slice(&[0xaa, 0x06, 0x03, b'r']);
        } else {
            // Unknown fields are intentional preservation witnesses for a
            // source-preserving reference rewrite.
            reference_payload
                .extend_from_slice(&[0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'r', b'e', b'f']);
        }

        let storage = tswp::StorageArchive {
            kind: Some(tswp::storage_archive::KindType::Footnote as i32),
            text: vec![format!("{FOOTNOTE_CONTENT_PREFIX}{text}")],
            table_attachment: Some(tswp::ObjectAttributeTable {
                entries: vec![tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(reference(marker_id)),
                }],
            }),
            ..tswp::StorageArchive::default()
        };

        let mut marker_payload = tswp::TextualAttachmentArchive {
            string_equivalent: Some("*".to_owned()),
            kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
        }
        .encode_to_vec();
        marker_payload.extend_from_slice(&[0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'm', b'a', b'r']);

        Ok(vec![
            object(
                reference_id,
                FOOTNOTE_REFERENCE_MESSAGE_TYPE,
                reference_payload,
            )?,
            object(storage_id, STORAGE_MESSAGE_TYPE, storage.encode_to_vec())?,
            object(marker_id, TEXTUAL_ATTACHMENT_MESSAGE_TYPE, marker_payload)?,
        ])
    };

    let mut objects = vec![
        object(
            DOCUMENT_OBJECT,
            DOCUMENT_MESSAGE_TYPE,
            tp::DocumentArchive {
                body_storage: Some(reference(BODY_OBJECT)),
                ..tp::DocumentArchive::default()
            }
            .encode_to_vec(),
        )?,
        object(BODY_OBJECT, STORAGE_MESSAGE_TYPE, body.encode_to_vec())?,
    ];
    objects.extend(note_objects(
        FIRST_REFERENCE_OBJECT,
        FIRST_STORAGE_OBJECT,
        FIRST_MARKER_OBJECT,
        "First",
        None,
        malformed_reference,
    )?);
    objects.extend(note_objects(
        SECOND_REFERENCE_OBJECT,
        SECOND_STORAGE_OBJECT,
        SECOND_MARKER_OBJECT,
        "Second",
        Some("†"),
        false,
    )?);

    let decompressed = Archive { objects }.to_bytes()?;
    let compressed = SnappyStream::compress(&decompressed)?;
    let bytes = to_bytes(
        [("Index/Document.iwa", compressed.as_slice())],
        ArchiveLimits::default(),
    )?;
    Ok(Fixture {
        bytes,
        decompressed_iwa_bytes: decompressed.len(),
    })
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn selectors_read_and_stage_text_and_custom_marker_presence() -> TestResult {
    let fixture = build_fixture(false)?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let footnotes = package.body_footnotes()?;

    assert_eq!(footnotes.len(), 2);
    assert_eq!(footnotes[0].position, Position::from_utf16_index(3)?);
    assert_eq!(footnotes[0].text.as_ref(), "First");
    assert_eq!(footnotes[0].custom_mark, None);
    assert_eq!(footnotes[1].position, Position::from_utf16_index(5)?);
    assert_eq!(footnotes[1].custom_mark.as_deref(), Some("†"));

    let by_index = package.edit_body_footnote_text(Selector::index(1))?;
    assert_eq!(by_index.position(), Position::from_utf16_index(5)?);
    assert_eq!(by_index.before(), &footnotes[1]);
    assert_eq!(by_index.custom_mark(), Some("†"));

    let mut by_position = package.edit_body_footnote_text(Selector::at(footnotes[1].position))?;
    by_position
        .set("UpdatedSecond")?
        .set_custom_mark(Some("‡"))?;
    assert_eq!(by_position.text(), "UpdatedSecond");
    assert_eq!(by_position.custom_mark(), Some("‡"));
    assert!(matches!(
        package.edit_body_footnote_text(Selector::index(2)),
        Err(FootnoteTextError::NotFound)
    ));
    Ok(())
}

#[test]
fn exact_noop_changed_inverse_and_patch_conflict_preserve_bytes() -> TestResult {
    let fixture = build_fixture(false)?;
    let package = Package::from_bytes(&fixture.bytes)?;

    let mut noop = package.edit_body_footnote_text(Selector::index(1))?;
    noop.set("Second")?.set_custom_mark(Some("†"))?;
    let noop_commit = noop.commit()?;
    assert!(!noop_commit.diagnostics().changed());
    assert!(noop_commit.patch().is_noop());
    assert_eq!(exact_bytes(noop_commit.package())?, fixture.bytes);

    let mut edit = package.edit_body_footnote_text(Selector::index(1))?;
    edit.set("UpdatedSecond")?.set_custom_mark(Some("‡"))?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_ne!(exact_bytes(commit.package())?, fixture.bytes);
    assert_eq!(commit.patch().before().text.as_ref(), "Second");
    assert_eq!(commit.patch().after().text.as_ref(), "UpdatedSecond");
    assert_eq!(commit.patch().after().custom_mark.as_deref(), Some("‡"));

    let conflict = commit
        .package()
        .apply_body_footnote_text(commit.patch())
        .expect_err("a patch cannot be applied to its already-published target");
    assert!(matches!(conflict, FootnoteTextError::PatchConflict));

    let restored = commit
        .package()
        .apply_body_footnote_text(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, fixture.bytes);
    assert_eq!(
        restored.package().body_footnotes()?,
        package.body_footnotes()?
    );
    Ok(())
}

#[test]
fn malformed_graph_rejects_selection_without_mutating_source() -> TestResult {
    let fixture = build_fixture(true)?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let before = exact_bytes(&package)?;

    let error = package
        .edit_body_footnote_text(Selector::index(0))
        .expect_err("a truncated reference field must not stage an edit");
    assert!(matches!(error, FootnoteTextError::InvalidSource));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn low_iwa_limit_rejects_commit_atomically() -> TestResult {
    let fixture = build_fixture(false)?;
    let defaults = ArchiveLimits::default();
    let limits = litchi_iwa_archive::Limits::new(
        u64::try_from(fixture.bytes.len())?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        fixture.decompressed_iwa_bytes,
    )?;
    let package = Package::from_bytes_with_limits(&fixture.bytes, limits)?;
    let before = exact_bytes(&package)?;

    let mut edit = package.edit_body_footnote_text(Selector::index(1))?;
    edit.set("This replacement is deliberately longer than the source note")?;
    let error = edit
        .commit()
        .expect_err("the original decompressed-IWA ceiling must reject growth");
    assert!(matches!(error, FootnoteTextError::LimitExceeded { .. }));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}
