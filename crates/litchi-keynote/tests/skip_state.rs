use std::io;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{
    decode_varint_from_bytes, encode_varint_into,
    wire::{WireView, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use litchi_keynote::{EditError, Limits, Package, ReadError, SlideSelector, SlideSelectorError};
use prost::Message;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const FIRST_NODE: u64 = 3;

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

trait ExactPackageBytes {
    fn exact_bytes(&self) -> &'static [u8];
}

impl ExactPackageBytes for Package {
    fn exact_bytes(&self) -> &'static [u8] {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts every package byte");
        Box::leak(bytes.into_boxed_slice())
    }
}

#[derive(Debug, Clone, Copy)]
enum FirstNodeEncoding {
    Canonical,
    MissingField,
    DuplicateField,
    WrongWireType,
    NoncanonicalKey,
    NoncanonicalValue,
    InvalidBoolean,
    DuplicateTypePayload,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    let archive = Archive { objects };
    let bytes = archive.to_bytes()?;
    Ok(SnappyStream::compress(&bytes)?)
}

fn component_with_unknown_header(
    objects: Vec<ArchiveObject>,
    target_identifier: u64,
) -> TestResult<Vec<u8>> {
    let archive = Archive { objects };
    let bytes = archive.to_bytes()?;
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
    append_varint_field(&mut header, 99, 9_999)?;

    let mut with_unknown = Vec::with_capacity(bytes.len().saturating_add(8));
    with_unknown.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut with_unknown, u64::try_from(header.len())?);
    with_unknown.extend_from_slice(&header);
    with_unknown.extend_from_slice(&bytes[data_offset..]);

    let reparsed = Archive::parse(&with_unknown)?;
    assert_eq!(reparsed.to_bytes()?, with_unknown);
    Ok(SnappyStream::compress(&with_unknown)?)
}

#[allow(deprecated, reason = "native schema requires legacy cache fields")]
fn canonical_node(slide_identifier: u64, is_skipped: bool) -> TestResult<Vec<u8>> {
    let mut data = kn::SlideNodeArchive {
        slide: Some(reference(slide_identifier)),
        is_skipped,
        has_builds: false,
        has_transition: false,
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut data, 99, 73)?;
    Ok(data)
}

fn replace_field_four(source: &[u8], replacement: Option<&[u8]>) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::with_capacity(source.len().saturating_add(3));
    for field in view.fields() {
        if field.number() == 4 {
            if let Some(replacement_bytes) = replacement {
                output.extend_from_slice(replacement_bytes);
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    Ok(output)
}

fn first_node_payload(encoding: FirstNodeEncoding) -> TestResult<Vec<u8>> {
    let canonical = canonical_node(4, false)?;
    Ok(match encoding {
        FirstNodeEncoding::Canonical | FirstNodeEncoding::DuplicateTypePayload => canonical,
        FirstNodeEncoding::MissingField => replace_field_four(&canonical, None)?,
        FirstNodeEncoding::DuplicateField => {
            let mut duplicate = canonical;
            duplicate.extend_from_slice(&[0x20, 0x01]);
            duplicate
        },
        FirstNodeEncoding::WrongWireType => {
            replace_field_four(&canonical, Some(&[0x22, 0x01, 0x00]))?
        },
        FirstNodeEncoding::NoncanonicalKey => {
            replace_field_four(&canonical, Some(&[0xa0, 0x00, 0x00]))?
        },
        FirstNodeEncoding::NoncanonicalValue => {
            replace_field_four(&canonical, Some(&[0x20, 0x80, 0x00]))?
        },
        FirstNodeEncoding::InvalidBoolean => replace_field_four(&canonical, Some(&[0x20, 0x02]))?,
    })
}

fn slide(name: &str) -> Vec<u8> {
    kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        name: Some(name.to_owned()),
        in_document: true,
        ..Default::default()
    }
    .encode_to_vec()
}

fn synthetic_package(
    encoding: FirstNodeEncoding,
    first_name: &str,
    second_name: &str,
) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(2),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(FIRST_NODE), reference(5)],
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..Default::default()
    };
    let first_payload = first_node_payload(encoding)?;
    let mut first_messages = vec![RawMessage {
        type_: 4,
        data: first_payload.clone(),
    }];
    if matches!(encoding, FirstNodeEncoding::DuplicateTypePayload) {
        first_messages.push(RawMessage {
            type_: 4,
            data: first_payload,
        });
    }
    let document_component = component_with_unknown_header(
        vec![
            object(1, 1, document.encode_to_vec())?,
            object(2, 2, show.encode_to_vec())?,
            ArchiveObject::new(FIRST_NODE, first_messages)?,
            object(5, 4, canonical_node(6, true)?)?,
        ],
        FIRST_NODE,
    )?;
    let first_slide = component(vec![object(4, 5, slide(first_name))?])?;
    let second_slide = component(vec![object(6, 5, slide(second_name))?])?;
    let untouched = b"opaque synthetic sentinel";

    // Keep the edited component last so every untouched local offset and
    // central-directory record can remain byte-for-byte identical.
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", untouched.as_slice()),
            ("Index/Slide-4.iwa", first_slide.as_slice()),
            ("Index/Slide-6.iwa", second_slide.as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn node_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("synthetic package has no document member"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    Ok(archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("synthetic package has no requested slide node"))?
        .messages
        .iter()
        .find(|message| message.type_ == 4)
        .ok_or_else(|| io::Error::other("synthetic slide node has no type-4 message"))?
        .data
        .clone())
}

fn decompressed_document(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("synthetic package has no document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.as_bytes().to_vec())
}

fn fields_except_skip(data: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(data)?
        .fields()
        .filter(|field| field.number() != 4)
        .map(|field| field.raw().to_vec())
        .collect())
}

#[test]
fn semantic_state_and_exact_noop_are_selector_first() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = synthetic_package(FirstNodeEncoding::Canonical, "Intro", "Agenda")?;
    let package = Package::from_bytes(&bytes)?;
    let show = package.show()?;
    assert_eq!(show.slides().len(), 2);
    assert!(!show.slides()[0].is_skipped());
    assert!(show.slides()[1].is_skipped());
    assert_eq!(show.select_slide("Agenda")?, Some(&show.slides()[1]));

    let source_snapshot = package.exact_bytes();
    let mut edit = package.edit();
    edit.include_slide(SlideSelector::name("Intro"))?;
    let commit = edit.commit()?;
    assert_eq!(commit.package().exact_bytes(), bytes);
    assert_eq!(commit.package().exact_bytes(), source_snapshot);
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(package.exact_bytes(), bytes);

    assert!(matches!(
        package.edit().commit(),
        Err(EditError::NoStagedOperation)
    ));
    let mut bounded = package.edit();
    bounded.skip_slide(SlideSelector::index(0))?;
    assert!(matches!(
        bounded.include_slide(SlideSelector::index(1)),
        Err(EditError::OperationAlreadyStaged)
    ));
    let bounded_commit = bounded.commit()?;
    assert!(bounded_commit.package().show()?.slides()[0].is_skipped());
    assert!(bounded_commit.package().show()?.slides()[1].is_skipped());
    assert_eq!(package.exact_bytes(), bytes);
    Ok(())
}

#[test]
fn selector_misses_and_ambiguity_are_typed() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = synthetic_package(FirstNodeEncoding::Canonical, "Same", "Same")?;
    let package = Package::from_bytes(&bytes)?;

    let mut by_position = package.edit();
    assert!(matches!(
        by_position.skip_slide(SlideSelector::position(Position::new(9))),
        Err(EditError::SlidePositionNotFound { position }) if position == Position::new(9)
    ));

    let mut by_name = package.edit();
    let Err(missing_name) = by_name.skip_slide("Missing") else {
        return Err(io::Error::other("missing name was accepted").into());
    };
    assert!(matches!(&missing_name, EditError::SlideNameNotFound));
    assert!(!missing_name.to_string().contains("Missing"));

    let mut ambiguous = package.edit();
    assert!(matches!(
        ambiguous.skip_slide("Same"),
        Err(EditError::Selector(SlideSelectorError::DuplicateSlideName { name }))
            if name.as_ref() == "Same"
    ));
    Ok(())
}

#[test]
fn commit_preserves_unrelated_wire_and_zip_bytes_and_patch_is_reversible()
-> Result<(), Box<dyn std::error::Error>> {
    let bytes = synthetic_package(FirstNodeEncoding::Canonical, "Intro", "Agenda")?;
    let package = Package::from_bytes(&bytes)?;
    let source_copy = package.exact_bytes().to_vec();
    let before_node = node_payload(&bytes, FIRST_NODE)?;

    let mut edit = package.edit();
    edit.skip_slide(0usize)?;
    let commit = edit.commit()?;
    assert!(commit.package().show()?.slides()[0].is_skipped());
    assert_eq!(package.exact_bytes(), source_copy);
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().position(), Position::new(0));
    assert!(!commit.patch().before());
    assert!(commit.patch().after());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());

    let after_node = node_payload(commit.package().exact_bytes(), FIRST_NODE)?;
    assert_eq!(
        fields_except_skip(&after_node)?,
        fields_except_skip(&before_node)?
    );
    let unknown_before = WireView::parse(&before_node)?
        .fields()
        .find(|field| field.number() == 99)
        .ok_or_else(|| io::Error::other("source node has no unknown sentinel"))?
        .raw()
        .to_vec();
    let unknown_after = WireView::parse(&after_node)?
        .fields()
        .find(|field| field.number() == 99)
        .ok_or_else(|| io::Error::other("target node has no unknown sentinel"))?
        .raw()
        .to_vec();
    assert_eq!(unknown_after, unknown_before);

    let decompressed_before = decompressed_document(&bytes)?;
    let decompressed_after = decompressed_document(commit.package().exact_bytes())?;
    assert_eq!(decompressed_after.len(), decompressed_before.len());
    let changed_iwa_bytes = decompressed_before
        .iter()
        .zip(&decompressed_after)
        .filter(|(before, after)| before != after)
        .map(|(before, after)| (*before, *after))
        .collect::<Vec<_>>();
    assert_eq!(changed_iwa_bytes, [(0, 1)]);

    let before_catalog = Catalog::from_bytes(&bytes)?;
    let after_catalog = Catalog::from_bytes(commit.package().exact_bytes())?;
    let mut changed = 0;
    for (before, after) in before_catalog.iter().zip(after_catalog.iter()) {
        assert_eq!(before.name(), after.name());
        assert_eq!(before.raw_name(), after.raw_name());
        if before.data() == after.data() {
            assert_eq!(before.metadata(), after.metadata());
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record()
            );
            assert_eq!(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record()
            );
        } else {
            changed += 1;
            assert_eq!(before.name(), DOCUMENT_MEMBER);
        }
    }
    assert_eq!(changed, 1);

    let applied = package.apply(commit.patch())?;
    assert_eq!(
        applied.package().exact_bytes(),
        commit.package().exact_bytes()
    );
    let inverse = commit.patch().inverse();
    assert!(inverse.before());
    assert!(!inverse.after());
    assert_eq!(
        inverse.source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(inverse.inverse(), commit.patch().clone());
    let reverted = commit.package().apply(&inverse)?;
    assert_eq!(reverted.package().exact_bytes(), bytes);
    assert!(!reverted.package().show()?.slides()[0].is_skipped());
    let forwarded_again = reverted.package().apply(commit.patch())?;
    assert_eq!(
        forwarded_again.package().exact_bytes(),
        commit.package().exact_bytes()
    );
    assert_eq!(forwarded_again.patch(), commit.patch());

    let unrelated_bytes = synthetic_package(FirstNodeEncoding::Canonical, "Different", "Agenda")?;
    let unrelated = Package::from_bytes(&unrelated_bytes)?;
    assert!(matches!(
        unrelated.apply(commit.patch()),
        Err(EditError::PatchConflict)
    ));
    let redacted_debug = format!("{:?}", commit.patch());
    assert!(!redacted_debug.contains("Index/"));
    assert!(!redacted_debug.contains("fingerprint"));
    assert!(!redacted_debug.contains("bytes"));
    Ok(())
}

#[test]
fn include_commit_changes_true_to_false_and_is_deterministic()
-> Result<(), Box<dyn std::error::Error>> {
    let bytes = synthetic_package(FirstNodeEncoding::Canonical, "Intro", "Agenda")?;
    let package = Package::from_bytes(&bytes)?;
    assert!(package.show()?.slides()[1].is_skipped());

    let mut edit = package.edit();
    edit.include_slide(1usize)?;
    let commit = edit.commit()?;
    assert!(!commit.package().show()?.slides()[1].is_skipped());
    assert_eq!(package.exact_bytes(), bytes);
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().position(), Position::new(1));
    assert!(commit.patch().before());
    assert!(!commit.patch().after());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());

    let mut repeated_edit = package.edit();
    repeated_edit.include_slide(1usize)?;
    let repeated = repeated_edit.commit()?;
    assert_eq!(
        repeated.package().exact_bytes(),
        commit.package().exact_bytes()
    );
    assert_eq!(repeated.patch(), commit.patch());
    assert_eq!(repeated.diagnostics(), commit.diagnostics());
    Ok(())
}

#[test]
fn malformed_required_skip_field_and_payload_ambiguity_are_rejected()
-> Result<(), Box<dyn std::error::Error>> {
    for encoding in [
        FirstNodeEncoding::MissingField,
        FirstNodeEncoding::DuplicateField,
        FirstNodeEncoding::WrongWireType,
        FirstNodeEncoding::NoncanonicalKey,
        FirstNodeEncoding::NoncanonicalValue,
        FirstNodeEncoding::InvalidBoolean,
        FirstNodeEncoding::DuplicateTypePayload,
    ] {
        let bytes = synthetic_package(encoding, "Intro", "Agenda")?;
        let package = Package::from_bytes(&bytes)?;
        assert!(
            matches!(
                package.show(),
                Err(ReadError::Decode(_) | ReadError::InvalidFormat(_))
            ),
            "malformed encoding should fail strict semantic decode: {encoding:?}"
        );
    }
    Ok(())
}

#[test]
fn public_transaction_values_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<litchi_keynote::Commit>();
    assert_send_sync::<litchi_keynote::Patch>();
    assert_send_sync::<litchi_keynote::Diagnostics>();
    assert_send_sync::<EditError>();

    // Keep the exact-source Arc implementation observable to the compiler:
    // the public Patch itself must remain cheap to clone across task boundaries.
    assert_send_sync::<Arc<[u8]>>();
}
