use std::io;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsp};
use litchi_pages::{BodyDrawableOrderError, DrawableLayerMove, Limits, Package, Position};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SENTINEL_MEMBER: &str = "Data/drawable-order-sentinel.bin";
const PREVIEW_MEMBERS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const ROOT_IDENTIFIER: u64 = 1;
const ZORDER_IDENTIFIER: u64 = 46;
const BODY_STORAGE_IDENTIFIER: u64 = 47;
const DRAWABLE_IDENTIFIERS: [u64; 3] = [101, 102, 103];

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn order_payload(order: &[u64]) -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    for (index, identifier) in order.iter().copied().enumerate() {
        let mut nested = reference(identifier).encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(
            &mut nested,
            99,
            u64::try_from(index)?.saturating_add(700),
        )?;
        litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 1, &nested)?;
    }
    // The root unknown field must remain byte-for-byte untouched by a reorder.
    litchi_iwa_common::wire::append_varint_field(&mut payload, 77, 0xfeed)?;
    Ok(payload)
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn fixture_bytes(order: &[u64], sentinel: &[u8]) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        drawables_zorder: Some(reference(ZORDER_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let objects = vec![
        object(ROOT_IDENTIFIER, 10_000, root.encode_to_vec())?,
        object(ZORDER_IDENTIFIER, 10_015, order_payload(order)?)?,
        object(DRAWABLE_IDENTIFIERS[0], 9_000, b"back".to_vec())?,
        object(DRAWABLE_IDENTIFIERS[1], 9_000, b"middle".to_vec())?,
        object(DRAWABLE_IDENTIFIERS[2], 9_000, b"front".to_vec())?,
    ];
    let archive = Archive { objects };
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (SENTINEL_MEMBER, sentinel),
            (PREVIEW_MEMBERS[0], b"full preview".as_slice()),
            (PREVIEW_MEMBERS[1], b"micro preview".as_slice()),
            (PREVIEW_MEMBERS[2], b"web preview".as_slice()),
            (DOCUMENT_MEMBER, compressed.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn fixture_bytes_with_native_body_slot(sentinel: &[u8]) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_STORAGE_IDENTIFIER)),
        drawables_zorder: Some(reference(ZORDER_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let mut native_order = vec![BODY_STORAGE_IDENTIFIER];
    native_order.extend(DRAWABLE_IDENTIFIERS);
    let objects = vec![
        object(ROOT_IDENTIFIER, 10_000, root.encode_to_vec())?,
        object(ZORDER_IDENTIFIER, 10_015, order_payload(&native_order)?)?,
        object(BODY_STORAGE_IDENTIFIER, 2_001, Vec::new())?,
        object(DRAWABLE_IDENTIFIERS[0], 9_000, b"back".to_vec())?,
        object(DRAWABLE_IDENTIFIERS[1], 9_000, b"middle".to_vec())?,
        object(DRAWABLE_IDENTIFIERS[2], 9_000, b"front".to_vec())?,
    ];
    let archive = Archive { objects };
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (SENTINEL_MEMBER, sentinel),
            (DOCUMENT_MEMBER, compressed.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn member_payload(package: &Package, member: &str) -> TestResult<Vec<u8>> {
    let bytes = exact_bytes(package)?;
    let catalog = Catalog::from_bytes(&bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("fixture member is missing"))?;
    Ok(entry.data().to_vec())
}

fn zorder_payload_from_package(package: &Package) -> TestResult<Vec<u8>> {
    let bytes = member_payload(package, DOCUMENT_MEMBER)?;
    let archive = Archive::parse(&SnappyStream::decompress(&bytes)?.into_bytes())?;
    let object = archive
        .object(ZORDER_IDENTIFIER)
        .ok_or_else(|| io::Error::other("fixture z-order object is missing"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == 10_015)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("fixture z-order message is missing").into())
}

#[test]
fn selector_first_move_preserves_unknowns_and_unrelated_members() -> TestResult<()> {
    let source_bytes = fixture_bytes(&DRAWABLE_IDENTIFIERS, b"unchanged sentinel")?;
    let source = Package::from_bytes(&source_bytes)?;
    let handles = source.body_drawable_order()?;
    assert_eq!(handles.len(), 3);
    assert_eq!(handles[0].position(), Position::new(0));
    assert!(!format!("{:?}", handles[0]).contains("101"));

    let mut edit = source.edit_body_drawable_order()?;
    assert!(edit.move_drawable(&handles[0], DrawableLayerMove::ToFront)?);
    let commit = edit.commit()?;
    assert_eq!(
        commit
            .package()
            .body_drawable_order()?
            .iter()
            .map(|handle| handle.position())
            .collect::<Vec<_>>(),
        [Position::new(0), Position::new(1), Position::new(2)]
    );
    assert!(commit.diagnostics().full_reparse_performed());
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.diagnostics().deleted_previews(),
        PREVIEW_MEMBERS.len()
    );

    let rewritten = zorder_payload_from_package(commit.package())?;
    let view = WireView::parse(&rewritten)?;
    let (snapshot, _report) =
        litchi_iwa_protos::pages_drawable_order_codec::decode_drawable_order_with_report(
            &rewritten,
            litchi_iwa_protos::pages_drawable_order_codec::DecodeOptions::for_source(&rewritten),
        )?;
    assert_eq!(
        snapshot.identifiers().collect::<Vec<_>>(),
        [
            DRAWABLE_IDENTIFIERS[1],
            DRAWABLE_IDENTIFIERS[2],
            DRAWABLE_IDENTIFIERS[0]
        ]
    );
    assert!(
        view.fields()
            .any(|field| field.number() == 77 && field.wire_type() == 0)
    );
    for field in view.fields().filter(|field| field.number() == 1) {
        assert!(
            WireView::parse(field.payload())?
                .fields()
                .any(|nested| nested.number() == 99)
        );
    }
    assert_eq!(
        member_payload(commit.package(), SENTINEL_MEMBER)?,
        b"unchanged sentinel"
    );
    let candidate_catalog = Catalog::from_bytes(&exact_bytes(commit.package())?)?;
    assert!(
        PREVIEW_MEMBERS
            .iter()
            .all(|name| candidate_catalog.iter().all(|entry| entry.name() != *name))
    );

    let restored = commit
        .package()
        .apply_body_drawable_order(&commit.patch().inverse())?;
    assert_eq!(restored.diagnostics().deleted_previews(), 0);
    assert_eq!(exact_bytes(&source)?, exact_bytes(restored.package())?);
    assert_eq!(
        source
            .body_drawable_order()?
            .iter()
            .map(|handle| handle.position())
            .collect::<Vec<_>>(),
        [Position::new(0), Position::new(1), Position::new(2)]
    );
    Ok(())
}

#[test]
fn native_body_storage_slot_is_preserved_without_becoming_a_drawable() -> TestResult<()> {
    let source = Package::from_bytes(&fixture_bytes_with_native_body_slot(b"native slot")?)?;
    let handles = source.body_drawable_order()?;
    assert_eq!(handles.len(), DRAWABLE_IDENTIFIERS.len());

    let mut edit = source.edit_body_drawable_order()?;
    assert!(edit.move_drawable(&handles[0], DrawableLayerMove::ToFront)?);
    let commit = edit.commit()?;
    let payload = zorder_payload_from_package(commit.package())?;
    let (snapshot, _report) =
        litchi_iwa_protos::pages_drawable_order_codec::decode_drawable_order_with_report(
            &payload,
            litchi_iwa_protos::pages_drawable_order_codec::DecodeOptions::for_source(&payload),
        )?;
    assert_eq!(
        snapshot.identifiers().collect::<Vec<_>>(),
        [
            BODY_STORAGE_IDENTIFIER,
            DRAWABLE_IDENTIFIERS[1],
            DRAWABLE_IDENTIFIERS[2],
            DRAWABLE_IDENTIFIERS[0],
        ]
    );
    assert_eq!(commit.package().body_drawable_order()?.len(), handles.len());
    Ok(())
}

#[test]
fn no_op_and_boundary_moves_are_exact_identity() -> TestResult<()> {
    let source = Package::from_bytes(&fixture_bytes(&DRAWABLE_IDENTIFIERS, b"sentinel")?)?;
    let handles = source.body_drawable_order()?;
    let mut edit = source.edit_body_drawable_order()?;
    assert!(!edit.move_drawable(&handles[0], DrawableLayerMove::ToBack)?);
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(exact_bytes(&source)?, exact_bytes(commit.package())?);
    Ok(())
}

#[test]
fn empty_order_is_a_valid_exact_no_op() -> TestResult<()> {
    let source = Package::from_bytes(&fixture_bytes(&[], b"empty sentinel")?)?;
    assert!(source.body_drawable_order()?.is_empty());

    let mut edit = source.edit_body_drawable_order()?;
    edit.set_order(&[])?;
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert!(commit.patch().is_noop());
    assert_eq!(exact_bytes(&source)?, exact_bytes(commit.package())?);
    Ok(())
}

#[test]
fn positions_handles_and_patches_fail_closed() -> TestResult<()> {
    let source = Package::from_bytes(&fixture_bytes(&DRAWABLE_IDENTIFIERS, b"source")?)?;
    let handles = source.body_drawable_order()?;
    assert!(matches!(
        source.body_drawable(Position::new(99)),
        Err(BodyDrawableOrderError::PositionNotFound { .. })
    ));

    let mut duplicate = source.edit_body_drawable_order()?;
    assert!(matches!(
        duplicate.set_order(&[handles[0].clone(), handles[0].clone(), handles[2].clone()]),
        Err(BodyDrawableOrderError::InvalidOrder)
    ));

    let foreign = Package::from_bytes(&fixture_bytes(&DRAWABLE_IDENTIFIERS, b"foreign")?)?;
    let foreign_handles = foreign.body_drawable_order()?;
    let mut foreign_edit = source.edit_body_drawable_order()?;
    assert!(matches!(
        foreign_edit.set_order(&foreign_handles),
        Err(BodyDrawableOrderError::HandleSourceMismatch)
    ));

    let mut changed = source.edit_body_drawable_order()?;
    changed.move_drawable(0_usize, DrawableLayerMove::ToFront)?;
    let commit = changed.commit()?;
    assert!(matches!(
        commit.package().apply_body_drawable_order(commit.patch()),
        Err(BodyDrawableOrderError::PatchConflict)
    ));
    Ok(())
}
