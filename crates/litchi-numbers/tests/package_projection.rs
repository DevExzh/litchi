//! Focused semantic coverage for the generated-free Numbers package ingress.
//!
//! These tests deliberately construct only the small native graph needed by
//! the package reader.  The wire messages are authored with the generated
//! types in this test-only oracle; production ingress must project the same
//! values through its bounded Buffa/raw-wire paths.

use std::error::Error as StdError;

use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsa, tsk, tsp, tswp};
use litchi_numbers::{
    MAX_MATERIALIZED_CELLS, Package, PackageError, PackageLimits, PackageReadOptions,
    PackageSemanticLimits, PackageSemanticPath, SemanticLimitKind,
};
use prost::Message as _;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn numbers_root(sheet_ids: impl IntoIterator<Item = u64>) -> tn::DocumentArchive {
    tn::DocumentArchive {
        // These required native authorities make the synthetic root classify
        // unambiguously as Numbers.  The package projection only consumes the
        // sheet and sidebar references from this message.
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        sheets: sheet_ids.into_iter().map(reference).collect(),
        stylesheet: reference(100),
        sidebar_order: reference(101),
        theme: reference(102),
        ..Default::default()
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

fn package_bytes(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    let iwa = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", iwa.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn form_sheet_payload(name: &str, drawable_ids: &[u64]) -> Vec<u8> {
    tn::FormBasedSheetArchive {
        super_: tn::SheetArchive {
            name: name.to_owned(),
            drawable_infos: drawable_ids.iter().copied().map(reference).collect(),
            ..Default::default()
        },
        ..Default::default()
    }
    .encode_to_vec()
}

fn package_with_form_sheet() -> TestResult<Vec<u8>> {
    let mut root = numbers_root([2]).encode_to_vec();
    // Unknown root fields are part of the preservation contract and must not
    // affect the small sheet-order projection.
    append_varint_field(&mut root, 99, 7)?;

    let mut sheet = form_sheet_payload("Form projection", &[]);
    append_varint_field(&mut sheet, 99, 11)?;

    package_bytes(vec![
        object(1, DOCUMENT_MESSAGE_TYPE, root)?,
        object(2, FORM_BASED_SHEET_MESSAGE_TYPE, sheet)?,
    ])
}

#[test]
fn form_sheet_projection_matches_standard_sheet_semantics_and_ignores_unknowns() -> TestResult {
    let package = Package::from_bytes(&package_with_form_sheet()?)?;
    assert_eq!(package.sheets().len(), 1);
    assert_eq!(package.sheets()[0].index(), 0);
    assert_eq!(package.sheets()[0].name(), "Form projection");
    assert!(package.sheets()[0].tables().next().is_none());
    Ok(())
}

#[test]
fn package_text_preserves_storage_fragment_boundaries_and_skips_malformed_storage() -> TestResult {
    let root = numbers_root([]).encode_to_vec();
    let first = tswp::StorageArchive {
        text: vec!["first fragment".to_owned(), "second fragment".to_owned()],
        ..Default::default()
    };
    let empty = tswp::StorageArchive::default();
    let second = tswp::StorageArchive {
        text: vec!["last fragment".to_owned()],
        ..Default::default()
    };
    let bytes = package_bytes(vec![
        object(1, DOCUMENT_MESSAGE_TYPE, root)?,
        object(2, 200, first.encode_to_vec())?,
        // A malformed compatibility storage is intentionally ignorable.
        object(3, 200, vec![0xff])?,
        object(4, 200, empty.encode_to_vec())?,
        object(5, 201, second.encode_to_vec())?,
    ])?;

    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.text()?,
        "first fragment\nsecond fragment\nlast fragment"
    );
    Ok(())
}

#[test]
fn package_text_enforces_the_selected_aggregate_output_budget() -> TestResult {
    let root = numbers_root([]).encode_to_vec();
    let storage = tswp::StorageArchive {
        text: vec!["12345".to_owned(), "67890".to_owned()],
        ..Default::default()
    };
    let bytes = package_bytes(vec![
        object(1, DOCUMENT_MESSAGE_TYPE, root)?,
        object(2, 200, storage.encode_to_vec())?,
    ])?;
    let semantic = PackageSemanticLimits::default()
        .with_projection_limits(MAX_MATERIALIZED_CELLS, "12345\n67890".len() - 1)?;
    let package = Package::from_bytes_with_options(
        &bytes,
        PackageReadOptions::new(PackageLimits::default(), semantic),
    )?;

    assert!(matches!(
        package.text(),
        Err(PackageError::SemanticLimit {
            kind: SemanticLimitKind::OutputTextBytes,
            observed: 11,
            maximum: 10,
            path: PackageSemanticPath::Package,
        })
    ));
    Ok(())
}

#[test]
fn malformed_form_sheet_payload_does_not_publish_a_partial_document() -> TestResult {
    let bytes = package_bytes(vec![
        object(1, DOCUMENT_MESSAGE_TYPE, numbers_root([2]).encode_to_vec())?,
        object(2, FORM_BASED_SHEET_MESSAGE_TYPE, vec![0xff])?,
    ])?;
    assert!(matches!(
        Package::from_bytes(&bytes),
        Err(PackageError::MalformedPayload {
            path: PackageSemanticPath::Sheet { index: 0 },
        })
    ));
    Ok(())
}

#[test]
fn duplicate_sheet_payload_ownership_is_rejected_before_projection() -> TestResult {
    let standard = tn::SheetArchive {
        name: "standard".to_owned(),
        ..Default::default()
    }
    .encode_to_vec();
    let form = tn::FormBasedSheetArchive {
        super_: tn::SheetArchive {
            name: "form".to_owned(),
            ..Default::default()
        },
        ..Default::default()
    }
    .encode_to_vec();
    let bytes = package_bytes(vec![
        object(1, DOCUMENT_MESSAGE_TYPE, numbers_root([2]).encode_to_vec())?,
        ArchiveObject::new(
            2,
            vec![
                RawMessage {
                    type_: SHEET_MESSAGE_TYPE,
                    data: standard,
                },
                RawMessage {
                    type_: FORM_BASED_SHEET_MESSAGE_TYPE,
                    data: form,
                },
            ],
        )?,
    ])?;
    assert!(matches!(
        Package::from_bytes(&bytes),
        Err(PackageError::InvalidFormat(message))
            if message.contains("ambiguous sheet payload ownership")
    ));
    Ok(())
}

#[test]
fn length_delimited_wire_fields_are_not_accidentally_treated_as_scalars() -> TestResult {
    let mut root = numbers_root([]).encode_to_vec();
    // Field 99 is unknown, but its payload is deliberately framed as a
    // length-delimited value.  The root projection must skip it without
    // recursing into arbitrary bytes or changing semantic output.
    append_length_delimited_field(&mut root, 99, b"opaque root bytes")?;
    let package = Package::from_bytes(&package_bytes(vec![object(
        1,
        DOCUMENT_MESSAGE_TYPE,
        root,
    )?])?)?;
    assert!(package.sheets().is_empty());
    assert_eq!(package.text()?, "");
    Ok(())
}
