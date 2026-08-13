use litchi_iwa_archive::{Limits, package};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsp, tswp};
use prost::Message as _;

use super::{Error, LimitKind, Package, production_test_attempt};

const DOCUMENT: u64 = 1;
const BODY: u64 = 2;
const SECTION: u64 = 3;
const FIRST_FILLER: u64 = 10_000;
const FIXED_OBJECTS: usize = 3;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

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

fn topology_package(object_count: usize) -> TestResult<Vec<u8>> {
    assert!(object_count >= FIXED_OBJECTS);
    let mut document = object(
        DOCUMENT,
        10_000,
        tp::DocumentArchive {
            body_storage: Some(reference(BODY)),
            ..tp::DocumentArchive::default()
        }
        .encode_to_vec(),
    )?;
    document.archive_info.message_infos[0].object_references = vec![BODY];
    let mut body = object(
        BODY,
        2_001,
        tswp::StorageArchive {
            text: vec!["scale".to_owned()],
            table_section: Some(tswp::ObjectAttributeTable {
                entries: vec![tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(reference(SECTION)),
                }],
            }),
            ..tswp::StorageArchive::default()
        }
        .encode_to_vec(),
    )?;
    body.archive_info.message_infos[0].object_references = vec![SECTION];
    let section = object(
        SECTION,
        10_011,
        tp::SectionArchive {
            name: Some("Scale".to_owned()),
            ..tp::SectionArchive::default()
        }
        .encode_to_vec(),
    )?;
    let mut objects = Vec::new();
    objects.try_reserve_exact(object_count)?;
    objects.extend([document, body, section]);
    for offset in 0..object_count - FIXED_OBJECTS {
        objects.push(object(
            FIRST_FILLER.saturating_add(u64::try_from(offset)?),
            99_999,
            vec![0x08, 0x00],
        )?);
    }
    let compressed = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(package::to_bytes(
        [("Index/Document.iwa", compressed.as_slice())],
        Limits::default(),
    )?)
}

fn requested_background() -> crate::section::Background {
    crate::section::Background::Solid(
        Rgba::new(0.125, 0.25, 0.5, 1.0, RgbColorSpace::Srgb).expect("fixed test color is valid"),
    )
}

fn attempt(
    object_count: usize,
    maximum_transaction_work: Option<usize>,
) -> TestResult<(Result<(), Error>, super::super::section_transaction::Usage)> {
    let bytes = topology_package(object_count)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.stats().total_objects(), object_count);
    Ok(production_test_attempt(
        &package,
        requested_background(),
        maximum_transaction_work,
    ))
}

fn assert_at_most_2_01x(label: &str, small: usize, large: usize) {
    assert_ne!(small, 0, "{label} small counter must be exercised");
    assert_ne!(large, 0, "{label} large counter must be exercised");
    assert!(
        large.saturating_mul(100) <= small.saturating_mul(201),
        "{label} grew by more than 2.01x: 4K={small}, 8K={large}",
    );
}

#[test]
fn component_population_work_scales_and_required_minus_one_refuses_before_publication() -> TestResult
{
    let (small_result, small) = attempt(4_096, None)?;
    small_result?;
    let (large_result, large) = attempt(8_192, None)?;
    large_result?;

    assert_eq!(small.fields, large.fields);
    assert_eq!(small.work, large.work);
    assert_eq!(small.references, large.references);
    assert_at_most_2_01x(
        "transaction work",
        small.transaction_work,
        large.transaction_work,
    );
    assert_eq!(small.output_allocations, 1);
    assert_eq!(large.output_allocations, 1);
    assert_eq!(small.candidate_reopens, 1);
    assert_eq!(large.candidate_reopens, 1);

    let maximum = large
        .transaction_work
        .checked_sub(1)
        .expect("successful transaction work is nonzero");
    let (rejected, rejected_usage) = attempt(8_192, Some(maximum))?;
    assert!(matches!(
        rejected,
        Err(Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            maximum: reported_maximum,
            ..
        }) if reported_maximum == u64::try_from(maximum)?
    ));
    assert_eq!(rejected_usage.output_allocations, 0);
    assert_eq!(rejected_usage.candidate_reopens, 0);
    Ok(())
}
