use litchi_core::Position;
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use prost::Message as _;

use super::{Error, LimitKind, Package, production_test_attempt};

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const FIRST_NODE: u64 = 3;
const FIRST_SLIDE: u64 = 4;
const SECOND_NODE: u64 = 5;
const SECOND_SLIDE: u64 = 6;
const FIRST_FILLER: u64 = 10_000;
const METADATA_OBJECT: u64 = 300;
const FIXED_OBJECTS: usize = 7;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: Some(-1),
        deprecated_is_external: Some(false),
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn referenced_object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    path: &[u32],
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = object(identifier, type_, data)?;
    let info = &mut object.archive_info.message_infos[0];
    info.object_references.extend_from_slice(references);
    let mut field = FieldInfo::new(path.to_vec());
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references.extend_from_slice(references);
    info.field_infos.push(field);
    Ok(object)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn node(slide_identifier: u64) -> Vec<u8> {
    #[allow(
        deprecated,
        reason = "native schema retains required legacy cache fields"
    )]
    kn::SlideNodeArchive {
        slide: Some(reference(slide_identifier)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..Default::default()
    }
    .encode_to_vec()
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

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(1_000),
            upper: identifier.saturating_add(2_000),
        },
    }
}

fn external_reference(
    component_identifier: u64,
    object_identifier: u64,
) -> tsp::ComponentExternalReference {
    tsp::ComponentExternalReference {
        component_identifier,
        object_identifier: Some(object_identifier),
        is_weak: None,
    }
}

fn package_metadata() -> tsp::PackageMetadata {
    tsp::PackageMetadata {
        last_object_identifier: FIRST_FILLER.saturating_add(8_192),
        components: vec![
            tsp::ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                object_uuid_map_entries: vec![uuid_entry(FIRST_NODE), uuid_entry(SECOND_NODE)],
                external_references: vec![
                    external_reference(FIRST_SLIDE, FIRST_SLIDE),
                    external_reference(SECOND_SLIDE, SECOND_SLIDE),
                ],
                ..Default::default()
            },
            tsp::ComponentInfo {
                identifier: FIRST_SLIDE,
                preferred_locator: "Slide".to_owned(),
                locator: Some("Slide-4".to_owned()),
                object_uuid_map_entries: vec![uuid_entry(FIRST_SLIDE)],
                ..Default::default()
            },
            tsp::ComponentInfo {
                identifier: SECOND_SLIDE,
                preferred_locator: "Slide".to_owned(),
                locator: Some("Slide-6".to_owned()),
                object_uuid_map_entries: vec![uuid_entry(SECOND_SLIDE)],
                ..Default::default()
            },
        ],
        ..Default::default()
    }
}

fn topology_package(
    total_objects: usize,
    reference_occurrences: usize,
    hostile_reference: bool,
) -> TestResult<Vec<u8>> {
    assert!(total_objects > FIXED_OBJECTS);
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
            slides: vec![reference(FIRST_NODE), reference(SECOND_NODE)],
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..Default::default()
    };
    let document_objects = vec![
        referenced_object(1, 1, document.encode_to_vec(), &[2], &[2])?,
        referenced_object(
            2,
            2,
            show.encode_to_vec(),
            &[3, 2],
            &[FIRST_NODE, SECOND_NODE],
        )?,
        referenced_object(FIRST_NODE, 4, node(FIRST_SLIDE), &[2], &[FIRST_SLIDE])?,
        referenced_object(SECOND_NODE, 4, node(SECOND_SLIDE), &[2], &[SECOND_SLIDE])?,
    ];

    let filler_count = total_objects - FIXED_OBJECTS;
    let mut filler = Vec::new();
    filler.try_reserve_exact(filler_count)?;
    for offset in 0..filler_count {
        let identifier = FIRST_FILLER + u64::try_from(offset)?;
        let per_object = if offset == 0 {
            reference_occurrences
        } else {
            0
        };
        let mut references = Vec::new();
        references.try_reserve_exact(per_object)?;
        let ordinary_target = FIRST_FILLER + u64::try_from((offset + 1) % filler_count)?;
        references.extend(std::iter::repeat_n(ordinary_target, per_object));
        if hostile_reference && offset == 0 {
            references.push(FIRST_SLIDE);
        }
        filler.push(referenced_object(
            identifier,
            99_999,
            b"unrelated opaque payload".to_vec(),
            &[1],
            &references,
        )?);
    }

    let entries = [
        (DOCUMENT_MEMBER.to_owned(), component(document_objects)?),
        (
            "Index/Slide-4.iwa".to_owned(),
            component(vec![object(FIRST_SLIDE, 5, slide("Delete"))?])?,
        ),
        (
            "Index/Slide-6.iwa".to_owned(),
            component(vec![object(SECOND_SLIDE, 5, slide("Keep"))?])?,
        ),
        (
            "Index/Adversarial-References.iwa".to_owned(),
            component(filler)?,
        ),
        (
            "Index/Metadata.iwa".to_owned(),
            component(vec![object(
                METADATA_OBJECT,
                11_006,
                package_metadata().encode_to_vec(),
            )?])?,
        ),
    ];
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        litchi_iwa_archive::Limits::default(),
    )?)
}

#[test]
fn cross_component_inbound_reference_refuses_before_publication() -> TestResult {
    let bytes = topology_package(128, 8, true)?;
    let package = Package::from_bytes(&bytes)?;
    let (attempt, report) = production_test_attempt(&package, Position::new(0), None);
    assert!(matches!(attempt, Err(Error::AmbiguousOwnership)));
    assert_ne!(report.objects_scanned(), 0);
    assert_ne!(report.references_scanned(), 0);
    assert_eq!(report.component_deletions(), 0);
    assert_eq!(report.output_allocations(), 0);
    assert_eq!(report.candidate_reopens(), 0);
    Ok(())
}

fn assert_at_most_2_20x(label: &str, small: u64, large: u64) {
    assert_ne!(small, 0, "{label} small counter must be exercised");
    assert_ne!(large, 0, "{label} large counter must be exercised");
    assert!(
        large.saturating_mul(100) <= small.saturating_mul(220),
        "{label} grew by more than 2.20x: 4K={small}, 8K={large}"
    );
}

#[test]
fn four_to_eight_k_objects_remain_bounded() -> TestResult {
    fn measured(total_objects: usize) -> TestResult<super::TestReport> {
        let bytes = topology_package(total_objects, 0, false)?;
        let package = Package::from_bytes(&bytes)?;
        assert_eq!(package.stats()?.total_objects, total_objects);
        let (result, report) = production_test_attempt(&package, Position::new(0), None);
        result?;
        assert_eq!(report.output_allocations(), 1);
        assert_eq!(report.candidate_reopens(), 1);
        assert_eq!(report.component_deletions(), 0);
        Ok(report)
    }

    let small = measured(4_096)?;
    let large = measured(8_192)?;
    assert!(small.objects_scanned() >= 4_096);
    assert!(large.objects_scanned() >= 8_192);
    assert_at_most_2_20x(
        "objects scanned",
        small.objects_scanned(),
        large.objects_scanned(),
    );
    assert_at_most_2_20x("transaction work", small.work(), large.work());
    assert_at_most_2_20x(
        "allocation events",
        small.allocation_events(),
        large.allocation_events(),
    );
    assert_at_most_2_20x(
        "peak scratch bytes",
        small.peak_scratch_bytes(),
        large.peak_scratch_bytes(),
    );
    Ok(())
}

#[test]
fn four_to_eight_k_reference_occurrences_remain_bounded() -> TestResult {
    fn measured(reference_occurrences: usize) -> TestResult<super::TestReport> {
        let bytes = topology_package(FIXED_OBJECTS + 1, reference_occurrences, false)?;
        let package = Package::from_bytes(&bytes)?;
        let (result, report) = production_test_attempt(&package, Position::new(0), None);
        result?;
        assert_eq!(report.output_allocations(), 1);
        assert_eq!(report.candidate_reopens(), 1);
        assert_eq!(report.component_deletions(), 0);
        Ok(report)
    }

    let small = measured(4_096)?;
    let large = measured(8_192)?;
    assert_eq!(small.objects_scanned(), large.objects_scanned());
    assert!(small.references_scanned() >= u64::try_from(2 * 4_096)?);
    assert!(large.references_scanned() >= u64::try_from(2 * 8_192)?);
    assert_at_most_2_20x(
        "references scanned",
        small.references_scanned(),
        large.references_scanned(),
    );
    assert_at_most_2_20x("transaction work", small.work(), large.work());
    assert_at_most_2_20x(
        "allocation events",
        small.allocation_events(),
        large.allocation_events(),
    );
    assert_at_most_2_20x(
        "peak scratch bytes",
        small.peak_scratch_bytes(),
        large.peak_scratch_bytes(),
    );
    Ok(())
}

#[test]
fn max_minus_one_refuses_before_deletion_output_or_reopen() -> TestResult {
    let bytes = topology_package(8_192, 1, false)?;
    let package = Package::from_bytes(&bytes)?;
    let (success, observed) = production_test_attempt(&package, Position::new(0), None);
    success?;
    let maximum = observed
        .work()
        .checked_sub(1)
        .expect("successful work is nonzero");

    let (attempt, rejected) = production_test_attempt(&package, Position::new(0), Some(maximum));
    assert!(matches!(
        attempt,
        Err(Error::LimitExceeded {
            kind: LimitKind::Work,
            observed: attempted,
            maximum: reported_maximum,
            ..
        }) if attempted == observed.work() && reported_maximum == maximum
    ));
    assert_eq!(rejected.component_deletions(), 0);
    assert_eq!(rejected.output_allocations(), 0);
    assert_eq!(rejected.candidate_reopens(), 0);
    Ok(())
}
