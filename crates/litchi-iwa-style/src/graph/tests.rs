use litchi_iwa_core::{Archive, ArchiveObject, Limits, RawMessage};

use super::{
    ConstructionError, ConstructionLimitKind, ConstructionLimits, FinishError, FinishLimitKind,
    FinishLimits, Workspace,
};

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
        Limits::default(),
    )
    .expect("valid object")
}

fn components() -> Vec<(String, Archive)> {
    vec![
        (
            "Index/Document.iwa".to_owned(),
            Archive {
                objects: vec![object(1, 1, vec![0x08, 0x01]), object(7, 2, Vec::new())],
            },
        ),
        (
            "Index/Metadata.iwa".to_owned(),
            Archive {
                // last_object_identifier = 9.
                objects: vec![object(2, 11_006, vec![0x08, 0x09])],
            },
        ),
    ]
}

#[test]
fn objects_are_copy_on_write_and_identifiers_are_fresh() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    assert_eq!(workspace.next_object_identifier().expect("next"), 10);
    assert_eq!(workspace.last_object_identifier().expect("last"), Some(9));
    workspace.remove(7).expect("removes");
    assert!(workspace.object(7).is_err());
    let changes = workspace.finish();
    assert_eq!(changes.edits.len(), 1);
    assert_eq!(changes.edits[0].0, "Index/Document.iwa");
    assert!(source[0].1.object(7).is_some());
}

#[test]
fn the_watermark_is_raised_and_lowered() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.set_last_object_identifier(12).expect("raises");
    assert_eq!(workspace.last_object_identifier().expect("last"), Some(12));
    workspace.set_last_object_identifier(7).expect("lowers");
    assert_eq!(workspace.last_object_identifier().expect("last"), Some(7));
    // Object 7 still exists, so releasing it must fail closed.
    let mut again = Workspace::new(&source, Limits::default()).expect("workspace");
    again.set_last_object_identifier(7).expect("lowers");
    assert!(again.release_identifier_suffix(&[7]).is_err());
}

#[test]
fn a_package_without_registry_skips_registry_edits() {
    let source = vec![components().remove(0)];
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    assert!(!workspace.has_registry());
    workspace.add_object_uuids(1, &[7]).expect("no-op");
    workspace.set_last_object_identifier(40).expect("no-op");
    assert!(workspace.finish().is_empty());
}

#[test]
fn registry_inspection_work_only_charges_an_uncached_state() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    let initial = workspace
        .registry_inspection_requirements()
        .expect("requirements")
        .expect("registry")
        .work_bytes();

    // Preparing the report is reused, but the first lookup still needs the
    // actual two-pass inspection that populates the root cache.
    assert_eq!(workspace.registry_inspection_work(), initial);
    assert_eq!(
        workspace
            .component_for_member("Index/Document.iwa")
            .expect("component lookup"),
        None
    );
    assert_eq!(workspace.registry_inspection_work(), 0);

    // A metadata rewrite invalidates both caches; the next lookup must be
    // admitted again for the new payload rather than reusing stale facts.
    workspace
        .set_last_object_identifier(10)
        .expect("watermark rewrite");
    assert!(workspace.registry_inspection_work() > 0);
}

#[test]
fn bounded_finish_admits_name_copies_before_materializing_changes() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.object_mut(1).expect("object").messages[0]
        .data
        .push(0x10);
    let requirements = workspace.finish_requirements().expect("requirements");
    assert_eq!(requirements.edit_count(), 1);
    assert_eq!(requirements.deletion_count(), 0);
    assert_eq!(requirements.copy_bytes(), "Index/Document.iwa".len());
    assert!(requirements.retained_bytes() >= requirements.copy_bytes());

    let refused = workspace.finish_bounded(FinishLimits::new(
        requirements.work(),
        requirements.allocations().saturating_sub(1),
        requirements.retained_bytes(),
        requirements.scratch_bytes(),
    ));
    assert!(matches!(
        refused,
        Err(FinishError::Limit {
            kind: FinishLimitKind::Allocations,
            ..
        })
    ));

    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.object_mut(1).expect("object").messages[0]
        .data
        .push(0x10);
    let requirements = workspace.finish_requirements().expect("requirements");
    let (changes, report) = workspace
        .finish_bounded(FinishLimits::new(
            requirements.work(),
            requirements.allocations(),
            requirements.retained_bytes(),
            requirements.scratch_bytes(),
        ))
        .expect("bounded finish");
    assert_eq!(report.copy_bytes(), requirements.copy_bytes());
    assert_eq!(changes.edits.len(), 1);
    assert_eq!(changes.edits[0].0, "Index/Document.iwa");
}

#[test]
fn bounded_finish_accounts_deleted_names_and_moves_insertions() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.object_mut(1).expect("object");
    workspace
        .delete_member(super::Location::Source(0))
        .expect("delete");
    let requirements = workspace.finish_requirements().expect("requirements");
    assert_eq!(requirements.edit_count(), 0);
    assert_eq!(requirements.deletion_count(), 1);
    assert_eq!(requirements.insertion_count(), 0);
    assert_eq!(requirements.copy_bytes(), "Index/Document.iwa".len());
    let (changes, _) = workspace
        .finish_bounded(FinishLimits::new(
            requirements.work(),
            requirements.allocations(),
            requirements.retained_bytes(),
            requirements.scratch_bytes(),
        ))
        .expect("bounded deletion finish");
    assert!(changes.edits.is_empty());
    assert_eq!(changes.deletions, vec!["Index/Document.iwa"]);
}

#[test]
fn bounded_finish_rejects_zero_work_and_copy_retention_before_allocating() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.object_mut(1).expect("object");
    let requirements = workspace.finish_requirements().expect("requirements");

    let zero_work = workspace.finish_bounded(FinishLimits::new(
        requirements.work().saturating_sub(1),
        requirements.allocations(),
        requirements.retained_bytes(),
        requirements.scratch_bytes(),
    ));
    assert!(matches!(
        zero_work,
        Err(FinishError::Limit {
            kind: FinishLimitKind::Work,
            ..
        })
    ));

    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.object_mut(1).expect("object");
    let requirements = workspace.finish_requirements().expect("requirements");
    let zero_copy_retention = workspace.finish_bounded(FinishLimits::new(
        requirements.work(),
        requirements.allocations(),
        requirements.retained_bytes().saturating_sub(1),
        requirements.scratch_bytes(),
    ));
    assert!(matches!(
        zero_copy_retention,
        Err(FinishError::Limit {
            kind: FinishLimitKind::RetainedBytes,
            ..
        })
    ));
}

#[test]
fn bounded_finish_rejects_requirement_arithmetic_overflow() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.finish.copy_bytes = usize::MAX;
    workspace.finish.edit_count = 1;
    assert!(matches!(
        workspace.finish_requirements(),
        Err(FinishError::Overflow(_))
    ));
}

#[test]
fn bounded_finish_moves_nonempty_insertions_without_copying_archive_buffers() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    let archive = Archive {
        objects: vec![object(99, 7, vec![1, 2, 3, 4])],
    };
    let before = archive.objects[0].messages[0].data.as_ptr();
    workspace
        .create_member("Index/Created.iwa", archive, None)
        .expect("insert member");
    let requirements = workspace.finish_requirements().expect("requirements");
    assert_eq!(requirements.insertion_count(), 1);
    let (changes, _) = workspace
        .finish_bounded(FinishLimits::new(
            requirements.work(),
            requirements.allocations(),
            requirements.retained_bytes(),
            requirements.scratch_bytes(),
        ))
        .expect("bounded insertion finish");
    let after = changes.insertions[0].archive.objects[0].messages[0]
        .data
        .as_ptr();
    assert_eq!(before, after);
}

#[test]
fn bounded_workspace_construction_reports_and_checks_index_requirements() {
    let source = components();
    let requirements =
        Workspace::construction_requirements(&source, ConstructionLimits::unlimited())
            .expect("construction requirements");
    assert_eq!(requirements.component_count(), 2);
    assert_eq!(requirements.object_count(), 3);
    assert!(requirements.source_work() > 5);
    assert_eq!(requirements.builder_work(), requirements.source_work() * 2);
    assert_eq!(requirements.total_work(), requirements.source_work() * 3);
    assert_eq!(requirements.index_allocations(), 2);
    assert!(requirements.index_retained_bytes() > 0);
    assert!(requirements.component_view_bytes() > 0);
    assert_eq!(requirements.component_view_allocations(), 1);

    let zero_work = Workspace::construction_requirements(
        &source,
        ConstructionLimits::new(
            requirements.source_work().saturating_sub(1),
            requirements.index_allocations(),
            requirements.index_retained_bytes(),
            requirements.index_scratch_bytes(),
        ),
    );
    assert!(matches!(
        zero_work,
        Err(ConstructionError::Limit {
            kind: ConstructionLimitKind::Work,
            ..
        })
    ));

    let zero_allocations = Workspace::construction_requirements(
        &source,
        ConstructionLimits::new(
            requirements.source_work(),
            requirements.index_allocations().saturating_sub(1),
            requirements.index_retained_bytes(),
            requirements.index_scratch_bytes(),
        ),
    );
    assert!(matches!(
        zero_allocations,
        Err(ConstructionError::Limit {
            kind: ConstructionLimitKind::Allocations,
            ..
        })
    ));

    let zero_retained = Workspace::construction_requirements(
        &source,
        ConstructionLimits::new(
            requirements.source_work(),
            requirements.index_allocations(),
            requirements.index_retained_bytes().saturating_sub(1),
            requirements.index_scratch_bytes(),
        ),
    );
    assert!(matches!(
        zero_retained,
        Err(ConstructionError::Limit {
            kind: ConstructionLimitKind::RetainedBytes,
            ..
        })
    ));

    let workspace = Workspace::new_bounded(
        &source,
        Limits::default(),
        requirements,
        ConstructionLimits::new(
            requirements.builder_work().saturating_sub(1),
            requirements.index_allocations(),
            requirements.index_retained_bytes(),
            requirements.index_scratch_bytes(),
        ),
    )
    .expect_err("two builder passes must be admitted");
    assert!(matches!(
        workspace,
        ConstructionError::Limit {
            kind: ConstructionLimitKind::Work,
            ..
        }
    ));

    let workspace = Workspace::new_bounded(
        &source,
        Limits::default(),
        requirements,
        ConstructionLimits::new(
            requirements.builder_work(),
            requirements.index_allocations(),
            requirements.index_retained_bytes(),
            requirements.index_scratch_bytes(),
        ),
    )
    .expect("bounded workspace");
    assert_eq!(workspace.source_work(), 0);
    assert_eq!(workspace.next_object_identifier().expect("identifier"), 10);
}

#[test]
fn bounded_workspace_rejects_same_count_name_and_identifier_changes() {
    let mut renamed = components();
    let requirements =
        Workspace::construction_requirements(&renamed, ConstructionLimits::unlimited())
            .expect("requirements");
    renamed[0].0 = "Index/Renamed.iwa".to_owned();
    let error = Workspace::new_bounded(
        &renamed,
        Limits::default(),
        requirements,
        ConstructionLimits::unlimited(),
    )
    .expect_err("renamed source must not reuse a census");
    assert!(matches!(error, ConstructionError::Graph(_)));

    let mut reidentified = components();
    let requirements =
        Workspace::construction_requirements(&reidentified, ConstructionLimits::unlimited())
            .expect("requirements");
    reidentified[0].1.objects[0].archive_info.identifier = Some(999);
    let error = Workspace::new_bounded(
        &reidentified,
        Limits::default(),
        requirements,
        ConstructionLimits::unlimited(),
    )
    .expect_err("reidentified source must not reuse a census");
    assert!(matches!(error, ConstructionError::Graph(_)));
}

#[test]
fn bounded_workspace_rejects_count_growth_before_index_population() {
    let mut source = components();
    let requirements =
        Workspace::construction_requirements(&source, ConstructionLimits::unlimited())
            .expect("requirements");
    source[0].1.objects.push(object(99, 7, vec![1, 2, 3]));
    let error = Workspace::new_bounded(
        &source,
        Limits::default(),
        requirements,
        ConstructionLimits::unlimited(),
    )
    .expect_err("grown source must not populate beyond reserved indexes");
    assert!(matches!(error, ConstructionError::Graph(_)));
}
