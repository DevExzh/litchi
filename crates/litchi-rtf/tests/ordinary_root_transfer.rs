#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Document, TableCellPath,
    edit::{
        Composition, CompositionLimits, Error, History, HistoryLimits, MergePlan, MergeResolution,
        TransferPlan,
    },
    style::Kind as StyleKind,
};

fn durable_limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        4 * 1024 * 1024,
        max_operations,
        8,
        2 * 1024 * 1024,
        4 * 1024 * 1024,
    )
}

#[test]
fn field_transfer_is_durable_reversible_historical_and_mergeable() {
    let source = Document::parse(
        r#"{\rtf1\ansi{\field{\*\fldinst HYPERLINK "https://example.com"}{\fldrslt Link}}}"#,
    )
    .unwrap();
    let target = Document::parse(r"{\rtf1\ansi Target}").unwrap();
    let transfer = TransferPlan::field(&source, 0, &target).unwrap();
    assert!(transfer.is_dependency_free());
    let commit = transfer.commit().unwrap();
    assert_eq!(commit.snapshot().fields().len(), 1);
    assert_eq!(
        commit.snapshot().fields()[0].instruction,
        source.fields()[0].instruction
    );

    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();
    let applied = target.apply_durable(&durable).unwrap();
    assert_eq!(
        applied.to_bytes().unwrap(),
        commit.snapshot().to_bytes().unwrap()
    );
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), target.to_bytes().unwrap());

    let mut history = History::new(target.clone(), HistoryLimits::new(2, 1024 * 1024));
    history.record_commit(&commit).unwrap();
    assert!(history.undo());
    assert!(history.current().same_snapshot(&target));
    assert!(history.redo());
    assert_eq!(history.current().fields().len(), 1);

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let left_plan = TransferPlan::field(&source, 0, &target).unwrap();
    let right_source = Document::parse(r"{\rtf1{\field{\*\fldinst PAGE}{\fldrslt 1}}}").unwrap();
    let right_plan = TransferPlan::field(&right_source, 0, &target).unwrap();
    let mut left = Composition::new(&target, limits);
    left.join(left_plan.into_edit().into_sub_edit("left", limits).unwrap())
        .unwrap();
    let mut right = Composition::new(&target, limits);
    right
        .join(
            right_plan
                .into_edit()
                .into_sub_edit("right", limits)
                .unwrap(),
        )
        .unwrap();
    let mut merge = MergePlan::new(left, right).unwrap();
    assert_eq!(merge.conflicts().len(), 1);
    merge.resolve(MergeResolution::Right);
    let merged = merge.finish().unwrap().commit().unwrap();
    assert_eq!(
        merged.snapshot().fields()[0].field_type,
        litchi_rtf::FieldType::Page
    );
}

#[test]
fn nested_table_style_list_and_object_transfers_reopen_with_dependencies() {
    let nested_source = Document::parse(
        r"{\rtf1\trowd\cellx5000\intbl\itap1 Before \intbl\itap2 Inner\nestcell{\*\nesttableprops\trowd\cellx1000\nestrow}{\nonesttables ignored}\intbl\itap1 After\cell\row}",
    )
    .unwrap();
    let table_target = Document::parse(r"{\rtf1\trowd\cellx1000\intbl Target\cell\row}").unwrap();
    let nested = TransferPlan::nested_table(
        &nested_source,
        &TableCellPath::outer(0, 0, 0),
        0,
        &table_target,
        TableCellPath::outer(0, 0, 0),
    )
    .unwrap()
    .commit()
    .unwrap()
    .into_snapshot();
    assert_eq!(
        nested.tables()[0].rows()[0].cells()[0].nested_tables()[0]
            .table
            .rows()[0]
            .cells()[0]
            .text(),
        "Inner"
    );
    Document::from_bytes(&nested.to_bytes().unwrap()).unwrap();

    let style_source =
        Document::parse(r"{\rtf1{\stylesheet{\s0 Normal;}{\s1\sbasedon0 Heading;}}Body}").unwrap();
    let resource_target = Document::parse(r"{\rtf1 Target}").unwrap();
    let style_plan =
        TransferPlan::style(&style_source, StyleKind::Paragraph, 1, &resource_target).unwrap();
    assert_eq!(style_plan.dependency_count(), 1);
    let styled = style_plan.commit().unwrap().into_snapshot();
    assert!(styled.styles().iter().any(|style| style.id == 0));
    assert!(styled.styles().iter().any(|style| style.id == 1));

    let list_source = Document::parse(
        r"{\rtf1{\*\listtable{\list\listtemplateid7\listsimple{\listlevel\levelnfc0\levelstartat1\f0{\leveltext\'02\'00.;}{\levelnumbers\'01;}}\listid7}}{\*\listoverridetable{\listoverride\listid7\listoverridecount0\ls3}}Body}",
    )
    .unwrap();
    let list_plan = TransferPlan::list(&list_source, 7, &resource_target).unwrap();
    assert_eq!(list_plan.dependency_count(), 1);
    let listed = list_plan.commit().unwrap().into_snapshot();
    assert!(listed.lists().iter().any(|list| list.id == 7));
    assert!(
        listed
            .list_overrides()
            .iter()
            .any(|entry| entry.list_id == 7)
    );

    let object_source = Document::parse(
        r"{\rtf1 A{\object\objemb{\*\objclass Package}{\*\objdata 0102}{\result fallback{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}}}",
    )
    .unwrap();
    let object_plan = TransferPlan::object(&object_source, 0, &resource_target).unwrap();
    assert_eq!(object_plan.dependency_count(), 1);
    let object = object_plan.commit().unwrap().into_snapshot();
    assert_eq!(object.pictures().len(), 1);
    assert_eq!(object.objects().len(), 1);
    Document::from_bytes(&object.to_bytes().unwrap()).unwrap();
}

#[test]
fn opaque_and_active_dependencies_are_refused_before_publication() {
    let source =
        Document::parse(r"{\rtf1{\field{\*\fldinst INCLUDETEXT external}{\fldrslt x}}}").unwrap();
    let target = Document::parse(r"{\rtf1 Target}").unwrap();
    assert!(matches!(
        TransferPlan::field(&source, 0, &target),
        Err(Error::UnsupportedSource(_))
    ));

    let plain = Document::parse(r"{\rtf1{\field{\*\fldinst PAGE}{\fldrslt 1}}}").unwrap();
    let opaque_target = Document::parse(r"{\rtf1{\*\vendor opaque}Target}").unwrap();
    let exact = opaque_target.to_bytes().unwrap();
    assert!(matches!(
        TransferPlan::field(&plain, 0, &opaque_target),
        Err(Error::UnsupportedSource(_))
    ));
    assert_eq!(opaque_target.to_bytes().unwrap(), exact);
}
