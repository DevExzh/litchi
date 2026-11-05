//! The facade retains the operations owner's explicit transaction types.
#![cfg(feature = "docx")]

#[test]
fn facade_exposes_bounded_edit_results_and_reversible_patches() {
    use litchi::transaction::{Commit, Edit, SemanticLimits};
    use litchi::{FileFormat, NewNode, create};
    let file = create(FileFormat::Docx).unwrap();
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let commit: Commit<_> = file
        .edit_with_options(
            litchi::EditOptions::new().semantic_limits(limits),
            |edit: &mut Edit| edit.add("/body", NewNode::new("paragraph").prop("text", "facade")),
        )
        .unwrap();
    assert!(
        file.capabilities()
            .verb(litchi::capabilities::Verb::Edit)
            .is_supported()
    );
    assert!(
        file.capabilities()
            .verb(litchi::capabilities::Verb::Batch)
            .is_supported()
    );
    assert!(commit.file().find(commit.value()).unwrap().is_some());
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        file.to_bytes().unwrap()
    );
}

#[test]
fn facade_exposes_command_results_for_atomic_and_best_effort_batches() {
    use litchi::{BatchCommit, BatchMode, ItemResult};
    let file = litchi::create(litchi::FileFormat::Docx).unwrap();
    let items = litchi::batch::from_bytes(
        br#"[
      {"command":"add","parent":"/body","type":"paragraph","props":{"text":"facade"}},
      {"command":"unknown"}
    ]"#,
        &litchi::batch::SourceLimits::default(),
    )
    .unwrap();
    let commit: BatchCommit = file.batch(&items, BatchMode::BestEffort).unwrap();
    assert!(matches!(commit.value()[0], ItemResult::Applied(_)));
    assert!(matches!(commit.value()[1], ItemResult::Failed(_)));
    let applied: litchi::transaction::Commit<litchi::Applied> = file.apply(&items[0]).unwrap();
    assert!(applied.value().path.is_some());
}
