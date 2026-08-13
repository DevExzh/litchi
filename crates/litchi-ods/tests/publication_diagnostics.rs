use litchi_core::Result;
use litchi_ods::document::{LogicalRowEdit, PublicationRoute, Snapshot};
use litchi_ods::names::{Expression, Scope};
use litchi_ods::{Builder, Cell, CellValue, Sheet};

fn source() -> Result<Vec<u8>> {
    let mut builder = Builder::new();
    builder.add_sheet(Sheet::new("Data")?)?;
    builder.set_cell(
        "Data",
        0,
        0,
        Cell::new(CellValue::Text("before".to_string()), "before"),
    )?;
    builder.build()
}

#[test]
fn unified_commit_diagnostics_attribute_noop_splice_and_replay() -> Result<()> {
    let snapshot = Snapshot::from_bytes(source()?)?;

    let unchanged = snapshot.edit().commit()?;
    let unchanged_diagnostics = unchanged.diagnostics();
    assert_eq!(
        unchanged_diagnostics.publication_route(),
        PublicationRoute::NoOp
    );
    assert!(!unchanged_diagnostics.changed());
    assert!(!unchanged_diagnostics.candidate_reopened());
    assert_eq!(unchanged_diagnostics.operation_count(), 0);

    let mut edit = snapshot.edit();
    edit.worksheets(|worksheets| {
        worksheets
            .set_cell(
                "Data",
                0,
                0,
                Cell::new(CellValue::Text("after".to_string()), "after"),
            )?
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat("missing diagnostic test sheet".to_string())
            })?;
        Ok(())
    })?;
    let commit = edit.commit()?;
    let diagnostics = commit.diagnostics();
    assert_eq!(
        diagnostics.publication_route(),
        PublicationRoute::ProvenanceSplice
    );
    assert!(diagnostics.changed());
    assert!(diagnostics.candidate_reopened());
    assert_eq!(diagnostics.operation_count(), 1);
    assert_ne!(commit.snapshot().as_bytes(), snapshot.as_bytes());

    let replay = commit.patch().apply(&snapshot)?;
    assert_eq!(
        replay.diagnostics().publication_route(),
        PublicationRoute::PatchReplay
    );
    assert_eq!(replay.snapshot().as_bytes(), commit.snapshot().as_bytes());
    let restored = commit.patch().inverse().apply(replay.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), snapshot.as_bytes());
    Ok(())
}

#[test]
fn logical_rebuild_is_not_reclassified_by_a_later_noop_splice() -> Result<()> {
    let snapshot = Snapshot::from_bytes(source()?)?;
    let mut edit = snapshot.edit();
    edit.definitions(|definitions| {
        definitions.add(Expression::new("Input", "of:=1+1", Scope::Global)?.into())
    })?;

    // A same-position move is an explicit no-op. The advanced source-splice
    // helper therefore returns the current candidate bytes unchanged.
    edit.edit_logical_rows(
        "Data",
        &[LogicalRowEdit::Move {
            at: 0,
            count: 1,
            to: 0,
        }],
    )?;
    let commit = edit.commit()?;
    let diagnostics = commit.diagnostics();
    assert_eq!(
        diagnostics.publication_route(),
        PublicationRoute::LogicalRebuild
    );
    assert!(diagnostics.changed());
    assert!(diagnostics.candidate_reopened());
    assert_eq!(diagnostics.operation_count(), 1);
    Ok(())
}
