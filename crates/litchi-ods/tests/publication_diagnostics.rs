use litchi_core::Result;
use litchi_ods::document::{LogicalRowOperation, PublicationRoute};
use litchi_ods::names::{Expression, Scope};
use litchi_ods::{Builder, Cell, CellValue, Package, Sheet};

fn source() -> Result<Vec<u8>> {
    let mut builder = Builder::new();
    builder.add_sheet(Sheet::new("Data")?)?;
    builder.set_cell(
        "Data",
        0,
        0,
        Cell::new(CellValue::Text("before".to_string()), "before"),
    )?;
    Ok(builder.build()?.to_bytes()?)
}

#[test]
fn unified_commit_diagnostics_attribute_noop_splice_and_replay() -> Result<()> {
    let snapshot = Package::from_bytes(&source()?)?;

    let unchanged = snapshot.edit()?.commit()?;
    let unchanged_diagnostics = unchanged.diagnostics();
    assert_eq!(
        unchanged_diagnostics.publication_route(),
        PublicationRoute::NoOp
    );
    assert!(!unchanged_diagnostics.changed());
    assert!(!unchanged_diagnostics.candidate_reopened());
    assert_eq!(unchanged_diagnostics.operation_count(), 0);

    let mut edit = snapshot.edit()?;
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
    assert_ne!(commit.package().as_bytes(), snapshot.as_bytes());

    let replay = snapshot.apply(commit.patch())?.into_package();
    assert_eq!(replay.as_bytes(), commit.package().as_bytes());
    let restored = replay.apply(&commit.patch().inverse())?;
    assert_eq!(restored.into_package().as_bytes(), snapshot.as_bytes());
    Ok(())
}

#[test]
fn logical_rebuild_is_not_reclassified_by_a_later_noop_splice() -> Result<()> {
    let snapshot = Package::from_bytes(&source()?)?;
    let mut edit = snapshot.edit()?;
    edit.definitions(|definitions| {
        definitions.add(Expression::new("Input", "of:=1+1", Scope::Global)?.into())
    })?;

    // A same-position move is an explicit no-op. The advanced source-splice
    // helper therefore returns the current candidate bytes unchanged.
    edit.edit_logical_rows(
        "Data",
        &[LogicalRowOperation::Move {
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
