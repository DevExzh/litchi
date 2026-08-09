use litchi_core::Position;
use litchi_ods::{
    Builder, Cell, CellValue, MutableSpreadsheet, Sheet, Spreadsheet, worksheet::Snapshot,
};

#[test]
fn worksheet_patch_round_trips_cells_order_and_inverse() -> litchi_core::Result<()> {
    let source = Builder::new().build()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.add(Sheet::new("First")?)?;
    edit.add(Sheet::new("Second")?)?;
    edit.set_cell(
        "Second",
        0,
        0,
        Cell::new(CellValue::Text("value".to_string()), "value"),
    )?;
    edit.move_to("Second", Position::new(0))?;
    let commit = edit.commit()?;

    assert!(commit.changed());
    assert_eq!(commit.snapshot().sheets()[0].name, "Second");
    let stored = commit.snapshot().sheets()[0].cell(0, 0).ok_or_else(|| {
        litchi_core::Error::InvalidFormat("the stored cell is missing".to_string())
    })?;
    assert_eq!(stored.text, "value");
    let spreadsheet = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    litchi_odf_common::compact_xml::validate(spreadsheet.content_xml().as_bytes())?;

    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), source);
    assert!(commit.patch().apply(&snapshot).is_ok());
    Ok(())
}

#[test]
fn failed_worksheet_staging_does_not_mutate_the_draft() -> litchi_core::Result<()> {
    let mut builder = Builder::new();
    builder.add_sheet(Sheet::new("Data")?)?;
    let source = builder.build()?;
    let snapshot = Snapshot::from_bytes(source)?;
    let mut edit = snapshot.edit();
    let before = edit.sheets().to_vec();

    assert!(edit.add(Sheet::new("Data")?).is_err());
    assert_eq!(edit.sheets(), before);
    assert!(edit.remove("Missing")?.is_none());
    assert!(edit.move_to("Data", Position::new(2)).is_err());
    assert_eq!(edit.sheets(), before);
    Ok(())
}

#[test]
fn mutable_facade_publishes_batched_worksheet_edits() -> litchi_core::Result<()> {
    let source = Builder::new().build()?;
    let mut mutable = MutableSpreadsheet::from_bytes(source)?;
    mutable.edit_worksheets(|edit| {
        edit.add(Sheet::new("Data")?)?;
        edit.set_formula("Data", 0, 0, "of:=1")?;
        Ok(())
    })?;

    let sheet = mutable
        .sheet("Data")
        .ok_or_else(|| litchi_core::Error::InvalidFormat("the new sheet is missing".to_string()))?;
    let cell = sheet.cell(0, 0).ok_or_else(|| {
        litchi_core::Error::InvalidFormat("the formula cell is missing".to_string())
    })?;
    assert_eq!(cell.formula.as_deref(), Some("of:=1"));
    Ok(())
}
