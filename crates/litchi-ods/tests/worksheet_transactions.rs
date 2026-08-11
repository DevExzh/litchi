use litchi_core::Position;
use litchi_odf_common::{constants, core::PackageWriter, package::raw_identical_members};
use litchi_ods::{
    Builder, Cell, CellValue, MutableSpreadsheet, Sheet, Spreadsheet, worksheet::Snapshot,
};

const OPAQUE_ROW: &str = concat!(
    r#"<table:table-row><table:table-cell office:value-type="string">"#,
    r#"<text:p>opaque owner</text:p><office:annotation office:name="vendor-note">"#,
    r#"<text:p>retain exactly</text:p></office:annotation>"#,
    r#"</table:table-cell></table:table-row>"#,
);

fn row_splice_content() -> String {
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
            r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
            r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
            r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4">"#,
            r#"<office:body><office:spreadsheet><table:table table:name="Data">"#,
            "{OPAQUE_ROW}",
            r#"<table:table-row><table:table-cell office:value-type="string">"#,
            r#"<text:p>replace me</text:p></table:table-cell></table:table-row>"#,
            r#"</table:table></office:spreadsheet></office:body></office:document-content>"#,
        ),
        OPAQUE_ROW = OPAQUE_ROW,
    )
}

fn row_splice_package() -> litchi_core::Result<Vec<u8>> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET)?;
    writer.add_file("content.xml", row_splice_content().as_bytes())?;
    let opaque_payload = vec![0x5a; 128 * 1024];
    writer.add_file_with_media_type(
        "Pictures/opaque.bin",
        &opaque_payload,
        "application/octet-stream",
    )?;
    writer.finish_to_bytes()
}

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

#[test]
fn row_local_commit_preserves_untouched_opaque_rows_and_refuses_touched_ones()
-> litchi_core::Result<()> {
    let source = row_splice_package()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.set_cell(
        "Data",
        1,
        0,
        Cell::new(CellValue::Text("changed".to_string()), "changed"),
    )?
    .ok_or_else(|| {
        litchi_core::Error::InvalidFormat("the selected worksheet is missing".to_string())
    })?;
    let commit = edit.commit()?;
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert!(reopened.content_xml().contains(OPAQUE_ROW));
    assert!(matches!(
        reopened.cell("Data", 1, 0),
        Some(litchi_ods::CellView::Stored(cell)) if cell.text == "changed"
    ));
    let identical =
        raw_identical_members(&source, commit.snapshot().as_bytes()).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("raw ODS comparison failed".to_string())
        })?;
    assert!(identical.contains("mimetype"));
    assert!(identical.contains("META-INF/manifest.xml"));
    assert!(identical.contains("Pictures/opaque.bin"));
    assert!(!identical.contains("content.xml"));
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );

    let untouched = Snapshot::from_bytes(source.clone())?;
    let mut refused = untouched.edit();
    refused
        .set_cell(
            "Data",
            0,
            0,
            Cell::new(CellValue::Text("unsafe".to_string()), "unsafe"),
        )?
        .ok_or_else(|| {
            litchi_core::Error::InvalidFormat("the selected worksheet is missing".to_string())
        })?;
    assert!(refused.commit().is_err());
    assert_eq!(untouched.as_bytes(), source);
    Ok(())
}

#[test]
fn unified_row_local_commit_retains_raw_package_members() -> litchi_core::Result<()> {
    let source = row_splice_package()?;
    let snapshot = litchi_ods::document::Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.worksheets(|worksheets| {
        worksheets
            .set_cell(
                "Data",
                1,
                0,
                Cell::new(CellValue::Text("unified".to_string()), "unified"),
            )?
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(
                    "the selected unified worksheet is missing".to_string(),
                )
            })?;
        Ok(())
    })?;
    let commit = edit.commit()?;
    let identical =
        raw_identical_members(&source, commit.snapshot().as_bytes()).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("raw unified ODS comparison failed".to_string())
        })?;
    assert!(identical.contains("mimetype"));
    assert!(identical.contains("META-INF/manifest.xml"));
    assert!(identical.contains("Pictures/opaque.bin"));
    assert!(!identical.contains("content.xml"));

    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert!(matches!(
        reopened.cell("Data", 1, 0),
        Some(litchi_ods::CellView::Stored(cell)) if cell.text == "unified"
    ));
    Ok(())
}
