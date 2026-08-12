use litchi_odf_common::{constants, core::PackageWriter, package::raw_identical_members};
use litchi_ods::{
    Builder, Cell, CellValue, Row, Sheet, Spreadsheet,
    worksheet::{CellChange, MAX_CELL_CHANGES, Snapshot},
};

fn text(value: &str) -> Cell {
    Cell::new(CellValue::Text(value.to_string()), value)
}

fn repeated_source() -> litchi_core::Result<Vec<u8>> {
    let mut row = Row::repeated(3)?;
    row.push_cell(Cell::repeated(
        CellValue::Text("old".to_string()),
        "old",
        4,
    )?)?;
    let mut sheet = Sheet::new("Data")?;
    sheet.push_row(row)?;
    let mut builder = Builder::new();
    builder.add_sheet(sheet)?;
    builder.build()
}

const OPAQUE_ROW: &str = concat!(
    r#"<table:table-row><table:table-cell office:value-type="string">"#,
    r#"<text:p>opaque owner</text:p><office:annotation office:name="vendor-note">"#,
    r#"<text:p>retain exactly</text:p></office:annotation>"#,
    r#"</table:table-cell></table:table-row>"#,
);

fn opaque_package() -> litchi_core::Result<Vec<u8>> {
    let content = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
            r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
            r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
            r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4">"#,
            r#"<office:body><office:spreadsheet><table:table table:name="Data">"#,
            "{OPAQUE_ROW}",
            r#"<table:table-row><table:table-cell office:value-type="string" table:number-columns-repeated="4">"#,
            r#"<text:p>old</text:p></table:table-cell></table:table-row>"#,
            r#"</table:table></office:spreadsheet></office:body></office:document-content>"#,
        ),
        OPAQUE_ROW = OPAQUE_ROW,
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET)?;
    writer.add_file("content.xml", content.as_bytes())?;
    writer.add_file_with_media_type(
        "Pictures/opaque.bin",
        &vec![0x5a; 128 * 1024],
        "application/octet-stream",
    )?;
    writer.finish_to_bytes()
}

#[test]
fn batch_updates_multiple_repeated_rows_and_cells_deterministically() -> litchi_core::Result<()> {
    let source = repeated_source()?;
    let changes = vec![
        CellChange::new(2, 3, text("last")),
        CellChange::new(0, 1, text("first")),
        CellChange::new(1, 2, text("middle")),
    ];
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    assert_eq!(edit.set_cells("Data", changes.clone())?, Some(3));
    let commit = edit.commit()?;

    let sheet = &commit.snapshot().sheets()[0];
    assert_eq!(
        sheet.cell(0, 1).map(|cell| cell.text.as_str()),
        Some("first")
    );
    assert_eq!(
        sheet.cell(1, 2).map(|cell| cell.text.as_str()),
        Some("middle")
    );
    assert_eq!(
        sheet.cell(2, 3).map(|cell| cell.text.as_str()),
        Some("last")
    );
    assert_eq!(sheet.cell(0, 0).map(|cell| cell.text.as_str()), Some("old"));
    assert_eq!(sheet.cell(2, 0).map(|cell| cell.text.as_str()), Some("old"));

    let second = Snapshot::from_bytes(source)?;
    let mut reverse = second.edit();
    let mut reverse_changes = changes;
    reverse_changes.reverse();
    assert_eq!(reverse.set_cells("Data", reverse_changes)?, Some(3));
    assert_eq!(
        reverse.commit()?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    Ok(())
}

#[test]
fn batch_matches_scalar_edits_across_sparse_and_last_grid_coordinates() -> litchi_core::Result<()> {
    const LAST_LOGICAL_COORDINATE: usize = 1_048_575;

    let mut first_row = Row::new();
    first_row.push_cell(text("seed"))?;
    first_row.push_cell(text("replace"))?;
    let mut sheet = Sheet::new("Data")?;
    sheet.push_row(first_row)?;
    let mut builder = Builder::new();
    builder.add_sheet(sheet)?;
    let source = builder.build()?;
    let changes = vec![
        CellChange::new(0, 1, text("near")),
        CellChange::new(3, 7, text("sparse")),
        CellChange::new(
            LAST_LOGICAL_COORDINATE,
            LAST_LOGICAL_COORDINATE,
            text("last"),
        ),
    ];

    let batch_snapshot = Snapshot::from_bytes(source.clone())?;
    let mut batch = batch_snapshot.edit();
    assert_eq!(batch.set_cells("Data", changes.clone())?, Some(3));

    let scalar_snapshot = Snapshot::from_bytes(source)?;
    let mut scalar = scalar_snapshot.edit();
    for change in changes {
        scalar.set_cell("Data", change.row(), change.column(), change.cell().clone())?;
    }
    assert_eq!(batch.sheets(), scalar.sheets());
    assert_eq!(
        batch.commit()?.snapshot().as_bytes(),
        scalar.commit()?.snapshot().as_bytes()
    );
    Ok(())
}

#[test]
fn duplicate_and_late_invalid_changes_leave_the_edit_unchanged() -> litchi_core::Result<()> {
    let snapshot = Snapshot::from_bytes(repeated_source()?)?;
    let mut edit = snapshot.edit();
    let before = edit.sheets().to_vec();

    let duplicate = edit.set_cells(
        "Data",
        vec![
            CellChange::new(1, 1, text("one")),
            CellChange::new(1, 1, text("two")),
        ],
    );
    assert!(matches!(
        duplicate,
        Err(litchi_core::Error::InvalidFormat(message))
            if message == "ODS cell batch repeats logical coordinate (1, 1)"
    ));
    assert_eq!(edit.sheets(), before);

    let invalid = edit.set_cells(
        "Data",
        vec![
            CellChange::new(0, 0, text("valid")),
            CellChange::new(2, 2, Cell::new(CellValue::Number(f64::NAN), "NaN")),
        ],
    );
    assert!(invalid.is_err());
    assert_eq!(edit.sheets(), before);

    let invalid_coordinate = edit.set_cells(
        "Data",
        vec![
            CellChange::new(0, 0, text("valid")),
            CellChange::new(1_048_576, 0, text("outside")),
        ],
    );
    assert!(invalid_coordinate.is_err());
    assert_eq!(edit.sheets(), before);

    let repeated = edit.set_cells(
        "Data",
        vec![
            CellChange::new(0, 0, text("valid")),
            CellChange::new(
                2,
                2,
                Cell::repeated(CellValue::Text("x".to_string()), "x", 2)?,
            ),
        ],
    );
    assert!(repeated.is_err());
    assert_eq!(edit.sheets(), before);
    Ok(())
}

#[test]
fn batch_accepts_the_exact_operation_limit_and_rejects_one_more_first() -> litchi_core::Result<()> {
    let snapshot = Snapshot::from_bytes(repeated_source()?)?;
    let mut edit = snapshot.edit();
    let exact = (0..MAX_CELL_CHANGES)
        .map(|column| CellChange::new(0, column, text("value")))
        .collect();
    assert_eq!(edit.set_cells("Data", exact)?, Some(MAX_CELL_CHANGES));

    let before = edit.sheets().to_vec();
    let above = (0..=MAX_CELL_CHANGES)
        .map(|column| CellChange::new(0, column, Cell::new(CellValue::Number(f64::NAN), "invalid")))
        .collect();
    let error = edit.set_cells("Data", above);
    assert!(matches!(
        error,
        Err(litchi_core::Error::InvalidFormat(message))
            if message.contains("operation safety limit")
    ));
    assert_eq!(edit.sheets(), before);
    Ok(())
}

#[test]
fn empty_and_semantic_noop_batches_share_exact_source_bytes() -> litchi_core::Result<()> {
    let source = repeated_source()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let source_pointer = snapshot.as_bytes().as_ptr();
    let mut edit = snapshot.edit();

    assert_eq!(edit.set_cells("Data", Vec::new())?, Some(0));
    assert_eq!(
        edit.set_cells("Data", vec![CellChange::new(1, 2, text("old"))])?,
        Some(0)
    );
    assert_eq!(
        edit.set_cells("Missing", vec![CellChange::new(0, 0, text("ignored"))])?,
        None
    );
    let commit = edit.commit()?;
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().as_bytes(), source);
    assert_eq!(commit.snapshot().as_bytes().as_ptr(), source_pointer);
    Ok(())
}

#[test]
fn batch_commit_preserves_opaque_rows_and_members_and_has_exact_inverse() -> litchi_core::Result<()>
{
    let source = opaque_package()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    assert_eq!(
        edit.set_cells(
            "Data",
            vec![
                CellChange::new(1, 0, text("left")),
                CellChange::new(1, 3, text("right")),
            ],
        )?,
        Some(2)
    );
    let commit = edit.commit()?;
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert!(reopened.content_xml().contains(OPAQUE_ROW));
    assert_eq!(
        reopened.cell("Data", 1, 0).and_then(|view| match view {
            litchi_ods::CellView::Stored(cell) => Some(cell.text.as_str()),
            litchi_ods::CellView::Missing => None,
        }),
        Some("left")
    );
    assert_eq!(
        reopened.cell("Data", 1, 3).and_then(|view| match view {
            litchi_ods::CellView::Stored(cell) => Some(cell.text.as_str()),
            litchi_ods::CellView::Missing => None,
        }),
        Some("right")
    );
    let identical = raw_identical_members(&source, commit.snapshot().as_bytes())
        .ok_or_else(|| litchi_core::Error::InvalidFormat("raw comparison failed".to_string()))?;
    assert!(identical.contains("mimetype"));
    assert!(identical.contains("META-INF/manifest.xml"));
    assert!(identical.contains("Pictures/opaque.bin"));
    assert!(!identical.contains("content.xml"));
    assert_eq!(
        commit.patch().apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );

    let stale = Snapshot::from_bytes(Builder::new().build()?)?;
    assert!(commit.patch().apply(&stale).is_err());
    Ok(())
}
