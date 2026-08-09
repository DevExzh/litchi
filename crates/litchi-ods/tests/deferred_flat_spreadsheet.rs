use litchi_ods::{Cell, CellValue, CellView, FlatLimits, FlatSpreadsheet};
use std::sync::Arc;

const FLAT_SAMPLE: &str = include_str!("fixtures/odfdo-test-flat-lo-compact.fods");
const FLAT_IMAGE: &str = include_str!("fixtures/libreoffice-draw-image-link-compact.fods");
const FLAT_DDE: &str = include_str!("fixtures/odfpy-sheet-dde-source-compact.fods");
const PACKAGED: &[u8] = include_bytes!("../../../test-data/odf/corpus/calc-formulas.ods");

#[test]
fn reads_real_flat_calc_sheets_cells_metadata_and_source_exactly() {
    let flat = FlatSpreadsheet::from_bytes(FLAT_SAMPLE.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    assert_eq!(flat.sheets().len(), 2);
    assert_eq!(
        flat.sheet(0usize)
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .name,
        "Sheet1"
    );
    assert!(matches!(
        flat.cell("Sheet1", 0, 0)
            .expect("test fixture or operation should succeed"),
        Some(CellView::Stored(_))
    ));
    assert_eq!(
        flat.sheet("Sheet1")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .cell(0, 0)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Text("test".to_string())
    );
    assert_eq!(
        flat.sheet("Sheet1")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .cell(1, 1)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Number(123.0)
    );
    assert_eq!(
        flat.sheet("Sheet2")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .cell(0, 0)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Text("abc".to_string())
    );
    assert!(
        flat.odf_metadata()
            .generator
            .as_deref()
            .is_some_and(|value| value.contains("LibreOffice"))
    );
    assert_eq!(flat.as_bytes(), FLAT_SAMPLE.as_bytes());
}

#[test]
fn repeated_cells_are_logically_addressable_without_expansion() {
    let xml = concat!(
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body>"#,
        r#"<office:spreadsheet><table:table table:name="R"><table:table-row>"#,
        r#"<table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="3"><text:p>7</text:p></table:table-cell>"#,
        r#"<table:table-cell office:value-type="string"><text:p>end</text:p></table:table-cell>"#,
        r#"</table:table-row></table:table></office:spreadsheet></office:body></office:document>"#,
    );
    let flat = FlatSpreadsheet::from_bytes(xml.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let sheet = flat
        .sheet("R")
        .expect("test fixture or operation should succeed")
        .expect("test fixture or operation should succeed");
    assert_eq!(sheet.rows[0].cells.len(), 2, "physical runs stay compact");
    for column in 0..3 {
        assert_eq!(
            sheet
                .cell(0, column)
                .expect("test fixture or operation should succeed")
                .value,
            CellValue::Number(7.0)
        );
    }
    assert_eq!(
        sheet
            .cell(0, 3)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Text("end".to_string())
    );
}

#[test]
fn wrappers_reject_wrong_family_package_and_garbage() {
    let presentation = br#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.presentation"><office:body><office:presentation/></office:body></office:document>"#;
    assert!(FlatSpreadsheet::from_bytes(presentation.to_vec()).is_err());
    assert!(FlatSpreadsheet::from_bytes(PACKAGED.to_vec()).is_err());
    assert!(FlatSpreadsheet::from_bytes(Vec::new()).is_err());
    assert!(FlatSpreadsheet::from_bytes(b"not XML".to_vec()).is_err());
}

#[test]
fn real_flat_cell_edit_preserves_linked_shape_and_reopens() {
    let original = FlatSpreadsheet::from_bytes(FLAT_IMAGE.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let mut edit = original
        .transaction()
        .expect("test fixture or operation should succeed");
    edit.set_cell("Sheet1", 0, 1, Cell::new(CellValue::Number(42.0), "42"))
        .expect("test fixture or operation should succeed");
    let commit = edit
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(commit.changed());
    let edited = commit.into_spreadsheet();
    assert!(String::from_utf8_lossy(edited.as_bytes()).contains("tracking-pixel.png"));
    assert_eq!(
        edited
            .sheet("Sheet1")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .cell(0, 0)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Text("test".to_string())
    );
    assert_eq!(
        edited
            .sheet("Sheet1")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .cell(0, 1)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Number(42.0)
    );
    assert_eq!(original.as_bytes(), FLAT_IMAGE.as_bytes());
}

#[test]
fn odfpy_dde_edit_preserves_inert_source_without_inventing_metadata() {
    let original = FlatSpreadsheet::from_bytes(FLAT_DDE.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let mut edit = original
        .transaction()
        .expect("test fixture or operation should succeed");
    edit.set_cell(
        "Live Data",
        0,
        0,
        Cell::new(CellValue::Text("hello".to_string()), "hello"),
    )
    .expect("test fixture or operation should succeed");
    let edited = edit
        .commit()
        .expect("test fixture or operation should succeed")
        .into_spreadsheet();
    let output = String::from_utf8_lossy(edited.as_bytes());
    assert!(output.contains("never/contacted.ods"));
    assert!(!output.contains("<office:meta>"));
    assert_eq!(
        edited
            .dde()
            .expect("test fixture or operation should succeed")
            .sheet_sources()
            .len(),
        1
    );
    assert_eq!(
        edited
            .sheet("Live Data")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .cell(0, 0)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Text("hello".to_string())
    );
    assert_eq!(original.as_bytes(), FLAT_DDE.as_bytes());
}

#[test]
fn no_op_is_byte_exact_and_unmodeled_changed_rows_are_refused() {
    let original = FlatSpreadsheet::from_bytes(FLAT_IMAGE.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let commit = original
        .transaction()
        .expect("test fixture or operation should succeed")
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(!commit.changed());
    assert_eq!(commit.spreadsheet().as_bytes(), FLAT_IMAGE.as_bytes());

    let annotated = FLAT_IMAGE.replace(
        "<text:p>test</text:p>",
        "<text:p>test</text:p><office:annotation><text:p>keep</text:p></office:annotation>",
    );
    let source = FlatSpreadsheet::from_bytes(annotated.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let mut edit = source
        .transaction()
        .expect("test fixture or operation should succeed");
    edit.set_cell("Sheet1", 0, 0, Cell::new(CellValue::Number(1.0), "1"))
        .expect("test fixture or operation should succeed");
    assert!(edit.commit().is_err());
    assert_eq!(source.as_bytes(), annotated.as_bytes());
}

#[test]
fn duplicate_names_require_an_unambiguous_selector() {
    let xml = concat!(
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
        r#"office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body>"#,
        r#"<office:spreadsheet><table:table table:name="Same"><table:table-row/></table:table>"#,
        r#"<table:table table:name="Same"><table:table-row/></table:table>"#,
        r#"</office:spreadsheet></office:body></office:document>"#,
    );
    let snapshot = FlatSpreadsheet::from_bytes(xml.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    assert!(snapshot.sheet("Same").is_err());
    assert_eq!(
        snapshot
            .sheet(0usize)
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .name,
        "Same"
    );
    assert_eq!(
        snapshot
            .sheet(1usize)
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .name,
        "Same"
    );
    assert!(
        snapshot
            .sheet(2usize)
            .expect("test fixture or operation should succeed")
            .is_none()
    );

    let mut ambiguous = snapshot
        .transaction()
        .expect("test fixture or operation should succeed");
    assert!(
        ambiguous
            .set_cell("Same", 0, 0, Cell::new(CellValue::Number(1.0), "1"))
            .is_err()
    );
    let mut indexed = snapshot
        .transaction()
        .expect("test fixture or operation should succeed");
    indexed
        .set_cell(1usize, 0, 0, Cell::new(CellValue::Number(2.0), "2"))
        .expect("test fixture or operation should succeed");
    let committed = indexed
        .commit()
        .expect("test fixture or operation should succeed");
    assert_eq!(
        committed
            .snapshot()
            .sheet(1usize)
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .cell(0, 0)
            .expect("test fixture or operation should succeed")
            .value,
        CellValue::Number(2.0)
    );
}

#[test]
fn patch_is_exact_source_checked_reversible_and_stale_safe() {
    let source = FlatSpreadsheet::from_bytes(FLAT_DDE.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let mut transaction = source
        .transaction()
        .expect("test fixture or operation should succeed");
    transaction
        .set_cell(
            0usize,
            0,
            0,
            Cell::new(CellValue::Text("patched".to_string()), "patched"),
        )
        .expect("test fixture or operation should succeed");
    let commit = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(commit.changed());
    let applied = commit
        .patch()
        .apply(&source)
        .expect("test fixture or operation should succeed");
    assert_eq!(applied.as_bytes(), commit.snapshot().as_bytes());
    assert!(commit.patch().apply(commit.snapshot()).is_err());

    let inverse = commit.patch().inverse();
    let restored = inverse
        .apply(commit.snapshot())
        .expect("test fixture or operation should succeed");
    assert_eq!(restored.as_bytes(), source.as_bytes());
    assert!(inverse.apply(&source).is_err());
}

#[test]
fn input_output_and_patch_limits_enforce_exact_n_and_n_plus_one() {
    let input_len = FLAT_DDE.len();
    assert!(
        FlatSpreadsheet::from_bytes_with(
            FLAT_DDE.as_bytes().to_vec(),
            FlatLimits::default().with_input_bytes(input_len),
        )
        .is_ok()
    );
    assert!(
        FlatSpreadsheet::from_bytes_with(
            FLAT_DDE.as_bytes().to_vec(),
            FlatLimits::default().with_input_bytes(input_len - 1),
        )
        .is_err()
    );

    let source = FlatSpreadsheet::from_bytes(FLAT_DDE.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let stage = |limits| {
        let mut transaction = source
            .transaction_with(limits)
            .expect("test fixture or operation should succeed");
        transaction
            .set_cell(
                0usize,
                0,
                0,
                Cell::new(CellValue::Text("bounded".to_string()), "bounded"),
            )
            .expect("test fixture or operation should succeed");
        transaction.commit()
    };
    let reference = stage(FlatLimits::default()).expect("test fixture or operation should succeed");
    let output_len = reference.snapshot().as_bytes().len();
    assert!(
        stage(FlatLimits::default().with_output_bytes(output_len)).is_ok(),
        "N bytes must be accepted"
    );
    assert!(
        stage(FlatLimits::default().with_output_bytes(output_len - 1)).is_err(),
        "N+1 observed bytes must be rejected by an N-byte limit"
    );
    assert!(
        reference
            .patch()
            .apply_with(
                &source,
                FlatLimits::default().with_output_bytes(output_len - 1),
            )
            .is_err()
    );
}

#[test]
fn trailing_content_is_rejected_and_formatted_source_is_read_only() {
    assert!(FlatSpreadsheet::from_bytes(format!("{FLAT_DDE}trailing").into_bytes()).is_err());
    assert!(
        FlatSpreadsheet::from_bytes(format!("{FLAT_DDE}<![CDATA[trailing]]>").into_bytes())
            .is_err()
    );

    let formatted = FLAT_DDE.replacen("><", ">\n<", 1);
    let snapshot = FlatSpreadsheet::from_bytes(formatted.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let mut transaction = snapshot
        .transaction()
        .expect("test fixture or operation should succeed");
    transaction
        .set_cell(0usize, 0, 0, Cell::new(CellValue::Number(9.0), "9"))
        .expect("test fixture or operation should succeed");
    assert!(transaction.commit().is_err());
    assert_eq!(snapshot.as_bytes(), formatted.as_bytes());
}

#[test]
fn changed_publication_uses_quote_aware_shared_compactness_rules() {
    let formatted = [
        FLAT_DDE.replacen("<office:document ", "<office:document\n", 1),
        FLAT_DDE.replacen("<office:document ", "<office:document\t", 1),
        FLAT_DDE.replacen(" xmlns:table=", "  xmlns:table=", 1),
        FLAT_DDE.replacen("<table:table-row/>", "<table:table-row />", 1),
        FLAT_DDE.replacen("</table:table>", "</table:table >", 1),
        format!(" {FLAT_DDE}"),
    ];
    for source in formatted {
        let snapshot = FlatSpreadsheet::from_bytes(source.as_bytes().to_vec())
            .expect("test fixture or operation should succeed");
        let mut transaction = snapshot
            .transaction()
            .expect("test fixture or operation should succeed");
        assert_eq!(
            transaction
                .set_cell(0usize, 0, 0, Cell::new(CellValue::Number(8.0), "8"))
                .expect("test fixture or operation should succeed"),
            Some(())
        );
        assert!(transaction.commit().is_err(), "formatted XML was published");
        assert_eq!(snapshot.as_bytes(), source.as_bytes());
    }

    let quoted = FLAT_DDE.replacen("<office:document ", "<office:document quoted=\"> <\" ", 1);
    let snapshot = FlatSpreadsheet::from_bytes(quoted.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let mut transaction = snapshot
        .transaction()
        .expect("test fixture or operation should succeed");
    assert_eq!(
        transaction
            .set_cell(0usize, 0, 0, Cell::new(CellValue::Number(7.0), "7"))
            .expect("test fixture or operation should succeed"),
        Some(())
    );
    let commit = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(commit.changed());
    assert!(String::from_utf8_lossy(commit.snapshot().as_bytes()).contains("quoted=\"> <\""));
}

#[test]
fn transaction_shares_source_and_reports_absent_selectors() {
    let snapshot = FlatSpreadsheet::from_bytes(FLAT_DDE.as_bytes().to_vec())
        .expect("test fixture or operation should succeed");
    let first = snapshot.to_bytes();
    let second = snapshot.to_bytes();
    assert!(Arc::ptr_eq(&first, &second));
    let mut transaction = snapshot
        .transaction()
        .expect("test fixture or operation should succeed");
    let while_detached = snapshot.to_bytes();
    assert!(Arc::ptr_eq(&first, &while_detached));
    assert_eq!(
        transaction
            .set_cell("Missing", 0, 0, Cell::new(CellValue::Number(1.0), "1"),)
            .expect("test fixture or operation should succeed"),
        None
    );
    assert_eq!(
        transaction
            .set_cell(99usize, 0, 0, Cell::new(CellValue::Number(1.0), "1"),)
            .expect("test fixture or operation should succeed"),
        None
    );
    assert_eq!(
        transaction
            .set_cell(0usize, 0, 0, Cell::new(CellValue::Number(1.0), "1"))
            .expect("test fixture or operation should succeed"),
        Some(())
    );
    assert_eq!(snapshot.as_bytes(), FLAT_DDE.as_bytes());
}

#[test]
fn content_events_outside_the_document_root_are_rejected() {
    assert!(FlatSpreadsheet::from_bytes(format!("text{FLAT_DDE}").into_bytes()).is_err());
    assert!(FlatSpreadsheet::from_bytes(format!("<![CDATA[]]>{FLAT_DDE}").into_bytes()).is_err());
    assert!(FlatSpreadsheet::from_bytes(format!("{FLAT_DDE}<![CDATA[]]>").into_bytes()).is_err());
    assert!(FlatSpreadsheet::from_bytes(format!("{FLAT_DDE}<!--outside-->").into_bytes()).is_err());
    assert!(FlatSpreadsheet::from_bytes(format!("{FLAT_DDE} ").into_bytes()).is_err());
    assert!(FlatSpreadsheet::from_bytes(format!(" \n{FLAT_DDE}").into_bytes()).is_ok());
}
