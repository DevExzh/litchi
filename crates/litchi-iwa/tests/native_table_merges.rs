//! Native merge readback through the focused facades and migration host.

use std::error::Error;

use litchi_iwa::{keynote::KeynoteEditor, pages::PagesEditor};
use litchi_iwa_common::table::merge::Region;

const PAGES: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-merges-native.pages"
));
const KEYNOTE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/slide-table-merges-native.key"
));

type TestResult = Result<(), Box<dyn Error>>;

#[test]
fn native_pages_merge_reads_match_and_preserve_source() -> TestResult {
    let expected = vec![Region::new(3, 2, 1, 2)?];
    let focused = litchi_pages::Package::from_bytes(PAGES)?;
    assert_eq!(focused.body_table_merges("Table 1")?, expected);
    assert_eq!(focused.body_table_merges(0)?, expected);
    assert!(focused.body_table_merges("Table 2")?.is_empty());
    let mut output = Vec::new();
    focused.write_to(&mut output)?;
    assert_eq!(output, PAGES);

    let host = PagesEditor::from_bytes(PAGES)?;
    let tables = host.tables()?;
    assert_eq!(tables.len(), 2);
    assert_eq!(host.table_cell_merges(tables[0].model_object_id)?, expected);
    let table = host.table(tables[0].model_object_id)?;
    assert_eq!(table.merges(), expected);
    let focused_second = focused.body_table_merges(1)?;
    assert!(focused_second.is_empty());
    assert_eq!(
        host.table_cell_merges(tables[1].model_object_id)?,
        focused_second
    );
    assert_eq!(host.table(tables[1].model_object_id)?.merges(), &[]);
    let invalid_error = host
        .table_cell_merges(u64::MAX)
        .expect_err("an unknown Pages model identifier must be rejected");
    assert!(
        invalid_error
            .to_string()
            .contains("is not attached to the body")
    );
    assert_eq!(
        table.get_cell(1, 3),
        Some(&litchi_iwa::pages::PagesCellValue::Text("北京".into()))
    );
    assert_eq!(host.to_bytes()?, PAGES);
    Ok(())
}

#[test]
fn native_keynote_merge_reads_match_and_preserve_source() -> TestResult {
    let expected = vec![Region::new(3, 1, 2, 2)?];
    let focused = litchi_keynote::Package::from_bytes(KEYNOTE)?;
    assert_eq!(focused.slide_table_merges(0, 0)?, expected);
    let mut output = Vec::new();
    focused.write_to(&mut output)?;
    assert_eq!(output, KEYNOTE);

    let host = KeynoteEditor::from_bytes(KEYNOTE)?;
    let tables = host.slide_tables(0)?;
    assert_eq!(tables.len(), 1);
    assert_eq!(
        host.slide_table_cell_merges(0, tables[0].model_object_id)?,
        expected
    );
    let table = host.slide_table(0, tables[0].model_object_id)?;
    assert_eq!(table.merges(), expected);
    let invalid_error = host
        .slide_table_cell_merges(0, u64::MAX)
        .expect_err("an unknown Keynote model identifier must be rejected");
    assert!(
        invalid_error
            .to_string()
            .contains("is not owned by slide 0")
    );
    let invalid_slide_error = host
        .slide_table_cell_merges(1, tables[0].model_object_id)
        .expect_err("a table owned by another or missing slide must be rejected");
    assert!(invalid_slide_error.to_string().contains("slide index 1"));
    assert_eq!(host.to_bytes()?, KEYNOTE);
    Ok(())
}
