//! Native focused table values and comments compared with the migration host.

use std::error::Error;

use litchi_iwa::{keynote::KeynoteEditor, pages::PagesEditor};
use litchi_iwa_common::table::{
    cell::value::Value, coordinate::CellPosition, merge::Region, read::TableRead,
};

const PAGES: &[u8] = include_bytes!("../../../test-data/iwork/pages/body-table-read-native.pages");
const KEYNOTE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/slide-table-read-native.key");
type TestResult = Result<(), Box<dyn Error>>;

#[test]
fn native_pages_cross_table_formula_preserves_its_table_name() -> TestResult {
    let source =
        include_bytes!("../../../test-data/iwork/pages/body-table-cross-reference-native.pages");
    let package = litchi_pages::Package::from_bytes(source)?;
    let read = package.body_table_cells("Table 2")?;
    assert_eq!(
        read.get_a1("A3")?,
        Some(&Value::Formula("=(Table 1::B2+1)".into()))
    );
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    assert_eq!(output, source);
    Ok(())
}

fn cells(read: &TableRead) -> Vec<((usize, usize), Value)> {
    read.iter_cells()
        .map(|cell| {
            (
                (
                    cell.position().row() as usize,
                    cell.position().column() as usize,
                ),
                cell.value().clone(),
            )
        })
        .collect()
}

#[test]
fn native_pages_full_table_reads_match_host_and_preserve_source() -> TestResult {
    let focused = litchi_pages::Package::from_bytes(PAGES)?;
    let host = PagesEditor::from_bytes(PAGES)?;
    let catalog = host.tables()?;
    assert_eq!(catalog.len(), 2);
    for (position, info) in catalog.iter().enumerate() {
        let read = focused.body_table_cells(position)?;
        let legacy = host.table(info.model_object_id)?;
        assert_eq!(
            cells(&read),
            legacy
                .iter_cells()
                .map(|(p, v)| (p, v.clone()))
                .collect::<Vec<_>>()
        );
        assert_eq!(read.cell_count(), legacy.cell_count());
        assert_eq!(read.comment_count(), legacy.comment_count());
        for (position, comment) in legacy.iter_comments() {
            let actual = read
                .get_comment(CellPosition::try_from_usize(position.0, position.1)?)
                .expect("every native comment survives focused readback");
            assert_eq!(actual.text(), comment.text);
            assert_eq!(
                actual.timestamp().map(|v| v.as_f64()),
                comment.creation_date_seconds
            );
        }
    }
    let first = focused.body_table_cells("Table 1")?;
    assert_eq!(
        first.get_a1("B4")?,
        Some(&Value::Formula("=SUM(B2:B3)".into()))
    );
    assert_eq!(first.get_a1("D2")?, Some(&Value::Text("北京".into())));
    let comment = first.get_comment_a1("B4")?.expect("native formula comment");
    assert_eq!(comment.text(), "Focused Pages table read — Café 北京");
    assert_eq!(
        comment.author().and_then(|author| author.display_name()),
        Some("Ryker Zhu")
    );
    assert!(comment.timestamp().is_some());
    assert_eq!(comment.replies(), Some([].as_slice()));
    let second = focused.body_table_cells("Table 2")?;
    assert_eq!(second.get_a1("A2")?, Some(&Value::Formula("=(1/0)".into())));
    let mut output = Vec::new();
    focused.write_to(&mut output)?;
    assert_eq!(output, PAGES);
    assert_eq!(host.to_bytes()?, PAGES);
    Ok(())
}

#[test]
fn native_keynote_full_table_reads_match_host_and_preserve_merge() -> TestResult {
    let focused = litchi_keynote::Package::from_bytes(KEYNOTE)?;
    let host = KeynoteEditor::from_bytes(KEYNOTE)?;
    let catalog = host.slide_tables(0)?;
    assert_eq!(catalog.len(), 1);
    let legacy = host.slide_table(0, catalog[0].model_object_id)?;
    let read = focused.slide_table_cells(0, 0)?;
    assert_eq!(
        cells(&read),
        legacy
            .iter_cells()
            .map(|(p, v)| (p, v.clone()))
            .collect::<Vec<_>>()
    );
    assert_eq!(read.cell_count(), legacy.cell_count());
    assert_eq!(read.get_a1("A2")?, Some(&Value::Text("Café 北京".into())));
    assert_eq!(read.get_a1("B2")?, Some(&Value::number(12.5)?));
    assert_eq!(read.get_a1("C2")?, Some(&Value::Boolean(true)));
    assert_eq!(read.get_a1("D2")?, Some(&Value::Formula("=(B2+1)".into())));
    assert_eq!(read.comment_count(), 1);
    let comment = read.get_comment_a1("D2")?.expect("native formula comment");
    assert_eq!(comment.text(), "Focused Keynote table read — Café 北京");
    assert_eq!(
        comment.text(),
        legacy.get_comment(1, 3).expect("host comment").text
    );
    assert_eq!(
        comment.author().and_then(|author| author.display_name()),
        Some("Ryker Zhu")
    );
    assert!(comment.timestamp().is_some());
    assert_eq!(comment.replies(), Some([].as_slice()));
    let merges = focused.slide_table_merges(0, 0)?;
    assert_eq!(merges, [Region::new(3, 1, 2, 2)?]);
    assert_eq!(merges, legacy.merges());
    let mut output = Vec::new();
    focused.write_to(&mut output)?;
    assert_eq!(output, KEYNOTE);
    assert_eq!(host.to_bytes()?, KEYNOTE);
    Ok(())
}
