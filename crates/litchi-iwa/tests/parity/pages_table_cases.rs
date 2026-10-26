//! Pages body-table parity cases of wave W1-B (record 0371): cell values
//! through the leaf `table::values` family and the cell reads.

use litchi_iwa::pages::{self, CellValue, TableCellUpdate};
use litchi_pages::Package;
use litchi_pages::table::values::Value;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const CELLS: &str = "iwork/pages/body-table-cells-native.pages";
/// Native two-table source without formulas (the values family refuses
/// cells that take part in a formula graph, as `body-table-cells-native`
/// has).
const CATALOG: &str = "iwork/pages/body-table-catalog-native.pages";
const TABLE: usize = 0;

fn host_model(editor: &pages::Draft) -> TestResult<u64> {
    Ok(editor
        .tables()?
        .get(TABLE)
        .ok_or("body table missing")?
        .model_object_id)
}

fn host_cells(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let table = editor.table(host_model(&editor)?)?;
    let mut cells = table
        .iter_cells()
        .map(|((row, column), value)| format!("{row},{column}: {value:?}"))
        .collect::<Vec<_>>();
    cells.push(format!("count {}", table.cell_count()));
    Ok(cells)
}

fn leaf_cells(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?.body_table_cells(TABLE)?
    ))
}

fn host_value(value: &Value) -> TestResult<CellValue> {
    Ok(match value {
        Value::Text(text) => CellValue::Text(text.clone()),
        Value::Number(number) => CellValue::number(*number)?,
        Value::Boolean(flag) => CellValue::Boolean(*flag),
        Value::Empty => CellValue::Empty,
        _ => return Err("unmapped value".into()),
    })
}

fn values_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    updates: Vec<(u32, u32, Value)>,
) -> TestResult {
    let input = fixture(CATALOG)?;
    let updates = &updates;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            if let [(row, column, value)] = updates.as_slice() {
                editor.set_table_cell(
                    model,
                    *row as usize,
                    *column as usize,
                    host_value(value)?,
                )?;
            } else {
                let batch = updates
                    .iter()
                    .map(|(row, column, value)| {
                        Ok(TableCellUpdate::new(
                            *row as usize,
                            *column as usize,
                            host_value(value)?,
                        ))
                    })
                    .collect::<TestResult<Vec<_>>>()?;
                editor.set_table_cells(model, batch)?;
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(TABLE)?;
            for (row, column, value) in updates {
                edit.set(*row, *column, value.clone())?;
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_table_values(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host table cells", host_cells)
        .read("leaf body_table_cells", leaf_cells)
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_cell_parity() -> TestResult {
    values_case(
        "pages_body_table_cell_parity",
        Tier::A,
        Evidence::Queued,
        vec![(1, 3, Value::Text("Zürich".to_owned()))],
    )
}

#[test]
fn pages_body_table_cells_batch_parity() -> TestResult {
    values_case(
        "pages_body_table_cells_batch_parity",
        Tier::A,
        Evidence::Queued,
        vec![
            (2, 3, Value::Number(42.25)),
            (1, 2, Value::Boolean(false)),
            (2, 0, Value::Empty),
        ],
    )
}

#[test]
fn pages_body_table_cells_read_parity() -> TestResult {
    let input = fixture(CELLS)?;
    Case::new(
        "pages_body_table_cells_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Native("test-data/iwork/pages/body-table-cells-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host table cells", host_cells)
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    // The two reads agree value for value.
    let host = pages::Draft::from_bytes(&input.bytes)?;
    let host_table = host.table(host_model(&host)?)?;
    let leaf = Package::from_bytes(&input.bytes)?
        .body_table_cells(TABLE)?
        .into_table();
    for ((row, column), value) in host_table.iter_cells() {
        let at = litchi_iwa_common::table::coordinate::At::new(
            u32::try_from(row)?,
            u32::try_from(column)?,
        );
        assert_eq!(leaf.get(at), Some(value), "{row},{column}");
    }
    Ok(())
}

const COMMENTS: &str = "iwork/pages/body-table-comment-hidden-native-saved.pages";

/// The table list, merges and cell comments read the same through the host
/// (`tables`, `Table::{merges, get_comment, iter_comments, comment_count}`,
/// `table_cell_comment[_replies]`) and the leaf (`body_tables`,
/// `body_table_merges`, `body_table_cells`) on native fixtures.
#[test]
fn pages_body_table_reads_parity() -> TestResult {
    let input = fixture(COMMENTS)?;
    Case::new(
        "pages_body_table_reads_parity",
        Format::Pages,
        Tier::A,
        Evidence::Native(
            "test-data/iwork/pages/body-table-comment-hidden-native-saved-receipt.json",
        ),
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    for path in [
        COMMENTS,
        CELLS,
        CATALOG,
        "iwork/pages/body-table-merges-native.pages",
    ] {
        let input = fixture(path)?;
        let host = pages::Draft::from_bytes(&input.bytes)?;
        let package = Package::from_bytes(&input.bytes)?;
        let catalog = package.body_tables()?;
        let tables = host.tables()?;
        assert_eq!(tables.len(), catalog.len(), "{path}");
        for (index, info) in tables.iter().enumerate() {
            let entry = catalog.get(index).ok_or("catalog entry")?;
            assert_eq!(entry.name(), info.name, "{path}");
            let host_table = host.table(info.model_object_id)?;
            let leaf_read = package.body_table_cells(index)?;
            assert_eq!(
                host_table.comment_count(),
                leaf_read.comment_count(),
                "{path}"
            );
            for ((row, column), comment) in host_table.iter_comments() {
                let at = litchi_iwa_common::table::coordinate::At::new(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                );
                assert_eq!(leaf_read.get_comment(at), Some(comment), "{path}");
                assert_eq!(host_table.get_comment(row, column), Some(comment));
                let native = host
                    .table_cell_comment(info.model_object_id, row, column)?
                    .ok_or("host cell comment")?;
                assert_eq!(native.comment.text, comment.text(), "{path}");
                assert_eq!(
                    host.table_cell_comment_replies(info.model_object_id, row, column)?
                        .len(),
                    comment.reply_count(),
                    "{path}"
                );
            }
            assert_eq!(
                host_table.merges(),
                package.body_table_merges(index)?.as_slice(),
                "{path}"
            );
        }
    }
    Ok(())
}
