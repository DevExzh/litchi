//! Keynote slide-table cell parity cases (wave W1-C, record 0372): cell
//! writes through `slide::drawable::Edit::set_table_cell`.

use litchi_iwa::keynote;
use litchi_keynote::slide::drawable::{Kind, TableCellValue};
use litchi_keynote::{DrawableSelector, Package, SlideSelector, TableSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{leaf_position, leaf_run};

const TABLE: &str = "iwork/keynote/table-discovery.key";
const SLIDE: usize = 0;

fn host_model(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(SLIDE)?
        .first()
        .ok_or("no slide table")?
        .model_object_id)
}

/// Host read: every cell of the first slide table.
fn host_cells(bytes: &[u8]) -> TestResult<Vec<((usize, usize), String)>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let table = editor.slide_table(SLIDE, host_model(&editor)?)?;
    Ok(table
        .iter_cells()
        .map(|(at, value)| (at, format!("{value:?}")))
        .collect())
}

/// Leaf read: every cell of the first slide table.
fn leaf_cells(bytes: &[u8]) -> TestResult<Vec<((u32, u32), String)>> {
    let package = Package::from_bytes(bytes)?;
    let table = package.slide_table_cells(SlideSelector::index(SLIDE), TableSelector::index(0))?;
    Ok(table
        .iter_cells()
        .map(|cell| {
            (
                (cell.position().row(), cell.position().column()),
                format!("{:?}", cell.value()),
            )
        })
        .collect())
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    let input = fixture(TABLE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host cells", host_cells)
        .read("leaf cells", leaf_cells)
        .run()?;
    Ok(())
}

/// The first stored cell of the first slide table (a cleared absent cell
/// is not a clear).
fn first_stored(editor: &keynote::Draft) -> TestResult<(usize, usize)> {
    let table = editor.slide_table(SLIDE, host_model(editor)?)?;
    Ok(table
        .iter_cells()
        .map(|(position, _)| position)
        .next()
        .ok_or("table has no stored cell")?)
}

fn leaf_set(source: &[u8], cells: &[(u32, u32, TableCellValue)]) -> TestResult<LeafRun> {
    let table = leaf_position(source, Kind::Table)?;
    leaf_run(source, |edit| {
        for (row, column, value) in cells {
            edit.set_table_cell(
                SlideSelector::index(SLIDE),
                DrawableSelector::index(table),
                *row,
                *column,
                value.clone(),
            )?;
        }
        Ok(())
    })
}

#[test]
fn keynote_slide_table_cell_set_parity() -> TestResult {
    run(
        "keynote_slide_table_cell_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            editor.set_slide_table_cell(
                SLIDE,
                model,
                1,
                1,
                keynote::TableCellValue::Text("Parity".into()),
            )?;
            Ok(())
        },
        |source| leaf_set(source, &[(1, 1, TableCellValue::Text("Parity".into()))]),
    )
}

#[test]
fn keynote_slide_table_cells_set_parity() -> TestResult {
    run(
        "keynote_slide_table_cells_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            editor.set_slide_table_cells(
                SLIDE,
                model,
                [
                    keynote::TableCellUpdate::new(
                        1,
                        1,
                        keynote::TableCellValue::Number(42.5.try_into()?),
                    ),
                    keynote::TableCellUpdate::new(2, 1, keynote::TableCellValue::Boolean(true)),
                ],
            )?;
            Ok(())
        },
        |source| {
            leaf_set(
                source,
                &[
                    (1, 1, TableCellValue::Number(42.5)),
                    (2, 1, TableCellValue::Boolean(true)),
                ],
            )
        },
    )
}

#[test]
fn keynote_slide_table_cell_clear_parity() -> TestResult {
    run(
        "keynote_slide_table_cell_clear_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            let (row, column) = first_stored(editor)?;
            editor.clear_slide_table_cell(SLIDE, model, row, column)?;
            Ok(())
        },
        |source| {
            let (row, column) = first_stored(&keynote::Draft::from_bytes(source)?)?;
            leaf_set(
                source,
                &[(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                    TableCellValue::Empty,
                )],
            )
        },
    )
}
