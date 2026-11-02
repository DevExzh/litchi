//! Numbers table cell fill, text layout and borders (record 0412b): the
//! host's `{set,reset}_table_cell_{fill,layout}`,
//! `{set,clear}_table_cell_border` and the cell reads against
//! `litchi_numbers::table::cell_appearance` on the shared
//! `litchi_iwa_table::cell_appearance` engine (the Keynote cases of record
//! 0372, on Numbers workbooks).

use litchi_iwa::numbers;
use litchi_iwa::shapes::{RgbColorSpace, RgbaColor, ShapeFill};
use litchi_iwa_common::table::cell::layout::{Inset, Insets, Layout, TextWrap, VerticalAlignment};
use litchi_numbers::Package;
use litchi_numbers::table::cell_appearance::{BorderSide, Borders, Fill, Stroke};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const TABLES: &str = "iwork/numbers/basic.numbers";

fn color(red: f32, green: f32, blue: f32) -> TestResult<RgbaColor> {
    Ok(RgbaColor::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

fn host_fill(fill: &Fill) -> ShapeFill {
    match fill {
        Fill::None => ShapeFill::None,
        Fill::Solid(color) => ShapeFill::Solid(*color),
        Fill::Gradient(gradient) => ShapeFill::Gradient(gradient.clone()),
        Fill::Image(_) => unreachable!("image fills are refused"),
    }
}

fn leaf_fill(fill: &ShapeFill) -> Fill {
    match fill {
        ShapeFill::None => Fill::None,
        ShapeFill::Solid(color) => Fill::Solid(*color),
        ShapeFill::Gradient(gradient) => Fill::Gradient(gradient.clone()),
        ShapeFill::Image(_) => Fill::None,
    }
}

fn layout() -> TestResult<Layout> {
    Ok(Layout::new(
        TextWrap::Wrapped,
        VerticalAlignment::Middle,
        Insets::new(
            Inset::from_points(3.0)?,
            Inset::from_points(4.5)?,
            Inset::from_points(6.0)?,
            Inset::from_points(1.0)?,
        ),
    ))
}

fn model(editor: &numbers::Draft, table: usize) -> TestResult<u64> {
    let _ = table;
    crate::numbers_cases::host_first_table_model_id(editor)
}

/// A staged operation on both routes.
#[derive(Clone)]
enum Op {
    SetFill(u32, u32, Fill),
    ResetFill(u32, u32),
    SetLayout(u32, u32, Layout),
    ResetLayout(u32, u32),
    Border(u32, u32, BorderSide, Option<Stroke>),
}

fn host_route(source: &[u8], table: usize, ops: &[Op]) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let model = model(&editor, table)?;
    for op in ops {
        match op {
            Op::SetFill(row, column, fill) => {
                editor.set_table_cell_fill(
                    model,
                    *row as usize,
                    *column as usize,
                    &host_fill(fill),
                )?;
            },
            Op::ResetFill(row, column) => {
                editor.reset_table_cell_fill(model, *row as usize, *column as usize)?;
            },
            Op::SetLayout(row, column, layout) => {
                editor.set_table_cell_layout(model, *row as usize, *column as usize, *layout)?;
            },
            Op::ResetLayout(row, column) => {
                editor.reset_table_cell_layout(model, *row as usize, *column as usize)?;
            },
            Op::Border(row, column, side, Some(stroke)) => {
                editor.set_table_cell_border(
                    model,
                    *row as usize,
                    *column as usize,
                    *side,
                    *stroke,
                )?;
            },
            Op::Border(row, column, side, None) => {
                editor.clear_table_cell_border(model, *row as usize, *column as usize, *side)?;
            },
        }
    }
    Ok(editor.to_bytes()?)
}

fn leaf_commit(
    source: &[u8],
    table: usize,
    ops: &[Op],
) -> TestResult<litchi_numbers::table::cell_appearance::Commit> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_table_cell_appearance(0usize, table)?;
    for op in ops {
        match op {
            Op::SetFill(row, column, fill) => {
                edit.set_fill(*row, *column, fill.clone())?;
            },
            Op::ResetFill(row, column) => {
                edit.reset_fill(*row, *column)?;
            },
            Op::SetLayout(row, column, layout) => {
                edit.set_layout(*row, *column, *layout)?;
            },
            Op::ResetLayout(row, column) => {
                edit.reset_layout(*row, *column)?;
            },
            Op::Border(row, column, side, Some(stroke)) => {
                edit.set_border(*row, *column, *side, *stroke)?;
            },
            Op::Border(row, column, side, None) => {
                edit.clear_border(*row, *column, *side)?;
            },
        }
    }
    Ok(edit.commit()?)
}

fn leaf_route(source: &[u8], table: usize, ops: &[Op]) -> TestResult<LeafRun> {
    let commit = leaf_commit(source, table, ops)?;
    let restored = commit
        .package()
        .apply_table_cell_appearance(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Every cell's host fill and layout of table `table`.
fn dimensions(bytes: &[u8], table: usize) -> TestResult<(u32, u32)> {
    let package = Package::from_bytes(bytes)?;
    let info = package.workbook_ref().sheets()[0]
        .tables()
        .nth(table)
        .ok_or("no such table")?;
    Ok((info.row_count().min(8), info.column_count().min(6)))
}

/// The first cells' host fill, layout and borders of table `table`.
fn host_cells(bytes: &[u8], table: usize) -> TestResult<Vec<(Fill, Layout, Borders)>> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let model = model(&editor, table)?;
    let (rows, columns) = dimensions(bytes, table)?;
    let mut cells = Vec::new();
    for row in 0..rows as usize {
        for column in 0..columns as usize {
            cells.push((
                leaf_fill(&editor.table_cell_fill(model, row, column)?),
                editor.table_cell_layout(model, row, column)?,
                editor.table_cell_borders(model, row, column)?,
            ));
        }
    }
    Ok(cells)
}

fn leaf_cells(bytes: &[u8], table: usize) -> TestResult<Vec<(Fill, Layout, Borders)>> {
    let package = Package::from_bytes(bytes)?;
    let (rows, columns) = dimensions(bytes, table)?;
    let mut cells = Vec::new();
    for row in 0..rows {
        for column in 0..columns {
            let appearance = package.table_cell_appearance(0usize, table, row, column)?;
            cells.push((appearance.fill, appearance.layout, appearance.borders));
        }
    }
    Ok(cells)
}

fn run(name: &'static str, tier: Tier, input: &Fixture, table: usize, ops: &[Op]) -> TestResult {
    let host_ops = ops.to_vec();
    let leaf_ops = ops.to_vec();
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(move |source| host_route(source, table, &host_ops))
        .leaf(move |source| leaf_route(source, table, &leaf_ops))
        .read("host cells", move |bytes: &[u8]| host_cells(bytes, table))
        .read("leaf cells", move |bytes: &[u8]| leaf_cells(bytes, table))
        .run()?;
    Ok(())
}

/// A fixture the host already styled: cell (1, 1) of table 0 has a private
/// fill and layout variation.
fn styled() -> TestResult<Fixture> {
    let input = fixture(TABLES)?;
    let bytes = host_route(
        &input.bytes,
        0,
        &[
            Op::SetFill(1, 1, Fill::Solid(color(0.9, 0.2, 0.1)?)),
            Op::SetLayout(1, 1, layout()?),
        ],
    )?;
    Ok(Fixture {
        path: "iwork/numbers/basic.numbers (host-styled cell 1,1)",
        bytes,
    })
}

#[test]
fn numbers_table_cell_fill_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "numbers_table_cell_fill_set_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetFill(1, 1, Fill::Solid(color(0.2, 0.5, 0.8)?))],
    )
}

#[test]
fn numbers_table_cell_fill_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "numbers_table_cell_fill_reset_parity",
        Tier::B,
        &input,
        0,
        &[Op::ResetFill(1, 1), Op::SetFill(2, 2, Fill::None)],
    )
}

#[test]
fn numbers_table_cell_layout_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "numbers_table_cell_layout_set_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetLayout(2, 1, layout()?)],
    )
}

#[test]
fn numbers_table_cell_layout_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "numbers_table_cell_layout_reset_parity",
        Tier::B,
        &input,
        0,
        &[Op::ResetLayout(1, 1)],
    )
}

/// Resetting every override of a host-styled cell: its variation is
/// removed from the stylesheet, the registry and the style list, and the
/// cell falls back to the table's style.
#[test]
fn numbers_table_cell_style_removal_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "numbers_table_cell_style_removal_parity",
        Tier::B,
        &input,
        0,
        &[Op::ResetFill(1, 1), Op::ResetLayout(1, 1)],
    )
}

/// Two new variations in one edit, then a rewrite of the first.
#[test]
fn numbers_table_cell_appearance_sequence_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "numbers_table_cell_appearance_sequence_parity",
        Tier::B,
        &input,
        0,
        &[
            Op::SetFill(0, 0, Fill::Solid(color(0.3, 0.3, 0.3)?)),
            Op::SetLayout(2, 1, layout()?),
            Op::SetFill(0, 0, Fill::Solid(color(0.6, 0.1, 0.6)?)),
            Op::ResetLayout(2, 1),
        ],
    )
}

fn stroke(red: f32, width: f32) -> TestResult<Stroke> {
    Ok(Stroke::new(
        color(red, 0.2, 0.2)?,
        litchi_iwa_common::shape::stroke::Width::new(width)?,
        litchi_iwa_common::shape::stroke::Pattern::Solid,
    ))
}

/// Cell borders: a new layer, a second run in it, an in-place restroke of
/// the cell's own run, and an explicitly cleared edge.
#[test]
fn numbers_table_cell_border_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "numbers_table_cell_border_parity",
        Tier::B,
        &input,
        0,
        &[
            Op::Border(1, 1, BorderSide::Top, Some(stroke(0.9, 2.0)?)),
            Op::Border(1, 2, BorderSide::Top, Some(stroke(0.5, 1.0)?)),
            Op::Border(1, 1, BorderSide::Top, Some(stroke(0.1, 3.0)?)),
            Op::Border(2, 0, BorderSide::Left, None),
        ],
    )
}

/// Clearing one edge of a cell.
#[test]
fn numbers_table_cell_border_clear_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "numbers_table_cell_border_clear_parity",
        Tier::B,
        &input,
        0,
        &[Op::Border(0, 1, BorderSide::Bottom, None)],
    )
}

/// Rewriting a cell's own variation in place (the host-styled cell).
#[test]
fn numbers_table_cell_fill_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "numbers_table_cell_fill_rewrite_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetFill(1, 1, Fill::Solid(color(0.1, 0.8, 0.3)?))],
    )
}

/// The first cells' fill, layout and borders on the host-styled cell and
/// on the corpus tables both routes read.
#[test]
fn numbers_table_cell_appearance_read_parity() -> TestResult {
    let styled = styled()?;
    Case::new(
        "numbers_table_cell_appearance_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &styled,
    )
    .read_only()
    .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host cells", |bytes: &[u8]| host_cells(bytes, 0))
    .read("leaf cells", |bytes: &[u8]| leaf_cells(bytes, 0))
    .run()?;
    let mut compared = 0_usize;
    for path in crate::numbers_cases::NUMBERS_CORPUS {
        let document = fixture(path)?;
        let (Ok(host), Ok(leaf)) = (
            host_cells(&document.bytes, 0),
            leaf_cells(&document.bytes, 0),
        ) else {
            continue;
        };
        assert_eq!(host, leaf, "{path}");
        compared += 1;
    }
    eprintln!("cell appearance tables compared: {compared}");
    assert!(compared >= 3, "only {compared} tables compared");
    Ok(())
}

/// Callgrind probe (record 0412b): a solid fill on B2, host against leaf,
/// ABBA.
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn numbers_table_cell_appearance_cost_abba() -> TestResult {
    let input = fixture(TABLES)?;
    let ops = vec![Op::SetFill(1, 1, Fill::Solid(color(0.2, 0.5, 0.8)?))];
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, 0, &ops),
        |source| Ok(leaf_bytes!(leaf_commit(source, 0, &ops)?.package())),
    )
}
