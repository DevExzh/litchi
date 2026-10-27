//! Pages body-table cell fill and text layout (wave W1-B, record 0371): the
//! host's `{set,reset}_table_cell_{fill,layout}` and `table_cell_{fill,layout}`
//! against `litchi_pages::table::cell_appearance` on the shared
//! `litchi_iwa_table::cell_appearance` engine.

use litchi_iwa::pages;
use litchi_iwa::shapes::{RgbColorSpace, RgbaColor, ShapeFill};
use litchi_iwa_common::table::cell::layout::{Inset, Insets, Layout, TextWrap, VerticalAlignment};
use litchi_pages::Package;
use litchi_pages::table::cell_appearance::Fill;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const TABLES: &str = "iwork/pages/body-table-catalog-native.pages";

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

fn model(editor: &pages::Draft, table: usize) -> TestResult<u64> {
    Ok(editor
        .tables()?
        .get(table)
        .ok_or("no such table")?
        .model_object_id)
}

/// A staged operation on both routes.
#[derive(Clone)]
enum Op {
    SetFill(u32, u32, Fill),
    ResetFill(u32, u32),
    SetLayout(u32, u32, Layout),
    ResetLayout(u32, u32),
}

fn host_route(source: &[u8], table: usize, ops: &[Op]) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
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
        }
    }
    Ok(editor.to_bytes()?)
}

fn leaf_commit(
    source: &[u8],
    table: usize,
    ops: &[Op],
) -> TestResult<litchi_pages::table::cell_appearance::Commit> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_table_cell_appearance(table)?;
    for op in ops {
        match op {
            Op::SetFill(row, column, fill) => {
                edit.set_fill(*row, *column, fill.clone())?;
            },
            Op::ResetFill(row, column) => {
                edit.reset_fill(*row, *column);
            },
            Op::SetLayout(row, column, layout) => {
                edit.set_layout(*row, *column, *layout);
            },
            Op::ResetLayout(row, column) => {
                edit.reset_layout(*row, *column);
            },
        }
    }
    Ok(edit.commit()?)
}

fn leaf_route(source: &[u8], table: usize, ops: &[Op]) -> TestResult<LeafRun> {
    let commit = leaf_commit(source, table, ops)?;
    let restored = commit
        .package()
        .apply_body_table_cell_appearance(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Every cell's host fill and layout of table `table`.
fn host_cells(bytes: &[u8], table: usize) -> TestResult<Vec<(Fill, Layout)>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let info = editor.tables()?.get(table).ok_or("no such table")?.clone();
    let mut cells = Vec::new();
    for row in 0..info.rows {
        for column in 0..info.columns {
            cells.push((
                leaf_fill(&editor.table_cell_fill(info.model_object_id, row, column)?),
                editor.table_cell_layout(info.model_object_id, row, column)?,
            ));
        }
    }
    Ok(cells)
}

fn leaf_cells(bytes: &[u8], table: usize) -> TestResult<Vec<(Fill, Layout)>> {
    let package = Package::from_bytes(bytes)?;
    let catalog = package.body_tables()?;
    let entry = catalog.get(table).ok_or("no such table")?;
    let (rows, columns) = (entry.rows(), entry.columns());
    let mut cells = Vec::new();
    for row in 0..rows {
        for column in 0..columns {
            let appearance = package.body_table_cell_appearance(table, row, column)?;
            cells.push((appearance.fill, appearance.layout));
        }
    }
    Ok(cells)
}

fn run(name: &'static str, tier: Tier, input: &Fixture, table: usize, ops: &[Op]) -> TestResult {
    let host_ops = ops.to_vec();
    let leaf_ops = ops.to_vec();
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input);
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.host(move |source| host_route(source, table, &host_ops))
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
        path: "iwork/pages/body-table-catalog-native.pages (host-styled cell 1,1)",
        bytes,
    })
}

#[test]
fn pages_body_table_cell_fill_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "pages_body_table_cell_fill_set_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetFill(1, 1, Fill::Solid(color(0.2, 0.5, 0.8)?))],
    )
}

#[test]
fn pages_body_table_cell_fill_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "pages_body_table_cell_fill_reset_parity",
        Tier::A,
        &input,
        0,
        &[Op::ResetFill(1, 1), Op::SetFill(2, 2, Fill::None)],
    )
}

#[test]
fn pages_body_table_cell_layout_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "pages_body_table_cell_layout_set_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetLayout(2, 1, layout()?)],
    )
}

#[test]
fn pages_body_table_cell_layout_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "pages_body_table_cell_layout_reset_parity",
        Tier::A,
        &input,
        0,
        &[Op::ResetLayout(1, 1)],
    )
}

/// Resetting every override of a host-styled cell: its variation is
/// removed from the stylesheet, the registry and the style list, and the
/// cell falls back to the table's style.
#[test]
fn pages_body_table_cell_style_removal_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "pages_body_table_cell_style_removal_parity",
        Tier::A,
        &input,
        0,
        &[Op::ResetFill(1, 1), Op::ResetLayout(1, 1)],
    )
}

/// Two new variations in one edit, then a rewrite of the first.
#[test]
fn pages_body_table_cell_appearance_sequence_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "pages_body_table_cell_appearance_sequence_parity",
        Tier::B,
        &input,
        1,
        &[
            Op::SetFill(0, 0, Fill::Solid(color(0.3, 0.3, 0.3)?)),
            Op::SetLayout(2, 1, layout()?),
            Op::SetFill(0, 0, Fill::Solid(color(0.6, 0.1, 0.6)?)),
            Op::ResetLayout(2, 1),
        ],
    )
}

/// Rewriting a cell's own variation in place (the host-styled cell).
#[test]
fn pages_body_table_cell_fill_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "pages_body_table_cell_fill_rewrite_parity",
        Tier::A,
        &input,
        0,
        &[Op::SetFill(1, 1, Fill::Solid(color(0.1, 0.8, 0.3)?))],
    )
}

/// Every cell's fill and layout on four native fixtures.
#[test]
fn pages_body_table_cell_appearance_read_parity() -> TestResult {
    let styled = styled()?;
    Case::new(
        "pages_body_table_cell_appearance_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &styled,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host cells", |bytes: &[u8]| host_cells(bytes, 0))
    .read("leaf cells", |bytes: &[u8]| leaf_cells(bytes, 0))
    .run()?;
    for path in [
        TABLES,
        "iwork/pages/body-table-cells-native.pages",
        "iwork/pages/body-table-merges-native.pages",
        "iwork/pages/body-table-comment-hidden-native-saved.pages",
    ] {
        let document = fixture(path)?;
        let tables = pages::Draft::from_bytes(&document.bytes)?.tables()?.len();
        for table in 0..tables {
            assert_eq!(
                host_cells(&document.bytes, table)?,
                leaf_cells(&document.bytes, table)?,
                "{path} table {table}"
            );
        }
    }
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn cell_fill_set_cost_abba() -> TestResult {
    let ops = [Op::SetFill(1, 1, Fill::Solid(color(0.2, 0.5, 0.8)?))];
    crate::pages_cost_probes::abba(
        &fixture(TABLES)?,
        |source| host_route(source, 0, &ops),
        |source| Ok(leaf_bytes!(leaf_commit(source, 0, &ops)?.package())),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn cell_layout_reset_cost_abba() -> TestResult {
    let ops = [Op::ResetLayout(1, 1)];
    crate::pages_cost_probes::abba(
        &styled()?,
        |source| host_route(source, 0, &ops),
        |source| Ok(leaf_bytes!(leaf_commit(source, 0, &ops)?.package())),
    )
}
