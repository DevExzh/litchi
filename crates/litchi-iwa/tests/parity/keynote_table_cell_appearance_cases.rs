//! Keynote slide-table cell fill, text layout and borders (wave W1-C,
//! record 0372): the host's `{set,reset}_slide_table_cell_{fill,layout}`,
//! `{set,clear}_slide_table_cell_border` and the cell reads against
//! `litchi_keynote::slide::table::cell_appearance` on the shared
//! `litchi_iwa_table::cell_appearance` engine (the Pages cases of record
//! 0371, on Keynote decks).

use litchi_iwa::keynote;
use litchi_iwa::shapes::{RgbColorSpace, RgbaColor, ShapeFill};
use litchi_iwa_common::table::cell::layout::{Inset, Insets, Layout, TextWrap, VerticalAlignment};
use litchi_keynote::slide::table::TableSelector;
use litchi_keynote::slide::table::cell_appearance::{BorderSide, Borders, Fill, Stroke};
use litchi_keynote::{Package, SlideSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const TABLES: &str = "iwork/keynote/slide-table-read-native.key";
const SLIDE: usize = 0;

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

fn model(editor: &keynote::Draft, table: usize) -> TestResult<u64> {
    Ok(editor
        .slide_tables(SLIDE)?
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
    Border(u32, u32, BorderSide, Option<Stroke>),
}

fn host_route(source: &[u8], table: usize, ops: &[Op]) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = model(&editor, table)?;
    for op in ops {
        match op {
            Op::SetFill(row, column, fill) => {
                editor.set_slide_table_cell_fill(
                    SLIDE,
                    model,
                    *row as usize,
                    *column as usize,
                    &host_fill(fill),
                )?;
            },
            Op::ResetFill(row, column) => {
                editor.reset_slide_table_cell_fill(
                    SLIDE,
                    model,
                    *row as usize,
                    *column as usize,
                )?;
            },
            Op::SetLayout(row, column, layout) => {
                editor.set_slide_table_cell_layout(
                    SLIDE,
                    model,
                    *row as usize,
                    *column as usize,
                    *layout,
                )?;
            },
            Op::ResetLayout(row, column) => {
                editor.reset_slide_table_cell_layout(
                    SLIDE,
                    model,
                    *row as usize,
                    *column as usize,
                )?;
            },
            Op::Border(row, column, side, Some(stroke)) => {
                editor.set_slide_table_cell_border(
                    SLIDE,
                    model,
                    *row as usize,
                    *column as usize,
                    *side,
                    *stroke,
                )?;
            },
            Op::Border(row, column, side, None) => {
                editor.clear_slide_table_cell_border(
                    SLIDE,
                    model,
                    *row as usize,
                    *column as usize,
                    *side,
                )?;
            },
        }
    }
    Ok(editor.to_bytes()?)
}

fn leaf_commit(
    source: &[u8],
    table: usize,
    ops: &[Op],
) -> TestResult<litchi_keynote::slide::table::cell_appearance::Commit> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_table_cell_appearance(
        SlideSelector::index(SLIDE),
        TableSelector::index(table),
    )?;
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
            Op::Border(row, column, side, Some(stroke)) => {
                edit.set_border(*row, *column, *side, *stroke);
            },
            Op::Border(row, column, side, None) => {
                edit.clear_border(*row, *column, *side);
            },
        }
    }
    Ok(edit.commit()?)
}

fn leaf_route(source: &[u8], table: usize, ops: &[Op]) -> TestResult<LeafRun> {
    let commit = leaf_commit(source, table, ops)?;
    let restored = commit
        .package()
        .apply_slide_table_cell_appearance(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Every cell's host fill and layout of table `table`.
fn host_cells(bytes: &[u8], table: usize) -> TestResult<Vec<(Fill, Layout, Borders)>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let info = editor
        .slide_tables(SLIDE)?
        .get(table)
        .ok_or("no such table")?
        .clone();
    let mut cells = Vec::new();
    for row in 0..info.rows {
        for column in 0..info.columns {
            let model = info.model_object_id;
            cells.push((
                leaf_fill(&editor.slide_table_cell_fill(SLIDE, model, row, column)?),
                editor.slide_table_cell_layout(SLIDE, model, row, column)?,
                editor.slide_table_cell_borders(SLIDE, model, row, column)?,
            ));
        }
    }
    Ok(cells)
}

fn leaf_cells(bytes: &[u8], table: usize) -> TestResult<Vec<(Fill, Layout, Borders)>> {
    let package = Package::from_bytes(bytes)?;
    let info = keynote::Draft::from_bytes(bytes)?
        .slide_tables(SLIDE)?
        .get(table)
        .ok_or("no such table")?
        .clone();
    let (rows, columns) = (u32::try_from(info.rows)?, u32::try_from(info.columns)?);
    let mut cells = Vec::new();
    for row in 0..rows {
        for column in 0..columns {
            let appearance = package.slide_table_cell_appearance(
                SlideSelector::index(SLIDE),
                TableSelector::index(table),
                row,
                column,
            )?;
            cells.push((appearance.fill, appearance.layout, appearance.borders));
        }
    }
    Ok(cells)
}

fn run(name: &'static str, tier: Tier, input: &Fixture, table: usize, ops: &[Op]) -> TestResult {
    let host_ops = ops.to_vec();
    let leaf_ops = ops.to_vec();
    Case::new(name, Format::Keynote, tier, Evidence::Queued, input)
        .leaf_invalidates_root_previews()
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
        path: "iwork/keynote/slide-table-read-native.key (host-styled cell 1,1)",
        bytes,
    })
}

#[test]
fn keynote_slide_table_cell_fill_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "keynote_slide_table_cell_fill_set_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetFill(1, 1, Fill::Solid(color(0.2, 0.5, 0.8)?))],
    )
}

#[test]
fn keynote_slide_table_cell_fill_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "keynote_slide_table_cell_fill_reset_parity",
        Tier::B,
        &input,
        0,
        &[Op::ResetFill(1, 1), Op::SetFill(2, 2, Fill::None)],
    )
}

#[test]
fn keynote_slide_table_cell_layout_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "keynote_slide_table_cell_layout_set_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetLayout(2, 1, layout()?)],
    )
}

#[test]
fn keynote_slide_table_cell_layout_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "keynote_slide_table_cell_layout_reset_parity",
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
fn keynote_slide_table_cell_style_removal_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "keynote_slide_table_cell_style_removal_parity",
        Tier::B,
        &input,
        0,
        &[Op::ResetFill(1, 1), Op::ResetLayout(1, 1)],
    )
}

/// Two new variations in one edit, then a rewrite of the first.
#[test]
fn keynote_slide_table_cell_appearance_sequence_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "keynote_slide_table_cell_appearance_sequence_parity",
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
fn keynote_slide_table_cell_border_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "keynote_slide_table_cell_border_parity",
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
fn keynote_slide_table_cell_border_clear_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = fixture(TABLES)?;
    run(
        "keynote_slide_table_cell_border_clear_parity",
        Tier::B,
        &input,
        0,
        &[Op::Border(0, 1, BorderSide::Bottom, None)],
    )
}

/// Rewriting a cell's own variation in place (the host-styled cell).
#[test]
fn keynote_slide_table_cell_fill_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = styled()?;
    run(
        "keynote_slide_table_cell_fill_rewrite_parity",
        Tier::B,
        &input,
        0,
        &[Op::SetFill(1, 1, Fill::Solid(color(0.1, 0.8, 0.3)?))],
    )
}

/// Every cell's fill and layout on four native fixtures.
#[test]
fn keynote_slide_table_cell_appearance_read_parity() -> TestResult {
    let styled = styled()?;
    Case::new(
        "keynote_slide_table_cell_appearance_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &styled,
    )
    .read_only()
    .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host cells", |bytes: &[u8]| host_cells(bytes, 0))
    .read("leaf cells", |bytes: &[u8]| leaf_cells(bytes, 0))
    .run()?;
    for path in [
        TABLES,
        "iwork/keynote/table-discovery.key",
        "iwork/keynote/slide-table-merges-native.key",
    ] {
        let document = fixture(path)?;
        let tables = keynote::Draft::from_bytes(&document.bytes)?
            .slide_tables(SLIDE)?
            .len();
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
