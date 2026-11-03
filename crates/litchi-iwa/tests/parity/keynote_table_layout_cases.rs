//! Keynote slide-table lifecycle (record 0412g): the host's
//! `resize_slide_table` and `set_slide_table_geometry` against the leaf's
//! `edit_slide_table_layout` over the shared table engine
//! (`litchi_iwa_table::topology::{stage_resize, stage_frame,
//! stage_uniform_dimensions}`), and the host's `slide_tables` /
//! `slide_table` reads against the leaf's drawable, name, cell and merge
//! reads.

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};
use litchi_keynote::slide::drawable::{Geometry, Kind};
use litchi_keynote::{Package, SlideSelector, TableSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_table_cases::{SLIDE, TABLE, host_cells, host_model, leaf_cells};
use crate::leaf_bytes;

#[derive(Clone, Copy)]
enum Op {
    Resize(u32, u32),
    Frame(Geometry),
}

fn host(editor: &mut keynote::Draft, op: Op) -> TestResult {
    let info = editor
        .slide_tables(SLIDE)?
        .into_iter()
        .next()
        .ok_or("no slide table")?;
    match op {
        Op::Resize(rows, columns) => editor.resize_slide_table(
            SLIDE,
            info.model_object_id,
            rows as usize,
            columns as usize,
        )?,
        Op::Frame(geometry) => editor.set_slide_table_geometry(
            SLIDE,
            info.drawable_object_id,
            DrawableGeometry {
                position: Some(DrawablePoint {
                    x: geometry.x,
                    y: geometry.y,
                }),
                size: Some(DrawableSize {
                    width: geometry.width,
                    height: geometry.height,
                }),
                flags: info.geometry.flags,
                angle: info.geometry.angle,
            },
        )?,
    }
    Ok(())
}

fn leaf(package: &Package, op: Op) -> TestResult<litchi_keynote::slide::table::layout::Commit> {
    let edit =
        package.edit_slide_table_layout(SlideSelector::index(SLIDE), TableSelector::index(0))?;
    let edit = match op {
        Op::Resize(rows, columns) => edit.resize(rows, columns)?,
        Op::Frame(geometry) => edit.set_frame(geometry)?,
    };
    Ok(edit.commit()?)
}

/// Host read: the first table's size and frame.
fn host_layout(bytes: &[u8]) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let info = editor
        .slide_tables(SLIDE)?
        .into_iter()
        .next()
        .ok_or("no slide table")?;
    let position = info.geometry.position.ok_or("no position")?;
    let size = info.geometry.size.ok_or("no size")?;
    Ok(format!(
        "{}x{} at ({}, {}) size ({}, {})",
        info.rows, info.columns, position.x, position.y, size.width, size.height
    ))
}

/// Leaf read: the first table's size and frame.
fn leaf_layout(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let cells = package.slide_table_cells(SlideSelector::index(SLIDE), TableSelector::index(0))?;
    let geometry = package
        .slide_drawable_list(SlideSelector::index(SLIDE))?
        .iter()
        .find(|drawable| drawable.kind() == Kind::Table)
        .and_then(|drawable| drawable.geometry())
        .ok_or("no table frame")?;
    Ok(format!(
        "{}x{} at ({}, {}) size ({}, {})",
        cells.row_count(),
        cells.column_count(),
        geometry.x,
        geometry.y,
        geometry.width,
        geometry.height
    ))
}

fn layout_case(name: &'static str, tier: Tier, evidence: Evidence, op: Op) -> TestResult {
    let input = fixture(TABLE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(move |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor, op)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            let commit = leaf(&package, op)?;
            let restored = commit
                .package()
                .apply_slide_table_layout(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host layout", host_layout)
        .read("leaf layout", leaf_layout)
        .read("host cells", host_cells)
        .read("leaf cells", leaf_cells)
        .run()?;
    Ok(())
}

fn current(bytes: &[u8]) -> TestResult<(u32, u32, Geometry)> {
    let package = Package::from_bytes(bytes)?;
    let cells = package.slide_table_cells(SlideSelector::index(SLIDE), TableSelector::index(0))?;
    let geometry = package
        .slide_drawable_list(SlideSelector::index(SLIDE))?
        .iter()
        .find(|drawable| drawable.kind() == Kind::Table)
        .and_then(|drawable| drawable.geometry())
        .ok_or("no table frame")?;
    Ok((cells.row_count(), cells.column_count(), geometry))
}

#[test]
fn keynote_slide_table_resize_parity() -> TestResult {
    let (rows, columns, _) = current(&fixture(TABLE)?.bytes)?;
    layout_case(
        "keynote_slide_table_resize_parity",
        Tier::A,
        Evidence::Queued,
        Op::Resize(rows + 2, columns + 1),
    )
}

#[test]
fn keynote_slide_table_frame_parity() -> TestResult {
    let (_, _, geometry) = current(&fixture(TABLE)?.bytes)?;
    layout_case(
        "keynote_slide_table_frame_parity",
        Tier::A,
        Evidence::Queued,
        Op::Frame(Geometry::new(
            geometry.x + 12.0,
            geometry.y + 8.0,
            geometry.width + 60.0,
            geometry.height + 30.0,
        )),
    )
}

/// The slide's table list: names, sizes and frames read alike.
#[test]
fn keynote_slide_tables_read_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    for path in [TABLE, "iwork/keynote/slide-table-read-native.key"] {
        let input = fixture(path)?;
        let editor = keynote::Draft::from_bytes(&input.bytes)?;
        let package = Package::from_bytes(&input.bytes)?;
        let host = editor.slide_tables(SLIDE)?;
        let frames: Vec<Geometry> = package
            .slide_drawable_list(SlideSelector::index(SLIDE))?
            .iter()
            .filter(|drawable| drawable.kind() == Kind::Table)
            .filter_map(|drawable| drawable.geometry())
            .collect();
        assert_eq!(host.len(), frames.len(), "{path}: table count");
        for (index, (info, frame)) in host.iter().zip(&frames).enumerate() {
            let cells = package
                .slide_table_cells(SlideSelector::index(SLIDE), TableSelector::index(index))?;
            let name = package
                .slide_table_name(SlideSelector::index(SLIDE), TableSelector::index(index))?;
            assert_eq!(info.name, name.as_str(), "{path}: name");
            assert_eq!(
                (info.rows, info.columns),
                (cells.row_count() as usize, cells.column_count() as usize),
                "{path}: size"
            );
            let position = info.geometry.position.ok_or("no position")?;
            let size = info.geometry.size.ok_or("no size")?;
            assert_eq!(
                (position.x, position.y, size.width, size.height),
                (frame.x, frame.y, frame.width, frame.height),
                "{path}: frame"
            );
        }
    }
    Ok(())
}

/// One table's cells and merged regions read alike.
#[test]
fn keynote_slide_table_cells_and_merges_read_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let input = fixture(TABLE)?;
    let host = host_cells(&input.bytes)?;
    let leaf = leaf_cells(&input.bytes)?;
    assert_eq!(
        host.iter()
            .map(|((row, column), value)| ((*row as u32, *column as u32), value.clone()))
            .collect::<Vec<_>>(),
        leaf
    );
    let editor = keynote::Draft::from_bytes(&input.bytes)?;
    let model = host_model(&editor)?;
    let package = Package::from_bytes(&input.bytes)?;
    assert_eq!(
        editor.slide_table(SLIDE, model)?.merges(),
        package
            .slide_table_merges(SlideSelector::index(SLIDE), TableSelector::index(0))?
            .as_slice()
    );
    Ok(())
}

fn probe(op: Op) -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(TABLE)?,
        |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor, op)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            Ok(leaf_bytes!(leaf(&package, op)?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_resize_cost_abba() -> TestResult {
    let (rows, columns, _) = current(&fixture(TABLE)?.bytes)?;
    probe(Op::Resize(rows + 2, columns + 1))
}

#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_frame_cost_abba() -> TestResult {
    let (_, _, geometry) = current(&fixture(TABLE)?.bytes)?;
    probe(Op::Frame(Geometry::new(
        geometry.x + 12.0,
        geometry.y + 8.0,
        geometry.width + 60.0,
        geometry.height + 30.0,
    )))
}
