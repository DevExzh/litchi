//! Numbers and Pages chart-lifecycle parity (wave W4-D, record 0416): the
//! chart kind and series direction (`sheet::chart::lifecycle`,
//! `chart::lifecycle`), the complete data (`sheet::chart::data`,
//! `chart::data`) and the host's chart lists.
//!
//! The host decodes and re-encodes the chart drawable with Prost, so the
//! writes are Tier B (the leaf rewrites the two fields in place). None of
//! these host routes has native evidence; every route is macOS-queued.

use litchi_iwa::numbers;
use litchi_iwa::pages;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::{ChartData, Direction};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{chart_fixture, first_sheet, host_drawable};

fn regrid() -> TestResult<ChartData> {
    Ok(ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned()],
        vec![vec![Some(5.0), Some(9.5)], vec![Some(-2.0), None]],
    )?)
}

#[derive(Clone)]
enum Op {
    Kind(Kind),
    Direction(Direction),
    Data(ChartData),
}

// ---------------------------------------------------------------------------
// Numbers

fn numbers_host(source: &[u8], op: &Op) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let (sheet, id) = host_drawable(&editor, litchi_numbers::sheet::drawable::Kind::Chart, 0)?;
    match op {
        Op::Kind(kind) => editor.set_sheet_chart_kind(sheet, id, *kind)?,
        Op::Direction(direction) => editor.set_sheet_chart_direction(sheet, id, *direction)?,
        Op::Data(data) => editor.set_sheet_chart_data(sheet, id, data.clone())?,
    }
    Ok(editor.to_bytes()?)
}

fn numbers_leaf(source: &[u8], op: &Op) -> TestResult<LeafRun> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    if let Op::Data(data) = op {
        let commit = package
            .edit_sheet_chart_data(0, 0)?
            .set(data.clone())
            .commit()?;
        let restored = commit
            .package()
            .apply_sheet_chart_data(&commit.patch().inverse())?;
        return Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ));
    }
    let mut edit = package.edit_sheet_chart_lifecycle(0, 0)?;
    match op {
        Op::Kind(kind) => drop(edit.set_kind(*kind)?),
        Op::Direction(direction) => drop(edit.set_direction(*direction)?),
        Op::Data(_) => return Err("unreachable".into()),
    }
    let commit = edit.commit()?;
    if !commit.diagnostics().changed() {
        return Err("leaf chart-lifecycle edit was a no-op".into());
    }
    let restored = commit
        .package()
        .apply_sheet_chart_lifecycle(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// The host's chart list: kind, direction, data, frame and Arrange flags.
pub(crate) fn numbers_host_list(bytes: &[u8]) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let sheet = first_sheet(&editor)?;
    let mut out = String::new();
    for chart in editor.sheet_charts(sheet)? {
        out.push_str(&format!(
            "{:?} {:?} {:?} {:?} {:?}\n",
            chart.kind, chart.direction, chart.data, chart.geometry, chart.arrangement
        ));
    }
    Ok(out)
}

pub(crate) fn numbers_leaf_list(bytes: &[u8]) -> TestResult<String> {
    use litchi_numbers::sheet::drawable::{DrawableSelector, Kind as DrawableKind};
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    let count = package
        .sheet_drawable_list(0)?
        .iter()
        .filter(|drawable| drawable.kind() == DrawableKind::Chart)
        .count();
    let mut out = String::new();
    for index in 0..count {
        let geometry = package
            .sheet_drawable_properties(0, DrawableSelector::chart(index))?
            .geometry()
            .map(crate::numbers_drawable_cases::host_geometry)
            .ok_or("chart without frame")?;
        out.push_str(&format!(
            "{:?} {:?} {:?} {:?} {:?}\n",
            package.sheet_chart_metadata(0, index)?.kind(),
            package.sheet_chart_direction(0, index)?,
            package.sheet_chart_data(0, index)?,
            geometry,
            package.sheet_chart_arrangement(0, index)?
        ));
    }
    Ok(out)
}

fn numbers_case(name: &'static str, op: Op) -> TestResult {
    let input = chart_fixture()?;
    let (host_op, leaf_op) = (op.clone(), op);
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, &input)
        .host(move |source| numbers_host(source, &host_op))
        .leaf(move |source| numbers_leaf(source, &leaf_op))
        .read("host sheet_charts", numbers_host_list)
        .read("leaf chart reads", numbers_leaf_list)
        .run()?;
    Ok(())
}

#[test]
fn numbers_sheet_chart_kind_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in [Kind::Bar2d, Kind::Line2d, Kind::Pie2d] {
        numbers_case("numbers_sheet_chart_kind_parity", Op::Kind(kind))
            .map_err(|error| format!("{kind:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_direction_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    numbers_case(
        "numbers_sheet_chart_direction_parity",
        Op::Direction(Direction::Columns),
    )
}

#[test]
fn numbers_sheet_chart_data_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    numbers_case("numbers_sheet_chart_data_set_parity", Op::Data(regrid()?))
}

#[test]
fn numbers_sheet_charts_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = chart_fixture()?;
    let (host, leaf) = (
        numbers_host_list(&input.bytes)?,
        numbers_leaf_list(&input.bytes)?,
    );
    if host.is_empty() || host != leaf {
        return Err(format!("host {host:?}, leaf {leaf:?}").into());
    }
    Case::new(
        "numbers_sheet_charts_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| {
        Ok(LeafRun::output(leaf_bytes!(
            litchi_numbers::Package::from_bytes(source)?
        )))
    })
    .read("host", numbers_host_list)
    .read("leaf", numbers_leaf_list)
    .run()?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Pages

const PAGES_NATIVE: &str = "iwork/pages/chart-data-native.pages";

fn pages_fixture() -> TestResult<Fixture> {
    let mut editor = pages::Draft::from_bytes(&fixture(PAGES_NATIVE)?.bytes)?;
    editor.add_body_chart(
        0,
        Kind::Column2d,
        ChartData::new(
            vec!["Region 1".to_owned(), "Region 2".to_owned()],
            vec!["April".to_owned(), "May".to_owned()],
            vec![vec![Some(17.0), Some(26.0)], vec![Some(55.0), Some(43.0)]],
        )?,
        litchi_iwa::shapes::DrawablePoint { x: 120.0, y: 140.0 },
        litchi_iwa::shapes::DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    Ok(Fixture {
        path: "chart-data-native.pages + host-added Column2d chart",
        bytes: editor.to_bytes()?,
    })
}

fn pages_host_geometry(
    geometry: litchi_pages::drawable::Geometry,
) -> litchi_iwa::shapes::DrawableGeometry {
    litchi_iwa::shapes::DrawableGeometry {
        position: geometry
            .position
            .map(|point| litchi_iwa::shapes::DrawablePoint {
                x: point.x,
                y: point.y,
            }),
        size: geometry.size.map(|size| litchi_iwa::shapes::DrawableSize {
            width: size.width,
            height: size.height,
        }),
        flags: geometry
            .flags
            .map(litchi_pages::drawable::GeometryFlags::bits),
        angle: geometry.angle,
    }
}

fn pages_host(source: &[u8], op: &Op) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let id = editor
        .body_charts()?
        .first()
        .map(|chart| chart.drawable_object_id)
        .ok_or("host chart missing")?;
    match op {
        Op::Kind(kind) => editor.set_body_chart_kind(id, *kind)?,
        Op::Direction(direction) => editor.set_body_chart_direction(id, *direction)?,
        Op::Data(data) => editor.set_body_chart_data(id, data.clone())?,
    }
    Ok(editor.to_bytes()?)
}

fn pages_leaf(source: &[u8], op: &Op) -> TestResult<LeafRun> {
    let package = litchi_pages::Package::from_bytes(source)?;
    if let Op::Data(data) = op {
        let commit = package
            .edit_body_chart_data(0)?
            .set(data.clone())
            .commit()?;
        let restored = commit
            .package()
            .apply_body_chart_data(&commit.patch().inverse())?;
        return Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ));
    }
    let mut edit = package.edit_body_chart_lifecycle(0)?;
    match op {
        Op::Kind(kind) => drop(edit.set_kind(*kind)?),
        Op::Direction(direction) => drop(edit.set_direction(*direction)?),
        Op::Data(_) => return Err("unreachable".into()),
    }
    let commit = edit.commit()?;
    if !commit.diagnostics().changed() {
        return Err("leaf chart-lifecycle edit was a no-op".into());
    }
    let restored = commit
        .package()
        .apply_body_chart_lifecycle(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn pages_host_list(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let mut out = String::new();
    for chart in editor.body_charts()? {
        out.push_str(&format!(
            "{:?} {:?} {:?} {:?} {:?}\n",
            chart.kind, chart.direction, chart.data, chart.geometry, chart.arrangement
        ));
    }
    Ok(out)
}

fn pages_leaf_list(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_pages::Package::from_bytes(bytes)?;
    let mut out = String::new();
    for index in 0..package.body_chart_arrangements()?.len() {
        let geometry = package
            .body_chart_geometry(index)?
            .map(pages_host_geometry)
            .ok_or("chart without frame")?;
        out.push_str(&format!(
            "{:?} {:?} {:?} {:?} {:?}\n",
            package.body_chart_metadata(index)?.kind(),
            package.body_chart_direction(index)?,
            package.body_chart_data(index)?,
            geometry,
            package.body_chart_arrangement(index)?
        ));
    }
    Ok(out)
}

fn pages_case(name: &'static str, op: Op) -> TestResult {
    let input = pages_fixture()?;
    let (host_op, leaf_op) = (op.clone(), op);
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, &input)
        .host(move |source| pages_host(source, &host_op))
        .leaf(move |source| pages_leaf(source, &leaf_op))
        .read("host body_charts", pages_host_list)
        .read("leaf chart reads", pages_leaf_list)
        .run()?;
    Ok(())
}

#[test]
fn pages_body_chart_kind_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in [Kind::Bar2d, Kind::Line2d, Kind::Pie2d] {
        pages_case("pages_body_chart_kind_parity", Op::Kind(kind))
            .map_err(|error| format!("{kind:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_chart_direction_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    pages_case(
        "pages_body_chart_direction_parity",
        Op::Direction(Direction::Columns),
    )
}

#[test]
fn pages_body_chart_data_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    pages_case("pages_body_chart_data_set_parity", Op::Data(regrid()?))
}

#[test]
fn pages_body_charts_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = pages_fixture()?;
    let (host, leaf) = (
        pages_host_list(&input.bytes)?,
        pages_leaf_list(&input.bytes)?,
    );
    if host.is_empty() || host != leaf {
        return Err(format!("host {host:?}, leaf {leaf:?}").into());
    }
    Case::new(
        "pages_body_charts_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| {
        Ok(LeafRun::output(leaf_bytes!(
            litchi_pages::Package::from_bytes(source)?
        )))
    })
    .read("host", pages_host_list)
    .read("leaf", pages_leaf_list)
    .run()?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Paired host-vs-leaf cost probe (record 0416 §4): open + one operation +
// commit + serialize, ABBA; read the inclusive cost of `w4d_chart_host` and
// `w4d_chart_leaf` under callgrind (`W4D_CHART_ROUTE=<route>`).

const CHART_ROUTES: &[&str] = &[
    "numbers_kind",
    "numbers_direction",
    "numbers_data",
    "pages_kind",
    "pages_direction",
    "pages_data",
];

fn chart_op(route: &str) -> TestResult<Op> {
    Ok(match route.split_once('_').map(|(_, op)| op) {
        Some("kind") => Op::Kind(Kind::Bar2d),
        Some("direction") => Op::Direction(Direction::Columns),
        _ => Op::Data(regrid()?),
    })
}

#[inline(never)]
fn w4d_chart_host(route: &str, source: &[u8], op: &Op) -> TestResult<Vec<u8>> {
    if route.starts_with("numbers") {
        numbers_host(source, op)
    } else {
        pages_host(source, op)
    }
}

#[inline(never)]
fn w4d_chart_leaf(route: &str, source: &[u8], op: &Op) -> TestResult<Vec<u8>> {
    Ok(if route.starts_with("numbers") {
        numbers_leaf(source, op)?.output
    } else {
        pages_leaf(source, op)?.output
    })
}

#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn w4d_chart_route_cost_abba() -> TestResult {
    let selected = std::env::var("W4D_CHART_ROUTE").ok();
    let numbers_input = chart_fixture()?.bytes;
    let pages_input = pages_fixture()?.bytes;
    for route in CHART_ROUTES {
        if selected
            .as_deref()
            .is_some_and(|selected| selected != *route)
        {
            continue;
        }
        let input = if route.starts_with("numbers") {
            &numbers_input
        } else {
            &pages_input
        };
        let op = chart_op(route)?;
        let a1 = std::hint::black_box(w4d_chart_host(route, input, &op)?);
        let b1 = std::hint::black_box(w4d_chart_leaf(route, input, &op)?);
        let b2 = std::hint::black_box(w4d_chart_leaf(route, input, &op)?);
        let a2 = std::hint::black_box(w4d_chart_host(route, input, &op)?);
        assert_eq!(b1, b2, "{route}: the leaf route is deterministic");
        assert!(!a1.is_empty() && !a2.is_empty());
    }
    Ok(())
}
