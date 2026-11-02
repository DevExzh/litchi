//! Numbers chart creation, duplication and removal parity (wave W4-E,
//! record 0418): `sheet::drawable::Edit::{add_chart, duplicate, remove}`
//! against the host's `add_sheet_chart`, `duplicate_sheet_chart` and
//! `remove_sheet_chart`.
//!
//! The new graph comes from the shared `chart_lifecycle::author` and takes
//! the host's identifiers in the host's order. The host re-encodes the
//! stylesheet and theme with Prost and draws its own UUIDs, so the cases
//! are Tier B; none of the host routes has native evidence, so each is
//! macOS-queued.

use litchi_iwa::numbers;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_numbers::sheet::drawable::{DrawableSelector, Geometry, Kind as DrawableKind};

use crate::chart_lifecycle_cases::{numbers_host_list, numbers_leaf_list};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{chart_fixture, first_sheet, host_drawable};

const POSITION: DrawablePoint = DrawablePoint { x: 60.0, y: 300.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 420.0,
    height: 260.0,
};

fn data() -> TestResult<ChartData> {
    Ok(ChartData::new(
        vec!["East".to_owned(), "West".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(1.0), Some(2.5), None],
            vec![Some(4.0), Some(-1.0), Some(7.0)],
        ],
    )?)
}

#[derive(Clone)]
pub(crate) enum Op {
    Add(Kind),
    Duplicate,
    Remove,
}

pub(crate) fn numbers_host(source: &[u8], op: &Op) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match op {
        Op::Add(kind) => {
            let sheet = first_sheet(&editor)?;
            editor.add_sheet_chart(sheet, *kind, data()?, POSITION, SIZE)?;
        },
        Op::Duplicate => {
            let (sheet, id) = host_drawable(&editor, DrawableKind::Chart, 0)?;
            editor.duplicate_sheet_chart(sheet, id)?;
        },
        Op::Remove => {
            let (sheet, id) = host_drawable(&editor, DrawableKind::Chart, 0)?;
            editor.remove_sheet_chart(sheet, id)?;
        },
    }
    Ok(editor.to_bytes()?)
}

pub(crate) fn numbers_leaf(source: &[u8], op: &Op) -> TestResult<LeafRun> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawables(0)?;
    match op {
        Op::Add(kind) => drop(edit.add_chart(
            *kind,
            data()?,
            Geometry::new(
                litchi_numbers::sheet::drawable::Point {
                    x: POSITION.x,
                    y: POSITION.y,
                },
                litchi_numbers::sheet::drawable::Size {
                    width: SIZE.width,
                    height: SIZE.height,
                },
            ),
        )?),
        Op::Duplicate => drop(edit.duplicate(DrawableSelector::chart(0))?),
        Op::Remove => drop(edit.remove(DrawableSelector::chart(0))?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn numbers_case(
    (name, tier, evidence): (&'static str, Tier, Evidence),
    input: &crate::harness::Fixture,
    op: Op,
) -> TestResult {
    let (host_op, leaf_op) = (op.clone(), op);
    Case::new(name, Format::Numbers, tier, evidence, input)
        .host(move |source| numbers_host(source, &host_op))
        .leaf(move |source| numbers_leaf(source, &leaf_op))
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host sheet_charts", numbers_host_list)
        .read("leaf chart reads", numbers_leaf_list)
        .run()?;
    Ok(())
}

#[test]
fn numbers_sheet_chart_add_parity() -> TestResult {
    numbers_case(
        ("numbers_sheet_chart_add_parity", Tier::B, Evidence::Queued),
        &fixture("iwork/numbers/basic.numbers")?,
        Op::Add(Kind::Line2d),
    )
}

#[test]
fn numbers_sheet_chart_add_beside_a_chart_parity() -> TestResult {
    numbers_case(
        (
            "numbers_sheet_chart_add_beside_a_chart_parity",
            Tier::B,
            Evidence::Queued,
        ),
        &chart_fixture()?,
        Op::Add(Kind::Scatter2d),
    )
}

#[test]
fn numbers_sheet_chart_duplicate_parity() -> TestResult {
    numbers_case(
        (
            "numbers_sheet_chart_duplicate_parity",
            Tier::B,
            Evidence::Queued,
        ),
        &chart_fixture()?,
        Op::Duplicate,
    )
}

#[test]
fn numbers_sheet_chart_remove_parity() -> TestResult {
    numbers_case(
        (
            "numbers_sheet_chart_remove_parity",
            Tier::B,
            Evidence::Queued,
        ),
        &chart_fixture()?,
        Op::Remove,
    )
}

#[test]
fn numbers_sheet_chart_add_refuses_unformable_data() -> TestResult {
    let package = litchi_numbers::Package::from_bytes(&chart_fixture()?.bytes)?;
    let mut edit = package.edit_sheet_drawables(0)?;
    let refused = edit.add_chart(
        Kind::Bubble2d,
        data()?,
        Geometry::new(
            litchi_numbers::sheet::drawable::Point { x: 0.0, y: 0.0 },
            litchi_numbers::sheet::drawable::Size {
                width: 10.0,
                height: 10.0,
            },
        ),
    );
    assert!(matches!(
        refused,
        Err(litchi_numbers::sheet::drawable::Error::Refused(
            litchi_numbers::sheet::drawable::Refusal::InvalidChart
        ))
    ));
    Ok(())
}

const ROUTES: [&str; 4] = [
    "numbers_add",
    "numbers_duplicate",
    "numbers_remove",
    "keynote_add",
];

fn route_op(route: &str) -> Op {
    match route {
        "numbers_add" => Op::Add(Kind::Line2d),
        "numbers_duplicate" => Op::Duplicate,
        _ => Op::Remove,
    }
}

fn keynote_frame() -> litchi_keynote::slide::drawable::Geometry {
    litchi_keynote::slide::drawable::Geometry::new(POSITION.x, POSITION.y, SIZE.width, SIZE.height)
}

#[inline(never)]
fn w4e_chart_host(route: &str, source: &[u8], op: &Op) -> TestResult<Vec<u8>> {
    if route == "keynote_add" {
        let mut editor = litchi_iwa::keynote::Draft::from_bytes(source)?;
        editor.add_slide_chart(0, Kind::Line2d, data()?, POSITION, SIZE)?;
        return Ok(editor.to_bytes()?);
    }
    numbers_host(source, op)
}

#[inline(never)]
fn w4e_chart_leaf(route: &str, source: &[u8], op: &Op) -> TestResult<Vec<u8>> {
    if route == "keynote_add" {
        let package = litchi_keynote::Package::from_bytes(source)?;
        let mut edit = package.edit_slide_chart_lifecycle();
        edit.add(
            litchi_keynote::SlideSelector::index(0),
            Kind::Line2d,
            data()?,
            keynote_frame(),
        )?;
        return Ok(leaf_bytes!(edit.commit()?.package()));
    }
    Ok(numbers_leaf(source, op)?.output)
}

/// Callgrind probe: two host and two leaf runs of each route (open, edit,
/// commit, serialize) in ABBA order; `W4E_CHART_ROUTE` selects one.
#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn w4e_chart_create_cost_abba() -> TestResult {
    let selected = std::env::var("W4E_CHART_ROUTE").ok();
    let numbers_input = chart_fixture()?.bytes;
    let keynote_input = fixture("iwork/keynote/basic.key")?.bytes;
    for route in ROUTES {
        if selected
            .as_deref()
            .is_some_and(|selected| selected != route)
        {
            continue;
        }
        let input = if route == "keynote_add" {
            &keynote_input
        } else {
            &numbers_input
        };
        let op = route_op(route);
        let a1 = std::hint::black_box(w4e_chart_host(route, input, &op)?);
        let b1 = std::hint::black_box(w4e_chart_leaf(route, input, &op)?);
        let b2 = std::hint::black_box(w4e_chart_leaf(route, input, &op)?);
        let a2 = std::hint::black_box(w4e_chart_host(route, input, &op)?);
        assert_eq!(b1, b2, "{route}: the leaf route is deterministic");
        assert!(!a1.is_empty() && !a2.is_empty());
    }
    Ok(())
}
