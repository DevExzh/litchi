//! Pages chart duplication and removal parity (wave W4-M, record 0426):
//! `body_media::Edit::{duplicate_chart, remove_chart}` against the host's
//! `duplicate_body_chart` and `remove_body_chart`.
//!
//! The copy takes the host's identifiers in the host's graph order. The
//! host re-encodes the chart drawable, stylesheet, theme and metadata with
//! Prost and draws random UUIDs, so the cases are Tier B; neither host
//! route has native evidence, so both are macOS-queued.

use litchi_iwa::pages;

use crate::chart_lifecycle_cases::{pages_fixture, pages_host_list, pages_leaf_list};
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

#[derive(Clone, Copy)]
pub(crate) enum Op {
    /// Copy chart `.0` to body position `.1`.
    Duplicate(usize, usize),
    /// Remove chart `.0`.
    Remove(usize),
}

fn host_chart(editor: &pages::Draft, position: usize) -> TestResult<u64> {
    Ok(editor
        .body_charts()?
        .get(position)
        .ok_or("host chart missing")?
        .drawable_object_id)
}

pub(crate) fn host(source: &[u8], op: Op) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match op {
        Op::Duplicate(chart, anchor) => {
            let id = host_chart(&editor, chart)?;
            editor.duplicate_body_chart(id, anchor)?;
        },
        Op::Remove(chart) => {
            let id = host_chart(&editor, chart)?;
            editor.remove_body_chart(id)?;
        },
    }
    Ok(editor.to_bytes()?)
}

pub(crate) fn leaf(source: &[u8], op: Op) -> TestResult<LeafRun> {
    let package = litchi_pages::Package::from_bytes(source)?;
    let mut edit = package.edit_body_media();
    match op {
        Op::Duplicate(chart, anchor) => drop(
            edit.duplicate_chart(chart, litchi_pages::TextPosition::from_utf16_index(anchor)?)?,
        ),
        Op::Remove(chart) => drop(edit.remove_chart(chart)?),
    }
    let commit = edit.commit()?;
    if !commit.diagnostics().changed() {
        return Err("leaf chart lifecycle was a no-op".into());
    }
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn case(
    (name, tier, evidence): (&'static str, Tier, Evidence),
    input: &Fixture,
    op: Op,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(move |source| host(source, op))
        .leaf(move |source| leaf(source, op))
        .leaf_structural_metadata()
        .read("host body_charts", pages_host_list)
        .read("leaf chart reads", pages_leaf_list)
        .run()?;
    Ok(())
}

/// The native chart document with a second, host-authored chart (its own
/// preset and sixteen registered styles).
fn two_charts() -> TestResult<Fixture> {
    pages_fixture()
}

/// The fixture with a commented chart: the host's own drawable comment on
/// the host-authored chart.
fn commented() -> TestResult<Fixture> {
    let source = two_charts()?.bytes;
    let id = host_chart(&pages::Draft::from_bytes(&source)?, 0)?;
    let mut comments =
        litchi_iwa::comments::Draft::from_package(litchi_iwa::Package::from_bytes(&source)?)?;
    comments.set_comment(
        litchi_iwa_common::comment::DrawableId::from_raw(id)?,
        "Check the April figures",
    )?;
    Ok(Fixture {
        path: "chart-data-native.pages + host chart + host comment",
        bytes: comments.into_package().to_bytes()?,
    })
}

#[test]
fn pages_body_chart_duplicate_parity() -> TestResult {
    let input = two_charts()?;
    for (chart, anchor) in [(0, 0), (0, 5), (1, 0)] {
        case(
            (
                "pages_body_chart_duplicate_parity",
                Tier::B,
                Evidence::Queued,
            ),
            &input,
            Op::Duplicate(chart, anchor),
        )
        .map_err(|error| format!("chart {chart} at {anchor}: {error}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_chart_duplicate_native_parity() -> TestResult {
    case(
        (
            "pages_body_chart_duplicate_native_parity",
            Tier::B,
            Evidence::Queued,
        ),
        &fixture("iwork/pages/chart-data-native.pages")?,
        Op::Duplicate(0, 0),
    )
}

#[test]
fn pages_body_chart_remove_parity() -> TestResult {
    let input = two_charts()?;
    for chart in [0, 1] {
        case(
            ("pages_body_chart_remove_parity", Tier::B, Evidence::Queued),
            &input,
            Op::Remove(chart),
        )
        .map_err(|error| format!("chart {chart}: {error}"))?;
    }
    case(
        ("pages_body_chart_remove_parity", Tier::B, Evidence::Queued),
        &fixture("iwork/pages/chart-data-native.pages")?,
        Op::Remove(0),
    )
}

#[test]
fn pages_body_chart_remove_commented_parity() -> TestResult {
    case(
        (
            "pages_body_chart_remove_commented_parity",
            Tier::B,
            Evidence::Queued,
        ),
        &commented()?,
        Op::Remove(0),
    )
}

/// A duplicate of a duplicate, then removing both copies, restores the
/// chart reads of the source (the leaf alone; the copies' graphs are the
/// source graph's).
#[test]
fn pages_body_chart_duplicate_then_remove_round_trips_reads() -> TestResult {
    let input = two_charts()?;
    let once = leaf(&input.bytes, Op::Duplicate(0, 0))?.output;
    let twice = leaf(&once, Op::Duplicate(0, 0))?.output;
    let removed = leaf(&leaf(&twice, Op::Remove(0))?.output, Op::Remove(0))?.output;
    if pages_leaf_list(&removed)? != pages_leaf_list(&input.bytes)? {
        return Err("removing the copies did not restore the chart reads".into());
    }
    litchi_pages::Package::from_bytes(&removed)?;
    Ok(())
}

#[test]
fn pages_body_chart_lifecycle_refusals() -> TestResult {
    let package = litchi_pages::Package::from_bytes(&two_charts()?.bytes)?;
    let mut edit = package.edit_body_media();
    edit.remove_chart(7)?;
    assert!(matches!(
        edit.commit(),
        Err(litchi_pages::drawable::lifecycle::Error::Refused(
            litchi_pages::drawable::lifecycle::Refusal::NotFound { kind: "chart", .. }
        ))
    ));
    let mut edit = package.edit_body_media();
    edit.remove_chart(0)?;
    assert!(matches!(
        edit.duplicate_chart(0, litchi_pages::TextPosition::ZERO),
        Err(litchi_pages::drawable::lifecycle::Error::Refused(
            litchi_pages::drawable::lifecycle::Refusal::OperationAlreadyStaged
        ))
    ));
    Ok(())
}

// ---------------------------------------------------------------------------
// Paired host-vs-leaf cost probe: open + operation + commit + serialize,
// ABBA; read the inclusive cost of `w4m_chart_host` and `w4m_chart_leaf`
// under callgrind (`W4M_CHART_ROUTE=duplicate|remove`).

#[inline(never)]
fn w4m_chart_host(source: &[u8], op: Op) -> TestResult<Vec<u8>> {
    host(source, op)
}

#[inline(never)]
fn w4m_chart_leaf(source: &[u8], op: Op) -> TestResult<Vec<u8>> {
    Ok(leaf(source, op)?.output)
}

#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn w4m_chart_lifecycle_cost_abba() -> TestResult {
    let selected = std::env::var("W4M_CHART_ROUTE").ok();
    let input = two_charts()?.bytes;
    for (route, op) in [
        ("duplicate", Op::Duplicate(0, 0)),
        ("remove", Op::Remove(0)),
    ] {
        if selected
            .as_deref()
            .is_some_and(|selected| selected != route)
        {
            continue;
        }
        let a1 = std::hint::black_box(w4m_chart_host(&input, op)?);
        let b1 = std::hint::black_box(w4m_chart_leaf(&input, op)?);
        let b2 = std::hint::black_box(w4m_chart_leaf(&input, op)?);
        let a2 = std::hint::black_box(w4m_chart_host(&input, op)?);
        assert_eq!(b1, b2, "{route}: the leaf route is deterministic");
        assert!(!a1.is_empty() && !a2.is_empty());
    }
    Ok(())
}
