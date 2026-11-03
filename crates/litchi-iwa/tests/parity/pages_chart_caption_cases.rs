//! Pages body chart caption parity (wave W4-N, record 0427):
//! `litchi_pages::chart::caption` on the shared `litchi_iwa_style::caption`
//! engine against the host's `body_chart_caption`, `set_body_chart_caption`
//! and `remove_body_chart_caption`.
//!
//! Creating a caption authors the native inline graph and clearing one a
//! stand-in; both register the new objects' UUIDs, which the host draws at
//! random and the leaf deterministically, so those cases accept the leaf
//! structural metadata policy (tier B). Every other byte is compared,
//! including the chart message's reference lists, where the host moves the
//! caption reference to the end.

use litchi_iwa::pages;
use litchi_pages::{BodyChartSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// The native Apple-saved chart document (chart 0 has no caption).
const CHARTS: &str = "iwork/pages/chart-data-native.pages";
const CHART: usize = 0;

fn host_chart(editor: &pages::Draft) -> TestResult<u64> {
    Ok(editor
        .body_charts()?
        .get(CHART)
        .ok_or("host chart missing")?
        .drawable_object_id)
}

fn host_caption(bytes: &[u8]) -> TestResult<Option<String>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let chart = host_chart(&editor)?;
    Ok(editor.body_chart_caption(chart)?)
}

fn leaf_caption(bytes: &[u8]) -> TestResult<Option<String>> {
    Ok(Package::from_bytes(bytes)?.body_chart_caption(BodyChartSelector::index(CHART))?)
}

/// Run the host's set (`Some`) or remove (`None`) route.
fn host_route(source: &[u8], text: Option<&str>) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let chart = host_chart(&editor)?;
    match text {
        Some(text) => editor.set_body_chart_caption(chart, text)?,
        None => drop(editor.remove_body_chart_caption(chart)?),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_commit(source: &[u8], text: Option<&str>) -> TestResult<(Package, Package)> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_chart_caption(BodyChartSelector::index(CHART))?;
    match text {
        Some(text) => drop(edit.set(text)?),
        None => drop(edit.clear()?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_chart_caption(&commit.patch().inverse())?
        .into_package();
    Ok((commit.into_package(), restored))
}

fn leaf_route(source: &[u8], text: Option<&str>) -> TestResult<LeafRun> {
    let (output, restored) = leaf_commit(source, text)?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(&output),
        leaf_bytes!(&restored),
    ))
}

fn run(name: &'static str, input: &Fixture, text: Option<&str>) -> TestResult {
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(|source| host_route(source, text))
        .leaf(|source| leaf_route(source, text))
        .read("host body_chart_caption", host_caption)
        .read("leaf body_chart_caption", leaf_caption)
        .run()?;
    Ok(())
}

/// The native chart after the host gave it a caption.
fn labelled() -> TestResult<Fixture> {
    Ok(Fixture {
        path: "chart-data-native.pages + host caption \"Revenue report\"",
        bytes: host_route(&fixture(CHARTS)?.bytes, Some("Revenue report"))?,
    })
}

/// Create the caption on the uncaptioned native chart (a new inline graph),
/// then replace the text of the present caption (only its storage
/// changes).
#[test]
fn pages_body_chart_caption_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    run(
        "pages_body_chart_caption_set_parity",
        &fixture(CHARTS)?,
        Some("Zürich — 2026"),
    )?;
    run(
        "pages_body_chart_caption_set_parity",
        &labelled()?,
        Some("Updated label"),
    )
}

/// Remove a present caption (fresh stand-in, the old graph kept).
#[test]
fn pages_body_chart_caption_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    run("pages_body_chart_caption_remove_parity", &labelled()?, None)
}

/// The reads: the uncaptioned native chart and the host-captioned chart
/// read the host's caption; removing an absent caption is an exact no-op.
#[test]
fn pages_body_chart_caption_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for input in [fixture(CHARTS)?, labelled()?] {
        if host_caption(&input.bytes)? != leaf_caption(&input.bytes)? {
            return Err(format!("chart captions differ on {}", input.path).into());
        }
    }
    let source = fixture(CHARTS)?;
    let package = Package::from_bytes(&source.bytes)?;
    let mut edit = package.edit_body_chart_caption(BodyChartSelector::index(CHART))?;
    edit.clear()?;
    let commit = edit.commit()?;
    if commit.diagnostics().changed() || !commit.patch().is_noop() {
        return Err("clearing an absent chart caption changed the package".into());
    }
    Ok(())
}

/// One measured route: (name, text, start captioned).
const COST_ROUTES: &[(&str, Option<&str>, bool)] = &[
    ("pages_chart_caption_create", Some("Zürich — 2026"), false),
    ("pages_chart_caption_replace", Some("Updated label"), true),
    ("pages_chart_caption_remove", None, true),
];

#[inline(never)]
fn w4n_host_route(source: &[u8], text: Option<&str>) -> TestResult<Vec<u8>> {
    host_route(source, text)
}

#[inline(never)]
fn w4n_leaf_route(source: &[u8], text: Option<&str>) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_chart_caption(BodyChartSelector::index(CHART))?;
    match text {
        Some(text) => drop(edit.set(text)?),
        None => drop(edit.clear()?),
    }
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Paired host-versus-leaf cost probe (record 0427): open + one caption
/// operation + commit + serialize, ABBA. Run under callgrind with
/// `W4N_ROUTE=<name>` and read the inclusive cost of `w4n_host_route` and
/// `w4n_leaf_route`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_chart_caption_route_cost_abba() -> TestResult {
    let selected = std::env::var("W4N_ROUTE").ok();
    for &(name, text, captioned) in COST_ROUTES {
        if selected.as_deref().is_some_and(|route| route != name) {
            continue;
        }
        let source = if captioned {
            labelled()?
        } else {
            fixture(CHARTS)?
        }
        .bytes;
        std::hint::black_box(w4n_host_route(&source, text)?);
        std::hint::black_box(w4n_leaf_route(&source, text)?);
        std::hint::black_box(w4n_leaf_route(&source, text)?);
        std::hint::black_box(w4n_host_route(&source, text)?);
    }
    Ok(())
}
