//! Numbers sheet drawable title and caption parity (wave W4-L, record
//! 0425): `litchi_numbers::sheet::drawable::caption` on the shared
//! `litchi_iwa_style::caption` engine against the host's
//! `{set,remove}_sheet_{shape,image,movie}_{title,caption}`,
//! `sheet_{shape,image,movie}_title_caption` and
//! `{set,remove}_sheet_chart_caption`/`sheet_chart_caption`.
//!
//! Creating a label authors a native graph and clearing one a stand-in;
//! both register the new objects' UUIDs, which the host draws at random and
//! the leaf deterministically, so those cases accept the leaf structural
//! metadata policy (tier B). Every other byte is compared.

use litchi_iwa::numbers;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::caption::Label;
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{authored_fixture, chart_fixture, host_drawable};

const IMAGES: &str = "iwork/numbers/image-adjustments-native.numbers";

fn selector(kind: Kind) -> DrawableSelector {
    DrawableSelector::of(kind, 0)
}

fn host_labels(bytes: &[u8], kind: Kind) -> TestResult<(Option<String>, Option<String>)> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let (sheet, id) = host_drawable(&editor, kind, 0)?;
    let labels = match kind {
        Kind::Shape => editor.sheet_shape_title_caption(sheet, id)?,
        Kind::Image => editor.sheet_image_title_caption(sheet, id)?,
        Kind::Movie => editor.sheet_movie_title_caption(sheet, id)?,
        _ => return Ok((None, editor.sheet_chart_caption(sheet, id)?)),
    };
    Ok((labels.title, labels.caption))
}

fn leaf_labels(bytes: &[u8], kind: Kind) -> TestResult<(Option<String>, Option<String>)> {
    let package = Package::from_bytes(bytes)?;
    let title = if kind == Kind::Chart {
        None
    } else {
        package.sheet_drawable_caption(0, selector(kind), Label::Title)?
    };
    Ok((
        title,
        package.sheet_drawable_caption(0, selector(kind), Label::Caption)?,
    ))
}

/// Run the host's set (`Some`) or remove (`None`) route.
fn host_route(source: &[u8], kind: Kind, label: Label, text: Option<&str>) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let (sheet, id) = host_drawable(&e, kind, 0)?;
    match (kind, label, text) {
        (Kind::Shape, Label::Title, Some(text)) => e.set_sheet_shape_title(sheet, id, text)?,
        (Kind::Shape, Label::Caption, Some(text)) => e.set_sheet_shape_caption(sheet, id, text)?,
        (Kind::Shape, Label::Title, None) => drop(e.remove_sheet_shape_title(sheet, id)?),
        (Kind::Shape, Label::Caption, None) => drop(e.remove_sheet_shape_caption(sheet, id)?),
        (Kind::Image, Label::Title, Some(text)) => e.set_sheet_image_title(sheet, id, text)?,
        (Kind::Image, Label::Caption, Some(text)) => e.set_sheet_image_caption(sheet, id, text)?,
        (Kind::Image, Label::Title, None) => drop(e.remove_sheet_image_title(sheet, id)?),
        (Kind::Image, Label::Caption, None) => drop(e.remove_sheet_image_caption(sheet, id)?),
        (Kind::Movie, Label::Title, Some(text)) => e.set_sheet_movie_title(sheet, id, text)?,
        (Kind::Movie, Label::Caption, Some(text)) => e.set_sheet_movie_caption(sheet, id, text)?,
        (Kind::Movie, Label::Title, None) => drop(e.remove_sheet_movie_title(sheet, id)?),
        (Kind::Movie, Label::Caption, None) => drop(e.remove_sheet_movie_caption(sheet, id)?),
        (Kind::Chart, Label::Caption, Some(text)) => e.set_sheet_chart_caption(sheet, id, text)?,
        (Kind::Chart, Label::Caption, None) => drop(e.remove_sheet_chart_caption(sheet, id)?),
        _ => return Err("no host route".into()),
    }
    Ok(e.to_bytes()?)
}

fn leaf_route(source: &[u8], kind: Kind, label: Label, text: Option<&str>) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawable_caption(0, selector(kind), label)?;
    match text {
        Some(text) => drop(edit.set(text)?),
        None => drop(edit.clear()?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_drawable_caption(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(
    name: &'static str,
    input: &Fixture,
    kind: Kind,
    label: Label,
    text: Option<&str>,
) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(|source| host_route(source, kind, label, text))
        .leaf(|source| leaf_route(source, kind, label, text))
        .read("host title/caption", move |bytes| host_labels(bytes, kind))
        .read("leaf sheet_drawable_caption", move |bytes| {
            leaf_labels(bytes, kind)
        })
        .run()?;
    Ok(())
}

fn input(kind: Kind) -> TestResult<Fixture> {
    match kind {
        Kind::Image => fixture(IMAGES),
        Kind::Chart => chart_fixture(),
        _ => authored_fixture(),
    }
}

/// `input` after the host gave the drawable its labels.
fn labelled(kind: Kind) -> TestResult<Fixture> {
    let mut bytes = input(kind)?.bytes;
    if kind != Kind::Chart {
        bytes = host_route(&bytes, kind, Label::Title, Some("Quarterly portrait"))?;
    }
    Ok(Fixture {
        path: "fixture + host title \"Quarterly portrait\" and caption \"Revenue report\"",
        bytes: host_route(&bytes, kind, Label::Caption, Some("Revenue report"))?,
    })
}

/// Create the label on an unlabelled drawable (a new graph), then replace
/// the text of the present label (only its storage changes).
fn set_cases(name: &'static str, kind: Kind, label: Label) -> TestResult {
    run(name, &input(kind)?, kind, label, Some("Zürich — 2026"))?;
    run(name, &labelled(kind)?, kind, label, Some("Updated label"))
}

/// Remove a present label (fresh stand-in).
fn remove_case(name: &'static str, kind: Kind, label: Label) -> TestResult {
    run(name, &labelled(kind)?, kind, label, None)
}

#[test]
fn numbers_sheet_shape_title_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    set_cases(
        "numbers_sheet_shape_title_set_parity",
        Kind::Shape,
        Label::Title,
    )
}

#[test]
fn numbers_sheet_shape_title_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "numbers_sheet_shape_title_remove_parity",
        Kind::Shape,
        Label::Title,
    )
}

#[test]
fn numbers_sheet_shape_caption_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    set_cases(
        "numbers_sheet_shape_caption_set_parity",
        Kind::Shape,
        Label::Caption,
    )
}

#[test]
fn numbers_sheet_shape_caption_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "numbers_sheet_shape_caption_remove_parity",
        Kind::Shape,
        Label::Caption,
    )
}

#[test]
fn numbers_sheet_image_title_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    set_cases(
        "numbers_sheet_image_title_set_parity",
        Kind::Image,
        Label::Title,
    )
}

#[test]
fn numbers_sheet_image_title_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "numbers_sheet_image_title_remove_parity",
        Kind::Image,
        Label::Title,
    )
}

#[test]
fn numbers_sheet_image_caption_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    set_cases(
        "numbers_sheet_image_caption_set_parity",
        Kind::Image,
        Label::Caption,
    )
}

#[test]
fn numbers_sheet_image_caption_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "numbers_sheet_image_caption_remove_parity",
        Kind::Image,
        Label::Caption,
    )
}

#[test]
fn numbers_sheet_movie_title_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    set_cases(
        "numbers_sheet_movie_title_set_parity",
        Kind::Movie,
        Label::Title,
    )
}

#[test]
fn numbers_sheet_movie_title_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "numbers_sheet_movie_title_remove_parity",
        Kind::Movie,
        Label::Title,
    )
}

#[test]
fn numbers_sheet_movie_caption_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    set_cases(
        "numbers_sheet_movie_caption_set_parity",
        Kind::Movie,
        Label::Caption,
    )
}

#[test]
fn numbers_sheet_movie_caption_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "numbers_sheet_movie_caption_remove_parity",
        Kind::Movie,
        Label::Caption,
    )
}

/// Replacing a present chart caption's text (the host-added caption).
/// Authoring a new caption on a chart stored beside the calculation engine
/// needs the host's componentized layout, which the leaf refuses
/// (`numbers_sheet_chart_caption_componentized_refuses`), so
/// `set_sheet_chart_caption` stays host-only.
#[test]
fn numbers_sheet_chart_caption_replace_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    run(
        "numbers_sheet_chart_caption_replace_parity",
        &labelled(Kind::Chart)?,
        Kind::Chart,
        Label::Caption,
        Some("Updated label"),
    )
}

/// The leaf refuses a new caption in the componentized layout (typed),
/// leaving the source unchanged, where the host authors one.
#[test]
fn numbers_sheet_chart_caption_componentized_refuses() -> TestResult {
    use litchi_numbers::sheet::drawable::caption::{Error, Refusal};
    let source = input(Kind::Chart)?;
    host_route(&source.bytes, Kind::Chart, Label::Caption, Some("New"))?;
    let package = Package::from_bytes(&source.bytes)?;
    let mut edit = package.edit_sheet_drawable_caption(0, selector(Kind::Chart), Label::Caption)?;
    edit.set("New")?;
    match edit.commit() {
        Err(Error::Refused(Refusal::ComponentizedLayout)) => Ok(()),
        other => Err(format!("expected the componentized refusal, got {other:?}").into()),
    }
}

#[test]
fn numbers_sheet_chart_caption_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "numbers_sheet_chart_caption_remove_parity",
        Kind::Chart,
        Label::Caption,
    )
}

/// The reads: every labelled and unlabelled drawable reads the host's
/// labels.
#[test]
fn numbers_sheet_drawable_caption_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for kind in [Kind::Shape, Kind::Image, Kind::Movie, Kind::Chart] {
        for input in [input(kind)?, labelled(kind)?] {
            if host_labels(&input.bytes, kind)? != leaf_labels(&input.bytes, kind)? {
                return Err(format!("{kind:?} labels differ on {}", input.path).into());
            }
        }
    }
    Ok(())
}

/// One measured caption route: (name, drawable kind, label, text, start
/// labelled).
const COST_ROUTES: &[(&str, Kind, Label, Option<&str>, bool)] = &[
    (
        "shape_title_create",
        Kind::Shape,
        Label::Title,
        Some("Zürich — 2026"),
        false,
    ),
    (
        "image_caption_create",
        Kind::Image,
        Label::Caption,
        Some("Zürich — 2026"),
        false,
    ),
    (
        "movie_caption_replace",
        Kind::Movie,
        Label::Caption,
        Some("Updated label"),
        true,
    ),
    ("image_title_remove", Kind::Image, Label::Title, None, true),
    (
        "chart_caption_remove",
        Kind::Chart,
        Label::Caption,
        None,
        true,
    ),
];

#[inline(never)]
fn w4l_host_route(
    source: &[u8],
    kind: Kind,
    label: Label,
    text: Option<&str>,
) -> TestResult<Vec<u8>> {
    host_route(source, kind, label, text)
}

#[inline(never)]
fn w4l_leaf_route(
    source: &[u8],
    kind: Kind,
    label: Label,
    text: Option<&str>,
) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawable_caption(0, selector(kind), label)?;
    match text {
        Some(text) => drop(edit.set(text)?),
        None => drop(edit.clear()?),
    }
    Ok(leaf_bytes!(edit.commit()?.package()))
}

/// Paired host-versus-leaf cost probe (record 0425): open + one label
/// operation + commit + serialize, ABBA. Run under callgrind with
/// `W4L_ROUTE=<name>` and read the inclusive cost of `w4l_host_route` and
/// `w4l_leaf_route`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_caption_route_cost_abba() -> TestResult {
    let selected = std::env::var("W4L_ROUTE").ok();
    for &(name, kind, label, text, labelled_start) in COST_ROUTES {
        if selected.as_deref().is_some_and(|route| route != name) {
            continue;
        }
        let source = if labelled_start {
            labelled(kind)?
        } else {
            input(kind)?
        }
        .bytes;
        std::hint::black_box(w4l_host_route(&source, kind, label, text)?);
        std::hint::black_box(w4l_leaf_route(&source, kind, label, text)?);
        std::hint::black_box(w4l_leaf_route(&source, kind, label, text)?);
        std::hint::black_box(w4l_host_route(&source, kind, label, text)?);
    }
    Ok(())
}
