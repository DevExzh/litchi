//! Pages body drawable title and caption parity cases of wave W1-B (record
//! 0371): `litchi_pages::drawable::caption` against the host's
//! `{set,remove}_body_{image,movie,shape}_{title,caption}` and
//! `body_{image,movie,shape}_title_caption`.
//!
//! Creating a label authors a native graph and clearing one a stand-in;
//! both register the new objects' UUIDs, which the host draws at random and
//! the leaf deterministically, so those cases accept the leaf structural
//! metadata policy (tier B). Every other byte is compared.

use litchi_iwa::pages;
use litchi_iwa_common::comment::DrawableId;
use litchi_pages::drawable::caption::Kind;
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::pages_drawable_cases::{IMAGES, movie_fixture, shape_fixture};

/// The host's drawable identifier of `selector`.
fn host_id(editor: &pages::Draft, selector: DrawableSelector) -> TestResult<u64> {
    Ok(match selector {
        DrawableSelector::Image(position) => {
            editor
                .body_images()?
                .get(position.get())
                .ok_or("host image missing")?
                .drawable_object_id
        },
        DrawableSelector::Movie(position) => {
            editor
                .body_movies()?
                .get(position.get())
                .ok_or("host movie missing")?
                .drawable_object_id
        },
        DrawableSelector::Shape(position) => {
            editor
                .body_shapes()?
                .get(position.get())
                .ok_or("host shape missing")?
                .drawable_object_id
        },
        _ => return Err("no host labels".into()),
    })
}

fn host_labels(
    bytes: &[u8],
    selector: DrawableSelector,
) -> TestResult<(Option<String>, Option<String>)> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, selector)?;
    let labels = match selector {
        DrawableSelector::Image(_) => editor.body_image_title_caption(DrawableId::try_from(id)?)?,
        DrawableSelector::Movie(_) => editor.body_movie_title_caption(id)?,
        _ => editor.body_shape_title_caption(id)?,
    };
    Ok((labels.title, labels.caption))
}

fn leaf_labels(
    bytes: &[u8],
    selector: DrawableSelector,
) -> TestResult<(Option<String>, Option<String>)> {
    let package = Package::from_bytes(bytes)?;
    Ok((
        package.body_drawable_caption(selector, Kind::Title)?,
        package.body_drawable_caption(selector, Kind::Caption)?,
    ))
}

/// Run the host's set (`Some`) or remove (`None`) route.
fn host_route(
    source: &[u8],
    selector: DrawableSelector,
    kind: Kind,
    text: Option<&str>,
) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let id = host_id(&editor, selector)?;
    match (selector, kind, text) {
        (DrawableSelector::Image(_), Kind::Title, Some(text)) => {
            editor.set_body_image_title(DrawableId::try_from(id)?, text)?;
        },
        (DrawableSelector::Image(_), Kind::Caption, Some(text)) => {
            editor.set_body_image_caption(DrawableId::try_from(id)?, text)?;
        },
        (DrawableSelector::Image(_), Kind::Title, None) => {
            editor.remove_body_image_title(DrawableId::try_from(id)?)?;
        },
        (DrawableSelector::Image(_), Kind::Caption, None) => {
            editor.remove_body_image_caption(DrawableId::try_from(id)?)?;
        },
        (DrawableSelector::Movie(_), Kind::Title, Some(text)) => {
            editor.set_body_movie_title(id, text)?;
        },
        (DrawableSelector::Movie(_), Kind::Caption, Some(text)) => {
            editor.set_body_movie_caption(id, text)?;
        },
        (DrawableSelector::Movie(_), Kind::Title, None) => {
            editor.remove_body_movie_title(id)?;
        },
        (DrawableSelector::Movie(_), Kind::Caption, None) => {
            editor.remove_body_movie_caption(id)?;
        },
        (_, Kind::Title, Some(text)) => editor.set_body_shape_title(id, text)?,
        (_, Kind::Caption, Some(text)) => editor.set_body_shape_caption(id, text)?,
        (_, Kind::Title, None) => drop(editor.remove_body_shape_title(id)?),
        (_, Kind::Caption, None) => drop(editor.remove_body_shape_caption(id)?),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(
    source: &[u8],
    selector: DrawableSelector,
    kind: Kind,
    text: Option<&str>,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_drawable_caption(selector, kind)?;
    match text {
        Some(text) => drop(edit.set(text)?),
        None => drop(edit.clear()?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_drawable_caption(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    target: (DrawableSelector, Kind),
    text: Option<&str>,
) -> TestResult {
    let (selector, kind) = target;
    let mut case = Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| host_route(source, selector, kind, text))
        .leaf(|source| leaf_route(source, selector, kind, text))
        .read("host title/caption", move |bytes| {
            host_labels(bytes, selector)
        })
        .read("leaf body_drawable_caption", move |bytes| {
            leaf_labels(bytes, selector)
        });
    if matches!(tier, Tier::B) {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

/// `input` after the host gave the drawable both labels.
fn labelled(input: &Fixture, selector: DrawableSelector) -> TestResult<Fixture> {
    let titled = host_route(
        &input.bytes,
        selector,
        Kind::Title,
        Some("Quarterly portrait"),
    )?;
    Ok(Fixture {
        path: "fixture + host title \"Quarterly portrait\" and caption \"Revenue report\"",
        bytes: host_route(&titled, selector, Kind::Caption, Some("Revenue report"))?,
    })
}

fn input(selector: DrawableSelector) -> TestResult<Fixture> {
    match selector {
        DrawableSelector::Image(_) => fixture(IMAGES),
        DrawableSelector::Movie(_) => movie_fixture(),
        _ => shape_fixture(),
    }
}

/// Create the label on an unlabelled drawable (a new graph), then replace
/// the text of the present label (only its storage changes; the run is
/// byte-identical, so the tier-B comparison is exact there too).
fn set_cases(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    target: (DrawableSelector, Kind),
) -> TestResult {
    let source = input(target.0)?;
    run(name, tier, evidence, &source, target, Some("Zürich — 2026"))?;
    let labelled = labelled(&source, target.0)?;
    run(
        name,
        tier,
        evidence,
        &labelled,
        target,
        Some("Updated label"),
    )
}

/// Remove a present label (fresh stand-in, tier B).
fn remove_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    target: (DrawableSelector, Kind),
) -> TestResult {
    let source = labelled(&input(target.0)?, target.0)?;
    run(name, tier, evidence, &source, target, None)
}

#[test]
fn pages_body_image_title_set_parity() -> TestResult {
    set_cases(
        "pages_body_image_title_set_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::image(0), Kind::Title),
    )
}

#[test]
fn pages_body_image_title_remove_parity() -> TestResult {
    remove_case(
        "pages_body_image_title_remove_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::image(0), Kind::Title),
    )
}

#[test]
fn pages_body_image_caption_set_parity() -> TestResult {
    set_cases(
        "pages_body_image_caption_set_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::image(0), Kind::Caption),
    )
}

#[test]
fn pages_body_image_caption_remove_parity() -> TestResult {
    remove_case(
        "pages_body_image_caption_remove_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::image(0), Kind::Caption),
    )
}

#[test]
fn pages_body_movie_title_set_parity() -> TestResult {
    set_cases(
        "pages_body_movie_title_set_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::movie(0), Kind::Title),
    )
}

#[test]
fn pages_body_movie_title_remove_parity() -> TestResult {
    remove_case(
        "pages_body_movie_title_remove_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::movie(0), Kind::Title),
    )
}

#[test]
fn pages_body_movie_caption_set_parity() -> TestResult {
    set_cases(
        "pages_body_movie_caption_set_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::movie(0), Kind::Caption),
    )
}

#[test]
fn pages_body_movie_caption_remove_parity() -> TestResult {
    remove_case(
        "pages_body_movie_caption_remove_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::movie(0), Kind::Caption),
    )
}

#[test]
fn pages_body_shape_title_set_parity() -> TestResult {
    set_cases(
        "pages_body_shape_title_set_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::shape(0), Kind::Title),
    )
}

#[test]
fn pages_body_shape_title_remove_parity() -> TestResult {
    remove_case(
        "pages_body_shape_title_remove_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::shape(0), Kind::Title),
    )
}

#[test]
fn pages_body_shape_caption_set_parity() -> TestResult {
    set_cases(
        "pages_body_shape_caption_set_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::shape(0), Kind::Caption),
    )
}

#[test]
fn pages_body_shape_caption_remove_parity() -> TestResult {
    remove_case(
        "pages_body_shape_caption_remove_parity",
        Tier::B,
        Evidence::Queued,
        (DrawableSelector::shape(0), Kind::Caption),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn image_title_create_cost_abba() -> TestResult {
    let target = (DrawableSelector::image(0), Kind::Title);
    crate::pages_cost_probes::abba(
        &fixture(IMAGES)?,
        |source| host_route(source, target.0, target.1, Some("Zürich — 2026")),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_caption(target.0, target.1)?;
            edit.set("Zürich — 2026")?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn image_caption_remove_cost_abba() -> TestResult {
    let target = (DrawableSelector::image(0), Kind::Caption);
    crate::pages_cost_probes::abba(
        &labelled(&fixture(IMAGES)?, target.0)?,
        |source| host_route(source, target.0, target.1, None),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_caption(target.0, target.1)?;
            edit.clear()?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
