//! Pages body-media lifecycle parity cases of wave W1-B (record 0371): the
//! host's `remove_body_{image,movie,audio}` against the leaf
//! `drawable::lifecycle` family.

use litchi_iwa::pages;
use litchi_iwa_common::comment::DrawableId;
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::pages_drawable_cases::{IMAGES, audio_fixture, movie_fixture};

const PNG: &str = "iwork/keynote/media-replacement-assets/poster.png";

/// A source-built document with one body image (`add_body_image`).
pub(crate) fn image_fixture() -> TestResult<Fixture> {
    let data = fixture(PNG)?.bytes;
    let mut editor = pages::Draft::create_with_text("Quarterly report")?;
    let anchor = "Quarterly report".encode_utf16().count();
    let options = litchi_pages::image::Options::new(
        litchi_pages::drawable::Point { x: 72.0, y: 96.0 },
        litchi_pages::drawable::Size {
            width: 160.0,
            height: 120.0,
        },
    )?;
    editor.add_body_image(anchor, "poster.png", &data, options)?;
    Ok(Fixture {
        path: "source-built: Draft::create_with_text + add_body_image (poster.png)",
        bytes: editor.to_bytes()?,
    })
}

fn host_media(bytes: &[u8]) -> TestResult<(usize, usize, usize, usize)> {
    let editor = pages::Draft::from_bytes(bytes)?;
    Ok((
        editor.body_images()?.len(),
        editor.body_movies()?.len(),
        editor.body_audio()?.len(),
        editor.media_assets()?.len(),
    ))
}

fn leaf_text(bytes: &[u8]) -> TestResult<String> {
    Ok(Package::from_bytes(bytes)?.text()?)
}

fn host_remove(source: &[u8], selector: DrawableSelector) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match selector {
        DrawableSelector::Image(position) => {
            let id = editor.body_images()?[position.get()].drawable_object_id;
            drop(editor.remove_body_image(DrawableId::try_from(id)?)?);
        },
        DrawableSelector::Movie(position) => {
            let id = editor.body_movies()?[position.get()].drawable_object_id;
            drop(editor.remove_body_movie(id)?);
        },
        DrawableSelector::Audio(position) => {
            let id = editor.body_audio()?[position.get()].drawable_object_id;
            drop(editor.remove_body_audio(id)?);
        },
        _ => return Err("not media".into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_remove(source: &[u8], selector: DrawableSelector) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_media();
    edit.remove(selector)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn remove_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    selector: DrawableSelector,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| host_remove(source, selector))
        .leaf(|source| leaf_remove(source, selector))
        .read("host media", host_media)
        .read("leaf plain text", leaf_text)
        .run()?;
    Ok(())
}

/// Tier B: deleting the native image's four data members leaves the
/// leaf's ZIP layout (local headers, member offsets) different from the
/// host's rewrite; every member's content is equal.
#[test]
fn pages_body_image_remove_parity() -> TestResult {
    remove_case(
        "pages_body_image_remove_parity",
        Tier::B,
        Evidence::Queued,
        &image_fixture()?,
        DrawableSelector::image(0),
    )?;
    remove_case(
        "pages_body_image_remove_parity",
        Tier::B,
        Evidence::Queued,
        &fixture(IMAGES)?,
        DrawableSelector::image(0),
    )
}

#[test]
fn pages_body_movie_remove_parity() -> TestResult {
    remove_case(
        "pages_body_movie_remove_parity",
        Tier::A,
        Evidence::Queued,
        &movie_fixture()?,
        DrawableSelector::movie(0),
    )
}

#[test]
fn pages_body_audio_remove_parity() -> TestResult {
    remove_case(
        "pages_body_audio_remove_parity",
        Tier::A,
        Evidence::Queued,
        &audio_fixture()?,
        DrawableSelector::audio(0),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn image_remove_cost_abba() -> TestResult {
    let selector = DrawableSelector::image(0);
    crate::pages_cost_probes::abba(
        &fixture(IMAGES)?,
        |source| host_remove(source, selector),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_media();
            edit.remove(selector)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn movie_remove_cost_abba() -> TestResult {
    let selector = DrawableSelector::movie(0);
    crate::pages_cost_probes::abba(
        &movie_fixture()?,
        |source| host_remove(source, selector),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_media();
            edit.remove(selector)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

fn host_duplicate(source: &[u8], selector: DrawableSelector, anchor: usize) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match selector {
        DrawableSelector::Image(position) => {
            let id = editor.body_images()?[position.get()].drawable_object_id;
            drop(editor.duplicate_body_image(DrawableId::try_from(id)?, anchor)?);
        },
        DrawableSelector::Movie(position) => {
            let id = editor.body_movies()?[position.get()].drawable_object_id;
            drop(editor.duplicate_body_movie(id, anchor)?);
        },
        DrawableSelector::Audio(position) => {
            let id = editor.body_audio()?[position.get()].drawable_object_id;
            drop(editor.duplicate_body_audio(id, anchor)?);
        },
        _ => return Err("not media".into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_duplicate(source: &[u8], selector: DrawableSelector, anchor: usize) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_media();
    edit.duplicate(
        selector,
        litchi_pages::TextPosition::from_utf16_index(anchor)?,
    )?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn duplicate_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    selector: DrawableSelector,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| host_duplicate(source, selector, 0))
        .leaf(|source| leaf_duplicate(source, selector, 0))
        .read("host media", host_media)
        .read("leaf plain text", leaf_text)
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

#[test]
fn pages_body_image_duplicate_parity() -> TestResult {
    duplicate_case(
        "pages_body_image_duplicate_parity",
        Tier::B,
        Evidence::Queued,
        &image_fixture()?,
        DrawableSelector::image(0),
    )?;
    duplicate_case(
        "pages_body_image_duplicate_parity",
        Tier::B,
        Evidence::Queued,
        &fixture(IMAGES)?,
        DrawableSelector::image(0),
    )
}

#[test]
fn pages_body_movie_duplicate_parity() -> TestResult {
    duplicate_case(
        "pages_body_movie_duplicate_parity",
        Tier::B,
        Evidence::Queued,
        &movie_fixture()?,
        DrawableSelector::movie(0),
    )
}

#[test]
fn pages_body_audio_duplicate_parity() -> TestResult {
    duplicate_case(
        "pages_body_audio_duplicate_parity",
        Tier::B,
        Evidence::Queued,
        &audio_fixture()?,
        DrawableSelector::audio(0),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn image_duplicate_cost_abba() -> TestResult {
    let selector = DrawableSelector::image(0);
    crate::pages_cost_probes::abba(
        &fixture(IMAGES)?,
        |source| host_duplicate(source, selector, 0),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_media();
            edit.duplicate(selector, litchi_pages::TextPosition::from_utf16_index(0)?)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
