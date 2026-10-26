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

const REPLACEMENT_PNG: &str = "officecli/examples/ppt/ole/thumb-xlsx.png";
const REPLACEMENT_MOVIE: &[u8] = b"\0\0\0\x18ftypqt  replacement-pages-movie";
const REPLACEMENT_AUDIO: &[u8] = b"FORM\0\0\0\x10AIFCreplacement-pages-audio";

#[derive(Clone, Copy)]
enum Replace {
    Data,
    Poster,
}

fn host_replace(
    source: &[u8],
    selector: DrawableSelector,
    asset: Replace,
    data: &[u8],
) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match (selector, asset) {
        (DrawableSelector::Image(position), _) => {
            let id = editor.body_images()?[position.get()].drawable_object_id;
            drop(editor.replace_body_image_data(DrawableId::try_from(id)?, data)?);
        },
        (DrawableSelector::Movie(position), Replace::Data) => {
            let id = editor.body_movies()?[position.get()].drawable_object_id;
            drop(editor.replace_body_movie_data(id, data)?);
        },
        (DrawableSelector::Movie(position), Replace::Poster) => {
            let id = editor.body_movies()?[position.get()].drawable_object_id;
            drop(editor.replace_body_movie_poster(id, data)?);
        },
        (DrawableSelector::Audio(position), _) => {
            let id = editor.body_audio()?[position.get()].drawable_object_id;
            drop(editor.replace_body_audio_data(id, data)?);
        },
        _ => return Err("not media".into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_replace(
    source: &[u8],
    selector: DrawableSelector,
    asset: Replace,
    data: &[u8],
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_media();
    match asset {
        Replace::Data => drop(edit.replace_data(selector, data)?),
        Replace::Poster => drop(edit.replace_poster(selector, data)?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn replace_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    target: (DrawableSelector, Replace),
    data: &[u8],
) -> TestResult {
    let (selector, asset) = target;
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| host_replace(source, selector, asset, data))
        .leaf(|source| leaf_replace(source, selector, asset, data))
        .read("host media", host_media)
        .run()?;
    Ok(())
}

/// Tier B: on the native image the rewritten members' ZIP layout differs
/// (content equal), as for the native removal.
#[test]
fn pages_body_image_replace_data_parity() -> TestResult {
    let data = fixture(REPLACEMENT_PNG)?.bytes;
    replace_case(
        "pages_body_image_replace_data_parity",
        Tier::B,
        Evidence::Queued,
        &image_fixture()?,
        (DrawableSelector::image(0), Replace::Data),
        &data,
    )?;
    replace_case(
        "pages_body_image_replace_data_parity",
        Tier::B,
        Evidence::Queued,
        &fixture(IMAGES)?,
        (DrawableSelector::image(0), Replace::Data),
        &data,
    )
}

#[test]
fn pages_body_movie_replace_data_parity() -> TestResult {
    replace_case(
        "pages_body_movie_replace_data_parity",
        Tier::A,
        Evidence::Queued,
        &movie_fixture()?,
        (DrawableSelector::movie(0), Replace::Data),
        REPLACEMENT_MOVIE,
    )
}

#[test]
fn pages_body_movie_replace_poster_parity() -> TestResult {
    let data = fixture(REPLACEMENT_PNG)?.bytes;
    replace_case(
        "pages_body_movie_replace_poster_parity",
        Tier::A,
        Evidence::Queued,
        &movie_fixture()?,
        (DrawableSelector::movie(0), Replace::Poster),
        &data,
    )
}

#[test]
fn pages_body_audio_replace_data_parity() -> TestResult {
    replace_case(
        "pages_body_audio_replace_data_parity",
        Tier::A,
        Evidence::Queued,
        &audio_fixture()?,
        (DrawableSelector::audio(0), Replace::Data),
        REPLACEMENT_AUDIO,
    )
}

/// The host's data identifier as a number (its accessor is crate-private).
fn raw_id(id: litchi_iwa::media::MediaAssetId) -> u64 {
    id.to_string().parse().unwrap_or(0)
}

type AssetFacts = (
    u64,
    String,
    Option<String>,
    Option<u64>,
    Option<u64>,
    Vec<u8>,
    u64,
    usize,
    usize,
    bool,
);

fn host_assets(bytes: &[u8]) -> TestResult<Vec<AssetFacts>> {
    let mut assets = pages::Draft::from_bytes(bytes)?
        .media_assets()?
        .into_iter()
        .map(|asset| {
            (
                raw_id(asset.data_identifier),
                asset.preferred_filename,
                asset.package_path,
                asset.size,
                asset.declared_size,
                asset.digest,
                asset.component_reference_count,
                asset.component_reference_record_count,
                asset.message_reference_count,
                asset.has_data_metadata,
            )
        })
        .collect::<Vec<_>>();
    assets.sort();
    Ok(assets)
}

fn leaf_assets(bytes: &[u8]) -> TestResult<Vec<AssetFacts>> {
    let mut assets = Package::from_bytes(bytes)?
        .media_assets()?
        .into_iter()
        .map(|asset| {
            (
                asset.data().get(),
                asset.preferred_file_name().to_owned(),
                asset.package_path().map(str::to_owned),
                asset.size(),
                asset.declared_size(),
                asset.digest().to_vec(),
                asset.component_reference_count(),
                asset.component_reference_record_count(),
                asset.message_reference_count(),
                asset.has_data_metadata(),
            )
        })
        .collect::<Vec<_>>();
    assets.sort();
    Ok(assets)
}

/// `media_assets` and `extract_media` agree on native and host-built media
/// documents.
#[test]
fn pages_media_assets_read_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    Case::new(
        "pages_media_assets_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host media_assets", host_assets)
    .read("leaf media_assets", leaf_assets)
    .run()?;
    for document in [
        fixture(IMAGES)?,
        fixture(crate::pages_drawable_cases::AUDIO)?,
        fixture("iwork/pages/body-sections-unicode.pages")?,
        image_fixture()?,
        movie_fixture()?,
        audio_fixture()?,
    ] {
        let host = host_assets(&document.bytes)?;
        assert_eq!(host, leaf_assets(&document.bytes)?, "{}", document.path);
        let editor = pages::Draft::from_bytes(&document.bytes)?;
        let package = Package::from_bytes(&document.bytes)?;
        for (index, section) in editor.sections().iter().enumerate() {
            let mut host = editor
                .section_media_assets(section.object_id)?
                .into_iter()
                .map(|asset| raw_id(asset.data_identifier))
                .collect::<Vec<_>>();
            host.sort_unstable();
            let leaf = package
                .section_media_assets(litchi_pages::SectionSelector::index(index))?
                .into_iter()
                .map(|asset| asset.data().get())
                .collect::<Vec<_>>();
            assert_eq!(host, leaf, "{} section {index}", document.path);
        }
        for asset in editor.media_assets()? {
            if asset.package_path.is_none() {
                continue;
            }
            let id =
                litchi_iwa_common::media::embedded::AssetId::new(raw_id(asset.data_identifier))?;
            assert_eq!(
                editor.extract_media(asset.data_identifier)?,
                package.extract_media(id)?,
                "{}",
                document.path
            );
        }
    }
    Ok(())
}

#[test]
fn pages_replace_media_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    let data = fixture(REPLACEMENT_PNG)?.bytes;
    let target = host_assets(&input.bytes)?
        .iter()
        .find(|asset| asset.2.is_some())
        .map(|asset| asset.0)
        .ok_or("no asset")?;
    let data = &data;
    Case::new(
        "pages_replace_media_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let asset = editor
            .media_assets()?
            .into_iter()
            .find(|asset| raw_id(asset.data_identifier) == target)
            .ok_or("asset")?;
        drop(editor.replace_media(asset.data_identifier, data)?);
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_media();
        edit.replace_asset(
            litchi_iwa_common::media::embedded::AssetId::new(target)?,
            data,
        )?;
        let commit = edit.commit()?;
        let restored = commit
            .package()
            .apply_body_media(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host media_assets", host_assets)
    .read("leaf media_assets", leaf_assets)
    .run()?;
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn image_replace_data_cost_abba() -> TestResult {
    let data = fixture(REPLACEMENT_PNG)?.bytes;
    let selector = DrawableSelector::image(0);
    crate::pages_cost_probes::abba(
        &fixture(IMAGES)?,
        |source| host_replace(source, selector, Replace::Data, &data),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_media();
            edit.replace_data(selector, &data)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn media_assets_read_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(IMAGES)?,
        |source| Ok(format!("{:?}", host_assets(source)?).into_bytes()),
        |source| Ok(format!("{:?}", leaf_assets(source)?).into_bytes()),
    )
}
