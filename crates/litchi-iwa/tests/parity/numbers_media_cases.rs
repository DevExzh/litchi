//! Numbers embedded-media parity (wave W2-A, record 0378): the host
//! `media_assets`, `sheet_media_assets`, `extract_media`, `replace_media`
//! and `replace_sheet_{image,movie,audio}_data` /
//! `replace_sheet_movie_poster` routes against
//! `litchi_numbers::Package::{media_assets, sheet_media_assets,
//! extract_media, edit_media}`.

use litchi_iwa::numbers;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::media::AssetId;
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{authored_fixture, first_sheet, host_drawable};

const IMAGE: &str = "iwork/numbers/image-adjustments-native.numbers";
const AUDIO: &str = "iwork/numbers/audio-playback-native.numbers";
const REPLACEMENT_PNG: &str = "officecli/examples/ppt/ole/thumb-xlsx.png";

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
);

fn host_assets(bytes: &[u8]) -> TestResult<Vec<AssetFacts>> {
    let mut assets: Vec<AssetFacts> = numbers::Draft::from_bytes(bytes)?
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
            )
        })
        .collect();
    assets.sort();
    Ok(assets)
}

fn leaf_assets(bytes: &[u8]) -> TestResult<Vec<AssetFacts>> {
    let mut assets: Vec<AssetFacts> = Package::from_bytes(bytes)?
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
            )
        })
        .collect();
    assets.sort();
    Ok(assets)
}

fn documents() -> TestResult<Vec<Fixture>> {
    Ok(vec![fixture(IMAGE)?, fixture(AUDIO)?, authored_fixture()?])
}

#[test]
fn numbers_media_assets_read_parity() -> TestResult {
    for document in documents()? {
        assert_eq!(
            host_assets(&document.bytes)?,
            leaf_assets(&document.bytes)?,
            "{}",
            document.path
        );
        let editor = numbers::Draft::from_bytes(&document.bytes)?;
        let package = Package::from_bytes(&document.bytes)?;
        let mut host: Vec<u64> = editor
            .sheet_media_assets(first_sheet(&editor)?)?
            .into_iter()
            .map(|asset| raw_id(asset.data_identifier))
            .collect();
        host.sort_unstable();
        let leaf: Vec<u64> = package
            .sheet_media_assets(0)?
            .into_iter()
            .map(|asset| asset.data().get())
            .collect();
        assert_eq!(host, leaf, "{}", document.path);
        for asset in package.media_assets()? {
            if asset.package_path().is_none() {
                continue;
            }
            let id = litchi_iwa::media::MediaAssetId::try_from(asset.data().get())?;
            assert_eq!(
                editor.extract_media(id)?,
                package.extract_media(asset.data())?,
                "{}",
                document.path
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum Target {
    Asset,
    Data(Kind),
    Poster,
}

fn host_change(source: &[u8], target: Target, asset: u64, data: &[u8]) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    match target {
        Target::Asset => {
            e.replace_media(litchi_iwa::media::MediaAssetId::try_from(asset)?, data)?;
        },
        Target::Data(Kind::Image) => {
            let (s, d) = host_drawable(&e, Kind::Image, 0)?;
            e.replace_sheet_image_data(
                litchi_iwa_index::ObjectId::try_from(s)?,
                litchi_iwa_common::comment::DrawableId::from_raw(d)?,
                data,
            )?;
        },
        Target::Data(Kind::Movie) => {
            let (s, d) = host_drawable(&e, Kind::Movie, 0)?;
            e.replace_sheet_movie_data(s, d, data)?;
        },
        Target::Data(_) => {
            let (s, d) = host_drawable(&e, Kind::Audio, 0)?;
            e.replace_sheet_audio_data(s, d, data)?;
        },
        Target::Poster => {
            let (s, d) = host_drawable(&e, Kind::Movie, 0)?;
            e.replace_sheet_movie_poster(s, d, data)?;
        },
    }
    Ok(e.to_bytes()?)
}

fn leaf_change(source: &[u8], target: Target, asset: u64, data: &[u8]) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_media();
    match target {
        Target::Asset => edit.replace_asset(AssetId::new(asset)?, data)?,
        Target::Data(kind) => edit.replace_data(0, DrawableSelector::of(kind, 0), data)?,
        Target::Poster => edit.replace_poster(0, DrawableSelector::movie(0), data)?,
    };
    let commit = edit.commit()?;
    let restored = commit.package().apply_media(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Tier B on the native image: the rewritten members' ZIP layout differs
/// (content equal), as for the Pages native image.
fn run(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    target: Target,
    asset: u64,
    data: &[u8],
) -> TestResult {
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(|source| host_change(source, target, asset, data))
        .leaf(|source| leaf_change(source, target, asset, data))
        .read("host media", host_assets)
        .read("leaf media", leaf_assets)
        .run()
        .map_err(|error| format!("{target:?}: {error}"))?;
    Ok(())
}

/// The data asset of the first `kind` drawable and its bytes.
fn data_of(input: &Fixture, kind: Kind, poster: bool) -> TestResult<(u64, Vec<u8>)> {
    let package = Package::from_bytes(&input.bytes)?;
    let before: Vec<u64> = package
        .media_assets()?
        .iter()
        .filter(|asset| asset.package_path().is_some())
        .map(|asset| asset.data().get())
        .collect();
    // Replace through the drawable with a marker and read back which asset
    // changed: the drawable's data identifier stays private to the leaf.
    let mut edit = package.edit_media();
    let marker = b"litchi-marker".to_vec();
    if poster {
        edit.replace_poster(0, DrawableSelector::movie(0), &marker)?;
    } else {
        edit.replace_data(0, DrawableSelector::of(kind, 0), &marker)?;
    }
    let changed = edit.commit()?.into_package();
    for asset in before {
        let id = AssetId::new(asset)?;
        if changed.extract_media(id)? == marker {
            return Ok((asset, package.extract_media(id)?));
        }
    }
    Err("no asset changed".into())
}

#[test]
fn numbers_media_replace_parity() -> TestResult {
    let _ = Evidence::Queued;
    let png = fixture(REPLACEMENT_PNG)?.bytes;
    let image = fixture(IMAGE)?;
    let (asset, _) = data_of(&image, Kind::Image, false)?;
    run(
        "numbers_media_replace_parity",
        Tier::B,
        &image,
        Target::Asset,
        asset,
        &png,
    )?;
    Ok(())
}

#[test]
fn numbers_sheet_media_replace_data_parity() -> TestResult {
    let _ = Evidence::Queued;
    let png = fixture(REPLACEMENT_PNG)?.bytes;
    let name = "numbers_sheet_media_replace_data_parity";
    run(
        name,
        Tier::B,
        &fixture(IMAGE)?,
        Target::Data(Kind::Image),
        0,
        &png,
    )?;
    for (input, kind) in [
        (authored_fixture()?, Kind::Movie),
        (fixture(AUDIO)?, Kind::Audio),
    ] {
        let (_, mut data) = data_of(&input, kind, false)?;
        data.extend_from_slice(b"litchi-replacement");
        run(name, Tier::B, &input, Target::Data(kind), 0, &data)?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_movie_replace_poster_parity() -> TestResult {
    let png = fixture(REPLACEMENT_PNG)?.bytes;
    run(
        "numbers_sheet_movie_replace_poster_parity",
        Tier::A,
        &authored_fixture()?,
        Target::Poster,
        0,
        &png,
    )?;
    let _ = Evidence::Queued;
    Ok(())
}
