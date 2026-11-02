//! Replacing one embedded media asset by its handle (record 0417):
//! `Package::edit_media_asset_data` keeps the asset's identity, rewrites its
//! bytes, digest and length, reads back, replays and inverts exactly.

use std::num::NonZeroUsize;

use litchi_keynote::Package;
use litchi_keynote::slide::media::asset_data::{EditOptions, Error, Refusal};
use litchi_keynote::slide::media::assets::{Kind, MediaAsset};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-native.key");
const WAV: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/tone.wav");
const MOV: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/striped.mov");
const PNG: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/poster.png");

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

#[test]
fn replace_reads_back_replays_and_inverts() -> TestResult {
    let package = Package::from_bytes(SOURCE)?;
    let assets = package.media_assets()?;
    for asset in assets.iter().filter(|asset| asset.is_materialized()) {
        let replacement = match asset.kind() {
            Kind::Audio => WAV,
            Kind::Video => MOV,
            _ => PNG,
        };
        if package.media_asset_data(asset)? != replacement {
            replace_one(&package, asset, replacement)?;
        }
    }
    Ok(())
}

fn replace_one(package: &Package, asset: &MediaAsset, replacement: &[u8]) -> TestResult {
    let assets = package.media_assets()?;
    let mut edit = package.edit_media_asset_data(asset)?;
    assert_eq!(edit.before(), package.media_asset_data(asset)?);
    edit.set(replacement)?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let edited = commit.package();
    edited.validate()?;
    let after = edited
        .media_assets()?
        .into_iter()
        .find(|candidate| candidate.package_path() == asset.package_path())
        .ok_or("asset vanished")?;
    assert_eq!(edited.media_asset_data(&after)?, replacement);
    assert_eq!(after.size(), Some(replacement.len() as u64));
    assert_eq!(after.declared_size(), Some(replacement.len() as u64));
    assert_eq!(edited.media_assets()?.len(), assets.len());

    let replayed = package.apply_media_asset_data(commit.patch())?;
    assert_eq!(bytes(replayed.package())?, bytes(edited)?);
    let restored = edited.apply_media_asset_data(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, SOURCE);
    assert!(matches!(
        edited.apply_media_asset_data(commit.patch()),
        Err(Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn identical_bytes_are_a_no_op() -> TestResult {
    let package = Package::from_bytes(SOURCE)?;
    let asset = package
        .media_assets()?
        .into_iter()
        .find(|asset| asset.is_materialized())
        .ok_or("asset")?;
    let current = package.media_asset_data(&asset)?.to_vec();
    let mut edit = package.edit_media_asset_data(&asset)?;
    edit.set(&current)?;
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package())?, SOURCE);
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let package = Package::from_bytes(SOURCE)?;
    let asset = package
        .media_assets()?
        .into_iter()
        .find(|asset| asset.is_materialized())
        .ok_or("asset")?;
    let mut edit = package.edit_media_asset_data(&asset)?;
    assert!(matches!(
        edit.set(&[]),
        Err(Error::Refused(Refusal::EmptyReplacement))
    ));
    assert!(matches!(
        package.edit_media_asset_data(&asset)?.commit(),
        Err(Error::Refused(Refusal::EmptyReplacement))
    ));
    let mut limited = package.edit_media_asset_data_with_options(
        &asset,
        EditOptions::new().max_replacement_bytes(NonZeroUsize::MIN),
    )?;
    assert!(matches!(
        limited.set(PNG),
        Err(Error::Refused(Refusal::ReplacementTooLarge))
    ));
    assert_eq!(
        EditOptions::default().max_replacement_bytes_value().get(),
        1 << 30
    );

    // An asset of another document is not reachable here.
    let other = Package::new();
    assert!(matches!(
        other.edit_media_asset_data(&asset),
        Err(Error::Refused(Refusal::AssetNotReachable))
    ));
    Ok(())
}
