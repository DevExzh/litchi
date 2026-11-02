//! Keynote media-asset replacement parity (wave W4-R, record 0417): the
//! host's `keynote::Draft::replace_media` and the format-neutral
//! `media::Draft::{replace, set_max_replacement_bytes}` against
//! `litchi_keynote::Package::edit_media_asset_data`.
//!
//! Every materialized asset the document reaches is replaced in turn (images,
//! movie content and posters, audio) with bytes of its own media type; the
//! leaf additionally invalidates the root previews (its documented policy).

use std::num::NonZeroUsize;

use litchi_iwa::keynote;
use litchi_iwa_common::media::Type;
use litchi_keynote::Package;
use litchi_keynote::slide::media::asset_data::{EditOptions, Error, Refusal};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// Presentations whose reachable assets cover images, movies, posters and
/// audio.
const FIXTURES: &[&str] = &[
    "iwork/keynote/image-adjustments-native.key",
    "iwork/keynote/slide-movie-geometry-source-native.key",
    "iwork/keynote/media-properties-audio-focused-native.key",
];

/// Replacement candidates, one per media type.
const REPLACEMENTS: &[&str] = &[
    "officecli/examples/ppt/ole/thumb-xlsx.png",
    "iwork/directory/keynote/basic.key/preview-micro.jpg",
    "iwork/keynote/media-replacement-assets/striped.mov",
    "iwork/keynote/media-replacement-assets/tone.wav",
];

/// A replacement for `current` of the same detected type and file
/// extension family.
fn replacement_for(current: &[u8], path: &str) -> TestResult<Option<Vec<u8>>> {
    let kind = Type::from_bytes(current);
    let jpeg = |name: &str| {
        let lower = name.to_ascii_lowercase();
        lower.ends_with(".jpg") || lower.ends_with(".jpeg")
    };
    for candidate in REPLACEMENTS {
        let bytes = fixture(candidate)?.bytes;
        if Type::from_bytes(&bytes) == kind
            && bytes != current
            && (kind != Type::Image || jpeg(candidate) == jpeg(path))
        {
            return Ok(Some(bytes));
        }
    }
    Ok(None)
}

/// The reachable assets' (preferred name, path, size, digest), host-read.
fn host_assets(bytes: &[u8]) -> TestResult<Vec<String>> {
    Ok(keynote::Draft::from_bytes(bytes)?
        .media_assets()?
        .iter()
        .map(|asset| {
            format!(
                "{} {:?} {:?} {:?} {:02x?}",
                asset.preferred_filename,
                asset.package_path,
                asset.size,
                asset.declared_size,
                asset.digest
            )
        })
        .collect())
}

/// The same facts, leaf-read.
fn leaf_assets(bytes: &[u8]) -> TestResult<Vec<String>> {
    Ok(Package::from_bytes(bytes)?
        .media_assets()?
        .iter()
        .map(|asset| {
            format!(
                "{} {:?} {:?} {:?} {:02x?}",
                asset.preferred_file_name(),
                asset.package_path(),
                asset.size(),
                asset.declared_size(),
                asset.digest()
            )
        })
        .collect())
}

/// Host `media::Draft::replace` (`format_neutral`) or `keynote::Draft::replace_media`.
fn host_replace(
    source: &[u8],
    index: usize,
    data: &[u8],
    format_neutral: bool,
) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let asset = editor
        .media_assets()?
        .get(index)
        .ok_or("host asset")?
        .data_identifier;
    if format_neutral {
        let mut media = litchi_iwa::media::Draft::from_bytes(source)?;
        drop(media.replace(asset, data)?);
        return Ok(media.to_bytes()?);
    }
    drop(editor.replace_media(asset, data)?);
    Ok(editor.to_bytes()?)
}

fn leaf_replace(source: &[u8], index: usize, data: &[u8]) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let asset = package
        .media_assets()?
        .get(index)
        .cloned()
        .ok_or("leaf asset")?;
    let mut edit = package.edit_media_asset_data(&asset)?;
    edit.set(data)?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let restored = commit
        .package()
        .apply_media_asset_data(&commit.patch().inverse())?;
    let replayed = package.apply_media_asset_data(commit.patch())?;
    assert_eq!(
        leaf_bytes!(replayed.package()),
        leaf_bytes!(commit.package())
    );
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(name: &'static str, format_neutral: bool) -> TestResult<usize> {
    let mut cases = 0;
    for path in FIXTURES {
        let input = fixture(path)?;
        let assets = Package::from_bytes(&input.bytes)?.media_assets()?;
        for (index, asset) in assets.iter().enumerate() {
            let Some(member) = asset.package_path() else {
                continue;
            };
            let current = Package::from_bytes(&input.bytes)?
                .media_asset_data(asset)?
                .to_vec();
            let Some(data) = replacement_for(&current, member)? else {
                continue;
            };
            let data = &data;
            Case::new(name, Format::Keynote, Tier::B, Evidence::Queued, &input)
                .leaf_invalidates_root_previews()
                .host(move |source| host_replace(source, index, data, format_neutral))
                .leaf(move |source| leaf_replace(source, index, data))
                .read("host media_assets", host_assets)
                .read("leaf media_assets", leaf_assets)
                .run()
                .map_err(|error| format!("{path} asset {index} ({member}): {error}"))?;
            cases += 1;
        }
    }
    Ok(cases)
}

/// Host `keynote::Draft::replace_media` on every replaceable asset.
#[test]
fn keynote_media_asset_replace_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let cases = run("keynote_media_asset_replace_parity", false)?;
    assert!(cases >= 4, "only {cases} assets replaced");
    Ok(())
}

/// Host format-neutral `media::Draft::replace` on every replaceable asset.
#[test]
fn media_draft_replace_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let cases = run("media_draft_replace_parity", true)?;
    assert!(cases >= 4, "only {cases} assets replaced");
    Ok(())
}

/// Refusals: the replacement ceiling (host `set_max_replacement_bytes`,
/// leaf `EditOptions::max_replacement_bytes`), empty data and a mismatched
/// media type are refused by both routes, and an identical replacement is
/// an exact no-op.
#[test]
fn media_asset_replacement_limit_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture(FIXTURES[0])?;
    let package = Package::from_bytes(&input.bytes)?;
    let asset = package
        .media_assets()?
        .into_iter()
        .find(|asset| asset.is_materialized())
        .ok_or("asset")?;
    let current = package.media_asset_data(&asset)?.to_vec();
    let data = replacement_for(&current, asset.package_path().ok_or("path")?)?.ok_or("data")?;
    let limit = NonZeroUsize::new(data.len() - 1).ok_or("limit")?;

    let host_assets_list = litchi_iwa::media::Draft::from_bytes(&input.bytes)?;
    let host_asset = keynote::Draft::from_bytes(&input.bytes)?
        .media_assets()?
        .into_iter()
        .find(|candidate| candidate.package_path.as_deref() == asset.package_path())
        .ok_or("host asset")?
        .data_identifier;
    let mut host = host_assets_list.clone();
    host.set_max_replacement_bytes(limit.get())?;
    assert!(host.replace(host_asset, &data).is_err(), "host ceiling");
    let mut leaf = package.edit_media_asset_data_with_options(
        &asset,
        EditOptions::new().max_replacement_bytes(limit),
    )?;
    assert!(matches!(
        leaf.set(&data),
        Err(Error::Refused(Refusal::ReplacementTooLarge))
    ));

    let mut host = host_assets_list.clone();
    assert!(host.replace(host_asset, &[]).is_err(), "host empty");
    assert!(matches!(
        package.edit_media_asset_data(&asset)?.set(&[]),
        Err(Error::Refused(Refusal::EmptyReplacement))
    ));

    let other_type = fixture(REPLACEMENTS[3])?.bytes;
    let mut host = host_assets_list.clone();
    assert!(host.replace(host_asset, &other_type).is_err(), "host type");
    assert!(matches!(
        package.edit_media_asset_data(&asset)?.set(&other_type),
        Err(Error::Refused(Refusal::ReplacementType))
    ));

    let mut same = package.edit_media_asset_data(&asset)?;
    same.set(&current)?;
    let commit = same.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(leaf_bytes!(commit.package()), leaf_bytes!(&package));
    Ok(())
}

/// `KEYNOTE_MEDIA_ROUTE` is ignored; image data replacement, host
/// `replace_media` against the leaf family.
#[test]
#[ignore = "callgrind cost probe (record 0417)"]
fn keynote_media_asset_replace_cost_abba() -> TestResult {
    let input = fixture(FIXTURES[0])?;
    let package = Package::from_bytes(&input.bytes)?;
    let (index, asset) = package
        .media_assets()?
        .into_iter()
        .enumerate()
        .find(|(_, asset)| asset.is_materialized())
        .ok_or("asset")?;
    let current = package.media_asset_data(&asset)?.to_vec();
    let data = replacement_for(&current, asset.package_path().ok_or("path")?)?.ok_or("data")?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_replace(source, index, &data, false),
        |source| {
            let package = Package::from_bytes(source)?;
            let asset = package.media_assets()?.get(index).cloned().ok_or("asset")?;
            let mut edit = package.edit_media_asset_data(&asset)?;
            edit.set(&data)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

/// Fixtures that receive host-inserted unreferenced data records, per
/// format.
const PRUNE_FIXTURES: &[(&str, Format)] = &[
    (
        "iwork/keynote/image-adjustments-native.key",
        Format::Keynote,
    ),
    ("iwork/pages/basic.pages", Format::Pages),
    ("iwork/numbers/basic.numbers", Format::Numbers),
];

/// `source` with `count` unreferenced data records (host
/// `media::Draft::insert_unreferenced`, which the leaves never publish).
fn with_orphans(source: &[u8], count: usize) -> TestResult<Vec<u8>> {
    let mut draft = litchi_iwa::media::Draft::from_bytes(source)?;
    let png = fixture(REPLACEMENTS[0])?.bytes;
    let wav = fixture(REPLACEMENTS[3])?.bytes;
    for index in 0..count {
        let (name, data) = if index % 2 == 0 {
            ("orphan.png", &png)
        } else {
            ("orphan.wav", &wav)
        };
        drop(draft.insert_unreferenced(name, data)?);
    }
    Ok(draft.to_bytes()?)
}

/// Host: remove every unreferenced record one by one
/// (`media::Draft::remove_unreferenced`).
fn host_prune(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut draft = litchi_iwa::media::Draft::from_bytes(source)?;
    let orphans = draft
        .assets()
        .iter()
        .filter(|asset| !asset.is_referenced())
        .map(|asset| asset.data_identifier)
        .collect::<Vec<_>>();
    assert!(!orphans.is_empty());
    for identifier in orphans {
        drop(draft.remove_unreferenced(identifier)?);
    }
    Ok(draft.to_bytes()?)
}

/// Leaf: `prune_unreferenced` on each format's media family, with its
/// inverse.
fn leaf_prune(format: Format, source: &[u8]) -> TestResult<LeafRun> {
    Ok(match format {
        Format::Keynote => {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_media_assets();
            edit.prune_unreferenced();
            let commit = edit.commit()?;
            assert!(commit.diagnostics().pruned() > 0);
            let restored = commit
                .package()
                .apply_media_asset_data(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Format::Pages => {
            let package = litchi_pages::Package::from_bytes(source)?;
            let mut edit = package.edit_body_media();
            edit.prune_unreferenced()?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_media(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Format::Numbers => {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let mut edit = package.edit_media();
            edit.prune_unreferenced()?;
            let commit = edit.commit()?;
            let restored = commit.package().apply_media(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
    })
}

/// Host `media::Draft::remove_unreferenced` on every unreferenced record
/// against the leaf `prune_unreferenced` on all three formats (one and
/// three orphans), and the exact no-op on a package without orphans.
#[test]
fn media_prune_unreferenced_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for &(path, format) in PRUNE_FIXTURES {
        let base = fixture(path)?;
        for count in [1, 3] {
            let input = crate::harness::Fixture {
                path,
                bytes: with_orphans(&base.bytes, count)?,
            };
            let mut case = Case::new(
                "media_prune_unreferenced_parity",
                format,
                Tier::B,
                Evidence::Queued,
                &input,
            );
            if format == Format::Keynote {
                // Every Keynote structural publication invalidates the
                // root previews (the leaf's natively verified policy).
                case = case.leaf_invalidates_root_previews();
            }
            case.host(host_prune)
                .leaf(move |source| leaf_prune(format, source))
                .run()
                .map_err(|error| format!("{path} ({count}): {error}"))?;
        }
        let unchanged = leaf_bytes!(&Package::from_bytes(&fixture(FIXTURES[0])?.bytes)?);
        let package = Package::from_bytes(&unchanged)?;
        let mut edit = package.edit_media_assets();
        edit.prune_unreferenced();
        let commit = edit.commit()?;
        assert!(!commit.diagnostics().changed() && commit.patch().is_noop());
    }
    Ok(())
}

/// `source` with one unreferenced, unmaterialized data record: a host
/// orphan whose `Data/` member is then removed through the host raw edit.
fn with_unmaterialized_orphan(source: &[u8]) -> TestResult<Vec<u8>> {
    let bytes = with_orphans(source, 1)?;
    let member = litchi_iwa::media::Draft::from_bytes(&bytes)?
        .assets()
        .iter()
        .find(|asset| !asset.is_referenced())
        .and_then(|asset| asset.package_path.clone())
        .ok_or("orphan member")?;
    let package = litchi_iwa::Package::from_bytes(&bytes)?;
    let mut edit = package.edit();
    edit.remove_entry(&member).ok_or("member")?;
    let unmaterialized = edit.commit()?.into_package().to_bytes()?;
    let draft = litchi_iwa::media::Draft::from_bytes(&unmaterialized)?;
    assert!(
        draft
            .assets()
            .iter()
            .any(|asset| !asset.is_referenced() && asset.package_path.is_none())
    );
    Ok(unmaterialized)
}

/// An unreferenced record without a `Data/` member: the host drops its
/// `DataInfo`; every leaf prunes it too (Numbers through the shared
/// transition's `data_record_removals`).
#[test]
fn media_prune_unmaterialized_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for &(path, format) in PRUNE_FIXTURES {
        let input = crate::harness::Fixture {
            path,
            bytes: with_unmaterialized_orphan(&fixture(path)?.bytes)?,
        };
        let mut case = Case::new(
            "media_prune_unmaterialized_parity",
            format,
            Tier::B,
            Evidence::Queued,
            &input,
        );
        if format == Format::Keynote {
            case = case.leaf_invalidates_root_previews();
        }
        case.host(host_prune)
            .leaf(move |source| leaf_prune(format, source))
            .run()
            .map_err(|error| format!("{path}: {error}"))?;
    }
    Ok(())
}
