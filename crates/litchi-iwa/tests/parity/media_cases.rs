//! Shared media asset discovery (wave W1-E, record 0375): the host's
//! `MediaManager` reads against the archive-free
//! `litchi_iwa_common::media::Catalog` fed with the package members of the
//! leaf physical package catalog, on native-saved fixtures of the three formats.

use std::path::PathBuf;

use litchi_iwa::media::MediaManager;
use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_archive::package::Catalog as PackageCatalog;
use litchi_iwa_common::media::{Asset, Catalog, Type, asset::Limits};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// One asset as both sides render it: path, type, size, basename and the
/// image/video/audio predicates.
type Row = (PathBuf, Type, u64, String, [bool; 3]);

/// The engine catalog of a package: the ZIP member table of the leaf
/// physical catalog (no IWA component is parsed) fed to the common catalog.
fn engine_catalog(bytes: &[u8]) -> TestResult<Catalog> {
    let catalog = PackageCatalog::from_bytes(bytes, ArchiveLimits::default())?;
    let members: Vec<(String, u64)> = catalog
        .iter()
        .map(|entry| Ok((entry.name().to_owned(), u64::try_from(entry.data().len())?)))
        .collect::<TestResult<_>>()?;
    Ok(Catalog::from_members(
        members.iter().map(|(name, size)| (name.as_str(), *size)),
        Limits::default(),
    )?)
}

fn engine_rows(bytes: &[u8]) -> TestResult<Vec<Row>> {
    Ok(engine_catalog(bytes)?
        .assets()
        .iter()
        .map(|asset: &Asset| {
            (
                asset.path().to_path_buf(),
                asset.media_type(),
                asset.size(),
                asset.file_name().to_owned(),
                [asset.is_image(), asset.is_video(), asset.is_audio()],
            )
        })
        .collect())
}

fn host_rows(bytes: &[u8]) -> TestResult<Vec<Row>> {
    Ok(MediaManager::from_bytes(bytes)?
        .assets()
        .iter()
        .map(|asset| {
            (
                asset.path.clone(),
                asset.media_type,
                asset.size,
                asset.filename.clone(),
                [asset.is_image(), asset.is_video(), asset.is_audio()],
            )
        })
        .collect())
}

/// Lookups by basename and type, and the statistics, on both sides.
fn host_queries(bytes: &[u8]) -> TestResult<String> {
    let manager = MediaManager::from_bytes(bytes)?;
    let stats = manager.stats();
    let first = manager.assets().first().map(|asset| asset.filename.clone());
    let found = first
        .as_deref()
        .and_then(|name| manager.get(name))
        .map(|asset| asset.size);
    Ok(format!(
        "{} | {} | {:?} | {} {} {}",
        stats.summary(),
        stats.total_size_human(),
        found,
        manager.images().len(),
        manager.videos().len(),
        manager.audio().len()
    ))
}

fn engine_queries(bytes: &[u8]) -> TestResult<String> {
    let catalog = engine_catalog(bytes)?;
    let stats = catalog.stats();
    let first = catalog
        .assets()
        .first()
        .map(|asset| asset.file_name().to_owned());
    let found = first
        .as_deref()
        .and_then(|name| catalog.get(name))
        .map(Asset::size);
    Ok(format!(
        "{} | {} | {:?} | {} {} {}",
        stats.summary(),
        stats.total_size_human(),
        found,
        catalog.by_type(Type::Image).count(),
        catalog.by_type(Type::Video).count(),
        catalog.by_type(Type::Audio).count()
    ))
}

fn media_case(
    name: &'static str,
    format: Format,
    path: &'static str,
    tier: Tier,
    evidence: Evidence,
) -> TestResult {
    let input = fixture(path)?;
    if engine_catalog(&input.bytes)?.is_empty() {
        return Err(format!("{path} has no media").into());
    }
    Case::new(name, format, tier, evidence, &input)
        .read_only()
        .host(|source| Ok(litchi_iwa::Package::from_bytes(source)?.to_bytes()?))
        .leaf(move |source| {
            let bytes = match format {
                Format::Numbers => leaf_bytes!(litchi_numbers::Package::from_bytes(source)?),
                Format::Pages => leaf_bytes!(litchi_pages::Package::from_bytes(source)?),
                Format::Keynote => leaf_bytes!(litchi_keynote::Package::from_bytes(source)?),
            };
            Ok(LeafRun::output(bytes))
        })
        .read("host media assets", host_rows)
        .read("engine media assets", engine_rows)
        .read("host media queries", host_queries)
        .read("engine media queries", engine_queries)
        .read("host and engine agree", |bytes| {
            Ok(host_rows(bytes)? == engine_rows(bytes)?
                && host_queries(bytes)? == engine_queries(bytes)?)
        })
        .run()?;
    Ok(())
}

#[test]
fn keynote_media_catalog_parity() -> TestResult {
    media_case(
        "keynote_media_catalog_parity",
        Format::Keynote,
        "iwork/keynote/slide-movie-geometry-source-native.key",
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn pages_media_catalog_parity() -> TestResult {
    media_case(
        "pages_media_catalog_parity",
        Format::Pages,
        "iwork/pages/image-adjustments-native.pages",
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn numbers_media_catalog_parity() -> TestResult {
    media_case(
        "numbers_media_catalog_parity",
        Format::Numbers,
        "iwork/numbers/table-relocation-native.numbers",
        Tier::A,
        Evidence::Queued,
    )
}

#[inline(never)]
fn host_media_route(bytes: &[u8]) -> TestResult<String> {
    Ok(format!("{:?}{}", host_rows(bytes)?, host_queries(bytes)?))
}

#[inline(never)]
fn leaf_media_route(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}{}",
        engine_rows(bytes)?,
        engine_queries(bytes)?
    ))
}

/// ABBA cost probe of the media catalog read (open + list + queries, record
/// 0375 §5); toggle `*media_cases::host_media_route*` or
/// `*media_cases::leaf_media_route*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn media_catalog_cost_abba() -> TestResult {
    for path in [
        "iwork/keynote/slide-movie-geometry-source-native.key",
        "iwork/pages/image-adjustments-native.pages",
        "iwork/numbers/table-relocation-native.numbers",
    ] {
        let input = fixture(path)?;
        std::hint::black_box(host_media_route(&input.bytes)?);
        std::hint::black_box(leaf_media_route(&input.bytes)?);
        std::hint::black_box(leaf_media_route(&input.bytes)?);
        std::hint::black_box(host_media_route(&input.bytes)?);
    }
    Ok(())
}
