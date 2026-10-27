//! Numbers sheet media creation parity (wave W2-A, record 0378): the host
//! `add_sheet_{image,movie,audio}` routes against
//! `litchi_numbers::sheet::drawable::Edit::{add_image, add_movie,
//! add_audio}`.

use std::time::Duration;

use litchi_iwa::numbers;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::{AudioOptions, ImageOptions, MovieOptions, Point, Size};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::numbers_drawable_cases::first_sheet;

const BASIC: &str = "iwork/numbers/basic.numbers";
const PNG: &str = "officecli/examples/ppt/ole/thumb-xlsx.png";
const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  litchi-numbers-movie";
const AUDIO: &[u8] = b"FORM\0\0\0\x10AIFClitchi-numbers-audio";

#[derive(Debug, Clone, Copy)]
enum New {
    Image,
    Movie,
    Audio,
}

fn host_add(source: &[u8], new: New, png: &[u8]) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let sheet = first_sheet(&e)?;
    let position = DrawablePoint { x: 120.0, y: 90.0 };
    let size = DrawableSize {
        width: 160.0,
        height: 120.0,
    };
    let natural = DrawableSize {
        width: 320.0,
        height: 240.0,
    };
    match new {
        New::Image => drop(e.add_sheet_image(
            sheet,
            "picture.png",
            png,
            numbers::SheetImageOptions::new(position, size).with_natural_size(natural),
        )?),
        New::Movie => drop(
            e.add_sheet_movie(
                sheet,
                "clip.mov",
                MOVIE,
                "poster.png",
                png,
                numbers::SheetMovieOptions::new(position, size, Duration::from_millis(2_500))
                    .with_natural_size(natural),
            )?,
        ),
        New::Audio => drop(e.add_sheet_audio(
            sheet,
            "sound.aiff",
            AUDIO,
            numbers::SheetAudioOptions::new(position, Duration::from_secs(4)),
        )?),
    }
    Ok(e.to_bytes()?)
}

fn leaf_add(source: &[u8], new: New, png: &[u8]) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawables(0)?;
    let position = Point { x: 120.0, y: 90.0 };
    let size = Size {
        width: 160.0,
        height: 120.0,
    };
    let natural = Size {
        width: 320.0,
        height: 240.0,
    };
    match new {
        New::Image => edit.add_image(
            "picture.png",
            png,
            ImageOptions::new(position, size).with_natural_size(natural),
        )?,
        New::Movie => edit.add_movie(
            ("clip.mov", MOVIE),
            ("poster.png", png),
            MovieOptions::new(position, size, Duration::from_millis(2_500))
                .with_natural_size(natural),
        )?,
        New::Audio => edit.add_audio(
            "sound.aiff",
            AUDIO,
            AudioOptions::new(position, Duration::from_secs(4)),
        )?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn leaf_list(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    Ok(format!(
        "{:?} {:?}",
        package
            .sheet_drawable_list(0)?
            .iter()
            .map(|item| (item.kind(), item.geometry()))
            .collect::<Vec<_>>(),
        package
            .media_assets()?
            .iter()
            .map(|asset| (asset.data().get(), asset.package_path().map(str::to_owned)))
            .collect::<Vec<_>>()
    ))
}

fn run(name: &'static str, input: &Fixture, new: New, png: &[u8]) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .host(|source| host_add(source, new, png))
        .leaf(|source| leaf_add(source, new, png))
        .read("leaf drawables and media", leaf_list)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{new:?}: {error}"))?;
    Ok(())
}

#[test]
fn numbers_sheet_media_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let png = fixture(PNG)?.bytes;
    let input = fixture(BASIC)?;
    for new in [New::Image, New::Movie, New::Audio] {
        run("numbers_sheet_media_add_parity", &input, new, &png)?;
    }
    Ok(())
}
