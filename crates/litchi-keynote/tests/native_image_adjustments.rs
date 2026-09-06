//! Native Keynote integration coverage for selector-first image adjustments.
//!
//! The source presentation was authored and reopened by Keynote 14.4.  Its
//! first slide owns one file-backed image whose basic inspector controls have
//! neutral display values.  Native exposure and saturation are omitted while
//! enhancement is explicitly disabled.  The same native record carries a
//! sharpness adjustment outside the public basic-controls projection, and the
//! image has a title marker and an embedded asset used by the preservation
//! assertions below.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tsd;
use litchi_keynote::slide::image::{
    ImageAdjustment, ImageAdjustments, ImageEnhancement, ImageSelector,
};
use litchi_keynote::{Package, SlideImageAdjustmentsError, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SLIDE_MEMBER: &str = "Index/Slide-2652150.iwa";
const IMAGE_ASSET_MEMBER: &str = "Data/abstract1-9073.jpg";
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const IMAGE_ASSET_IDENTIFIER: u64 = 9_073;
const TITLE_MARKER: &[u8] = b"Native image adjustment marker";
const RETIREMENT_MARKER: &[u8] = b"Focused Keynote image adjustment retirement saved";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/image-adjustments-native.key")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/image-adjustments-native-resaved.key")
}

fn retirement_resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/image-adjustments-retirement-resaved.key")
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn baseline_adjustments() -> ImageAdjustments {
    ImageAdjustments::new().with_enhancement(Some(ImageEnhancement::Disabled))
}

fn replacement_adjustments() -> TestResult<ImageAdjustments> {
    Ok(ImageAdjustments::new()
        .with_exposure(Some(ImageAdjustment::new(0.25)?))
        .with_saturation(Some(ImageAdjustment::new(-0.2)?))
        .with_enhancement(Some(ImageEnhancement::Disabled)))
}

fn image_archive(package: &[u8]) -> TestResult<tsd::ImageArchive> {
    let catalog = Catalog::from_bytes(package)?;
    let slide = catalog
        .iter()
        .find(|entry| entry.name() == SLIDE_MEMBER)
        .ok_or_else(|| io::Error::other("native image slide member is missing"))?;
    let bytes = SnappyStream::decompress(slide.data())?.into_bytes();
    let archive = Archive::parse(&bytes)?;
    let mut selected = None;
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ != IMAGE_MESSAGE_TYPE {
                continue;
            }
            let image = tsd::ImageArchive::decode(message.data.as_slice())?;
            if image
                .data
                .as_ref()
                .is_some_and(|reference| reference.identifier == IMAGE_ASSET_IDENTIFIER)
            {
                if selected.is_some() {
                    return Err(io::Error::other("native image asset is ambiguous").into());
                }
                selected = Some(image);
            }
        }
    }
    selected.ok_or_else(|| io::Error::other("native image asset is missing").into())
}

fn slide_contains_marker(package: &[u8], marker: &[u8]) -> TestResult<bool> {
    let catalog = Catalog::from_bytes(package)?;
    let slide = catalog
        .iter()
        .find(|entry| entry.name() == SLIDE_MEMBER)
        .ok_or_else(|| io::Error::other("native image slide member is missing"))?;
    let bytes = SnappyStream::decompress(slide.data())?.into_bytes();
    let archive = Archive::parse(&bytes)?;
    Ok(archive.objects.iter().any(|object| {
        object.messages.iter().any(|message| {
            message
                .data
                .windows(marker.len())
                .any(|window| window == marker)
        })
    }))
}

fn assert_native_image_preserved(source: &[u8], candidate: &[u8]) -> TestResult {
    let source_image = image_archive(source)?;
    let candidate_image = image_archive(candidate)?;
    assert_eq!(candidate_image.data, source_image.data);
    assert_eq!(candidate_image.super_.title, source_image.super_.title);

    let source_advanced = source_image
        .image_adjustments
        .as_ref()
        .ok_or_else(|| io::Error::other("native image adjustments are missing"))?;
    let candidate_advanced = candidate_image
        .image_adjustments
        .as_ref()
        .ok_or_else(|| io::Error::other("candidate image adjustments are missing"))?;
    assert_eq!(source_advanced.sharpness, Some(0.25));
    assert_eq!(candidate_advanced.sharpness, source_advanced.sharpness);
    assert_eq!(candidate_advanced.contrast, source_advanced.contrast);
    assert_eq!(candidate_advanced.highlights, source_advanced.highlights);
    assert_eq!(candidate_advanced.shadows, source_advanced.shadows);
    assert_eq!(candidate_advanced.denoise, source_advanced.denoise);
    assert_eq!(candidate_advanced.temperature, source_advanced.temperature);
    assert_eq!(candidate_advanced.tint, source_advanced.tint);
    assert_eq!(
        candidate_advanced.bottom_level,
        source_advanced.bottom_level
    );
    assert_eq!(candidate_advanced.top_level, source_advanced.top_level);
    assert_eq!(candidate_advanced.gamma, source_advanced.gamma);
    assert_eq!(
        candidate_advanced.represents_sage_adjustments,
        source_advanced.represents_sage_adjustments
    );

    assert!(slide_contains_marker(candidate, TITLE_MARKER)?);
    Ok(())
}

fn assert_asset_unchanged(source: &[u8], candidate: &[u8]) -> TestResult {
    let source_catalog = Catalog::from_bytes(source)?;
    let candidate_catalog = Catalog::from_bytes(candidate)?;
    let source_asset = source_catalog
        .iter()
        .find(|entry| entry.name() == IMAGE_ASSET_MEMBER)
        .ok_or_else(|| io::Error::other("native image asset member is missing"))?;
    let candidate_asset = candidate_catalog
        .iter()
        .find(|entry| entry.name() == IMAGE_ASSET_MEMBER)
        .ok_or_else(|| io::Error::other("candidate image asset member is missing"))?;
    assert_eq!(candidate_asset.data(), source_asset.data());
    assert_eq!(
        candidate_asset.raw_record().local_record(),
        source_asset.raw_record().local_record()
    );
    Ok(())
}

fn assert_exact_locality(source: &[u8], candidate: &[u8]) -> TestResult {
    let source_catalog = Catalog::from_bytes(source)?;
    let candidate_catalog = Catalog::from_bytes(candidate)?;
    let mut changed = Vec::new();
    for source_entry in source_catalog.iter() {
        let candidate_entry = candidate_catalog
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or_else(|| io::Error::other("candidate removed a native package member"))?;
        if source_entry.data() != candidate_entry.data() {
            changed.push(source_entry.name().to_owned());
        } else {
            assert_eq!(
                candidate_entry.raw_record().local_record(),
                source_entry.raw_record().local_record(),
                "unchanged member {} lost its exact local ZIP record",
                source_entry.name()
            );
        }
    }
    changed.sort_unstable();
    assert_eq!(changed, [SLIDE_MEMBER]);
    assert_eq!(source_catalog.len(), candidate_catalog.len());
    Ok(())
}

fn assert_basic_adjustments(package: &Package, expected: ImageAdjustments) -> TestResult {
    assert_eq!(
        package.slide_image_adjustments(SlideSelector::index(0), ImageSelector::index(0))?,
        expected
    );
    Ok(())
}

#[test]
fn native_image_adjustments_read_exact_noop_and_preserve_context() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let original =
        package.slide_image_adjustments(SlideSelector::index(0), ImageSelector::index(0))?;
    assert_eq!(original, baseline_adjustments());
    assert_native_image_preserved(&source, &source)?;
    assert_asset_unchanged(&source, &source)?;

    let no_op = package
        .edit_slide_image_adjustments(SlideSelector::index(0), ImageSelector::index(0))?
        .set(original)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_basic_adjustments(no_op.package(), original)?;
    Ok(())
}

#[test]
fn native_image_adjustments_replace_reopen_preserve_and_inverse_exactly() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let replacement = replacement_adjustments()?;
    let changed = package
        .edit_slide_image_adjustments(SlideSelector::index(0), ImageSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let candidate = exact_bytes(changed.package())?;

    assert_basic_adjustments(changed.package(), replacement)?;
    assert_native_image_preserved(&source, &candidate)?;
    assert_asset_unchanged(&source, &candidate)?;
    assert_exact_locality(&source, &candidate)?;

    let reopened = Package::from_bytes(&candidate)?;
    assert_basic_adjustments(&reopened, replacement)?;
    assert_native_image_preserved(&source, &candidate)?;
    let restored = changed
        .package()
        .apply_slide_image_adjustments(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_basic_adjustments(restored.package(), baseline_adjustments())?;
    Ok(())
}

#[test]
fn native_image_adjustments_resaved_fixture_reopens_with_candidate_values() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let expected = replacement_adjustments()?;
    assert_basic_adjustments(&package, expected)?;
    assert_native_image_preserved(&source, &source)?;
    assert_asset_unchanged(&source, &source)?;
    Ok(())
}

#[test]
fn native_image_adjustments_retirement_fixture_reads_profile_and_keeps_exact_noop() -> TestResult {
    let source = std::fs::read(retirement_resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let expected = ImageAdjustments::new()
        .with_exposure(Some(ImageAdjustment::new(-0.2)?))
        .with_saturation(Some(ImageAdjustment::new(0.35)?))
        .with_enhancement(Some(ImageEnhancement::Enabled));

    assert_basic_adjustments(&package, expected)?;
    assert!(slide_contains_marker(&source, RETIREMENT_MARKER)?);
    let source_title = package
        .show()?
        .slides()
        .first()
        .and_then(|slide| slide.title())
        .map(str::to_owned);
    assert!(source_title.is_some(), "native slide title is missing");

    let no_op = package
        .edit_slide_image_adjustments(SlideSelector::index(0), ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    let candidate = exact_bytes(no_op.package())?;
    assert_eq!(candidate, source);

    let reopened = Package::from_bytes(&candidate)?;
    assert_basic_adjustments(&reopened, expected)?;
    assert_eq!(
        reopened
            .show()?
            .slides()
            .first()
            .and_then(|slide| slide.title())
            .map(str::to_owned),
        source_title
    );
    assert!(slide_contains_marker(&candidate, RETIREMENT_MARKER)?);
    Ok(())
}

#[test]
fn native_image_adjustments_reject_out_of_range_and_stale_transactions() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let error = package
        .slide_image_adjustments(SlideSelector::index(0), ImageSelector::index(1))
        .expect_err("a second image must not be synthesized");
    assert!(matches!(
        error,
        SlideImageAdjustmentsError::ImagePositionNotFound { .. }
    ));

    let replacement = replacement_adjustments()?;
    let changed = package
        .edit_slide_image_adjustments(SlideSelector::index(0), ImageSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let stale = changed
        .package()
        .apply_slide_image_adjustments(changed.patch())
        .expect_err("a patch must not apply to its own target as its source");
    assert!(matches!(stale, SlideImageAdjustmentsError::PatchConflict));
    Ok(())
}
