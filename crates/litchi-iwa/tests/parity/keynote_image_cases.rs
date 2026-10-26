//! Keynote slide-image lifecycle parity cases (wave W1-C, record 0372):
//! duplicate, flip, original size, properties, data replacement and reads
//! through `slide::image::lifecycle`.

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawableFlipAxis, DrawableProperties};
use litchi_keynote::slide::image::lifecycle::{Commit, Edit, FlipAxis, Info, MediaProperties};
use litchi_keynote::{ImageSelector, Package, SlideSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// `image-adjustments-native.key`: slide 0 owns one file image (JPEG).
const IMAGE: &str = "iwork/keynote/image-adjustments-native.key";
/// A small tracked JPEG (the micro preview of the `basic.key` directory
/// fixture) used as replacement image data.
const JPEG: &str = "iwork/directory/keynote/basic.key/preview-micro.jpg";
const SLIDE: usize = 0;

/// The host's native identifier of the first image of slide 0.
fn host_image(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_images(SLIDE)?
        .first()
        .ok_or("slide 0 has no image")?
        .drawable_object_id)
}

/// Leaf read: every image of slide 0.
fn leaf_images(bytes: &[u8]) -> TestResult<Vec<Info>> {
    Ok(Package::from_bytes(bytes)?.slide_image_list(SlideSelector::index(SLIDE))?)
}

/// Host read: per image of slide 0, its frame, flags, properties and sizes.
fn host_images(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    Ok(editor
        .slide_images(SLIDE)?
        .into_iter()
        .map(|image| {
            format!(
                "{:?} {:?} {:?} {:?} {:?}",
                image.kind,
                image.geometry,
                image.properties,
                image.original_size,
                image.natural_size
            )
        })
        .collect())
}

/// Leaf route helper: stage with `stage`, commit, check the inverse.
fn leaf_run(source: &[u8], stage: impl FnOnce(&mut Edit<'_>) -> TestResult) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_image_lifecycle();
    stage(&mut edit)?;
    let commit: Commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_image_lifecycle(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft, u64) -> TestResult,
    leaf: impl FnOnce(&mut Edit<'_>, SlideSelector<'static>, ImageSelector) -> TestResult,
) -> TestResult {
    let input = fixture(IMAGE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let image = host_image(&editor)?;
            host(&mut editor, image)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            leaf_run(source, |edit| {
                leaf(edit, SlideSelector::index(SLIDE), ImageSelector::index(0))
            })
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("leaf slide_image_list", leaf_images)
        .read("host slide_images", host_images)
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_image_duplicate_parity() -> TestResult {
    run(
        "keynote_slide_image_duplicate_parity",
        Tier::B,
        Evidence::Queued,
        |editor, image| {
            editor.duplicate_slide_image(SLIDE, image)?;
            Ok(())
        },
        |edit, slide, image| {
            edit.duplicate(slide, image)?;
            Ok(())
        },
    )
}

#[test]
fn keynote_slide_image_flip_parity() -> TestResult {
    run(
        "keynote_slide_image_flip_parity",
        Tier::B,
        Evidence::Queued,
        |editor, image| {
            editor.flip_slide_image(SLIDE, image, DrawableFlipAxis::Vertical)?;
            editor.flip_slide_image(SLIDE, image, DrawableFlipAxis::Horizontal)?;
            editor.flip_slide_image(SLIDE, image, DrawableFlipAxis::Vertical)?;
            Ok(())
        },
        |edit, slide, image| {
            edit.flip(slide, image, FlipAxis::Vertical)?;
            edit.flip(slide, image, FlipAxis::Horizontal)?;
            edit.flip(slide, image, FlipAxis::Vertical)?;
            Ok(())
        },
    )
}

#[test]
fn keynote_slide_image_original_size_parity() -> TestResult {
    run(
        "keynote_slide_image_original_size_parity",
        Tier::B,
        Evidence::Queued,
        |editor, image| {
            editor.restore_slide_image_original_size(SLIDE, image)?;
            Ok(())
        },
        |edit, slide, image| {
            edit.restore_original_size(slide, image)?;
            Ok(())
        },
    )
}

#[test]
fn keynote_slide_image_properties_parity() -> TestResult {
    run(
        "keynote_slide_image_properties_parity",
        Tier::B,
        Evidence::Queued,
        |editor, image| {
            editor.set_slide_image_properties(
                SLIDE,
                image,
                DrawableProperties {
                    hyperlink_url: Some("https://example.test/image".to_owned()),
                    locked: Some(true),
                    aspect_ratio_locked: None,
                    accessibility_description: Some("A described image".to_owned()),
                },
            )?;
            Ok(())
        },
        |edit, slide, image| {
            edit.set_properties(
                slide,
                image,
                MediaProperties::from_parts(
                    Some("https://example.test/image".to_owned()),
                    Some(true),
                    None,
                    Some("A described image".to_owned()),
                ),
            )?;
            Ok(())
        },
    )
}

#[test]
fn keynote_slide_image_data_parity() -> TestResult {
    let jpeg = fixture(JPEG)?.bytes;
    run(
        "keynote_slide_image_data_parity",
        Tier::B,
        Evidence::Queued,
        |editor, image| {
            editor.replace_slide_image_data(SLIDE, image, &jpeg)?;
            Ok(())
        },
        |edit, slide, image| {
            edit.set_data(slide, image, &jpeg)?;
            Ok(())
        },
    )
}

/// Host and leaf reads of every image of slide 0 agree: frame, reflection,
/// properties, original and natural size, and the embedded data.
fn images_agree(bytes: &[u8]) -> TestResult<usize> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let package = Package::from_bytes(bytes)?;
    let host = editor.slide_images(SLIDE)?;
    let leaf = package.slide_image_list(SlideSelector::index(SLIDE))?;
    assert_eq!(host.len(), leaf.len());
    for (index, (host, leaf)) in host.iter().zip(&leaf).enumerate() {
        let geometry = leaf.geometry().ok_or("leaf image has no frame")?;
        let position = host.geometry.position.ok_or("host image has no position")?;
        let size = host.geometry.size.ok_or("host image has no size")?;
        assert_eq!(
            (geometry.x, geometry.y, geometry.width, geometry.height),
            (position.x, position.y, size.width, size.height)
        );
        assert_eq!(geometry.rotation, host.geometry.angle);
        assert_eq!(
            leaf.is_reflected(),
            host.geometry.flags.is_some_and(|flags| flags & 4 != 0)
        );
        let properties = leaf.properties();
        assert_eq!(
            properties.hyperlink_url(),
            host.properties.hyperlink_url.as_deref()
        );
        assert_eq!(properties.locked(), host.properties.locked);
        assert_eq!(
            properties.aspect_ratio_locked(),
            host.properties.aspect_ratio_locked
        );
        assert_eq!(
            properties.accessibility_description(),
            host.properties.accessibility_description.as_deref()
        );
        assert_eq!(
            leaf.original_size().map(|size| (size.width, size.height)),
            host.original_size.map(|size| (size.width, size.height))
        );
        assert_eq!(
            leaf.natural_size().map(|size| (size.width, size.height)),
            host.natural_size.map(|size| (size.width, size.height))
        );
        assert_eq!(
            package.slide_image_data(SlideSelector::index(SLIDE), ImageSelector::index(index))?,
            editor.extract_media(host.image_data_identifier)?.as_slice()
        );
    }
    Ok(leaf.len())
}

#[test]
fn keynote_slide_image_read_parity() -> TestResult {
    let input = fixture(IMAGE)?;
    Case::new(
        "keynote_slide_image_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host and leaf images agree", images_agree)
    .run()?;
    Ok(())
}
