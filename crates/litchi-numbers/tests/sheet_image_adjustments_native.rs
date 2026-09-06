//! Default-feature native Numbers coverage for semantic image adjustments.
//!
//! The fixture was authored, saved, closed, and reopened by Numbers 14.4.
//! This test uses the public sheet/image selectors while independently
//! checking the native image payload, advanced sharpness field, asset bytes,
//! and physical component locality.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tsd;
use litchi_numbers::cell::Value;
use litchi_numbers::{
    ImageAdjustment, ImageAdjustments, ImageEnhancement, Package, SheetImageAdjustmentsError,
    SheetSelector, shape::image::ImageSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "Table 1";
const MARKER_POSITION: &str = "B2";
const EXPECTED_MARKER: &str = "Native image adjustment marker direct saved";
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const IMAGE_ADJUSTMENTS_FIELD: u32 = 14;
const ADJUSTMENT_FIELDS: [u32; 3] = [1, 2, 13];
const SHARPNESS_FIELD: u32 = 6;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/image-adjustments-direct-resaved.numbers")
}

fn retirement_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/image-adjustments-retirement-resaved.numbers")
}

#[derive(Debug, Clone)]
struct ImagePayload {
    member: String,
    bytes: Vec<u8>,
}

fn image_payload(source: &[u8]) -> TestResult<ImagePayload> {
    let catalog = Catalog::from_bytes(source)?;
    let mut found = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in archive.objects {
            for message in object.messages {
                if message.type_ != IMAGE_MESSAGE_TYPE {
                    continue;
                }
                let image = match tsd::ImageArchive::decode(message.data.as_slice()) {
                    Ok(image) => image,
                    Err(_) => continue,
                };
                if image.data.is_none() || image.image_adjustments.is_none() {
                    continue;
                }
                if found.is_some() {
                    return Err(io::Error::other(
                        "native fixture contains multiple ordinary image ImageArchives",
                    )
                    .into());
                }
                found = Some(ImagePayload {
                    member: entry.name().to_owned(),
                    bytes: message.data,
                });
            }
        }
    }
    found.ok_or_else(|| io::Error::other("native ordinary image ImageArchive is missing").into())
}

fn marker(package: &Package) -> TestResult<String> {
    let table = package
        .table(SHEET_NAME, TABLE_NAME)?
        .ok_or_else(|| io::Error::other("native image marker table is missing"))?;
    match table.get_a1(MARKER_POSITION)? {
        Some(Value::Text(value)) => Ok(value.clone()),
        value => {
            Err(io::Error::other(format!("native image marker is not text: {value:?}")).into())
        },
    }
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_member_locality(source: &[u8], target: &[u8], changed_member: &str) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("native image edit removed a package member"))?;
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged member {} lost its exact local ZIP record",
                entry.name()
            );
        }
    }
    for entry in after.iter() {
        assert!(
            before.iter().any(|other| other.name() == entry.name()),
            "native image edit inserted an unexpected package member {}",
            entry.name()
        );
    }
    changed.sort_unstable();
    assert_eq!(changed, [changed_member.to_owned()]);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_image_assets_untouched(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut assets = 0;
    for entry in before
        .iter()
        .filter(|entry| entry.name().starts_with("Data/"))
    {
        assets += 1;
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("native image asset was removed"))?;
        assert_eq!(entry.data(), candidate.data(), "image asset bytes changed");
        assert_eq!(
            entry.raw_record().local_record(),
            candidate.raw_record().local_record(),
            "image asset ZIP record changed"
        );
    }
    assert!(assets > 0, "native image fixture has no Data assets");
    Ok(())
}

fn unknown_image_wire(source: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let root = WireView::parse(source)?;
    let adjustments = root
        .fields()
        .find(|field| field.number() == IMAGE_ADJUSTMENTS_FIELD)
        .ok_or_else(|| io::Error::other("native image-adjustments field is missing"))?;
    let mut records = root
        .fields()
        .filter(|field| field.number() != IMAGE_ADJUSTMENTS_FIELD)
        .map(|field| field.raw().to_vec())
        .collect::<Vec<_>>();
    records.extend(
        WireView::parse(adjustments.payload())?
            .fields()
            .filter(|field| !ADJUSTMENT_FIELDS.contains(&field.number()))
            .map(|field| field.raw().to_vec()),
    );
    Ok(records)
}

fn sharpness_wire(source: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let root = WireView::parse(source)?;
    let adjustments = root
        .fields()
        .find(|field| field.number() == IMAGE_ADJUSTMENTS_FIELD)
        .ok_or_else(|| io::Error::other("native image-adjustments field is missing"))?;
    Ok(WireView::parse(adjustments.payload())?
        .fields()
        .filter(|field| field.number() == SHARPNESS_FIELD)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn sharpness_value(source: &[u8]) -> TestResult<Option<f32>> {
    let image = tsd::ImageArchive::decode(source)?;
    Ok(image
        .image_adjustments
        .as_ref()
        .and_then(|adjustments| adjustments.sharpness))
}

fn direct_adjustments() -> TestResult<ImageAdjustments> {
    Ok(ImageAdjustments::new()
        .with_exposure(Some(ImageAdjustment::new(0.5)?))
        .with_saturation(Some(ImageAdjustment::new(-0.4)?))
        .with_enhancement(Some(ImageEnhancement::Disabled)))
}

fn replacement_adjustments() -> TestResult<ImageAdjustments> {
    Ok(ImageAdjustments::new()
        .with_exposure(Some(ImageAdjustment::new(0.25)?))
        .with_saturation(Some(ImageAdjustment::new(-0.2)?))
        .with_enhancement(Some(ImageEnhancement::Disabled)))
}

fn retirement_adjustments() -> TestResult<ImageAdjustments> {
    Ok(ImageAdjustments::new()
        .with_exposure(Some(ImageAdjustment::new(-0.2)?))
        .with_saturation(Some(ImageAdjustment::new(0.35)?))
        .with_enhancement(Some(ImageEnhancement::Enabled)))
}

fn reset_adjustments() -> ImageAdjustments {
    ImageAdjustments::new()
}

fn public_image_adjustments(package: &Package) -> TestResult<ImageAdjustments> {
    Ok(package.sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?)
}

fn assert_raw_image_state(source: &[u8], target: &[u8]) -> TestResult {
    let source_location = image_payload(source)?;
    let target_location = image_payload(target)?;
    assert_eq!(target_location.member, source_location.member);
    assert_eq!(
        unknown_image_wire(&target_location.bytes)?,
        unknown_image_wire(&source_location.bytes)?,
        "unknown image fields changed"
    );
    assert_eq!(
        sharpness_wire(&target_location.bytes)?,
        sharpness_wire(&source_location.bytes)?,
        "advanced sharpness wire changed"
    );
    assert_eq!(sharpness_value(&target_location.bytes)?, Some(0.25));
    Ok(())
}

#[test]
fn native_image_adjustments_read_and_noop_preserve_exact_source() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let expected = direct_adjustments()?;
    assert_eq!(public_image_adjustments(&package)?, expected);
    assert_eq!(marker(&package)?, EXPECTED_MARKER);
    let location = image_payload(&source)?;
    assert_eq!(sharpness_value(&location.bytes)?, Some(0.25));
    assert!(!sharpness_wire(&location.bytes)?.is_empty());
    assert_image_assets_untouched(&source, &source)?;
    assert_eq!(exact_bytes(&package)?, source);

    let no_op = package
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert_eq!(public_image_adjustments(no_op.package())?, expected);
    assert_eq!(marker(no_op.package())?, EXPECTED_MARKER);
    Ok(())
}

#[test]
fn native_image_adjustments_set_reset_reopen_and_inverse_exactly() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let baseline = public_image_adjustments(&package)?;
    let expected = direct_adjustments()?;
    assert_eq!(baseline, expected);
    let replacement = replacement_adjustments()?;
    let source_location = image_payload(&source)?;

    let changed = package
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let changed_bytes = exact_bytes(changed.package())?;
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(public_image_adjustments(changed.package())?, replacement);
    assert_eq!(marker(changed.package())?, EXPECTED_MARKER);
    assert_raw_image_state(&source, &changed_bytes)?;
    assert_image_assets_untouched(&source, &changed_bytes)?;
    assert_member_locality(&source, &changed_bytes, &source_location.member)?;

    let reopened = Package::from_bytes(&changed_bytes)?;
    assert_eq!(public_image_adjustments(&reopened)?, replacement);
    assert_eq!(marker(&reopened)?, EXPECTED_MARKER);

    let restored = changed
        .package()
        .apply_sheet_image_adjustments(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(public_image_adjustments(restored.package())?, baseline);

    let forward_again = changed.patch().inverse().inverse();
    assert_eq!(forward_again, changed.patch().clone());
    let reapplied = package.apply_sheet_image_adjustments(&forward_again)?;
    assert_eq!(exact_bytes(reapplied.package())?, changed_bytes);

    let reset = reopened
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(reset_adjustments())?
        .commit()?;
    let reset_bytes = exact_bytes(reset.package())?;
    assert_eq!(
        public_image_adjustments(reset.package())?,
        reset_adjustments()
    );
    assert_eq!(marker(reset.package())?, EXPECTED_MARKER);
    assert_raw_image_state(&changed_bytes, &reset_bytes)?;
    assert_image_assets_untouched(&changed_bytes, &reset_bytes)?;
    assert_member_locality(&changed_bytes, &reset_bytes, &source_location.member)?;

    let reset_reopened = Package::from_bytes(&reset_bytes)?;
    assert_eq!(
        public_image_adjustments(&reset_reopened)?,
        reset_adjustments()
    );
    let reset_restored = reset
        .package()
        .apply_sheet_image_adjustments(&reset.patch().inverse())?;
    assert_eq!(exact_bytes(reset_restored.package())?, changed_bytes);
    assert_eq!(
        public_image_adjustments(reset_restored.package())?,
        replacement
    );
    Ok(())
}

#[test]
fn native_image_adjustments_reject_stale_patch_without_mutating_source() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let replacement = replacement_adjustments()?;
    let changed = package
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let candidate = exact_bytes(changed.package())?;
    let stale = changed
        .package()
        .apply_sheet_image_adjustments(changed.patch())
        .expect_err("a patch must not apply to its own target as its source");
    assert!(matches!(stale, SheetImageAdjustmentsError::PatchConflict));
    assert_eq!(exact_bytes(changed.package())?, candidate);
    assert_eq!(public_image_adjustments(changed.package())?, replacement);
    assert!(matches!(
        ImageAdjustment::new(1.01),
        Err(litchi_numbers::shape::image::Error::AdjustmentOutOfRange)
    ));
    assert!(matches!(
        ImageAdjustment::new(-1.01),
        Err(litchi_numbers::shape::image::Error::AdjustmentOutOfRange)
    ));
    Ok(())
}

#[test]
fn native_retirement_resaved_fixture_reads_profile_and_keeps_exact_noop() -> TestResult {
    let source = std::fs::read(retirement_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let expected = retirement_adjustments()?;
    assert_eq!(public_image_adjustments(&package)?, expected);
    assert_eq!(marker(&package)?, EXPECTED_MARKER);
    let location = image_payload(&source)?;
    assert_eq!(sharpness_value(&location.bytes)?, Some(0.25));
    assert!(!sharpness_wire(&location.bytes)?.is_empty());
    assert_image_assets_untouched(&source, &source)?;
    assert_eq!(exact_bytes(&package)?, source);

    let no_op = package
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert_eq!(public_image_adjustments(no_op.package())?, expected);
    assert_eq!(marker(no_op.package())?, EXPECTED_MARKER);
    Ok(())
}
