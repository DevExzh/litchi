#![cfg(feature = "internal-iwork-source")]

//! Native Numbers integration coverage for image-inspector adjustments.
//!
//! The fixtures are authored, saved, closed, and reopened by Numbers 14.4.
//! The test locates the ordinary typed `TSD.ImageArchive` by its native
//! message and image-data edge, then exercises the focused adjustment seam
//! without depending on native drawable identifiers or component names.

use std::{io, path::PathBuf};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{image_adjustments_codec as codec, tsd};
use litchi_numbers::cell::Value;
use litchi_numbers::{
    __decode_image_adjustments_payload, ImageAdjustmentsError, Package, SheetImageAdjustmentsError,
    SheetSelector,
    shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement, ImageSelector},
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "Table 1";
const MARKER_POSITION: &str = "B2";
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const IMAGE_ADJUSTMENTS_FIELD: u32 = 14;
const ADJUSTMENT_FIELDS: [u32; 3] = [1, 2, 13];
const SHARPNESS_FIELD: u32 = 6;
const DIRECT_RESAVED_MARKER: &str = "Native image adjustment marker direct saved";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/image-adjustments-native.numbers")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/image-adjustments-native-resaved.numbers")
}

fn direct_resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/image-adjustments-direct-resaved.numbers")
}

#[derive(Debug, Clone)]
struct ImagePayload {
    member: String,
    object_id: u64,
    message_index: usize,
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
            let object_id = object.archive_info.identifier;
            for (message_index, message) in object.messages.into_iter().enumerate() {
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
                    object_id: object_id
                        .ok_or_else(|| io::Error::other("native image object has no identifier"))?,
                    message_index,
                    bytes: message.data,
                });
            }
        }
    }
    found.ok_or_else(|| io::Error::other("native ordinary image ImageArchive is missing").into())
}

fn rewrite_image_payload(
    source: &[u8],
    location: &ImagePayload,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == location.member)
        .ok_or_else(|| io::Error::other("native image component is missing"))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .object_mut(location.object_id)
        .ok_or_else(|| io::Error::other("native image ImageArchive object is missing"))?;
    let message_type = object
        .messages
        .get(location.message_index)
        .map(|message| message.type_)
        .ok_or_else(|| io::Error::other("native image ImageArchive message is missing"))?;
    if message_type != IMAGE_MESSAGE_TYPE {
        return Err(io::Error::other("native image message type changed").into());
    }
    object.replace_message_preserving_header(
        location.message_index,
        RawMessage {
            type_: IMAGE_MESSAGE_TYPE,
            data: replacement.to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&location.member, &compressed)],
        Limits::default(),
    )?)
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

fn native_source_adjustments() -> ImageAdjustments {
    ImageAdjustments::new().with_enhancement(Some(ImageEnhancement::Disabled))
}

fn native_resaved_adjustments() -> TestResult<ImageAdjustments> {
    Ok(ImageAdjustments::new()
        .with_exposure(Some(ImageAdjustment::new(0.25)?))
        .with_saturation(Some(ImageAdjustment::new(-0.2)?))
        .with_enhancement(Some(ImageEnhancement::Disabled)))
}

fn replacement_adjustments(baseline: ImageAdjustments) -> TestResult<ImageAdjustments> {
    Ok(baseline
        .with_exposure(Some(ImageAdjustment::new(0.25)?))
        .with_saturation(Some(ImageAdjustment::new(-0.2)?)))
}

fn alternate_adjustments(baseline: ImageAdjustments) -> TestResult<ImageAdjustments> {
    Ok(baseline
        .with_exposure(Some(ImageAdjustment::new(0.5)?))
        .with_saturation(Some(ImageAdjustment::new(-0.4)?)))
}

fn public_image_adjustments(package: &Package) -> TestResult<ImageAdjustments> {
    Ok(package.sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?)
}

fn image_adjustments_codec_options(source: &[u8]) -> codec::DecodeOptions {
    let limits = WireLimits::default();
    codec::DecodeOptions::new(
        limits.max_input_bytes().min(source.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_output_bytes(limits.max_output_bytes())
}

fn image_adjustments_write(adjustments: ImageAdjustments) -> codec::ImageAdjustmentsWrite {
    codec::ImageAdjustmentsWrite::from_values(
        adjustments.exposure().map(ImageAdjustment::value),
        adjustments.saturation().map(ImageAdjustment::value),
        adjustments
            .enhancement()
            .map(|value| matches!(value, ImageEnhancement::Enabled)),
    )
}

fn rewrite_image_adjustments_payload(
    source: &[u8],
    adjustments: ImageAdjustments,
) -> Result<Vec<u8>, ImageAdjustmentsError> {
    codec::rewrite_image_adjustments(
        source,
        image_adjustments_write(adjustments),
        image_adjustments_codec_options(source),
    )
    .map_err(ImageAdjustmentsError::Codec)
}

#[test]
fn native_image_source_reads_typed_adjustments_and_has_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(exact_bytes(&package)?, source);

    let location = image_payload(&source)?;
    let baseline = __decode_image_adjustments_payload(&location.bytes, WireLimits::default())?;
    assert_eq!(baseline, native_source_adjustments());
    assert!(!sharpness_wire(&location.bytes)?.is_empty());
    assert_eq!(
        rewrite_image_adjustments_payload(&location.bytes, baseline)?,
        location.bytes
    );
    Ok(())
}

#[test]
fn native_image_source_rewrite_preserves_advanced_wire_marker_locality_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let source_marker = marker(&Package::from_bytes(&source)?)?;
    let location = image_payload(&source)?;
    let baseline = __decode_image_adjustments_payload(&location.bytes, WireLimits::default())?;
    assert_eq!(baseline, native_source_adjustments());
    let expected = replacement_adjustments(baseline)?;
    let changed_payload = rewrite_image_adjustments_payload(&location.bytes, expected)?;

    assert_ne!(changed_payload, location.bytes);
    assert_eq!(
        __decode_image_adjustments_payload(&changed_payload, WireLimits::default())?,
        expected
    );
    assert_eq!(
        unknown_image_wire(&changed_payload)?,
        unknown_image_wire(&location.bytes)?
    );
    assert_eq!(
        sharpness_wire(&changed_payload)?,
        sharpness_wire(&location.bytes)?
    );

    let changed_package = rewrite_image_payload(&source, &location, &changed_payload)?;
    let changed_location = image_payload(&changed_package)?;
    assert_eq!(changed_location.bytes, changed_payload);
    assert_eq!(
        marker(&Package::from_bytes(&changed_package)?)?,
        source_marker
    );
    assert_image_assets_untouched(&source, &changed_package)?;
    assert_member_locality(&source, &changed_package, &location.member)?;

    let restored_payload = rewrite_image_adjustments_payload(&changed_payload, baseline)?;
    assert_eq!(restored_payload, location.bytes);
    let restored_package =
        rewrite_image_payload(&changed_package, &changed_location, &restored_payload)?;
    assert_eq!(image_payload(&restored_package)?.bytes, location.bytes);
    assert_eq!(
        marker(&Package::from_bytes(&restored_package)?)?,
        source_marker
    );
    assert_image_assets_untouched(&source, &restored_package)?;
    Ok(())
}

#[test]
fn native_image_resaved_fixture_reads_changed_adjustments_and_roundtrips() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(exact_bytes(&package)?, source);
    let source_marker = marker(&package)?;

    let location = image_payload(&source)?;
    let baseline = __decode_image_adjustments_payload(&location.bytes, WireLimits::default())?;
    assert_eq!(baseline, native_resaved_adjustments()?);
    assert!(!sharpness_wire(&location.bytes)?.is_empty());
    assert_eq!(
        rewrite_image_adjustments_payload(&location.bytes, baseline)?,
        location.bytes
    );

    let changed = alternate_adjustments(baseline)?;
    let changed_payload = rewrite_image_adjustments_payload(&location.bytes, changed)?;
    assert_ne!(changed_payload, location.bytes);
    let changed_package = rewrite_image_payload(&source, &location, &changed_payload)?;
    assert_eq!(
        marker(&Package::from_bytes(&changed_package)?)?,
        source_marker
    );
    assert_image_assets_untouched(&source, &changed_package)?;
    assert_member_locality(&source, &changed_package, &location.member)?;

    let restored_payload = rewrite_image_adjustments_payload(&changed_payload, baseline)?;
    assert_eq!(restored_payload, location.bytes);
    let restored_package = rewrite_image_payload(
        &changed_package,
        &image_payload(&changed_package)?,
        &restored_payload,
    )?;
    assert_eq!(image_payload(&restored_package)?.bytes, location.bytes);
    assert_eq!(
        marker(&Package::from_bytes(&restored_package)?)?,
        source_marker
    );
    assert_image_assets_untouched(&source, &restored_package)?;
    Ok(())
}

#[test]
fn native_image_adjustments_reject_truncated_payload_with_typed_errors() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let location = image_payload(&source)?;
    let truncated = &location.bytes[..location.bytes.len().saturating_sub(1)];
    let error = __decode_image_adjustments_payload(truncated, WireLimits::default()).unwrap_err();
    assert!(matches!(error, ImageAdjustmentsError::Codec(_)));
    let error =
        rewrite_image_adjustments_payload(truncated, ImageAdjustments::default()).unwrap_err();
    assert!(matches!(error, ImageAdjustmentsError::Codec(_)));
    Ok(())
}

#[test]
fn native_image_public_api_reads_exactly_and_keeps_noop_bytes() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let baseline = native_source_adjustments();

    assert_eq!(public_image_adjustments(&package)?, baseline);
    let no_op = package
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(baseline)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_eq!(public_image_adjustments(no_op.package())?, baseline);
    assert_eq!(marker(no_op.package())?, marker(&package)?);
    Ok(())
}

#[test]
fn native_image_public_api_replaces_reopens_preserves_and_inverts_exactly() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let baseline = public_image_adjustments(&package)?;
    assert_eq!(baseline, native_source_adjustments());
    let replacement = alternate_adjustments(baseline)?;
    assert_eq!(replacement.enhancement(), baseline.enhancement());

    let changed = package
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(replacement)?
        .commit()?;
    let candidate = exact_bytes(changed.package())?;
    let source_location = image_payload(&source)?;
    let candidate_location = image_payload(&candidate)?;

    assert_eq!(public_image_adjustments(changed.package())?, replacement);
    assert_eq!(
        __decode_image_adjustments_payload(&candidate_location.bytes, WireLimits::default())?,
        replacement
    );
    assert_eq!(
        unknown_image_wire(&candidate_location.bytes)?,
        unknown_image_wire(&source_location.bytes)?
    );
    assert_eq!(
        sharpness_wire(&candidate_location.bytes)?,
        sharpness_wire(&source_location.bytes)?
    );
    assert_eq!(marker(changed.package())?, marker(&package)?);
    assert_image_assets_untouched(&source, &candidate)?;
    assert_member_locality(&source, &candidate, &source_location.member)?;
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());

    let reopened = Package::from_bytes(&candidate)?;
    assert_eq!(public_image_adjustments(&reopened)?, replacement);
    let restored = changed
        .package()
        .apply_sheet_image_adjustments(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(public_image_adjustments(restored.package())?, baseline);
    Ok(())
}

#[test]
fn native_image_public_api_rejects_stale_selection_and_out_of_range_values() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let replacement = alternate_adjustments(public_image_adjustments(&package)?)?;
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
fn native_image_public_api_reads_the_directly_resaved_fixture_and_keeps_noop() -> TestResult {
    let source = std::fs::read(direct_resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let expected = alternate_adjustments(native_source_adjustments())?;
    assert_eq!(public_image_adjustments(&package)?, expected);
    let location = image_payload(&source)?;
    assert_eq!(
        __decode_image_adjustments_payload(&location.bytes, WireLimits::default())?,
        expected
    );
    assert_eq!(marker(&package)?, DIRECT_RESAVED_MARKER);
    assert_eq!(sharpness_value(&location.bytes)?, Some(0.25));
    assert_image_assets_untouched(&source, &source)?;

    let no_op = package
        .edit_sheet_image_adjustments(SheetSelector::index(0), ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert_eq!(public_image_adjustments(no_op.package())?, expected);
    assert_eq!(marker(no_op.package())?, DIRECT_RESAVED_MARKER);
    assert_eq!(sharpness_value(&image_payload(&source)?.bytes)?, Some(0.25));
    Ok(())
}
