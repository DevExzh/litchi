#![cfg(feature = "internal-iwork-source")]

//! Native Pages integration coverage for image-inspector adjustments.
//!
//! The fixtures are authored, saved, closed, and reopened by Pages 14.4.
//! The test locates the ordinary typed `TSD.ImageArchive` by its native
//! message and image-data edge, then exercises the focused adjustment seam
//! without depending on native drawable identifiers or component names.

use std::{io, path::PathBuf};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{
    WireLimits,
    shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement},
    wire::WireView,
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::tsd;
use litchi_pages::{
    __decode_image_adjustments_payload, __rewrite_image_adjustments_payload, ImageAdjustmentsError,
    Package,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const IMAGE_ADJUSTMENTS_FIELD: u32 = 14;
const ADJUSTMENT_FIELDS: [u32; 3] = [1, 2, 13];
const SHARPNESS_FIELD: u32 = 6;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/image-adjustments-native.pages")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/image-adjustments-native-resaved.pages")
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
    let text = package.text()?;
    if !text.contains("Native Pages image adjustment marker") {
        return Err(io::Error::other("native Pages image adjustment marker is missing").into());
    }
    Ok(text)
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

fn native_source_adjustments() -> ImageAdjustments {
    ImageAdjustments::new()
        // Pages omits the neutral scalar controls in a freshly-created image;
        // their absence is distinct from an explicitly encoded zero.
        .with_exposure(None)
        .with_saturation(None)
        .with_enhancement(Some(ImageEnhancement::Disabled))
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
        __rewrite_image_adjustments_payload(&location.bytes, baseline, WireLimits::default())?,
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
    let changed_payload =
        __rewrite_image_adjustments_payload(&location.bytes, expected, WireLimits::default())?;

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

    let restored_payload =
        __rewrite_image_adjustments_payload(&changed_payload, baseline, WireLimits::default())?;
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
        __rewrite_image_adjustments_payload(&location.bytes, baseline, WireLimits::default())?,
        location.bytes
    );

    let changed = alternate_adjustments(baseline)?;
    let changed_payload =
        __rewrite_image_adjustments_payload(&location.bytes, changed, WireLimits::default())?;
    assert_ne!(changed_payload, location.bytes);
    let changed_package = rewrite_image_payload(&source, &location, &changed_payload)?;
    assert_eq!(
        marker(&Package::from_bytes(&changed_package)?)?,
        source_marker
    );
    assert_image_assets_untouched(&source, &changed_package)?;
    assert_member_locality(&source, &changed_package, &location.member)?;

    let restored_payload =
        __rewrite_image_adjustments_payload(&changed_payload, baseline, WireLimits::default())?;
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
    let error = __rewrite_image_adjustments_payload(
        truncated,
        ImageAdjustments::default(),
        WireLimits::default(),
    )
    .unwrap_err();
    assert!(matches!(error, ImageAdjustmentsError::Codec(_)));
    Ok(())
}
