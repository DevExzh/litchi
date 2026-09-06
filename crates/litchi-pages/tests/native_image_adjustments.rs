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
    WireLimits, decode_varint_from_bytes,
    shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement},
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{tsd, tswp};
use litchi_pages::{
    __decode_image_adjustments_payload, __rewrite_image_adjustments_payload,
    BodyImageAdjustmentsError, ImageAdjustmentsError, ImageSelector, Package, PackageError,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const STORAGE_MESSAGE_TYPES: [u32; 2] = [2_001, 2_022];
const IMAGE_ADJUSTMENTS_FIELD: u32 = 14;
const IMAGE_DATA_FIELD: u32 = 11;
const BODY_ATTACHMENTS_FIELD: u32 = 9;
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

fn direct_resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/image-adjustments-direct-resaved.pages")
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

#[derive(Debug, Clone)]
struct MessageLocation {
    member: String,
    object_id: u64,
    message_index: usize,
    message_type: u32,
    bytes: Vec<u8>,
}

fn rewrite_message_payload(
    source: &[u8],
    location: &MessageLocation,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == location.member)
        .ok_or_else(|| io::Error::other("native body component is missing"))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .object_mut(location.object_id)
        .ok_or_else(|| io::Error::other("native body object is missing"))?;
    let message_type = object
        .messages
        .get(location.message_index)
        .map(|message| message.type_)
        .ok_or_else(|| io::Error::other("native body message is missing"))?;
    if message_type != location.message_type {
        return Err(io::Error::other("native body message type changed").into());
    }
    object.replace_message_preserving_header(
        location.message_index,
        RawMessage {
            type_: location.message_type,
            data: replacement.to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&location.member, &compressed)],
        Limits::default(),
    )?)
}

fn image_attachment_location(source: &[u8], image_identifier: u64) -> TestResult<MessageLocation> {
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
        for object in &archive.objects {
            let Some(object_id) = object.archive_info.identifier else {
                continue;
            };
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != ATTACHMENT_MESSAGE_TYPE {
                    continue;
                }
                let attachment =
                    match tswp::DrawableAttachmentArchive::decode(message.data.as_slice()) {
                        Ok(attachment) => attachment,
                        Err(_) => continue,
                    };
                if attachment
                    .drawable
                    .as_ref()
                    .is_none_or(|reference| reference.identifier != image_identifier)
                {
                    continue;
                }
                if found.is_some() {
                    return Err(
                        io::Error::other("native image has multiple drawable attachments").into(),
                    );
                }
                found = Some(MessageLocation {
                    member: entry.name().to_owned(),
                    object_id,
                    message_index,
                    message_type: message.type_,
                    bytes: message.data.clone(),
                });
            }
        }
    }
    found.ok_or_else(|| io::Error::other("native image drawable attachment is missing").into())
}

fn body_storage_location(source: &[u8], attachment_identifier: u64) -> TestResult<MessageLocation> {
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
        for object in &archive.objects {
            let Some(object_id) = object.archive_info.identifier else {
                continue;
            };
            for (message_index, message) in object.messages.iter().enumerate() {
                if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
                    continue;
                }
                let storage = match tswp::StorageArchive::decode(message.data.as_slice()) {
                    Ok(storage) => storage,
                    Err(_) => continue,
                };
                let Some(table) = storage.table_attachment.as_ref() else {
                    continue;
                };
                if !table.entries.iter().any(|entry| {
                    entry
                        .object
                        .as_ref()
                        .is_some_and(|reference| reference.identifier == attachment_identifier)
                }) {
                    continue;
                }
                if found.is_some() {
                    return Err(io::Error::other(
                        "native image attachment has multiple body storage owners",
                    )
                    .into());
                }
                found = Some(MessageLocation {
                    member: entry.name().to_owned(),
                    object_id,
                    message_index,
                    message_type: message.type_,
                    bytes: message.data.clone(),
                });
            }
        }
    }
    found.ok_or_else(|| io::Error::other("native image body storage is missing").into())
}

fn replace_length_delimited_field(
    source: &[u8],
    number: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::with_capacity(source.len().saturating_add(replacement.len()));
    let mut replaced = false;
    for field in view.fields() {
        if field.number() == number {
            if replaced {
                return Err(
                    io::Error::other("native image fixture has duplicate mutation field").into(),
                );
            }
            append_length_delimited_field(&mut output, number, replacement)?;
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(io::Error::other("native image mutation field is missing").into());
    }
    Ok(output)
}

fn with_data_reference_extension(source: &[u8]) -> TestResult<Vec<u8>> {
    let location = image_payload(source)?;
    let image = WireView::parse(&location.bytes)?;
    let data = image
        .fields()
        .find(|field| field.number() == IMAGE_DATA_FIELD)
        .ok_or_else(|| io::Error::other("native image data reference is missing"))?;
    let mut data_payload = data.payload().to_vec();
    append_varint_field(&mut data_payload, 90, 0xfeed_beef)?;
    let replacement =
        replace_length_delimited_field(&location.bytes, IMAGE_DATA_FIELD, &data_payload)?;
    rewrite_image_payload(source, &location, &replacement)
}

fn with_wrong_image_parent(source: &[u8]) -> TestResult<Vec<u8>> {
    let location = image_payload(source)?;
    let image = WireView::parse(&location.bytes)?;
    let drawable = image
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("native image drawable envelope is missing"))?;
    let mut wrong_parent = Vec::new();
    append_varint_field(&mut wrong_parent, 1, 0xfeed_face)?;
    let drawable_payload = replace_length_delimited_field(drawable.payload(), 2, &wrong_parent)?;
    let replacement = replace_length_delimited_field(&location.bytes, 1, &drawable_payload)?;
    rewrite_image_payload(source, &location, &replacement)
}

fn with_duplicate_body_attachment(source: &[u8]) -> TestResult<Vec<u8>> {
    let image = image_payload(source)?;
    let attachment = image_attachment_location(source, image.object_id)?;
    let attachment_identifier = attachment.object_id;
    let body = body_storage_location(source, attachment_identifier)?;
    let body_view = WireView::parse(&body.bytes)?;
    let table = body_view
        .fields()
        .find(|field| field.number() == BODY_ATTACHMENTS_FIELD)
        .ok_or_else(|| io::Error::other("native body attachment table is missing"))?;
    let table_view = WireView::parse(table.payload())?;
    let mut table_payload = Vec::with_capacity(table.payload().len());
    let mut duplicated = false;
    for field in table_view.fields() {
        table_payload.extend_from_slice(field.raw());
        if field.number() != 1 || duplicated {
            continue;
        }
        let entry = WireView::parse(field.payload())?;
        let object = entry
            .fields()
            .find(|entry_field| entry_field.number() == 2)
            .ok_or_else(|| io::Error::other("native body attachment object is missing"))?;
        let reference = WireView::parse(object.payload())?;
        let identifier = reference
            .fields()
            .find(|reference_field| reference_field.number() == 1)
            .ok_or_else(|| io::Error::other("native body attachment identifier is missing"))?;
        let identifier = decode_varint_from_bytes(identifier.payload())?.0;
        if identifier == attachment_identifier {
            table_payload.extend_from_slice(field.raw());
            duplicated = true;
        }
    }
    if !duplicated {
        return Err(io::Error::other("selected image attachment entry is missing").into());
    }
    let body_payload =
        replace_length_delimited_field(&body.bytes, BODY_ATTACHMENTS_FIELD, &table_payload)?;
    rewrite_message_payload(source, &body, &body_payload)
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
fn native_image_public_body_read_projects_source_and_resaved_semantics() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_image_adjustments(ImageSelector::index(0))?,
        native_source_adjustments()
    );

    let resaved = std::fs::read(resaved_fixture_path())?;
    let package = Package::from_bytes(&resaved)?;
    assert_eq!(
        package.body_image_adjustments(ImageSelector::index(0))?,
        native_resaved_adjustments()?
    );
    Ok(())
}

#[test]
fn native_image_public_body_edit_preserves_graph_and_applies_exact_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let location = image_payload(&source)?;
    let baseline = package.body_image_adjustments(ImageSelector::index(0))?;
    let expected = alternate_adjustments(baseline)?;

    let changed = package
        .edit_body_image_adjustments(ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    assert!(changed.diagnostics().changed());
    assert!(!changed.patch().is_noop());
    assert_eq!(
        changed
            .package()
            .body_image_adjustments(ImageSelector::index(0))?,
        expected
    );

    let changed_bytes = exact_bytes(changed.package())?;
    assert_ne!(changed_bytes, source);
    let reopened = Package::from_bytes(&changed_bytes)?;
    assert_eq!(
        reopened.body_image_adjustments(ImageSelector::index(0))?,
        expected
    );
    assert_eq!(marker(&reopened)?, marker(&package)?);
    assert_image_assets_untouched(&source, &changed_bytes)?;
    assert_member_locality(&source, &changed_bytes, &location.member)?;
    assert_eq!(
        unknown_image_wire(&image_payload(&changed_bytes)?.bytes)?,
        unknown_image_wire(&location.bytes)?
    );
    assert_eq!(
        sharpness_wire(&image_payload(&changed_bytes)?.bytes)?,
        sharpness_wire(&location.bytes)?
    );

    let restored = changed
        .package()
        .apply_body_image_adjustments(&changed.patch().inverse())?;
    assert!(restored.diagnostics().changed());
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored
            .package()
            .body_image_adjustments(ImageSelector::index(0))?,
        baseline
    );
    Ok(())
}

#[test]
fn native_image_public_body_noop_and_conflicts_are_atomic() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let baseline = package.body_image_adjustments(ImageSelector::index(0))?;
    let noop = package
        .edit_body_image_adjustments(ImageSelector::index(0))?
        .set(baseline)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(exact_bytes(noop.package())?, source);

    let expected = alternate_adjustments(baseline)?;
    let changed = package
        .edit_body_image_adjustments(ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    let changed_bytes = exact_bytes(changed.package())?;
    let stale = changed
        .package()
        .apply_body_image_adjustments(changed.patch());
    assert!(matches!(
        stale,
        Err(BodyImageAdjustmentsError::PatchConflict)
    ));
    assert_eq!(exact_bytes(changed.package())?, changed_bytes);

    assert!(ImageAdjustment::new(1.01).is_err());
    assert!(ImageAdjustment::new(f32::NAN).is_err());
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn native_image_public_edit_preserves_unknown_data_reference_extensions() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let extended = with_data_reference_extension(&source)?;
    let package = Package::from_bytes(&extended)?;
    let location = image_payload(&extended)?;
    let baseline = package.body_image_adjustments(ImageSelector::index(0))?;
    let before_wire = unknown_image_wire(&location.bytes)?;
    let expected = alternate_adjustments(baseline)?;
    let changed = package
        .edit_body_image_adjustments(ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    let changed_bytes = exact_bytes(changed.package())?;

    assert_eq!(
        changed
            .package()
            .body_image_adjustments(ImageSelector::index(0))?,
        expected
    );
    assert_eq!(
        unknown_image_wire(&image_payload(&changed_bytes)?.bytes)?,
        before_wire
    );
    assert_member_locality(&extended, &changed_bytes, &location.member)?;
    assert_image_assets_untouched(&extended, &changed_bytes)?;
    Ok(())
}

#[test]
fn native_image_direct_resaved_fixture_reads_and_exactly_noops() -> TestResult {
    let source = std::fs::read(direct_resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let location = image_payload(&source)?;
    let expected = ImageAdjustments::new()
        .with_exposure(Some(ImageAdjustment::new(0.5)?))
        .with_saturation(Some(ImageAdjustment::new(-0.4)?))
        .with_enhancement(Some(ImageEnhancement::Disabled));

    assert_eq!(
        package.body_image_adjustments(ImageSelector::index(0))?,
        expected
    );
    assert_eq!(
        package.text()?,
        "Direct saved Native Pages image adjustment marker\u{fffc}"
    );
    assert_eq!(
        sharpness_wire(&location.bytes)?,
        vec![vec![0x35, 0x00, 0x00, 0x80, 0x3e]]
    );

    let noop = package
        .edit_body_image_adjustments(ImageSelector::index(0))?
        .set(expected)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    let noop_bytes = exact_bytes(noop.package())?;
    assert_eq!(noop_bytes, source);
    assert_eq!(
        Package::from_bytes(&noop_bytes)?.body_image_adjustments(ImageSelector::index(0))?,
        expected
    );
    assert_image_assets_untouched(&source, &noop_bytes)?;
    Ok(())
}

#[test]
fn native_image_public_ingress_rejects_duplicate_body_attachment_atomically() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let malformed = with_duplicate_body_attachment(&source)?;
    assert!(matches!(
        Package::from_bytes(&malformed),
        Err(PackageError::InvalidFormat(_))
    ));
    assert_eq!(exact_bytes(&Package::from_bytes(&source)?)?, source);
    Ok(())
}

#[test]
fn native_image_public_edit_rejects_wrong_parent_atomically() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let malformed = with_wrong_image_parent(&source)?;
    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.edit_body_image_adjustments(ImageSelector::index(0)),
        Err(BodyImageAdjustmentsError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
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
