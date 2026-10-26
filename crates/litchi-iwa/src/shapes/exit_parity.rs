//! Host-versus-leaf read parity for the drawable and number-attachment codecs
//! ported in record 0368 (`litchi_iwa_protos::{shape_properties_codec,
//! number_attachment_codec}`), on every tracked Pages, Numbers and Keynote
//! fixture: each `TSD.ShapeArchive` (3004) and `TSD.ImageArchive` (3005) drawable
//! and each `TSWP.NumberAttachmentArchive` (2043) the host decodes with Prost
//! must project, through the leaf codec, to the same value of every field the
//! codec exposes. Deleted with the host (W3).

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    clippy::float_cmp,
    reason = "differential parity tests compare exact decoded values"
)]

use std::path::{Path, PathBuf};

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::parse_wire_fields;
use litchi_iwa_protos::number_attachment_codec as leaf_attachment;
use litchi_iwa_protos::shape_properties_codec as leaf_shape;
use prost::Message;

use crate::HostPackage;
use crate::protobuf::{tsd, tswp};

const SHAPE: u32 = 3_004;
const IMAGE: u32 = 3_005;
const NUMBER_ATTACHMENT: u32 = 2_043;

fn fixtures() -> Vec<PathBuf> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork");
    let mut paths = Vec::new();
    for (format, extension) in [
        ("pages", "pages"),
        ("numbers", "numbers"),
        ("keynote", "key"),
    ] {
        let Ok(entries) = std::fs::read_dir(root.join(format)) else {
            continue;
        };
        for entry in entries {
            let path = entry.unwrap().path();
            if path.extension().is_some_and(|value| value == extension) && path.is_file() {
                paths.push(path);
            }
        }
    }
    paths.sort();
    paths
}

fn messages(path: &Path, types: &[u32]) -> Vec<(u32, Vec<u8>)> {
    let Ok(package) = HostPackage::from_bytes(&std::fs::read(path).unwrap()) else {
        return Vec::new();
    };
    let names = package
        .iwa_entry_names()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let mut found = Vec::new();
    for name in names {
        let Ok(archive) = package.archive(&name) else {
            continue;
        };
        for object in &archive.objects {
            for message in &object.messages {
                if types.contains(&message.type_) {
                    found.push((message.type_, message.data.clone()));
                }
            }
        }
    }
    found
}

/// The `super` drawable payload (field 1) of a shape or image archive.
fn drawable_payload(data: &[u8]) -> Option<Vec<u8>> {
    let fields = parse_wire_fields(data, WireLimits::DEFAULT).ok()?;
    let field = fields.iter().find(|field| field.number() == 1)?;
    Some(field.payload(data).ok()?.to_vec())
}

#[test]
fn drawable_and_attachment_reads_equal_the_host_on_every_fixture() {
    let mut drawables = 0usize;
    let mut attachments = 0usize;
    for path in fixtures() {
        let name = path.file_name().unwrap().to_string_lossy().into_owned();
        for (message_type, data) in messages(&path, &[SHAPE, IMAGE, NUMBER_ATTACHMENT]) {
            let context = format!("{name} type {message_type}");
            if message_type == NUMBER_ATTACHMENT {
                let host = tswp::NumberAttachmentArchive::decode(data.as_slice()).unwrap();
                let leaf = leaf_attachment::decode_number_attachment(
                    &data,
                    leaf_attachment::DecodeOptions::for_source(&data),
                )
                .unwrap_or_else(|error| panic!("{context}: {error}"));
                let textual = host.super_.as_ref();
                assert_eq!(
                    leaf.string_equivalent(),
                    textual.and_then(|value| value.string_equivalent.as_deref()),
                    "{context} string equivalent"
                );
                assert_eq!(
                    leaf.kind(),
                    textual.and_then(|value| value.kind),
                    "{context} kind"
                );
                assert_eq!(leaf.number_format(), host.number_format, "{context} format");
                assert_eq!(
                    leaf.string_value(),
                    host.string_value.as_deref(),
                    "{context} value"
                );
                assert_eq!(
                    leaf.number_format_name(),
                    host.number_format_name.as_deref(),
                    "{context} format name"
                );
                attachments += 1;
                continue;
            }
            let Some(payload) = drawable_payload(&data) else {
                continue;
            };
            let host = tsd::DrawableArchive::decode(payload.as_slice()).unwrap();
            let leaf = leaf_shape::decode_drawable_archive(
                &payload,
                leaf_shape::DecodeOptions::for_source(&payload),
            )
            .unwrap_or_else(|error| panic!("{context}: {error}"));
            let geometry = host.geometry.as_ref();
            assert_eq!(
                leaf.geometry().is_some(),
                geometry.is_some(),
                "{context} geometry"
            );
            if let (Some(leaf_geometry), Some(host_geometry)) = (leaf.geometry(), geometry) {
                assert_eq!(
                    leaf_geometry.position().map(|point| (point.x(), point.y())),
                    host_geometry
                        .position
                        .as_ref()
                        .map(|point| (point.x, point.y)),
                    "{context} position"
                );
                assert_eq!(
                    leaf_geometry
                        .size()
                        .map(|size| (size.width(), size.height())),
                    host_geometry
                        .size
                        .as_ref()
                        .map(|size| (size.width, size.height)),
                    "{context} size"
                );
                assert_eq!(
                    leaf_geometry.flags(),
                    host_geometry.flags,
                    "{context} flags"
                );
                assert_eq!(
                    leaf_geometry.angle(),
                    host_geometry.angle,
                    "{context} angle"
                );
            }
            let properties = leaf.properties();
            assert_eq!(
                properties.hyperlink_url(),
                host.hyperlink_url.as_deref(),
                "{context} url"
            );
            assert_eq!(properties.locked(), host.locked, "{context} locked");
            assert_eq!(
                properties.aspect_ratio_locked(),
                host.aspect_ratio_locked,
                "{context} aspect"
            );
            assert_eq!(
                properties.accessibility_description(),
                host.accessibility_description.as_deref(),
                "{context} description"
            );
            drawables += 1;
        }
    }
    eprintln!("drawable parity: {drawables} drawables, {attachments} number attachments");
    assert!(drawables > 0, "{drawables}");
}
