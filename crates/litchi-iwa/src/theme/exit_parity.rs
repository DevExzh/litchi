//! Host-versus-leaf read parity for the theme codecs ported in record 0368
//! (`litchi_iwa_protos::{theme_codec, keynote_theme_codec}`), on every
//! tracked Pages, Numbers and Keynote fixture.
//!
//! Pages (`TP.ThemeArchive`, 10001) and Numbers (`TN.ThemeArchive`, 12009)
//! wrap `TSS.ThemeArchive`; the host decodes them with
//! [`IWorkThemeArchive`]. Every known field the host reads — the theme
//! identifier, both stylesheet references, the colour presets and the five
//! preset extensions — must equal the leaf's source-preserving view (the
//! leaf extension bytes decode, with the test-only Prost oracle, to the
//! host's typed extension). Keynote's outer `KN.ThemeArchive` (10) is
//! compared field by field with its Prost decode. Deleted with the host (W3).

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "differential parity tests fail loudly on any divergence"
)]

use std::path::{Path, PathBuf};

use litchi_iwa_protos::keynote_theme_codec as leaf_keynote;
use litchi_iwa_protos::theme_codec as leaf;
use prost::Message;

use super::*;
use crate::HostPackage;
use crate::protobuf::{kn, tsp};

const KEYNOTE_THEME: u32 = 10;
const PAGES_THEME: u32 = 10_001;
const NUMBERS_THEME: u32 = 12_009;

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

/// Every theme payload of a fixture, with its message type.
fn themes(path: &Path) -> Vec<(u32, Vec<u8>)> {
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
                if matches!(message.type_, KEYNOTE_THEME | PAGES_THEME | NUMBERS_THEME) {
                    found.push((message.type_, message.data.clone()));
                }
            }
        }
    }
    found
}

fn reference(snapshot: Option<leaf::ThemeReferenceSnapshot>) -> Option<u64> {
    snapshot.map(|reference| reference.identifier)
}

fn host_reference(reference: Option<&tsp::Reference>) -> Option<u64> {
    reference.map(|reference| reference.identifier)
}

fn assert_extension<M: Message + Default + PartialEq + std::fmt::Debug>(
    leaf_bytes: Option<&[u8]>,
    host: Option<&M>,
    context: &str,
) {
    let leaf = leaf_bytes.map(|bytes| M::decode(bytes).expect("leaf extension bytes decode"));
    assert_eq!(leaf.as_ref(), host, "{context}");
}

#[test]
fn theme_reads_equal_the_host_on_every_fixture() {
    let mut wrapped = 0usize;
    let mut keynote = 0usize;
    for path in fixtures() {
        let name = path.file_name().unwrap().to_string_lossy().into_owned();
        for (message_type, data) in themes(&path) {
            let context = format!("{name} theme type {message_type}");
            if message_type == KEYNOTE_THEME {
                let snapshot = leaf_keynote::decode_theme_archive(
                    &data,
                    leaf_keynote::DecodeOptions::for_source(&data),
                )
                .unwrap_or_else(|error| panic!("{context}: {error}"));
                let host = kn::ThemeArchive::decode(data.as_slice()).unwrap();
                assert_eq!(snapshot.uuid(), host.uuid.as_deref(), "{context} uuid");
                assert_eq!(
                    snapshot.template_count(),
                    host.templates.len(),
                    "{context} templates"
                );
                assert_eq!(
                    snapshot
                        .default_template_slide_node()
                        .map(|value| value.identifier),
                    host_reference(host.default_template_slide_node.as_ref()),
                    "{context} default template"
                );
                assert_eq!(
                    snapshot
                        .live_video_source_collection()
                        .map(|value| value.identifier),
                    host_reference(host.live_video_source_collection.as_ref()),
                    "{context} live video"
                );
                assert_eq!(
                    tss::ThemeArchive::decode(snapshot.super_bytes()).unwrap(),
                    host.super_,
                    "{context} super"
                );
                keynote += 1;
                continue;
            }
            let host = IWorkThemeArchive::decode(&data).unwrap();
            let snapshot =
                leaf::decode_theme_archive(&data, leaf::DecodeOptions::for_source(&data))
                    .unwrap_or_else(|error| panic!("{context}: {error}"));
            assert_eq!(
                snapshot.theme_identifier(),
                host.base.theme_identifier.as_deref(),
                "{context} identifier"
            );
            assert_eq!(
                reference(snapshot.legacy_stylesheet_reference()),
                host_reference(host.base.legacy_stylesheet.as_ref()),
                "{context} legacy stylesheet"
            );
            assert_eq!(
                reference(snapshot.document_stylesheet_reference()),
                host_reference(host.base.document_stylesheet.as_ref()),
                "{context} document stylesheet"
            );
            assert_eq!(
                snapshot.color_preset_payloads().count(),
                host.base.color_presets.len(),
                "{context} colour presets"
            );
            assert_eq!(
                snapshot.old_uuid_payloads().count(),
                host.base.old_uuids_for_preset_replacements.len(),
                "{context} old preset UUIDs"
            );
            let extensions = &host.extensions;
            assert_extension(
                snapshot.drawing_presets(),
                extensions.drawing.as_ref(),
                &context,
            );
            assert_extension(snapshot.text_presets(), extensions.text.as_ref(), &context);
            assert_extension(
                snapshot.chart_presets(),
                extensions.chart.as_ref(),
                &context,
            );
            assert_extension(
                snapshot.table_presets(),
                extensions.table.as_ref(),
                &context,
            );
            assert_extension(
                snapshot.application_presets(),
                extensions.application.as_ref(),
                &context,
            );
            wrapped += 1;
        }
    }
    eprintln!("theme parity: {wrapped} Pages/Numbers themes, {keynote} Keynote themes");
    assert!(wrapped > 0 && keynote > 0, "{wrapped} {keynote}");
}
