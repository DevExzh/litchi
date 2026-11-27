//! Large single-selector PPTX shape sets stay atomic across native chunks.
//!
//! These are selector fanout tests.  They deliberately do not measure
//! performance; they exercise one selector against many shapes and validate
//! the resulting package through the native scene reader.
#![cfg(feature = "pptx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_ops::node::Value;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, OpenOptions, SourceMode, props};

const LARGE_SHAPE_COUNT: usize = 1_000;
const NATIVE_SCENE_SHAPE_COUNT: usize = 10_000;
const CHUNK_BOUNDARY: usize = 256;
const SELECTOR: &str = "/slide[1]/*";

fn fixture(shape_count: usize) -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    let slide = builder.presentation().unwrap().add_slide().unwrap();
    for index in 0..shape_count {
        let text = format!("original-{index}");
        slide.add_text_box(
            &text,
            i64::try_from(index).unwrap().saturating_mul(100),
            0,
            914_400,
            914_400,
        );
    }
    builder.build().unwrap().to_bytes().unwrap()
}

fn open(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}

fn shape_path(index: usize) -> String {
    format!("/slide[1]/shape[{index}]")
}

fn text(file: &File, index: usize) -> String {
    let path = shape_path(index);
    file.get(&path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}

fn length(file: &File, index: usize, key: &str) -> i64 {
    let path = shape_path(index);
    let node = file.get(&path).unwrap();
    let property = node
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .unwrap();
    let Value::Length(value) = property.value() else {
        panic!("expected a length property")
    };
    value.emu()
}

fn native_shape_values(bytes: &[u8]) -> Vec<(Option<String>, Option<(i64, i64, i64, i64)>)> {
    litchi_pptx::Package::from_bytes(bytes)
        .unwrap()
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .unwrap()
        .shapes()
        .unwrap()
        .iter()
        .map(|shape| {
            let bounds = shape
                .bounds()
                .map(|value| (value.x(), value.y(), value.width(), value.height()));
            (shape.text().map(ToOwned::to_owned), bounds)
        })
        .collect()
}

fn rewrite_final_transform(bytes: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut members = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let mut member = archive.by_index(index).unwrap();
        let name = member.name().expect("valid ZIP member name").into_owned();
        let mut payload = Vec::new();
        member.read_to_end(&mut payload).unwrap();
        if name == "ppt/slides/slide1.xml" {
            let mut xml = String::from_utf8(payload).unwrap();
            let start = xml
                .rfind("<a:xfrm>")
                .expect("the last shape has a local transform");
            xml.replace_range(
                start..start + "<a:xfrm>".len(),
                "<a:xfrm rot=\"private-invalid\">",
            );
            payload = xml.into_bytes();
        }
        members.push((name, payload));
    }
    drop(archive);

    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, payload) in members {
        writer.start_file(name, options).unwrap();
        writer.write_all(&payload).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

#[test]
fn one_selector_text_fanout_crosses_native_chunk_boundary_for_both_owners() {
    let bytes = fixture(LARGE_SHAPE_COUNT);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let updated = source
            .set(SELECTOR, props! { "text" => "fanout text" })
            .unwrap();

        for index in [1, CHUNK_BOUNDARY, CHUNK_BOUNDARY + 1, 512, 513, 1_000] {
            assert_eq!(text(&updated, index), "fanout text", "shape {index}");
        }
        assert_eq!(source.to_bytes().unwrap(), before);

        let updated_bytes = updated.to_bytes().unwrap();
        let noop = updated
            .set(SELECTOR, props! { "text" => "fanout text" })
            .unwrap();
        assert_eq!(noop.to_bytes().unwrap(), updated_bytes);

        let commit = source
            .edit(|edit| edit.set(SELECTOR, props! { "text" => "transaction text" }))
            .unwrap();
        assert_eq!(text(commit.file(), 1), "transaction text");
        assert_eq!(text(commit.file(), LARGE_SHAPE_COUNT), "transaction text");
        assert_eq!(
            commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
            commit.file().to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            before
        );
    }
}

#[test]
fn one_selector_mixed_text_and_geometry_fanout_crosses_native_chunk_boundary() {
    let bytes = fixture(LARGE_SHAPE_COUNT);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let properties = props! {
            "text" => "mixed fanout",
            "x" => "1in",
            "y" => "2in",
            "width" => "3in",
            "height" => "4in",
        };
        let updated = source.set(SELECTOR, properties).unwrap();

        for index in [1, CHUNK_BOUNDARY, CHUNK_BOUNDARY + 1, 512, 513, 1_000] {
            assert_eq!(text(&updated, index), "mixed fanout", "shape {index}");
            assert_eq!(length(&updated, index, "x"), 914_400, "shape {index}");
            assert_eq!(length(&updated, index, "y"), 1_828_800, "shape {index}");
            assert_eq!(length(&updated, index, "width"), 2_743_200, "shape {index}");
            assert_eq!(
                length(&updated, index, "height"),
                3_657_600,
                "shape {index}"
            );
        }
        assert_eq!(source.to_bytes().unwrap(), before);

        let updated_bytes = updated.to_bytes().unwrap();
        let noop = updated
            .set(
                SELECTOR,
                props! {
                    "text" => "mixed fanout",
                    "x" => "1in",
                    "y" => "2in",
                    "width" => "3in",
                    "height" => "4in",
                },
            )
            .unwrap();
        assert_eq!(noop.to_bytes().unwrap(), updated_bytes);

        let commit = source
            .edit(|edit| {
                edit.set(
                    SELECTOR,
                    props! {
                        "text" => "transaction mixed",
                        "x" => "5in",
                        "y" => "6in",
                        "width" => "7in",
                        "height" => "8in",
                    },
                )
            })
            .unwrap();
        assert_eq!(text(commit.file(), LARGE_SHAPE_COUNT), "transaction mixed");
        assert_eq!(length(commit.file(), LARGE_SHAPE_COUNT, "x"), 4_572_000);
        assert_eq!(
            length(commit.file(), LARGE_SHAPE_COUNT, "height"),
            7_315_200
        );
        assert_eq!(
            commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
            commit.file().to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            before
        );
    }
}

#[test]
fn ten_thousand_selector_mixed_updates_are_verified_by_one_native_scene_per_owner() {
    let bytes = fixture(NATIVE_SCENE_SHAPE_COUNT);
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let updated = source
            .set(
                SELECTOR,
                props! {
                    "text" => "ten-thousand mixed",
                    "x" => "1in",
                    "y" => "2in",
                    "width" => "3in",
                    "height" => "4in",
                },
            )
            .unwrap();

        let values = native_shape_values(&updated.to_bytes().unwrap());
        assert_eq!(values.len(), NATIVE_SCENE_SHAPE_COUNT);
        for (index, (shape_text, bounds)) in values.iter().enumerate() {
            assert_eq!(
                shape_text.as_deref(),
                Some("ten-thousand mixed"),
                "shape {index}"
            );
            assert_eq!(
                *bounds,
                Some((914_400, 1_828_800, 2_743_200, 3_657_600)),
                "shape {index}"
            );
        }
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn selector_fanout_splits_aggregate_text_cap_for_text_and_mixed_updates() {
    let text_bytes = litchi_pptx::opened::Limits::default()
        .max_text_bytes()
        .checked_div(3)
        .unwrap()
        .saturating_add(1);
    let large_text = "x".repeat(text_bytes);

    for mixed in [false, true] {
        for eager in [false, true] {
            let source = open(&fixture(3), eager);
            let before = source.to_bytes().unwrap();
            let updated = if mixed {
                source
                    .set(
                        SELECTOR,
                        props! {
                            "text" => large_text.as_str(),
                            "x" => "1in",
                            "y" => "2in",
                            "width" => "3in",
                            "height" => "4in",
                        },
                    )
                    .unwrap()
            } else {
                source
                    .set(SELECTOR, props! { "text" => large_text.as_str() })
                    .unwrap()
            };

            let updated_bytes = updated.to_bytes().unwrap_or_else(|error| {
                panic!(
                    "large selector output serialization failed (mixed={mixed}, eager={eager}): {error}; source={:?}", std::error::Error::source(&error)
                )
            });
            let values = native_shape_values(&updated_bytes);
            assert_eq!(values.len(), 3);
            for (index, (shape_text, bounds)) in values.iter().enumerate() {
                assert_eq!(
                    shape_text.as_deref(),
                    Some(large_text.as_str()),
                    "shape {index}"
                );
                if mixed {
                    assert_eq!(
                        *bounds,
                        Some((914_400, 1_828_800, 2_743_200, 3_657_600)),
                        "shape {index}"
                    );
                }
            }
            assert_eq!(source.to_bytes().unwrap(), before);
        }
    }
}

#[test]
fn selector_fanout_keeps_target_and_staged_byte_limits_atomic() {
    let bytes = fixture(LARGE_SHAPE_COUNT);
    let target_limit = SemanticLimits::builder()
        .max_expanded_targets(LARGE_SHAPE_COUNT - 1)
        .unwrap()
        .build()
        .unwrap();
    let staged_limit = SemanticLimits::builder()
        .max_staged_bytes(1)
        .unwrap()
        .build()
        .unwrap();

    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let target_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(target_limit),
                |edit| edit.set(SELECTOR, props! { "text" => "over target limit" }),
            )
            .unwrap_err();
        assert_eq!(target_error.code(), Code::LimitExceeded);
        assert_eq!(source.to_bytes().unwrap(), before);

        let staged_error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(staged_limit),
                |edit| edit.set(SELECTOR, props! { "text" => "over staged limit" }),
            )
            .unwrap_err();
        assert_eq!(staged_error.code(), Code::LimitExceeded);
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn malformed_final_transform_rejects_late_geometry_chunk_without_partial_output() {
    let bytes = rewrite_final_transform(&fixture(CHUNK_BOUNDARY + 1));
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let before = source.to_bytes().unwrap();
        let error = source
            .set(
                SELECTOR,
                props! {
                    "text" => "must roll back",
                    "x" => "1in",
                    "width" => "2in",
                },
            )
            .unwrap_err();
        assert!(
            matches!(error.code(), Code::CorruptFile | Code::MissingProperty),
            "unexpected malformed-transform code: {:?}",
            error.code()
        );
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn publication_token_limits_keep_their_code_and_native_source() {
    // Native staging permits 8 MiB of text, but publication bounds each XML
    // token to 4 MiB. The save failure must retain resource semantics.
    let large_text = "x".repeat(4 * 1024 * 1024 + 1);
    for eager in [false, true] {
        let source = open(&fixture(1), eager);
        let before = source.to_bytes().unwrap();
        let updated = source
            .set(SELECTOR, props! { "text" => large_text.as_str() })
            .unwrap();
        let error = updated.to_bytes().unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded);
        let leaf = std::error::Error::source(&error).unwrap();
        let native = leaf.downcast_ref::<litchi_pptx::Error>().unwrap();
        let litchi_pptx::Error::Opc(litchi_opc::OpcError::XmlPublication { source: audit, .. }) =
            native
        else {
            panic!("expected a native XML publication error: {native:?}")
        };
        assert!(audit.is_limit());
        assert!(!audit.is_allocation());
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}
