//! Unified PPTX shape geometry writes through both native owner routes.
#![cfg(feature = "pptx")]

use litchi_ops::node::Value;
use litchi_ops::{Code, File, OpenOptions, SourceMode, props};

const FIRST: &str = "/slide[1]/shape[1]";

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    for texts in [&["first", "second"][..], &["third"][..]] {
        let slide = builder.presentation().unwrap().add_slide().unwrap();
        for (index, text) in texts.iter().copied().enumerate() {
            slide.add_text_box(
                text,
                0,
                i64::try_from(index).unwrap() * 914_400,
                1_828_800,
                914_400,
            );
        }
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

fn length(file: &File, path: &str, key: &str) -> i64 {
    let node = file.get(path).unwrap();
    let property = node
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .unwrap();
    let Value::Length(value) = property.value() else {
        panic!("expected a length property");
    };
    value.emu()
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}

#[test]
fn partial_geometry_hydrates_existing_bounds_and_mixes_text_atomically() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let updated = source.set(FIRST, props! { "x" => "1in" }).unwrap();
        assert_eq!(length(&updated, FIRST, "x"), 914_400);
        assert_eq!(length(&updated, FIRST, "y"), 0);
        assert_eq!(length(&updated, FIRST, "width"), 1_828_800);
        assert_eq!(length(&updated, FIRST, "height"), 914_400);

        let mixed = updated
            .set(
                FIRST,
                props! {
                    "text" => "mixed",
                    "y" => "2in",
                    "height" => "2in",
                },
            )
            .unwrap();
        assert_eq!(text(&mixed, FIRST), "mixed");
        assert_eq!(length(&mixed, FIRST, "x"), 914_400);
        assert_eq!(length(&mixed, FIRST, "y"), 1_828_800);
        assert_eq!(length(&mixed, FIRST, "width"), 1_828_800);
        assert_eq!(length(&mixed, FIRST, "height"), 1_828_800);
        assert_eq!(text(&source, FIRST), "first");
        assert_eq!(source.to_bytes().unwrap(), bytes);

        let reopened = open(&mixed.to_bytes().unwrap(), eager);
        assert_eq!(text(&reopened, FIRST), "mixed");
        assert_eq!(length(&reopened, FIRST, "height"), 1_828_800);
    }
}

#[test]
fn geometry_batches_group_by_slide_and_reject_duplicates_or_nonpositive_extents() {
    let bytes = fixture();
    for eager in [false, true] {
        let source = open(&bytes, eager);
        let updated = source.set("/slide[1]/*", props! { "x" => "3in" }).unwrap();
        assert_eq!(length(&updated, "/slide[1]/shape[1]", "x"), 2_743_200);
        assert_eq!(length(&updated, "/slide[1]/shape[2]", "x"), 2_743_200);
        assert_eq!(length(&updated, "/slide[2]/shape[1]", "x"), 0);

        assert_eq!(
            source
                .set(FIRST, props! { "x" => "1in", "X" => "2in" })
                .unwrap_err()
                .code(),
            Code::InvalidArgument
        );
        assert_eq!(source.to_bytes().unwrap(), bytes);
        assert_eq!(
            source
                .set(FIRST, props! { "width" => 0 })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        assert_eq!(source.to_bytes().unwrap(), bytes);
        assert_eq!(
            source
                .set(FIRST, props! { "height" => "-1in" })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        assert_eq!(source.to_bytes().unwrap(), bytes);
    }
}
