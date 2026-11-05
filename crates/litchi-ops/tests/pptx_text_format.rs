//! Unified PPTX rich text writes (record 0252): run and paragraph
//! formatting, paragraph/run/cell text and table-cell fill, through both
//! native owner routes, read back through `get` and after a byte round trip.
#![cfg(feature = "pptx")]

use litchi_ops::{Code, File, OpenOptions, PropValue, SourceMode, props};

const RUN: &str = "/slide[1]/shape[1]/paragraph[1]/run[1]";
const PARAGRAPH: &str = "/slide[1]/shape[1]/paragraph[1]";
const CELL: &str = "/slide[1]/table[1]/tr[2]/tc[1]";

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder
        .presentation()
        .unwrap()
        .add_slide()
        .unwrap()
        .add_text_box("alpha", 0, 0, 1_828_800, 914_400);
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.add_table(
        0usize,
        &[
            vec!["a1".to_owned(), "b1".to_owned()],
            vec!["a2".to_owned(), "b2".to_owned()],
        ],
        (0, 1_000_000, 2_000_000, 600_000),
    )
    .unwrap();
    edit.commit().unwrap().into_package().to_bytes().unwrap()
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

fn reopen(file: &File, eager: bool) -> File {
    open(&file.to_bytes().unwrap(), eager)
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
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
fn run_formatting_reads_back_on_both_owners() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let updated = file
            .set(
                RUN,
                props! {
                    "bold" => true,
                    "italic" => true,
                    "size" => "24pt",
                    "color" => "1F4E79",
                    "font.latin" => "Georgia",
                },
            )
            .unwrap();
        for file in [&updated, &reopen(&updated, eager)] {
            assert_eq!(prop(file, RUN, "bold").as_deref(), Some("true"), "{eager}");
            assert_eq!(prop(file, RUN, "italic").as_deref(), Some("true"));
            assert_eq!(prop(file, RUN, "size").as_deref(), Some("24pt"));
            assert_eq!(prop(file, RUN, "color").as_deref(), Some("#1F4E79"));
            assert_eq!(prop(file, RUN, "font.latin").as_deref(), Some("Georgia"));
            assert_eq!(text(file, RUN), "alpha");
        }
        // The source handle is unchanged.
        assert_eq!(prop(&file, RUN, "bold"), None);
        // Null removes the local property.
        let cleared = updated
            .set(
                RUN,
                props! { "bold" => PropValue::Null, "font" => PropValue::Null },
            )
            .unwrap();
        assert_eq!(prop(&cleared, RUN, "bold"), None);
        assert_eq!(prop(&cleared, RUN, "font.latin"), None);
        assert_eq!(prop(&cleared, RUN, "italic").as_deref(), Some("true"));
        // Run text keeps the formatting.
        let retexted = updated.set(RUN, props! { "text" => "omega" }).unwrap();
        assert_eq!(text(&retexted, RUN), "omega");
        assert_eq!(prop(&retexted, RUN, "bold").as_deref(), Some("true"));
    }
}

#[test]
fn paragraph_text_alignment_and_shape_level_formatting() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let updated = file
            .set(
                PARAGRAPH,
                props! { "text" => "new words", "align" => "center", "bold" => true },
            )
            .unwrap();
        let reread = reopen(&updated, eager);
        assert_eq!(text(&reread, PARAGRAPH), "new words");
        assert_eq!(prop(&reread, PARAGRAPH, "align").as_deref(), Some("center"));
        assert_eq!(prop(&reread, RUN, "bold").as_deref(), Some("true"));

        let shaped = file
            .set(
                "/slide[1]/shape[1]",
                props! { "italic" => true, "size" => "18pt", "align" => "right" },
            )
            .unwrap();
        assert_eq!(prop(&shaped, RUN, "italic").as_deref(), Some("true"));
        assert_eq!(
            prop(&shaped, "/slide[1]/shape[1]", "size").as_deref(),
            Some("18pt")
        );
        assert_eq!(prop(&shaped, PARAGRAPH, "align").as_deref(), Some("right"));

        // Shape text and formatting stage in one native edit.
        let mixed = file
            .set(
                "/slide[1]/shape[1]",
                props! { "text" => "both", "bold" => true },
            )
            .unwrap();
        assert_eq!(text(&mixed, "/slide[1]/shape[1]"), "both");
        assert_eq!(prop(&mixed, RUN, "bold").as_deref(), Some("true"));
    }
}

#[test]
fn table_cell_text_and_fill() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        assert_eq!(prop(&file, CELL, "fill"), None);
        let updated = file
            .set(
                CELL,
                props! { "text" => "changed", "fill" => "FFEE00", "bold" => true },
            )
            .unwrap();
        let reread = reopen(&updated, eager);
        assert_eq!(text(&reread, CELL), "changed");
        assert_eq!(prop(&reread, CELL, "fill").as_deref(), Some("#FFEE00"));
        assert_eq!(
            prop(&reread, "/slide[1]/table[1]/tr[1]/tc[1]", "fill"),
            None
        );
        assert_eq!(text(&reread, "/slide[1]/table[1]/tr[1]/tc[2]"), "b1");
        let cleared = updated
            .set(CELL, props! { "fill" => PropValue::Null })
            .unwrap();
        assert_eq!(prop(&cleared, CELL, "fill"), None);
    }
}

#[test]
fn query_fans_out_over_every_run() {
    let bytes = fixture();
    let file = open(&bytes, false);
    let query = litchi_core::selector::Query::parse("slide[1] run")
        .unwrap()
        .with_scope(litchi_core::selector::Scope::Path(Box::new(
            litchi_core::selector::Path::parse("/").unwrap(),
        )));
    let updated = file.set(query, props! { "italic" => true }).unwrap();
    assert_eq!(prop(&updated, RUN, "italic").as_deref(), Some("true"));
    assert_eq!(
        prop(&updated, "/slide[1]/table[1]/tr[1]/tc[1]", "fill"),
        None,
        "cells are not runs of a shape"
    );
}

#[test]
fn typed_refusals_leave_the_handle_unchanged() {
    let bytes = fixture();
    let file = open(&bytes, false);
    let before = file.to_bytes().unwrap();
    let code = |result: litchi_ops::Result<File>| result.err().map(|error| error.code());
    assert_eq!(
        code(file.set(RUN, props! { "fill" => "FF0000" })),
        Some(Code::UnsupportedProperty)
    );
    assert_eq!(
        code(file.set(RUN, props! { "size" => "0.5pt" })),
        Some(Code::InvalidValue)
    );
    assert_eq!(
        code(file.set(RUN, props! { "color" => "accent1" })),
        Some(Code::UnsupportedCapability)
    );
    assert_eq!(
        code(file.set("/slide[1]/shape[1]/paragraph[9]", props! { "bold" => true })),
        Some(Code::PathNotFound)
    );
    assert_eq!(
        code(file.set(RUN, props! { "text" => "a\nb" })),
        Some(Code::InvalidValue)
    );
    assert_eq!(file.to_bytes().unwrap(), before);
}

#[test]
fn edit_closure_composes_text_writes_atomically() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let updated = file
            .edit(|edit| {
                edit.set(RUN, props! { "bold" => true })?;
                edit.set(CELL, props! { "fill" => "00FF00" })?;
                Ok(())
            })
            .unwrap()
            .into_file();
        assert_eq!(prop(&updated, RUN, "bold").as_deref(), Some("true"));
        assert_eq!(prop(&updated, CELL, "fill").as_deref(), Some("#00FF00"));
        let failed = file.edit(|edit| {
            edit.set(RUN, props! { "bold" => true })?;
            // A transformed theme colour is not writable (record 0333).
            edit.set(CELL, props! { "fill" => "accent1+lumMod75" })?;
            Ok(())
        });
        assert!(failed.is_err());
        assert_eq!(prop(&file, RUN, "bold"), None);
    }
}
