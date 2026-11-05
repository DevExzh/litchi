//! Unified PPTX table-cell properties and run underline/strike (record
//! 0333): theme and `none` fills (0330 F1), vertical anchor, insets, and
//! OfficeCLI's cell aliases, on both owners and after a byte round trip.
#![cfg(feature = "pptx")]

use litchi_ops::{Code, File, OpenOptions, PropValue, SourceMode, props};

const CELL: &str = "/slide[1]/table[1]/tr[1]/tc[2]";
const RUN: &str = "/slide[1]/shape[1]/paragraph[1]/run[1]";

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

#[test]
fn cell_fill_takes_theme_colours_and_none() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let themed = file.set(CELL, props! { "fill" => "accent1" }).unwrap();
        let themed = reopen(&themed, eager);
        assert_eq!(prop(&themed, CELL, "fill").as_deref(), Some("accent1"));
        let aliased = themed
            .set(CELL, props! { "background" => "4472C4" })
            .unwrap();
        assert_eq!(prop(&aliased, CELL, "fill").as_deref(), Some("#4472C4"));
        let none = aliased.set(CELL, props! { "fill" => "none" }).unwrap();
        assert_eq!(prop(&none, CELL, "fill"), None);
        let cleared = none
            .set(CELL, props! { "fill" => PropValue::Null })
            .unwrap();
        assert_eq!(prop(&cleared, CELL, "fill"), None);
        // Untouched cells keep no local fill.
        assert_eq!(
            prop(&cleared, "/slide[1]/table[1]/tr[1]/tc[1]", "fill"),
            None
        );
        // A transformed colour is refused, typed.
        let error = file
            .set(CELL, props! { "fill" => "accent1+lumMod75" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
    }
}

#[test]
fn cell_anchor_insets_and_text_formatting() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let set = file
            .set(
                CELL,
                props! {
                    "valign" => "center",
                    "padding" => "0.1in",
                    "margin.bottom" => "0.2in",
                    "underline" => "single",
                    "strike" => "double",
                    "font.bold" => true,
                },
            )
            .unwrap_err();
        // `padding` and `padding.bottom` name the same inset.
        assert_eq!(set.code(), Code::InvalidArgument, "{set}");
        let set = file
            .set(
                CELL,
                props! {
                    "valign" => "center",
                    "padding" => "0.1in",
                    "underline" => "single",
                    "strike" => "double",
                    "font.bold" => true,
                },
            )
            .unwrap();
        let set = reopen(&set, eager);
        assert_eq!(prop(&set, CELL, "valign").as_deref(), Some("center"));
        for side in ["left", "right", "top", "bottom"] {
            assert_eq!(
                prop(&set, CELL, &format!("padding.{side}")).as_deref(),
                Some("7.2pt"),
                "{side}"
            );
        }
        let run = format!("{CELL}/paragraph[1]/run[1]");
        assert_eq!(prop(&set, &run, "underline").as_deref(), Some("single"));
        assert_eq!(prop(&set, &run, "strike").as_deref(), Some("double"));
        assert_eq!(prop(&set, &run, "bold").as_deref(), Some("true"));
        let bottom = set
            .set(
                CELL,
                props! { "padding.bottom" => "0.5in", "valign" => PropValue::Null },
            )
            .unwrap();
        assert_eq!(
            prop(&bottom, CELL, "padding.bottom").as_deref(),
            Some("1.27cm")
        );
        assert_eq!(prop(&bottom, CELL, "padding.top").as_deref(), Some("7.2pt"));
        assert_eq!(prop(&bottom, CELL, "valign"), None);
    }
}

#[test]
fn runs_paragraphs_and_shapes_take_underline_and_strike() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let run = file
            .set(RUN, props! { "underline" => "true", "strike" => "single" })
            .unwrap();
        assert_eq!(prop(&run, RUN, "underline").as_deref(), Some("single"));
        assert_eq!(prop(&run, RUN, "strike").as_deref(), Some("single"));
        let shape = file
            .set("/slide[1]/shape[1]", props! { "underline" => "double" })
            .unwrap();
        assert_eq!(prop(&shape, RUN, "underline").as_deref(), Some("double"));
        let paragraph = shape
            .set(
                "/slide[1]/shape[1]/paragraph[1]",
                props! { "underline" => "none", "strike" => PropValue::Null },
            )
            .unwrap();
        assert_eq!(prop(&paragraph, RUN, "underline").as_deref(), None);
        assert_eq!(
            file.set(RUN, props! { "underline" => "zigzag" })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
    }
}
