#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Tests for `NormalViewSetInfo9` and `NotesTextViewInfo9` view preferences
//! against real `PowerPoint` fixtures.

use litchi_ppt::{NormalViewSetPayload, Package};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow")
        .join(name)
}

#[test]
fn reads_normal_view_set_info_from_real_presentations() {
    let mut found_layout = 0;
    let mut found_other = 0;
    for name in [
        "headers_footers.ppt",
        "basic_test_ppt_file.ppt",
        "datetime.ppt",
        "incorrect_slide_order.ppt",
        "41246-1.ppt",
        "WithComments.ppt",
    ] {
        let mut package = Package::open(fixture(name)).unwrap();
        let Some(view) = package
            .presentation()
            .unwrap()
            .normal_view_set_info()
            .unwrap_or_else(|error| panic!("{name}: {error}"))
        else {
            continue;
        };
        match view.payload() {
            NormalViewSetPayload::Layout(layout) => {
                found_layout += 1;
                let left = layout.left_portion();
                assert!(
                    left.numerator() >= 0
                        && left.denominator() > 0
                        && left.numerator() <= left.denominator(),
                    "{name}: leftPortion out of range"
                );
            },
            NormalViewSetPayload::Other(raw) => {
                found_other += 1;
                assert_eq!(raw.len(), 20, "{name}");
            },
        }
    }
    // At least one fixture carries each payload shape (POI timestamps are
    // preserved as opaque payloads).
    assert!(found_layout + found_other >= 2);
}

#[test]
fn reads_notes_text_view_info_from_real_presentations() {
    for name in [
        "headers_footers.ppt",
        "basic_test_ppt_file.ppt",
        "datetime.ppt",
    ] {
        let mut package = Package::open(fixture(name)).unwrap();
        if let Some(notes) = package
            .presentation()
            .unwrap()
            .notes_text_view_info()
            .unwrap_or_else(|error| panic!("{name}: {error}"))
        {
            let scale = notes.view_info().x_scale();
            assert!(scale.numerator() > 0 && scale.denominator() > 0, "{name}");
            return;
        }
    }
    panic!("no fixture carried a NotesTextViewInfo9 record");
}
