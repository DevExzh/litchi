//! Differential proof that the `<sheetData>` lane leaves the edit scanner's
//! layout and refusals unchanged.
//!
//! Every document is scanned twice, once with the lane disabled, and the
//! complete `Layout` (every span, tag, slot and guard) or error is compared by
//! its `Debug` and `Display` forms.

use super::{scan, scan_with_event_limit};
use crate::raw::worksheet::lane::corpus::{Lcg, MAIN, X14AC, generated_body, worksheet};
use crate::raw::worksheet::lane::route::{self, Pass};

/// Scan through both routes and require identical outcomes.
///
/// Returns whether the lane admitted the body.
fn assert_parity(document: &[u8], limit: Option<usize>) -> bool {
    let run = || match limit {
        Some(limit) => scan_with_event_limit(document, limit),
        None => scan(document),
    };
    route::reset();
    let fast = run();
    let admitted = route::admitted(Pass::Scan) > 0;
    let slow = route::without_lane(run);
    let text = String::from_utf8_lossy(document);
    match (&fast, &slow) {
        (Ok(fast), Ok(slow)) => {
            assert_eq!(format!("{fast:?}"), format!("{slow:?}"), "{text}");
        },
        (Err(fast), Err(slow)) => {
            assert_eq!(format!("{fast:?}"), format!("{slow:?}"), "{text}");
            assert_eq!(fast.to_string(), slow.to_string(), "{text}");
        },
        _ => panic!("routes disagree for {text}: {fast:?} vs {slow:?}"),
    }
    admitted
}

#[test]
fn generated_benign_bodies_take_the_lane_with_identical_layouts() {
    let mut random = Lcg(0x5CA7);
    let mut admitted = 0usize;
    for _ in 0..600 {
        let document = worksheet(&generated_body(&mut random, false));
        admitted += usize::from(assert_parity(document.as_bytes(), None));
    }
    assert_eq!(admitted, 600, "every benign body must take the lane");
}

#[test]
fn generated_hostile_bodies_keep_identical_layouts_and_refusals() {
    let mut random = Lcg(0x7AC5);
    let mut admitted = 0usize;
    let mut refused = 0usize;
    for _ in 0..1_500 {
        let document = worksheet(&generated_body(&mut random, true));
        admitted += usize::from(assert_parity(document.as_bytes(), None));
        refused += usize::from(scan(document.as_bytes()).is_err());
    }
    assert!(admitted > 1_000, "lane admitted {admitted}");
    assert!(refused > 100, "only {refused} documents were refused");
}

#[test]
fn event_limits_inside_and_after_the_body_keep_the_reader_refusal() {
    let body = "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" s=\"2\"/></row><row r=\"2\"/>";
    let document = worksheet(body);
    let bytes = document.as_bytes();
    let Ok(_layout) = scan(bytes) else {
        panic!("fixture must scan");
    };
    // Find the smallest limit that still succeeds, then probe around it and
    // at every limit that ends inside the body.
    let mut minimum = 1usize;
    while scan_with_event_limit(bytes, minimum).is_err() {
        minimum += 1;
    }
    let mut admitted = 0usize;
    for limit in 1..=minimum + 2 {
        admitted += usize::from(assert_parity(bytes, Some(limit)));
    }
    assert!(admitted > 0, "some limit past the body must admit the lane");
}

#[test]
fn invalid_utf8_in_a_value_keeps_the_reader_refusal() {
    let mut document =
        worksheet("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>xx</v></c></row>").into_bytes();
    let at = document
        .windows(2)
        .position(|window| window == b"xx")
        .expect("marker");
    document[at] = 0xFF;
    assert!(!assert_parity(&document, None));
}

#[test]
fn namespace_and_envelope_variants_keep_identical_layouts() {
    let body = "<row r=\"1\" x14ac:dyDescent=\"0.25\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" s=\"2\" t=\"n\"/></row>";
    let cases = [
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\" xmlns:x14ac=\"{X14AC}\"><sheetData>{body}</sheetData></worksheet>"
            ),
            true,
        ),
        (
            format!(
                "<x:worksheet xmlns:x=\"{MAIN}\" xmlns:x14ac=\"{X14AC}\"><x:sheetData xmlns=\"{MAIN}\">{body}</x:sheetData></x:worksheet>"
            ),
            true,
        ),
        (
            format!(
                "<x:worksheet xmlns:x=\"{MAIN}\" xmlns:x14ac=\"{X14AC}\"><x:sheetData>{body}</x:sheetData></x:worksheet>"
            ),
            false,
        ),
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData><sheetProtection sheet=\"1\"/></worksheet>"
            ),
            true,
        ),
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData><mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells></worksheet>"
            ),
            true,
        ),
        (
            format!("<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData><bad></worksheet>"),
            true,
        ),
        (
            format!("<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData>"),
            true,
        ),
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData><sheetData/></worksheet>"
            ),
            true,
        ),
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}<row r=\"2\"><c r=\"A2\"><f>A1</f></c></row></sheetData></worksheet>"
            ),
            false,
        ),
    ];
    for (document, expected) in cases {
        assert_eq!(
            assert_parity(document.as_bytes(), None),
            expected,
            "{document}"
        );
    }
}

#[test]
fn ordering_and_coordinate_refusals_are_unchanged() {
    for body in [
        "<row r=\"1\"><c r=\"B1\"/><c r=\"A1\"/></row>",
        "<row r=\"1\"><c r=\"A1\"/><c r=\"A1\"/></row>",
        "<row r=\"2\"/><row r=\"1\"/>",
        "<row r=\"1\"/><row r=\"1\"/>",
        "<row r=\"1\"><c r=\"A2\"/></row>",
        "<row r=\"1\"><c r=\"XFE1\"/></row>",
        "<row><c/><c/></row><row><c r=\"C2\"/></row>",
        "<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c><c t=\"n\" r=\"B1\"/><c s=\"1\"/></row>",
    ] {
        assert!(assert_parity(worksheet(body).as_bytes(), None), "{body}");
    }
}

#[test]
fn byte_order_marked_worksheets_take_the_lane_with_document_offsets() {
    // The reader drops a leading byte-order mark without counting it; the
    // scanner's spans are document offsets, so a marked document lays out
    // exactly like the same document behind three bytes of whitespace, and
    // both routes agree.
    let mut random = Lcg(0xB0C);
    for _ in 0..50 {
        let body = generated_body(&mut random, false);
        let document = worksheet(&body);
        let root = document
            .find("<worksheet")
            .expect("generated worksheet root");
        let undeclared = &document[root..];
        let marked = format!("\u{feff}{undeclared}");
        let spaced = format!("   {undeclared}");
        assert!(assert_parity(marked.as_bytes(), None), "{marked}");
        assert!(assert_parity(spaced.as_bytes(), None), "{spaced}");
        assert_eq!(
            format!("{:?}", scan(marked.as_bytes())),
            format!("{:?}", scan(spaced.as_bytes())),
            "{marked}"
        );
        let declared = format!("\u{feff}{document}");
        assert!(assert_parity(declared.as_bytes(), None), "{declared}");
    }
    // A second mark is character data before the root, which the scanner
    // ignores like any other text there; its spans stay document offsets.
    let doubled = format!("\u{feff}\u{feff}{}", worksheet("<row r=\"1\"/>"));
    assert!(assert_parity(doubled.as_bytes(), None), "{doubled}");
}
