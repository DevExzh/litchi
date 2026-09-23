//! Differential proof that the `<sheetData>` lane leaves the worksheet
//! parser's results unchanged.
//!
//! Every document is parsed twice through the public raw entry point, once
//! with the lane disabled, and the complete `Store` or error is compared by
//! its `Debug` and `Display` forms. The generated corpus mixes admitted bodies
//! with near misses the lane must decline and with values the shared parser
//! methods must refuse, so both the fast route and its refusals are covered.

use super::lane::corpus::{Lcg, MAIN, STRICT, X14AC, generated_body, shared_strings, worksheet};
use super::lane::route::{self, Pass};
use crate::cell::Text;

/// Parse through both routes and require identical outcomes.
///
/// Returns whether the lane admitted the body.
fn assert_parity(document: &str, strings: &[Text]) -> bool {
    route::reset();
    let fast = super::parse(document.as_bytes(), || Ok(Some(strings)));
    let admitted = route::admitted(Pass::Parse) > 0;
    let slow = route::without_lane(|| super::parse(document.as_bytes(), || Ok(Some(strings))));
    match (&fast, &slow) {
        (Ok(fast), Ok(slow)) => {
            assert_eq!(format!("{fast:?}"), format!("{slow:?}"), "{document}");
        },
        (Err(fast), Err(slow)) => {
            assert_eq!(format!("{fast:?}"), format!("{slow:?}"), "{document}");
            assert_eq!(fast.to_string(), slow.to_string(), "{document}");
        },
        _ => panic!("routes disagree for {document}: {fast:?} vs {slow:?}"),
    }
    admitted
}

#[test]
fn generated_benign_bodies_take_the_lane_with_identical_stores() {
    let strings = shared_strings();
    let mut random = Lcg(0x0744);
    let mut admitted = 0usize;
    for _ in 0..600 {
        let document = worksheet(&generated_body(&mut random, false));
        admitted += usize::from(assert_parity(&document, &strings));
    }
    assert_eq!(admitted, 600, "every benign body must take the lane");
}

#[test]
fn generated_hostile_bodies_keep_identical_values_and_refusals() {
    let strings = shared_strings();
    let mut random = Lcg(0x4470);
    let mut admitted = 0usize;
    let mut refused = 0usize;
    for _ in 0..1_500 {
        let document = worksheet(&generated_body(&mut random, true));
        admitted += usize::from(assert_parity(&document, &strings));
        refused += usize::from(super::parse(document.as_bytes(), || Ok(Some(&strings))).is_err());
    }
    // The comparison is only meaningful if both outcomes were exercised.
    assert!(admitted > 1_000, "lane admitted {admitted}");
    assert!(refused > 300, "only {refused} documents were refused");
}

#[test]
fn shared_string_refusals_keep_their_precedence() {
    // No shared-string part: the refusal comes from materialization after
    // the whole worksheet parsed, on both routes.
    let document = worksheet("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>");
    route::reset();
    let fast = super::parse(document.as_bytes(), || Ok(None));
    assert_eq!(route::admitted(Pass::Parse), 1);
    let slow = route::without_lane(|| super::parse(document.as_bytes(), || Ok(None)));
    assert_eq!(
        format!("{:?}", fast.expect_err("no strings")),
        format!("{:?}", slow.expect_err("no strings"))
    );
    // An out-of-range index and an unreadable table behave identically too.
    assert!(assert_parity(
        &worksheet("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>9</v></c></row>"),
        &shared_strings()
    ));
    let failing = || Err(crate::error::invalid("shared strings are unreadable"));
    let fast = super::parse(document.as_bytes(), failing);
    let slow = route::without_lane(|| super::parse(document.as_bytes(), failing));
    assert_eq!(
        format!("{:?}", fast.expect_err("table error")),
        format!("{:?}", slow.expect_err("table error"))
    );
}

#[test]
fn namespace_and_envelope_variants_keep_identical_outcomes() {
    let strings = shared_strings();
    let body = "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" s=\"2\"/></row>";
    let cases = [
        // Strict dialect: the lane admits it.
        (
            format!("<worksheet xmlns=\"{STRICT}\"><sheetData>{body}</sheetData></worksheet>"),
            true,
        ),
        // Prefixed dialect declared on the root, default bound on sheetData.
        (
            format!(
                "<x:worksheet xmlns:x=\"{MAIN}\"><x:sheetData xmlns=\"{MAIN}\">{body}</x:sheetData></x:worksheet>"
            ),
            true,
        ),
        // Prefixed dialect without a default namespace: unprefixed rows are
        // foreign elements the parser ignores, and the lane stays out.
        (
            format!(
                "<x:worksheet xmlns:x=\"{MAIN}\"><x:sheetData>{body}</x:sheetData></x:worksheet>"
            ),
            false,
        ),
        // Prefixed rows inside a prefixed dialect.
        (
            format!(
                "<x:worksheet xmlns:x=\"{MAIN}\"><x:sheetData><x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>"
            ),
            false,
        ),
        // A foreign default namespace on sheetData.
        (
            format!(
                "<x:worksheet xmlns:x=\"{MAIN}\"><x:sheetData xmlns=\"urn:other\">{body}</x:sheetData></x:worksheet>"
            ),
            false,
        ),
        // An empty body and an empty sheetData element.
        (
            format!("<worksheet xmlns=\"{MAIN}\"><sheetData></sheetData></worksheet>"),
            true,
        ),
        (
            format!("<worksheet xmlns=\"{MAIN}\"><sheetData/></worksheet>"),
            false,
        ),
        // A second sheetData after an admitted first one is still refused.
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData><sheetData/></worksheet>"
            ),
            true,
        ),
        // Malformed markup after an admitted body is still refused.
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
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData></worksheet><extra/>"
            ),
            true,
        ),
        // A merge range after the body still reaches the parser.
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}</sheetData><mergeCells><mergeCell ref=\"A1:B1\"/></mergeCells></worksheet>"
            ),
            true,
        ),
        // Truncated inside the body: the lane declines, the reader refuses.
        (
            format!("<worksheet xmlns=\"{MAIN}\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1"),
            false,
        ),
        // Non-row content inside sheetData.
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData><!-- c -->{body}</sheetData></worksheet>"
            ),
            false,
        ),
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}<extLst/></sheetData></worksheet>"
            ),
            false,
        ),
        // A formula cell keeps the whole body on the reader.
        (
            format!(
                "<worksheet xmlns=\"{MAIN}\"><sheetData>{body}<row r=\"2\"><c r=\"A2\"><f>A1+1</f><v>2</v></c></row></sheetData></worksheet>"
            ),
            false,
        ),
    ];
    for (document, expected) in cases {
        assert_eq!(assert_parity(&document, &strings), expected, "{document}");
    }
}

#[test]
fn row_descent_extensions_reach_admitted_rows() {
    let strings = shared_strings();
    let document = format!(
        "<worksheet xmlns=\"{MAIN}\" xmlns:x14ac=\"{X14AC}\"><sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>\
         <sheetData><row r=\"1\" x14ac:dyDescent=\"0.3\"><c r=\"A1\"><v>1</v></c></row>\
         <row r=\"2\"><c r=\"A2\"><v>2</v></c></row></sheetData></worksheet>"
    );
    assert!(assert_parity(&document, &strings));
    let store = super::parse(document.as_bytes(), || Ok(Some(&strings))).expect("valid");
    let first = store
        .row_entry(litchi_sheet::Row::new(0).expect("row"))
        .expect("row 1");
    assert!(first.properties.descent.is_some());
}

#[test]
fn large_and_oversized_values_keep_identical_outcomes() {
    let strings = shared_strings();
    let long = "9".repeat(300);
    assert!(assert_parity(
        &worksheet(&format!("<row r=\"1\"><c r=\"A1\"><v>{long}</v></c></row>")),
        &strings
    ));
    let text = "x".repeat(super::model::MAX_CELL_CHARACTERS + 1);
    assert!(assert_parity(
        &worksheet(&format!(
            "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>{text}</v></c></row>"
        )),
        &strings
    ));
    let encoded = "y".repeat(super::model::MAX_ENCODED_CELL_BYTES + 1);
    assert!(assert_parity(
        &worksheet(&format!(
            "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>{encoded}</v></c></row>"
        )),
        &strings
    ));
}

#[test]
fn duplicate_rows_cells_and_ordering_refusals_are_unchanged() {
    let strings = shared_strings();
    for body in [
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"A1\"><v>2</v></c></row>",
        "<row r=\"2\"/><row r=\"1\"/>",
        "<row r=\"1\"/><row r=\"1\"/>",
        "<row r=\"1\"><c r=\"B1\"/><c r=\"A1\"/></row>",
        "<row r=\"1\"><c r=\"XFE1\"/></row>",
        "<row r=\"1\"><c r=\"A0\"/></row>",
        "<row r=\"1\"><c r=\"1A\"/></row>",
        "<row r=\"1048577\"/>",
        "<row r=\"0\"/>",
        "<row r=\"1\"><c r=\"a1\"><v>1</v></c></row>",
    ] {
        assert!(assert_parity(&worksheet(body), &strings), "{body}");
    }
}

#[test]
fn inferred_coordinates_match_the_reader() {
    let strings = shared_strings();
    let mut body = String::from("<row>");
    for _ in 0..3 {
        body.push_str("<c><v>1</v></c>");
    }
    body.push_str("</row><row><c/><c r=\"D2\"/><c/></row>");
    assert!(assert_parity(&worksheet(&body), &strings));
}

#[test]
fn compact_lane_cells_match_every_value_form() {
    let strings = shared_strings();
    for body in [
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"/></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><v>1</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><v/></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"str\"><v/></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"str\"/></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>a_x0041_b</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"s\"/></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v></v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v> 2 </v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#N/A</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#VENDOR!</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"d\"><v>2026-01-02</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"b\"><v> 1 </v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"x\"><v>raw</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"x\"/></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"n\"/></row>",
        "<row r=\"1\"><c r=\"A1\" t=\"\"><v>1</v></c></row>",
        "<row r=\"1\"><c r=\"A1\"><v>  </v></c></row>",
        "<row r=\"1\"><c r=\"A1\"><v>-0.000</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" cm=\"3\" vm=\"4\"><v>1</v></c></row>",
        "<row r=\"1\"><c r=\"A1\" s=\"0\"/><c r=\"B1\" s=\"65490\"/></row>",
    ] {
        assert!(assert_parity(&worksheet(body), &strings), "{body}");
    }
}

#[test]
fn byte_order_marked_worksheets_keep_the_reader_with_identical_stores() {
    let strings = shared_strings();
    let mut random = Lcg(0xB0B);
    for _ in 0..50 {
        let document = format!("\u{feff}{}", worksheet(&generated_body(&mut random, false)));
        assert!(!assert_parity(&document, &strings), "{document}");
    }
}
