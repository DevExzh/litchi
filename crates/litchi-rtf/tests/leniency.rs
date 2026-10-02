#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

//! The tolerance contract: Word-compatible repairs are applied and recorded by
//! default, opt-in repairs only under a named `Leniency`, and conformant input
//! reports clean.

use litchi_rtf::leniency::{Leniency, ToleranceReport};
use litchi_rtf::read::{Limits, Options};
use litchi_rtf::{Document, RtfDocument, compress};

/// A conformant document touching the main destinations.
const CONFORMANT: &str = concat!(
    r"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fcharset0 Arial;}}",
    r"{\colortbl;\red255\green0\blue0;}",
    r"{\stylesheet{\s0 Normal;}{\s1\b Heading;}}",
    r"{\info{\title Conformant}}",
    r"\pard\s1\b\fs28 Title\b0\par",
    r"\pard\plain\fs24 caf\'e9 荤?\tab \i italic\i0\line next\par",
    r"\trowd\cellx2000\cellx4000\pard\intbl A\cell B\cell\row",
    r"\pard\hyphpar1\page{\*\unknowndest ignored}\par}",
);

fn assert_clean(report: &ToleranceReport) {
    assert!(report.is_clean(), "{report:?}");
    assert_eq!(report.total(), 0);
    assert_eq!(report.unrecorded(), 0);
    assert!(report.defects().is_empty());
}

#[test]
fn conformant_input_reports_clean_through_every_entry_point() {
    let strict = Options::new();
    assert_eq!(strict, Options::default());
    assert_eq!(strict.leniency(), Leniency::Strict);

    let document = Document::parse(CONFORMANT).unwrap();
    assert_clean(document.tolerance_report());
    assert!(document.text().contains("Title"));

    let document = Document::parse_with_options(CONFORMANT, strict).unwrap();
    assert_clean(document.tolerance_report());
    let document = Document::from_bytes_with_options(CONFORMANT.as_bytes(), strict).unwrap();
    assert_clean(document.tolerance_report());
    let document =
        Document::from_owned_bytes_with_options(CONFORMANT.as_bytes().to_vec(), strict).unwrap();
    assert_clean(document.tolerance_report());

    let raw = RtfDocument::parse_bytes_with_options(CONFORMANT.as_bytes(), strict).unwrap();
    assert_clean(raw.tolerance_report());

    let compressed = compress(CONFORMANT.as_bytes(), true).unwrap();
    let document = Document::from_bytes_with_options(&compressed, strict).unwrap();
    assert_clean(document.tolerance_report());
}

#[test]
fn conformant_non_ascii_transport_reports_clean() {
    let mut bytes = br"{\rtf1\ansi\ansicpg1252 caf".to_vec();
    bytes.push(0xE9);
    bytes.extend_from_slice(br" bar\par}");
    let document = Document::from_bytes(&bytes).unwrap();
    assert_clean(document.tolerance_report());
}

#[test]
fn options_carry_limits_into_the_parse() {
    let limits = Limits::default().with_max_tokens(2);
    let options = Options::new().with_limits(limits);
    assert_eq!(options.limits(), limits);
    assert!(Document::parse_with_options(CONFORMANT, options).is_err());
}

#[test]
fn opening_a_file_with_options_reports_clean() {
    let directory =
        std::env::temp_dir().join(format!("litchi-rtf-leniency-{}", std::process::id()));
    std::fs::create_dir_all(&directory).unwrap();
    let path = directory.join("conformant.rtf");
    std::fs::write(&path, CONFORMANT).unwrap();
    let document = Document::open_with_options(&path, Options::new()).unwrap();
    assert_clean(document.tolerance_report());
    std::fs::remove_dir_all(&directory).unwrap();
}
