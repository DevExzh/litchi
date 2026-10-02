#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

//! The tolerance contract: Word-compatible repairs are applied and recorded by
//! default, opt-in repairs only under a named `Leniency`, and conformant input
//! reports clean.

use litchi_rtf::leniency::{Defect, Leniency, SyntaxDefect, ToleranceReport, ToleratedDefect};
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

fn recorded(report: &ToleranceReport) -> Vec<(Defect, u64)> {
    report.iter().map(|d| (d.defect, d.offset)).collect()
}

// F9.4: undefined control symbols are ignored the way Word and LibreOffice
// ignore them, and every one is recorded.
#[test]
fn undefined_control_symbols_are_ignored_and_recorded_by_default() {
    let source = r"{\rtf1\ansi\pard before\!after \;x\par}";
    let document = Document::parse(source).unwrap();
    assert_eq!(document.text(), "beforeafter x\n");
    let unknown = Defect::Syntax(SyntaxDefect::UnknownControlSymbol);
    let first = source.find(r"\!").unwrap() as u64;
    let second = source.find(r"\;").unwrap() as u64;
    assert_eq!(
        recorded(document.tolerance_report()),
        vec![(unknown, first), (unknown, second)]
    );
    assert!(!SyntaxDefect::UnknownControlSymbol.repair().is_empty());
}

#[test]
fn undefined_control_symbols_are_tolerated_in_header_destinations() {
    let source = r"{\rtf1\ansi{\fonttbl{\f0\fswiss Arial\!;}}{\info{\title A\?B}}\pard x\1y\par}";
    let document = Document::parse(source).unwrap();
    assert_eq!(document.text(), "xy\n");
    assert_eq!(document.fonts().iter().next().unwrap().name(), "Arial");
    let report = document.tolerance_report();
    assert_eq!(report.total(), 3);
    assert!(report.iter().all(|tolerated| matches!(
        tolerated,
        ToleratedDefect {
            defect: Defect::Syntax(SyntaxDefect::UnknownControlSymbol),
            ..
        }
    )));
}

#[test]
fn defined_control_symbols_without_a_model_are_not_defects() {
    // `\|` (formula character) and `\:` (index subentry) are RTF 1.9.1
    // control symbols; they are retained inertly and the report stays clean.
    let document = Document::parse(r"{\rtf1\ansi\pard a\|b\:c\par}").unwrap();
    assert_eq!(document.text(), "abc\n");
    assert_clean(document.tolerance_report());
}

#[test]
fn undefined_control_symbol_offsets_count_transport_bytes() {
    let mut bytes = br"{\rtf1\ansi\ansicpg1252\pard caf".to_vec();
    bytes.push(0xE9);
    let offset = bytes.len() as u64;
    bytes.extend_from_slice(br"\!\par}");
    let document = Document::from_bytes(&bytes).unwrap();
    assert_eq!(document.text(), "caf\u{e9}\n");
    assert_eq!(
        recorded(document.tolerance_report()),
        vec![(Defect::Syntax(SyntaxDefect::UnknownControlSymbol), offset)]
    );
}
