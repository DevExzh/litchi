#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

//! The tolerance contract: Word-compatible repairs are applied and recorded by
//! default, opt-in repairs only under a named `Leniency`, and conformant input
//! reports clean.

use litchi_rtf::leniency::{
    Defect, Leniency, ParameterDefect, SyntaxDefect, ToleranceReport, ToleratedDefect,
    UnicodeDefect,
};
use litchi_rtf::read::{Limits, Options};
use litchi_rtf::{Document, EmphasisMark, RtfDocument, compress};

/// A conformant document touching the main destinations.
const CONFORMANT: &str = concat!(
    r"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fcharset0 Arial;}}",
    r"{\colortbl;\red255\green0\blue0;}",
    r"{\stylesheet{\s0 Normal;}{\s1\b Heading;}}",
    r"{\info{\title Conformant}}",
    r"\pard\s1\b\fs28 Title\b0\par",
    r"\pard\plain\fs24 caf\'e9 \u233?\tab \i italic\i0\line next\par",
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

// F9.6: unpaired UTF-16 surrogates decode to U+FFFD, as Word and LibreOffice
// show them, and each substitution is recorded.
#[test]
fn unpaired_surrogates_become_replacement_characters_and_are_recorded() {
    let source = r"{\rtf1\ansi\uc1\pard a\u-10179?b\u-9047?c\u-10179?\u-8694?d\u-10179?\u-10179?\u-8694?\par}";
    let document = Document::parse(source).unwrap();
    // Lone high, lone low, a valid pair (U+1F60A), then a high orphaned by
    // the next high that does pair.
    assert_eq!(
        document.text(),
        "a\u{fffd}b\u{fffd}c\u{1f60a}d\u{fffd}\u{1f60a}\n"
    );
    let surrogate = Defect::Unicode(UnicodeDefect::UnpairedSurrogate);
    let lone_high = source.find(r"\u-10179?b").unwrap() as u64;
    let lone_low = source.find(r"\u-9047").unwrap() as u64;
    let orphan = source.find(r"\u-10179?\u-10179").unwrap() as u64;
    assert_eq!(
        recorded(document.tolerance_report()),
        vec![
            (surrogate, lone_high),
            (surrogate, lone_low),
            (surrogate, orphan)
        ]
    );
}

#[test]
fn unpaired_surrogates_are_repaired_in_destination_text() {
    let source = concat!(
        r"{\rtf1\ansi{\fonttbl{\f0\fnil A\u-10179?;}}",
        r"{\info{\title T\u-10179?}}",
        r"{\stylesheet{\s0 N\u-9047?;}}",
        r"{\header H\u-10179?\par}",
        r"\pard x{\footnote F\u-9047?}\par}",
    );
    let document = Document::parse(source).unwrap();
    assert_eq!(document.fonts().iter().next().unwrap().name(), "A\u{fffd}");
    let report = document.tolerance_report();
    assert_eq!(report.total(), 5, "{report:?}");
    assert!(
        report
            .iter()
            .all(|tolerated| tolerated.defect == Defect::Unicode(UnicodeDefect::UnpairedSurrogate))
    );
}

#[test]
fn paired_surrogates_are_not_defects() {
    let document = Document::parse(r"{\rtf1\ansi\uc1\pard \u-10179?\u-8694?\u233?\par}").unwrap();
    assert_eq!(document.text(), "\u{1f60a}\u{e9}\n");
    assert_clean(document.tolerance_report());
}

#[test]
fn word_style_surrogate_pairs_with_fallbacks_decode_cleanly_in_destinations() {
    // Word writes each surrogate half with its own `\ucN` fallback.
    let pair = r"\u-10179?\u-8694?";
    let source = format!(
        r"{{\rtf1\ansi\uc1{{\fonttbl{{\f0\fnil A{pair};}}}}{{\info{{\title T{pair}}}}}\pard x{pair}y\par}}"
    );
    let document = Document::parse(&source).unwrap();
    assert_eq!(document.text(), "x\u{1f60a}y\n");
    assert_eq!(document.fonts().iter().next().unwrap().name(), "A\u{1f60a}");
    assert_eq!(document.info().title.as_deref(), Some("T\u{1f60a}"));
    assert_clean(document.tolerance_report());
}

// F9.5: malformed parameter shapes take Word's interpretation (LibreOffice's
// `RTFTokenizer::dispatchKeyword`) and every repair is recorded.
#[test]
fn malformed_parameter_shapes_take_words_interpretation_and_are_recorded() {
    let source =
        r"{\rtf1\ansi\hyphauto2\deff0{\fonttbl{\f0 Arial;}}\pard\fs A\page2 B\accdot0 C\par}";
    let document = RtfDocument::parse(source).unwrap();

    // A nonzero toggle parameter is on.
    assert_eq!(document.hyphenation().automatic, Some(true));
    // A missing value takes the keyword default (12 points).
    let block = |text: &str| {
        document
            .blocks()
            .iter()
            .find(|block| block.text == text)
            .unwrap_or_else(|| panic!("missing block {text:?}"))
    };
    assert_eq!(block("A").formatting.font_size.get(), 24);
    // A flag ignores its parameter.
    assert_eq!(document.page_breaks().count(), 1);
    // A zero parameter on a toggled selector turns it off, so `C` keeps the
    // formatting of `B` and shares its block.
    assert_eq!(block("BC").formatting.emphasis_mark, EmphasisMark::None);

    let at = |needle: &str| source.find(needle).unwrap() as u64;
    assert_eq!(
        recorded(document.tolerance_report()),
        vec![
            (
                Defect::Parameter(ParameterDefect::NonBinaryToggle),
                at(r"\hyphauto2")
            ),
            (
                Defect::Parameter(ParameterDefect::MissingParameter),
                at(r"\fs ")
            ),
            (
                Defect::Parameter(ParameterDefect::UnexpectedParameter),
                at(r"\page2")
            ),
            (
                Defect::Parameter(ParameterDefect::UnexpectedParameter),
                at(r"\accdot0")
            ),
        ]
    );
}

#[test]
fn well_formed_parameter_shapes_are_not_defects() {
    let source = r"{\rtf1\ansi\hyphauto1\deff0{\fonttbl{\f0 Arial;}}\pard\fs24 A\page B\accdot C\accnone\par}";
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.page_breaks().count(), 1);
    assert_clean(document.tolerance_report());
}

#[test]
fn parameter_shapes_in_destinations_are_repaired() {
    // Stylesheet, table-row, and section controls share the lexer's repair.
    let source = concat!(
        r"{\rtf1\ansi{\stylesheet{\s0\sbauto Normal;}{\*\cs10\fs Char;}}",
        r"\trowd\trgaph\trql2\cellx1000\pard\intbl X\cell\row",
        r"\sectd\vertsect1\pard Y\par}",
    );
    let document = RtfDocument::parse(source).unwrap();
    let report = document.tolerance_report();
    assert_eq!(report.total(), 5, "{report:?}");
    assert!(
        report
            .iter()
            .all(|tolerated| matches!(tolerated.defect, Defect::Parameter(_)))
    );
}
