//! Characterization of the worksheet readers' implicit state: the
//! inline-string item machine of a cell and the worksheet child order.
//!
//! Every row pins what the materialized parser answers when it repairs,
//! what it answers when it refuses repairs, and what the selected-cell scan
//! answers for the same bytes.

use std::io::Cursor;

use litchi_ooxml_common::mce::{Capabilities, StreamLimits};
use litchi_sheet::Cell as Address;

use super::selected::{ScanOutcome, scan};
use super::{parse, parse_refusing_repairs};
use crate::pinned_rows::assert_rows;

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn a1() -> Address {
    Address::at(0, 0).expect("A1")
}

fn materialized(xml: &str, repair: bool) -> String {
    let parsed = if repair {
        parse(xml.as_bytes(), || Ok(None))
    } else {
        parse_refusing_repairs(xml.as_bytes(), || Ok(None))
    };
    match parsed {
        Ok(store) => {
            let cell = store
                .entry(a1())
                .map(|entry| format!("{:?} rich={}", entry.cell, entry.inline_rich));
            format!(
                "ok cell={cell:?} merges={} repairs={}",
                store.merge_ranges().len(),
                store.tolerance_report().iter().count()
            )
        },
        Err(error) => error.to_string(),
    }
}

fn selected(xml: &str) -> String {
    let mut input = Cursor::new(xml.as_bytes());
    match scan(
        &mut input,
        &Capabilities::default(),
        &StreamLimits::default(),
        a1(),
    ) {
        Ok(ScanOutcome::Eligible(cell)) => format!("eligible {:?}", cell.cell),
        Ok(ScanOutcome::NotEligible(reason)) => format!("{reason:?}"),
        Err(error) => error.to_string(),
    }
}

fn outcome(xml: &str) -> String {
    format!(
        "{} | {} | {}",
        materialized(xml, true),
        materialized(xml, false),
        selected(xml)
    )
}

/// Inline-string bodies: valid simple and rich items, each mixed pair,
/// duplicates and runs without text.
const INLINE_STRINGS: &[(&str, &str)] = &[
    (
        r#""#,
        "ok cell=Some(\"Value(Text(\\\"\\\")) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"\\\")) rich=false\") merges=0 repairs=0 | eligible Some(Value(Text(\"\")))",
    ),
    (
        r#"<t>a</t>"#,
        "ok cell=Some(\"Value(Text(\\\"a\\\")) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"a\\\")) rich=false\") merges=0 repairs=0 | eligible Some(Value(Text(\"a\")))",
    ),
    (
        r#"<t/>"#,
        "ok cell=Some(\"Value(Text(\\\"\\\")) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"\\\")) rich=false\") merges=0 repairs=0 | eligible Some(Value(Text(\"\")))",
    ),
    (
        r#"<r><t>a</t></r>"#,
        "ok cell=Some(\"Value(Text(\\\"a\\\")) rich=true\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"a\\\")) rich=true\") merges=0 repairs=0 | RichInlineText",
    ),
    (
        r#"<r><t>a</t></r><r><t>b</t></r>"#,
        "ok cell=Some(\"Value(Text(\\\"ab\\\")) rich=true\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"ab\\\")) rich=true\") merges=0 repairs=0 | RichInlineText",
    ),
    (
        r#"<t>a</t><t>b</t>"#,
        "invalid XLSX structure: inline string mixes or duplicates text | invalid XLSX structure: inline string mixes or duplicates text | MCE stream observer failed (active observer: invalid XLSX structure: inline string mixes or duplicates text)",
    ),
    (
        r#"<t>a</t><r><t>b</t></r>"#,
        "invalid XLSX structure: inline string mixes simple and rich text | invalid XLSX structure: inline string mixes simple and rich text | RichInlineText",
    ),
    (
        r#"<r><t>a</t></r><t>b</t>"#,
        "invalid XLSX structure: inline string mixes or duplicates text | invalid XLSX structure: inline string mixes or duplicates text | RichInlineText",
    ),
    (
        r#"<r><t>a</t><t>b</t></r>"#,
        "invalid XLSX structure: rich-text run has duplicate text | invalid XLSX structure: rich-text run has duplicate text | RichInlineText",
    ),
    (
        r#"<r><t>a</t></r><r><t>b</t><t>c</t></r>"#,
        "invalid XLSX structure: rich-text run has duplicate text | invalid XLSX structure: rich-text run has duplicate text | RichInlineText",
    ),
    (
        r#"<t>a</t><r><t>b</t></r><r><t>c</t></r>"#,
        "invalid XLSX structure: inline string mixes simple and rich text | invalid XLSX structure: inline string mixes simple and rich text | RichInlineText",
    ),
    (
        r#"<r><t>a</t></r><t>b</t><r><t>c</t></r>"#,
        "invalid XLSX structure: inline string mixes or duplicates text | invalid XLSX structure: inline string mixes or duplicates text | RichInlineText",
    ),
    (
        r#"<r><t>a</t></r><t>b</t><t>c</t>"#,
        "invalid XLSX structure: inline string mixes or duplicates text | invalid XLSX structure: inline string mixes or duplicates text | RichInlineText",
    ),
    (
        r#"<t>a</t><r><t>b</t><t>c</t></r>"#,
        "invalid XLSX structure: inline string mixes simple and rich text | invalid XLSX structure: inline string mixes simple and rich text | RichInlineText",
    ),
    (
        r#"<t>a</t><t>b</t><r><t>c</t></r>"#,
        "invalid XLSX structure: inline string mixes or duplicates text | invalid XLSX structure: inline string mixes or duplicates text | MCE stream observer failed (active observer: invalid XLSX structure: inline string mixes or duplicates text)",
    ),
    (
        r#"<r/>"#,
        "ok cell=Some(\"Value(Text(\\\"\\\")) rich=true\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"\\\")) rich=true\") merges=0 repairs=0 | RichInlineText",
    ),
    (
        r#"<r></r>"#,
        "ok cell=Some(\"Value(Text(\\\"\\\")) rich=true\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"\\\")) rich=true\") merges=0 repairs=0 | RichInlineText",
    ),
    (
        r#"<r><rPr/></r>"#,
        "ok cell=Some(\"Value(Text(\\\"\\\")) rich=true\") merges=0 repairs=0 | ok cell=Some(\"Value(Text(\\\"\\\")) rich=true\") merges=0 repairs=0 | RichInlineText",
    ),
    (
        r#"<t>a</t><r/>"#,
        "invalid XLSX structure: inline string mixes simple and rich text | invalid XLSX structure: inline string mixes simple and rich text | RichInlineText",
    ),
];

#[test]
fn inline_string_item_outcomes_are_pinned() {
    assert_rows(INLINE_STRINGS, |body| {
        outcome(&format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is>{body}</is></c></row></sheetData></worksheet>"#
        ))
    });
}

/// Cell children: a value, an inline string, and their duplicates and mixes.
const CELL_CHILDREN: &[(&str, &str)] = &[
    (
        r#"<v>1</v>"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"<v>1</v><v>2</v>"#,
        "invalid XLSX structure: duplicate worksheet cell value | invalid XLSX structure: duplicate worksheet cell value | MCE stream observer failed (active observer: invalid XLSX structure: duplicate worksheet cell value)",
    ),
    (
        r#"<is><t>a</t></is><is><t>b</t></is>"#,
        "invalid XLSX structure: duplicate worksheet inline string | invalid XLSX structure: duplicate worksheet inline string | MCE stream observer failed (active observer: invalid XLSX structure: duplicate worksheet inline string)",
    ),
    (
        r#"<v>1</v><is><t>a</t></is>"#,
        "invalid XLSX structure: worksheet cell contains both inline text and a value | invalid XLSX structure: worksheet cell contains both inline text and a value | MCE stream observer failed (active observer: invalid XLSX structure: worksheet cell contains both inline text and a value)",
    ),
    (
        r#"<is><t>a</t></is><v>1</v>"#,
        "invalid XLSX structure: worksheet cell contains both inline text and a value | invalid XLSX structure: worksheet cell contains both inline text and a value | MCE stream observer failed (active observer: invalid XLSX structure: worksheet cell contains both inline text and a value)",
    ),
];

#[test]
fn cell_child_outcomes_are_pinned() {
    assert_rows(CELL_CHILDREN, |body| {
        outcome(&format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1">{body}</c></row></sheetData></worksheet>"#
        ))
    });
}

/// One worksheet child per letter of a child-order row.
fn worksheet_child(letter: char) -> &'static str {
    match letter {
        'V' => r#"<sheetViews><sheetView workbookViewId="0"/></sheetViews>"#,
        'D' => r#"<dimension ref="A1:B1"/>"#,
        'F' => r#"<sheetFormatPr defaultRowHeight="15"/>"#,
        'C' => r#"<cols><col min="1" max="1" width="9"/></cols>"#,
        'S' => r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
        'M' => r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
        'P' => r#"<phoneticPr fontId="0"/>"#,
        other => panic!("unknown worksheet child letter {other}"),
    }
}

/// Worksheet child orders: every single child and ordered pair of
/// `dimension`, `sheetFormatPr`, `cols`, `sheetData`, `mergeCells` and a
/// `mergeCells` successor, every order of `sheetData` and `mergeCells`
/// around a third child, and a few complete worksheets.
const CHILD_ORDERS: &[(&str, &str)] = &[
    (
        r#"D"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | MCE stream observer failed (active observer: invalid XLSX structure: worksheet XML is missing required sheetData)",
    ),
    (
        r#"F"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | MCE stream observer failed (active observer: invalid XLSX structure: worksheet XML is missing required sheetData)",
    ),
    (
        r#"C"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | Styles",
    ),
    (
        r#"S"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"M"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"P"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"DD"#,
        "ok cell=None merges=0 repairs=1 | invalid XLSX structure: worksheet has duplicate dimension elements | UnsupportedStructure",
    ),
    (
        r#"DF"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | MCE stream observer failed (active observer: invalid XLSX structure: worksheet XML is missing required sheetData)",
    ),
    (
        r#"DC"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | Styles",
    ),
    (
        r#"DS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"DM"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"DP"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"FD"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | MCE stream observer failed (active observer: invalid XLSX structure: worksheet XML is missing required sheetData)",
    ),
    (
        r#"FF"#,
        "invalid XLSX structure: worksheet has duplicate sheetFormatPr elements | invalid XLSX structure: worksheet has duplicate sheetFormatPr elements | UnsupportedStructure",
    ),
    (
        r#"FC"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | Styles",
    ),
    (
        r#"FS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"FM"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"FP"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"CD"#,
        "ok cell=None merges=0 repairs=1 | invalid XLSX structure: worksheet dimension appears after column or cell data | Styles",
    ),
    (
        r#"CF"#,
        "ok cell=None merges=0 repairs=1 | invalid XLSX structure: worksheet sheetFormatPr appears after column or cell data | Styles",
    ),
    (
        r#"CC"#,
        "invalid XLSX structure: worksheet has duplicate cols elements | invalid XLSX structure: worksheet has duplicate cols elements | Styles",
    ),
    (
        r#"CS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | Styles",
    ),
    (
        r#"CM"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | Styles",
    ),
    (
        r#"CP"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | Styles",
    ),
    (
        r#"SD"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=1 | invalid XLSX structure: worksheet dimension appears after column or cell data | UnsupportedStructure",
    ),
    (
        r#"SF"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=1 | invalid XLSX structure: worksheet sheetFormatPr appears after column or cell data | UnsupportedStructure",
    ),
    (
        r#"SC"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=1 | invalid XLSX structure: worksheet cols appears after sheetData | Styles",
    ),
    (
        r#"SS"#,
        "invalid XLSX structure: worksheet has duplicate sheetData | invalid XLSX structure: worksheet has duplicate sheetData | MCE stream observer failed (active observer: invalid XLSX structure: worksheet has duplicate sheetData)",
    ),
    (
        r#"SM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"SP"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"MD"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MF"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MC"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MM"#,
        "invalid XLSX structure: worksheet has duplicate mergeCells elements | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MP"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"PD"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"PF"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"PC"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"PS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"PM"#,
        "ok cell=None merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | UnsupportedStructure",
    ),
    (
        r#"PP"#,
        "ok cell=None merges=0 repairs=0 | ok cell=None merges=0 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"CMS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | Styles",
    ),
    (
        r#"CSM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | Styles",
    ),
    (
        r#"DMS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"DSM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"FMS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"FSM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"MCS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MDS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MFS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MMS"#,
        "invalid XLSX structure: worksheet has duplicate mergeCells elements | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MPS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MSC"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=2 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MSD"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=2 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MSF"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=2 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MSM"#,
        "invalid XLSX structure: worksheet has duplicate mergeCells elements | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MSP"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"MSS"#,
        "invalid XLSX structure: worksheet has duplicate sheetData | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
    (
        r#"PMS"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears before sheetData | UnsupportedStructure",
    ),
    (
        r#"PSM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"SCM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet cols appears after sheetData | Styles",
    ),
    (
        r#"SDM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet dimension appears after column or cell data | UnsupportedStructure",
    ),
    (
        r#"SFM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet sheetFormatPr appears after column or cell data | UnsupportedStructure",
    ),
    (
        r#"SMC"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet cols appears after sheetData | Styles",
    ),
    (
        r#"SMD"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet dimension appears after column or cell data | UnsupportedStructure",
    ),
    (
        r#"SMF"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet sheetFormatPr appears after column or cell data | UnsupportedStructure",
    ),
    (
        r#"SMM"#,
        "invalid XLSX structure: worksheet has duplicate mergeCells elements | invalid XLSX structure: worksheet has duplicate mergeCells elements | MCE stream observer failed (active observer: invalid XLSX structure: worksheet has duplicate mergeCells elements)",
    ),
    (
        r#"SMP"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | UnsupportedStructure",
    ),
    (
        r#"SMS"#,
        "invalid XLSX structure: worksheet has duplicate sheetData | invalid XLSX structure: worksheet has duplicate sheetData | MCE stream observer failed (active observer: invalid XLSX structure: worksheet has duplicate sheetData)",
    ),
    (
        r#"SPM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears after a schema successor | UnsupportedStructure",
    ),
    (
        r#"SSM"#,
        "invalid XLSX structure: worksheet has duplicate sheetData | invalid XLSX structure: worksheet has duplicate sheetData | MCE stream observer failed (active observer: invalid XLSX structure: worksheet has duplicate sheetData)",
    ),
    (
        r#"DFCSMP"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | Styles",
    ),
    (
        r#"DFCSPM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears after a schema successor | Styles",
    ),
    (
        r#"VDFCSM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | Styles",
    ),
    (
        r#"SVM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=0 | eligible Some(Value(Number(Number(\"1\"))))",
    ),
    (
        r#"SPMP"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears after a schema successor | UnsupportedStructure",
    ),
    (
        r#"SMPD"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet dimension appears after column or cell data | UnsupportedStructure",
    ),
    (
        r#"SPPM"#,
        "ok cell=Some(\"Value(Number(Number(\\\"1\\\"))) rich=false\") merges=1 repairs=1 | invalid XLSX structure: worksheet mergeCells appears after a schema successor | UnsupportedStructure",
    ),
    (
        r#"MSPM"#,
        "invalid XLSX structure: worksheet has duplicate mergeCells elements | invalid XLSX structure: worksheet mergeCells appears before sheetData | MergeSemantics",
    ),
];

#[test]
fn worksheet_child_order_outcomes_are_pinned() {
    assert_rows(CHILD_ORDERS, |order| {
        let body = order.chars().map(worksheet_child).collect::<String>();
        outcome(&format!(r#"<worksheet xmlns="{S}">{body}</worksheet>"#))
    });
}
