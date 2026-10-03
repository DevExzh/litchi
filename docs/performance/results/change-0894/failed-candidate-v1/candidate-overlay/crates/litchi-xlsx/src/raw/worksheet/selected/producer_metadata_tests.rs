//! Differential qualification for the producer metadata admission profile.
//!
//! These tests exercise the selected scanner directly.  The materialized
//! worksheet parser remains the semantic oracle: an eligible selected result
//! must contain the same stored addresses and cells, in source order, and a
//! clean parser tolerance report must remain possible.  Metadata that the
//! parser can repair or ignore is deliberately tested as a scanner refusal;
//! accepting it would make a future selected route hide the authoritative
//! fallback's tolerance report.

#![allow(clippy::unwrap_used, reason = "focused parser differential assertions")]

use std::io::{BufReader, Cursor};

use litchi_ooxml_common::mce::{Capabilities, StreamError, StreamLimits};
use litchi_sheet::{Cell as Address, Rect};

use super::super::{parse, parse_refusing_repairs};
use super::{NotEligibleReason, RangeScanOutcome, SelectedPayload, scan_range};
use crate::cell::{Cell, Text, Value};

const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14AC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceStyle {
    Transitional,
    TransitionalWithAliases,
    StrictWithAliases,
}

impl NamespaceStyle {
    fn root(self) -> &'static str {
        match self {
            Self::Transitional | Self::TransitionalWithAliases => "worksheet",
            Self::StrictWithAliases => "s:worksheet",
        }
    }

    fn element(self, local: &str) -> String {
        match self {
            Self::Transitional | Self::TransitionalWithAliases => local.to_owned(),
            Self::StrictWithAliases => format!("s:{local}"),
        }
    }

    fn aliases(self) -> &'static str {
        match self {
            Self::Transitional => {
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            },
            Self::TransitionalWithAliases => concat!(
                r#"xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
                r#"xmlns:xa="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" "#,
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
                r#"m:Ignorable="xa""#,
            ),
            Self::StrictWithAliases => concat!(
                r#"xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
                r#"xmlns:xa="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" "#,
                r#"xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" "#,
                r#"m:Ignorable="xa""#,
            ),
        }
    }

    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional | Self::TransitionalWithAliases => TRANSITIONAL,
            Self::StrictWithAliases => STRICT,
        }
    }

    fn x14ac(self) -> &'static str {
        match self {
            Self::Transitional => "",
            Self::TransitionalWithAliases | Self::StrictWithAliases => r#" xa:dyDescent="0.25""#,
        }
    }

    fn row_x14ac(self) -> &'static str {
        self.x14ac()
    }
}

fn metadata_body(style: NamespaceStyle, cells: &str) -> String {
    let dimension = style.element("dimension");
    let sheet_views = style.element("sheetViews");
    let sheet_view = style.element("sheetView");
    let selection = style.element("selection");
    let sheet_format = style.element("sheetFormatPr");
    let cols = style.element("cols");
    let col = style.element("col");
    let sheet_data = style.element("sheetData");
    let page_margins = style.element("pageMargins");
    let page_setup = style.element("pageSetup");
    let x14ac = style.x14ac();
    format!(
        r#"<{dimension} ref="F7:I8"/>
<{sheet_views}><{sheet_view} workbookViewId="3"><{selection} activeCell="G8" sqref="G8:H8"/></{sheet_view}></{sheet_views}>
<{sheet_format} defaultRowHeight="15"{x14ac}/>
<{cols}><{col} min="6" max="7" width="11.5703125" bestFit="1" customWidth="1"/><{col} min="8" max="9" width="12" customWidth="1"/></{cols}>
<{sheet_data}>{cells}</{sheet_data}>
<{page_margins} left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
<{page_setup} paperSize="9" orientation="portrait" r:id="rIdPrinter"/>"#,
    )
}

fn metadata_xml(style: NamespaceStyle, cells: &str) -> String {
    let root = style.root();
    let aliases = style.aliases();
    format!(
        r#"<{root} xmlns="{namespace}" xmlns:s="{STRICT}" {aliases}>{body}</{root}>"#,
        namespace = style.namespace(),
        body = metadata_body(style, cells),
    )
}

fn alternate_metadata_xml(choice: &str, fallback: &str) -> String {
    format!(
        r#"<worksheet xmlns="{TRANSITIONAL}" xmlns:m="{MC}" xmlns:xa="{X14AC}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" m:Ignorable="xa"><m:AlternateContent><m:Choice Requires="xa">{choice}</m:Choice><m:Fallback>{fallback}</m:Fallback></m:AlternateContent></worksheet>"#,
        choice = choice,
        fallback = fallback,
    )
}

fn process_content_metadata_xml(active: bool, body: &str) -> String {
    let process_target = if active {
        "future:payload"
    } else {
        "future:other"
    };
    let tail = if active {
        String::new()
    } else {
        r#"<sheetData><row r="1"><c r="A1"><v>17</v></c></row></sheetData>"#.to_owned()
    };
    format!(
        r#"<worksheet xmlns="{TRANSITIONAL}" xmlns:m="{MC}" xmlns:xa="{X14AC}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:future="urn:litchi:future" m:Ignorable="xa future" m:ProcessContent="{process_target}"><future:payload>{body}</future:payload>{tail}</worksheet>"#,
        process_target = process_target,
        body = body,
        tail = tail,
    )
}

fn scalar_cells(style: NamespaceStyle) -> String {
    let c = style.element("c");
    let v = style.element("v");
    let is = style.element("is");
    let t = style.element("t");
    let first = format!(
        r#"<{c} r="F7"><{v}>-0.000</{v}></{c}>
<{c} r="G7" t="b"><{v}>true</{v}></{c}>
<{c} r="H7" t="str"><{v}>plain&amp;text</{v}></{c}>
<{c} r="I7"/>"#,
    );
    let second = format!(
        r#"<{c} r="F8" t="e"><{v}>#DIV/0!</{v}></{c}>
<{c} r="G8" t="inlineStr"><{is}><{t}>inline text</{t}></{is}></{c}>"#,
    );
    rows(style, &first, &second)
}

fn shared_cells(style: NamespaceStyle) -> String {
    let c = style.element("c");
    let v = style.element("v");
    let first = format!(
        r#"<{c} r="F7" t="s"><{v}>0</{v}></{c}>
<{c} r="G7" t="s"><{v}>3</{v}></{c}>"#,
    );
    rows(style, &first, "")
}

fn rows(style: NamespaceStyle, first: &str, second: &str) -> String {
    let row = style.element("row");
    let row_x14ac = style.row_x14ac();
    if second.is_empty() {
        format!(r#"<{row} r="7" spans="6:9"{row_x14ac}>{first}</{row}>"#,)
    } else {
        format!(
            r#"<{row} r="7" spans="6:9"{row_x14ac}>{first}</{row}>
<{row} r="8" spans="6:9"{row_x14ac}>{second}</{row}>"#,
        )
    }
}

fn scan_xml(xml: &str) -> Result<RangeScanOutcome, StreamError<crate::Error, crate::Error>> {
    scan_xml_with(xml, &Capabilities::default())
}

fn scan_xml_with(
    xml: &str,
    capabilities: &Capabilities,
) -> Result<RangeScanOutcome, StreamError<crate::Error, crate::Error>> {
    let mut input = BufReader::new(Cursor::new(xml.as_bytes()));
    scan_range(
        &mut input,
        capabilities,
        &StreamLimits::default(),
        Rect::ALL,
    )
}

fn parse_cells(xml: &str, strings: &[Text]) -> Vec<(Address, Cell)> {
    let store = parse(xml.as_bytes(), || Ok(Some(strings))).expect("authoritative worksheet parse");
    assert!(
        store.tolerance_report().is_clean(),
        "positive metadata fixture unexpectedly required repairs: {:?}",
        store.tolerance_report()
    );
    assert!(
        parse_refusing_repairs(xml.as_bytes(), || Ok(Some(strings))).is_ok(),
        "write-path admission rejected clean metadata fixture"
    );
    store
        .entries()
        .iter()
        .map(|entry| (entry.address, entry.cell.clone()))
        .collect()
}

fn assert_differential(xml: &str, strings: &[Text]) {
    assert_differential_against(xml, xml, strings, &Capabilities::default());
}

fn assert_differential_against(
    selected_xml: &str,
    authoritative_xml: &str,
    strings: &[Text],
    capabilities: &Capabilities,
) {
    let expected = parse_cells(authoritative_xml, strings);
    let outcome = scan_xml_with(selected_xml, capabilities).expect("selected metadata scan");
    let RangeScanOutcome::Eligible(selected) = outcome else {
        panic!("clean producer metadata should be eligible: {outcome:?}");
    };

    let mut actual = Vec::with_capacity(selected.cells.len());
    let mut max_shared = None;
    for record in &selected.cells {
        let cell = match &record.payload {
            SelectedPayload::Cell(cell) => cell.clone(),
            SelectedPayload::SharedString(index) => {
                max_shared = Some(max_shared.map_or(*index, |current: u32| current.max(*index)));
                Cell::Value(Value::Text(
                    strings
                        .get(*index as usize)
                        .unwrap_or_else(|| panic!("selected shared-string index {index} missing"))
                        .clone(),
                ))
            },
        };
        actual.push((record.address, cell));
    }
    assert_eq!(
        actual, expected,
        "selected records diverged from materialized store"
    );
    assert_eq!(selected.dependencies.max_shared_string_index, max_shared);
    assert_eq!(selected.dependencies.max_direct_style_index, None);
    assert_eq!(selected.dependencies.target_shared_string_index, None);
}

fn assert_not_eligible(xml: &str, label: &str) -> NotEligibleReason {
    match scan_xml(xml) {
        Ok(RangeScanOutcome::NotEligible(reason)) => reason,
        Ok(RangeScanOutcome::Eligible(_)) => panic!("{label} was silently admitted"),
        Err(error) => panic!("{label} produced a stream error instead of fallback: {error}"),
    }
}

fn replace_once(xml: &str, old: &str, new: &str) -> String {
    assert_eq!(
        xml.matches(old).count(),
        1,
        "fixture mutation anchor is unique"
    );
    xml.replacen(old, new, 1)
}

#[test]
fn clean_metadata_variants_match_authoritative_values_order_and_dependencies() {
    let transitional = metadata_xml(
        NamespaceStyle::Transitional,
        &scalar_cells(NamespaceStyle::Transitional),
    );
    assert_differential(&transitional, &[]);

    let aliases = metadata_xml(
        NamespaceStyle::TransitionalWithAliases,
        &scalar_cells(NamespaceStyle::TransitionalWithAliases),
    );
    assert_differential(&aliases, &[]);

    let strict = metadata_xml(
        NamespaceStyle::StrictWithAliases,
        &scalar_cells(NamespaceStyle::StrictWithAliases),
    );
    assert_differential(&strict, &[]);

    // The prefixed strict worksheet remains authoritative even when a
    // producer leaves a transitional default namespace declaration on the
    // root. Keep this mismatch covered separately from the clean strict
    // package above: the selected route follows expanded names, as does the
    // materialized parser.
    let strict_mismatch = replace_once(
        &strict,
        &format!(r#"xmlns="{STRICT}""#),
        &format!(r#"xmlns="{TRANSITIONAL}""#),
    );
    assert_differential(&strict_mismatch, &[]);

    let shared = metadata_xml(
        NamespaceStyle::TransitionalWithAliases,
        &shared_cells(NamespaceStyle::TransitionalWithAliases),
    );
    let strings = [
        Text::from("shared-zero"),
        Text::from("unused-one"),
        Text::from("unused-two"),
        Text::from("shared-three"),
    ];
    assert_differential(&shared, &strings);
    let outcome = scan_xml(&shared).expect("shared-string metadata scan");
    let RangeScanOutcome::Eligible(selected) = outcome else {
        panic!("valid shared-string references should remain eligible");
    };
    assert_eq!(selected.cells.len(), 2);
    assert!(matches!(
        selected.cells[0].payload,
        SelectedPayload::SharedString(0)
    ));
    assert!(matches!(
        selected.cells[1].payload,
        SelectedPayload::SharedString(3)
    ));
    assert_eq!(selected.dependencies.max_shared_string_index, Some(3));
    assert_eq!(selected.dependencies.max_direct_style_index, None);
    assert_eq!(selected.dependencies.target_shared_string_index, None);
}

#[test]
fn mce_metadata_branches_and_process_content_follow_the_authoritative_route() {
    let style = NamespaceStyle::TransitionalWithAliases;
    let cells = scalar_cells(style);
    let clean = metadata_body(style, &cells);
    let authoritative = metadata_xml(style, &cells);

    let mut x14ac_capabilities = Capabilities::default();
    x14ac_capabilities.understand_namespace(X14AC);
    let active_choice = alternate_metadata_xml(
        &clean,
        r#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><sheetData/>"#,
    );
    assert_differential_against(&active_choice, &authoritative, &[], &x14ac_capabilities);

    let inactive_choice = alternate_metadata_xml(
        r#"<future:producerMetadata xmlns:future="urn:litchi:future"><sheetViews future="ignored"/></future:producerMetadata>"#,
        &clean,
    );
    assert_differential(&inactive_choice, &[]);

    let active_process = process_content_metadata_xml(true, &clean);
    assert_differential_against(
        &active_process,
        &authoritative,
        &[],
        &Capabilities::default(),
    );

    // Without a matching ProcessContent target, the future wrapper is
    // inactive and its metadata children disappear before either parser sees
    // them. The ordinary sheetData tail remains the complete cell oracle.
    let inactive_process = process_content_metadata_xml(
        false,
        r#"<sheetViews><sheetView workbookViewId="9"/></sheetViews><sheetData><row r="1"><c r="A1"><v>99</v></c></row></sheetData>"#,
    );
    let inactive_authoritative = format!(
        r#"<worksheet xmlns="{TRANSITIONAL}"><sheetData><row r="1"><c r="A1"><v>17</v></c></row></sheetData></worksheet>"#
    );
    assert_differential_against(
        &inactive_process,
        &inactive_authoritative,
        &[],
        &Capabilities::default(),
    );
}

#[test]
fn metadata_unknown_markup_and_style_inheritance_fall_back() {
    let style = NamespaceStyle::Transitional;
    let base = metadata_xml(style, &scalar_cells(style));
    let cases = [
        (
            "unknown sheetViews attribute",
            replace_once(&base, "<sheetViews>", r#"<sheetViews future="keep">"#),
        ),
        (
            "unknown sheetView attribute",
            replace_once(
                &base,
                r#"workbookViewId="3""#,
                r#"workbookViewId="3" future="keep""#,
            ),
        ),
        (
            "unknown sheetViews child",
            replace_once(
                &base,
                r#"<selection activeCell="G8" sqref="G8:H8"/>"#,
                r#"<selection activeCell="G8" sqref="G8:H8"><future/></selection>"#,
            ),
        ),
        (
            "unknown pageMargins attribute",
            replace_once(
                &base,
                r#"<pageMargins left="0.7""#,
                r#"<pageMargins future="keep" left="0.7""#,
            ),
        ),
        (
            "unknown pageSetup attribute",
            replace_once(
                &base,
                r#"<pageSetup paperSize="9""#,
                r#"<pageSetup future="keep" paperSize="9""#,
            ),
        ),
        (
            "styled producer column",
            replace_once(
                &base,
                r#"<col min="6" max="7""#,
                r#"<col min="6" max="7" style="2""#,
            ),
        ),
        (
            "overlapping producer columns",
            replace_once(&base, r#"<col min="8" max="9""#, r#"<col min="7" max="9""#),
        ),
        (
            "invalid row spans",
            replace_once(
                &base,
                r#"<row r="7" spans="6:9""#,
                r#"<row r="7" spans="0:9""#,
            ),
        ),
    ];
    for (label, xml) in cases {
        assert_not_eligible(&xml, label);
    }
}

#[test]
fn repairable_duplicate_and_misordered_metadata_never_becomes_selected_success() {
    let style = NamespaceStyle::Transitional;
    let base = metadata_xml(style, &scalar_cells(style));

    let dimension_after_cols = replace_once(
        &replace_once(&base, r#"<dimension ref="F7:I8"/>"#, ""),
        "</cols>",
        r#"</cols><dimension ref="F7:I8"/>"#,
    );
    assert_not_eligible(&dimension_after_cols, "dimension after cols");
    let dimension_after_cols_store = parse(dimension_after_cols.as_bytes(), || Ok(Some(&[])))
        .expect("read parser repairs dimension after cols");
    assert!(!dimension_after_cols_store.tolerance_report().is_clean());
    assert!(parse_refusing_repairs(dimension_after_cols.as_bytes(), || Ok(Some(&[]))).is_err());

    let invalid_dimension = replace_once(&base, r#"ref="F7:I8""#, r#"ref="not-a-range""#);
    assert_not_eligible(&invalid_dimension, "invalid dimension");
    let dimension_store = parse(&invalid_dimension.as_bytes(), || Ok(Some(&[])))
        .expect("read parser repairs invalid dimension");
    assert!(!dimension_store.tolerance_report().is_clean());
    assert!(parse_refusing_repairs(invalid_dimension.as_bytes(), || Ok(Some(&[]))).is_err());

    let invalid_default = replace_once(
        &base,
        r#"defaultRowHeight="15""#,
        r#"defaultRowHeight="NaN""#,
    );
    assert_not_eligible(&invalid_default, "repairable default row height");
    let default_store = parse(&invalid_default.as_bytes(), || Ok(Some(&[])))
        .expect("read parser repairs invalid default row height");
    assert!(!default_store.tolerance_report().is_clean());
    assert!(parse_refusing_repairs(invalid_default.as_bytes(), || Ok(Some(&[]))).is_err());

    let duplicate_view = replace_once(
        &base,
        "</sheetViews>",
        r#"</sheetViews><sheetViews><sheetView workbookViewId="3"/></sheetViews>"#,
    );
    assert_not_eligible(&duplicate_view, "duplicate sheetViews");

    let duplicate_format = replace_once(
        &base,
        r#"<cols>"#,
        r#"<sheetFormatPr defaultRowHeight="15"/><cols>"#,
    );
    assert_not_eligible(&duplicate_format, "duplicate sheetFormatPr");

    let late_views = replace_once(
        &replace_once(
            &base,
            r#"<sheetViews><sheetView workbookViewId="3"><selection activeCell="G8" sqref="G8:H8"/></sheetView></sheetViews>"#,
            "",
        ),
        r#"<sheetFormatPr defaultRowHeight="15"/>"#,
        r#"<sheetFormatPr defaultRowHeight="15"/><sheetViews><sheetView workbookViewId="3"/></sheetViews>"#,
    );
    assert_not_eligible(&late_views, "sheetViews after sheetFormatPr");

    let late_format = replace_once(
        &replace_once(&base, r#"<sheetFormatPr defaultRowHeight="15"/>"#, ""),
        "</sheetData>",
        r#"</sheetData><sheetFormatPr defaultRowHeight="15"/>"#,
    );
    assert_not_eligible(&late_format, "late sheetFormatPr");
    let late_store = parse(late_format.as_bytes(), || Ok(Some(&[])))
        .expect("read parser repairs late sheetFormatPr");
    assert!(!late_store.tolerance_report().is_clean());
    assert!(parse_refusing_repairs(late_format.as_bytes(), || Ok(Some(&[]))).is_err());

    let late_margins = replace_once(
        &replace_once(
            &base,
            r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
            "",
        ),
        "</cols>",
        r#"</cols><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
    );
    assert_not_eligible(&late_margins, "pageMargins before sheetData");

    let late_columns = replace_once(
        &replace_once(
            &base,
            r#"<cols><col min="6" max="7" width="11.5703125" bestFit="1" customWidth="1"/><col min="8" max="9" width="12" customWidth="1"/></cols>"#,
            "",
        ),
        "</sheetData>",
        r#"</sheetData><cols><col min="6" max="9" width="11"/></cols>"#,
    );
    assert_not_eligible(&late_columns, "late cols");
}

#[test]
fn malformed_extension_and_limits_remain_typed_stream_failures() {
    let style = NamespaceStyle::TransitionalWithAliases;
    let base = metadata_xml(style, &scalar_cells(style));

    let malformed = replace_once(&base, "</sheetViews>", "</sheetView>");
    match scan_xml(&malformed) {
        Err(error) => assert!(
            error.to_string().contains("XML") || error.to_string().contains("worksheet"),
            "malformed metadata lost its typed XML context: {error}"
        ),
        Ok(RangeScanOutcome::Eligible(_)) | Ok(RangeScanOutcome::NotEligible(_)) => {
            panic!("malformed metadata was converted into a capability result")
        },
    }

    let invalid_x14ac = replace_once(
        &base,
        r#"<sheetFormatPr defaultRowHeight="15" xa:dyDescent="0.25"/>"#,
        r#"<sheetFormatPr defaultRowHeight="15" xa:dyDescent="NaN"/>"#,
    );
    match scan_xml(&invalid_x14ac) {
        Err(StreamError::Callback { active_error, .. }) => {
            let error = active_error.expect("x14ac callback error");
            assert!(matches!(error, crate::Error::Descent(_)), "{error}");
        },
        Err(error) => panic!("invalid x14ac changed error class: {error}"),
        Ok(RangeScanOutcome::Eligible(_)) | Ok(RangeScanOutcome::NotEligible(_)) => {
            panic!("invalid x14ac was silently admitted")
        },
    }

    let mut limits = StreamLimits::default();
    limits.processing.max_input_bytes = base.len().saturating_sub(1);
    let mut input = BufReader::new(Cursor::new(base.as_bytes()));
    match scan_range(&mut input, &Capabilities::default(), &limits, Rect::ALL) {
        Err(error) => assert!(
            error.to_string().contains("limit") || error.to_string().contains("input"),
            "input limit error lost its resource context: {error}"
        ),
        Ok(RangeScanOutcome::Eligible(_)) | Ok(RangeScanOutcome::NotEligible(_)) => {
            panic!("input limit was converted into a successful scanner result")
        },
    }
}
