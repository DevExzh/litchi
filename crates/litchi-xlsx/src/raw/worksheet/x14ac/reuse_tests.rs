//! Differential controls for the `capture_inner` namespace-resolution change.
//!
//! The `reference_capture` implementation below is an intentional test-only
//! copy of the baseline `capture_inner`/`start`/`descent` path from commit
//! `3348479699` (the source-file SHA-256 is
//! `d226c59e7174c5977ffd2daa889eac89d147a22c404be72f8e9941650a0381fa`).  It
//! must not call the optimized implementation's helpers:
//! these tests are meant to catch a resolver-state or early-return change that
//! still happens to agree with a small hand-written fixture.

use std::{collections::BTreeMap, io::Cursor, mem, str};

use litchi_ooxml_common::mce::{Capabilities, StreamLimits};
use litchi_ooxml_common::xml::{attributes::BytesStartExt as _, unqualified_attribute_value};
use litchi_opc::OpcPackage;
use litchi_sheet::ROWS;
use quick_xml::{
    XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
    name::{Namespace, NamespaceResolver, ResolveResult},
};

use super::{Context, Values};
use crate::{
    error::{Error, Result, invalid},
    layout::Descent,
};

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14AC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_XML_DEPTH: usize = 256;

#[derive(Debug, PartialEq)]
struct Signature {
    defaults: Option<f64>,
    rows: Vec<(u32, f64)>,
}

fn signature(values: &Values) -> Signature {
    Signature {
        defaults: values.defaults.map(Descent::get),
        rows: values
            .rows
            .iter()
            .map(|(&row, &value)| (row, value.get()))
            .collect(),
    }
}

fn assert_matches_reference(name: &str, xml: &str, capture_rows: bool) {
    let expected = reference_capture(xml.as_bytes(), capture_rows);
    let actual = super::capture_inner(xml.as_bytes(), capture_rows);
    assert_same_outcome(name, expected, actual);
}

fn assert_same_outcome(name: &str, expected: Result<Values>, actual: Result<Values>) {
    match (expected, actual) {
        (Ok(expected), Ok(actual)) => {
            assert_eq!(signature(&actual), signature(&expected), "{name}");
        },
        (Err(expected), Err(actual)) => {
            assert_eq!(
                mem::discriminant(&actual),
                mem::discriminant(&expected),
                "{name}: error type changed: {actual:?}"
            );
            assert_eq!(actual.to_string(), expected.to_string(), "{name}");
        },
        (expected, actual) => {
            panic!("{name}: outcome changed: expected {expected:?}, got {actual:?}")
        },
    }
}

fn worksheet(body: &str) -> String {
    format!(r#"<worksheet xmlns="{S}" xmlns:x14ac="{X14AC}">{body}</worksheet>"#)
}

fn prefixed_worksheet(body: &str) -> String {
    format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:a="{X14AC}" xmlns="urn:default-other">{body}</x:worksheet>"#
    )
}

fn strict_worksheet(body: &str) -> String {
    format!(r#"<worksheet xmlns="{STRICT_S}" xmlns:a="{X14AC}">{body}</worksheet>"#)
}

fn alternate(choice: &str, fallback: &str) -> String {
    format!(
        r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x14ac="{X14AC}" mc:Ignorable="x14ac"><mc:AlternateContent><mc:Choice Requires="x14ac">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent></worksheet>"#
    )
}

#[test]
fn optimized_capture_matches_baseline_across_namespace_spellings() {
    let cases = [
        (
            "default",
            worksheet(
                r#"<sheetFormatPr x14ac:dyDescent="0.125"/><sheetData><row r="2" x14ac:dyDescent="0.25"/><row r="4" x14ac:dyDescent="0.5"/></sheetData>"#,
            ),
        ),
        (
            "strict",
            strict_worksheet(
                r#"<sheetFormatPr a:dyDescent="0.125"/><sheetData><row r="2" a:dyDescent="0.25"/></sheetData>"#,
            ),
        ),
        (
            "prefix",
            prefixed_worksheet(
                r#"<x:sheetFormatPr a:dyDescent="0.125"/><x:sheetData><x:row r="2" a:dyDescent="0.25"/></x:sheetData>"#,
            ),
        ),
        (
            "alias",
            format!(
                r#"<worksheet xmlns="{S}" xmlns:first="{X14AC}" xmlns:second="{X14AC}"><sheetFormatPr first:dyDescent="0.125"/><sheetData><row r="2" second:dyDescent="0.25"/></sheetData></worksheet>"#
            ),
        ),
    ];

    for (name, xml) in cases {
        assert_matches_reference(name, &xml, true);
        assert_matches_reference(&format!("{name}-defaults-only"), &xml, false);
    }
}

#[test]
fn optimized_capture_keeps_namespace_reader_state_after_rebinding() {
    let xml = format!(
        r#"<worksheet xmlns="{S}" xmlns:x14ac="{X14AC}" xmlns:a="{X14AC}"><sheetFormatPr x14ac:dyDescent="0.1"/><sheetData xmlns:a="urn:not-x14ac"><row r="1" a:dyDescent="0.2"/></sheetData><sheetData xmlns:a="{X14AC}"><row r="2" a:dyDescent="0.3"/></sheetData><sheetData xmlns:a="urn:other" xmlns:b="{X14AC}"><row r="3" a:dyDescent="0.4" b:dyDescent="0.5"/></sheetData></worksheet>"#
    );
    assert_matches_reference("rebind-and-restore", &xml, true);
    assert_matches_reference("rebind-and-restore-defaults-only", &xml, false);

    let actual = super::capture_inner(xml.as_bytes(), true).expect("rebind fixture");
    assert_eq!(
        signature(&actual),
        Signature {
            defaults: Some(0.1),
            rows: vec![(2, 0.3), (3, 0.5)]
        }
    );
}

#[test]
fn optimized_capture_matches_reference_for_non_structural_unknown_prefixes() {
    let cases = [
        worksheet(
            r#"<sheetFormatPr xmlns:unknown="urn:unknown" unknown:unused="ok" x14ac:dyDescent="0.2"/><sheetData><row xmlns:unknown="urn:unknown" unknown:unused="ok" r="1" x14ac:dyDescent="0.3"/></sheetData>"#,
        ),
        prefixed_worksheet(
            r#"<x:opaque xmlns:z="urn:unknown" z:unused="ok"/><x:sheetData><x:row r="1" a:dyDescent="0.3"/></x:sheetData>"#,
        ),
        format!(
            r#"<worksheet xmlns="{S}" xmlns:q="urn:unknown"><sheetData><row r="1" q:unused="ok"/><row r="2" x14ac:dyDescent="0.4"/></sheetData></worksheet>"#
        ),
    ];

    for (index, xml) in cases.iter().enumerate() {
        assert_matches_reference(&format!("unknown-prefix-{index}"), xml, true);
        assert_matches_reference(&format!("unknown-prefix-{index}-defaults-only"), xml, false);
    }
}

#[test]
fn optimized_capture_matches_reference_for_errors_and_limits() {
    let cases = [
        (
            "malformed-end",
            worksheet(r#"<sheetData><row r="1" x14ac:dyDescent="0.2"></sheetData></worksheet>"#),
        ),
        (
            "malformed-attribute",
            worksheet(
                r#"<sheetData><row r="1" x14ac:dyDescent="0.2" bad="unterminated></row></sheetData>"#,
            ),
        ),
        (
            "invalid-descent",
            worksheet(r#"<sheetFormatPr x14ac:dyDescent="NaN"/>"#),
        ),
        (
            "negative-descent",
            worksheet(r#"<sheetFormatPr x14ac:dyDescent="-0.1"/>"#),
        ),
        (
            "duplicate-default",
            worksheet(
                r#"<sheetFormatPr x14ac:dyDescent="0.1"/><sheetFormatPr x14ac:dyDescent="0.2"/>"#,
            ),
        ),
        (
            "duplicate-row",
            worksheet(
                r#"<sheetData><row r="1" x14ac:dyDescent="0.1"/><row r="1" x14ac:dyDescent="0.2"/></sheetData>"#,
            ),
        ),
        (
            "zero-row",
            worksheet(r#"<sheetData><row r="0" x14ac:dyDescent="0.1"/></sheetData>"#),
        ),
        (
            "row-overflow",
            format!(
                r#"<worksheet xmlns="{S}" xmlns:x14ac="{X14AC}"><sheetData><row r="{}" x14ac:dyDescent="0.1"/></sheetData></worksheet>"#,
                ROWS + 1
            ),
        ),
        (
            "inferred-row-overflow",
            format!(
                r#"<worksheet xmlns="{S}" xmlns:x14ac="{X14AC}"><sheetData><row r="{}"/><row x14ac:dyDescent="0.1"/></sheetData></worksheet>"#,
                ROWS
            ),
        ),
        (
            "reserved-marker",
            worksheet(r#"<sheetFormatPr litchi_x14ac_dyDescent="0.2"/>"#),
        ),
    ];

    for (name, xml) in cases {
        assert_matches_reference(name, &xml, true);
        assert_matches_reference(&format!("{name}-defaults-only"), &xml, false);
    }
}

#[test]
fn optimized_capture_matches_the_same_explicit_depth_boundary() {
    let mut body = String::new();
    // The worksheet root occupies level one, so 255 nested starts are the
    // exact 256-level boundary.  The empty probe does not add a stack level.
    for _ in 0..(MAX_XML_DEPTH - 1) {
        body.push_str("<opaque>");
    }
    body.push_str("<probe/>");
    for _ in 0..(MAX_XML_DEPTH - 1) {
        body.push_str("</opaque>");
    }
    let xml = worksheet(&body);
    assert_matches_reference("depth-at-limit", &xml, true);
    assert_matches_reference("depth-at-limit-defaults-only", &xml, false);
    for capture_rows in [true, false] {
        let values = super::capture_inner(xml.as_bytes(), capture_rows)
            .expect("exact maximum depth must remain accepted");
        assert_eq!(
            signature(&values),
            Signature {
                defaults: None,
                rows: vec![]
            }
        );
    }

    let mut too_deep = String::new();
    for _ in 0..MAX_XML_DEPTH {
        too_deep.push_str("<opaque>");
    }
    too_deep.push_str("<probe/>");
    for _ in 0..MAX_XML_DEPTH {
        too_deep.push_str("</opaque>");
    }
    let too_deep = worksheet(&too_deep);
    assert_matches_reference("depth-over-limit", &too_deep, true);
    assert_matches_reference("depth-over-limit-defaults-only", &too_deep, false);
    for capture_rows in [true, false] {
        let error = super::capture_inner(too_deep.as_bytes(), capture_rows)
            .expect_err("one level over the limit must be refused");
        assert!(matches!(
            error,
            Error::Invalid(message)
                if message == format!("worksheet extension XML exceeds {MAX_XML_DEPTH} levels")
        ));
    }
}

#[test]
fn optimized_capture_matches_reference_for_row_descent_edge_values() {
    let xml = worksheet(
        r#"<sheetData><row r="1" x14ac:dyDescent="0"/><row r="2" x14ac:dyDescent="1.25"/><row r="3"/><row x14ac:dyDescent="2.5"/></sheetData>"#,
    );
    assert_matches_reference("row-edge-values", &xml, true);
    assert_matches_reference("row-edge-values-defaults-only", &xml, false);

    let actual = super::capture_inner(xml.as_bytes(), true).expect("edge values");
    assert_eq!(
        signature(&actual),
        Signature {
            defaults: None,
            rows: vec![(1, 0.0), (2, 1.25), (4, 2.5)]
        }
    );
}

#[test]
fn mce_alternate_has_independent_selected_branch_controls() {
    let fallback = alternate(
        r#"<sheetFormatPr x14ac:dyDescent="NaN"/><sheetData><row r="0" x14ac:dyDescent="-1"/></sheetData>"#,
        r#"<sheetFormatPr x14ac:dyDescent="0.25"/><sheetData><row r="2" x14ac:dyDescent="0.35"/></sheetData>"#,
    );
    let values = super::capture(fallback.as_bytes()).expect("fallback branch");
    assert_eq!(
        signature(&values),
        Signature {
            defaults: Some(0.25),
            rows: vec![(2, 0.35)]
        }
    );

    let choice = alternate(
        r#"<sheetFormatPr x14ac:dyDescent="0.2"/><sheetData><row r="1" x14ac:dyDescent="0.3"/></sheetData>"#,
        r#"<sheetFormatPr x14ac:dyDescent="NaN"/><sheetData><row r="0" x14ac:dyDescent="-1"/></sheetData>"#,
    );
    // The default capability set does not understand x14ac, so the fallback
    // is selected.  Its malformed value is rejected by `selected_start`; the
    // raw observer only records the lexical candidate and does not parse it.
    let error = super::capture(choice.as_bytes()).expect_err("selected fallback validation");
    assert!(matches!(error, Error::Descent(_)), "{error:?}");

    let mut raw_choice = Capabilities::default();
    raw_choice.understand_namespace(X14AC);
    let mut input = Cursor::new(choice.as_bytes());
    let values =
        super::capture_stream_legacy(&mut input, &raw_choice, &StreamLimits::default(), true)
            .expect("supported choice branch");
    assert_eq!(
        signature(&values),
        Signature {
            defaults: Some(0.2),
            rows: vec![(1, 0.3)]
        }
    );
}

#[test]
fn mce_raw_errors_and_trailing_xml_keep_typed_error_controls() {
    let xml = alternate(
        r#"<sheetFormatPr litchi_x14ac_dyDescent="0.2"/>"#,
        r#"<sheetFormatPr/>"#,
    ) + "<tail>";
    let error = super::capture(xml.as_bytes()).expect_err("raw marker and malformed tail");
    assert!(matches!(error, Error::Invalid(_)) || matches!(error, Error::MarkupCompatibility(_)));
    assert!(error.to_string().contains("reserved internal marker"));
}

#[test]
fn actual_xlsx_worksheet_parts_match_the_baseline_reference() {
    let fixtures: &[(&str, usize, &[u8])] = &[
        (
            "ooxml-simple-multi-cell",
            3,
            include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/ooxml/xlsx/SimpleMultiCell.xlsx"
            )),
        ),
        (
            "ooxml-new-style-conditional-formattings",
            1,
            include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/ooxml/xlsx/NewStyleConditionalFormattings.xlsx"
            )),
        ),
        (
            "ooxml-structured-references",
            2,
            include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/ooxml/xlsx/StructuredReferences.xlsx"
            )),
        ),
        (
            "poi-with-chart-sheet",
            6,
            include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/poi/test-data/spreadsheet/WithChartSheet.xlsx"
            )),
        ),
        (
            "libreoffice-date-autofilter",
            1,
            include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/office-interop/libreoffice-resaved/date-autofilter-litchi.xlsx"
            )),
        ),
    ];

    let mut total = 0;
    for &(fixture_name, expected_count, bytes) in fixtures {
        let package = OpcPackage::from_bytes(bytes).expect(fixture_name);
        let mut count = 0;
        for part in package.try_iter_parts() {
            let part = part.expect("fixture part payload");
            let part_name = part.partname().as_str();
            if !part_name.starts_with("/xl/worksheets/") || !part_name.ends_with(".xml") {
                continue;
            }
            count += 1;
            total += 1;
            let xml = str::from_utf8(part.blob()).expect("worksheet XML UTF-8");
            for capture_rows in [true, false] {
                let name = format!("{fixture_name}:{part_name}:rows={capture_rows}");
                assert_matches_reference(&name, xml, capture_rows);

                // Exercise the public dispatch too.  This follows the exact
                // baseline branch predicate and uses the stream path only
                // for MCE AlternateContent, where `capture_inner` is not the
                // authoritative reader.
                let expected = reference_public_capture(xml.as_bytes(), capture_rows);
                let actual = if capture_rows {
                    super::capture(xml.as_bytes())
                } else {
                    super::capture_defaults(xml.as_bytes()).map(|defaults| Values {
                        defaults,
                        rows: BTreeMap::new(),
                    })
                };
                assert_same_outcome(&format!("{name}:public"), expected, actual);
            }
        }
        assert_eq!(count, expected_count, "{fixture_name}: worksheet census");
    }
    assert_eq!(total, 13, "fixture worksheet census");
}

fn reference_public_capture(content: &[u8], capture_rows: bool) -> Result<Values> {
    if super::has_markup_compatibility(content)
        && super::has_descent_attribute(content)
        && super::has_alternate_content(content)
    {
        let mut input = Cursor::new(content);
        if capture_rows {
            return super::capture_stream_legacy(
                &mut input,
                &Capabilities::default(),
                &StreamLimits::default(),
                true,
            );
        }
        return super::capture_stream_defaults_legacy(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
        )
        .map(|defaults| Values {
            defaults,
            rows: BTreeMap::new(),
        });
    }
    reference_capture(content, capture_rows)
}

// Baseline reference -------------------------------------------------------

fn reference_capture(content: &[u8], capture_rows: bool) -> Result<Values> {
    let mut reader = quick_xml::NsReader::from_reader(content);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::<Context>::new();
    let mut values = Values {
        defaults: None,
        rows: BTreeMap::new(),
    };
    let mut previous_row = 0u32;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid worksheet extension XML: {error}")))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let decoder = reader.decoder();
        let resolver = reader.resolver();
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "worksheet extension XML exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
                let context = reference_start(
                    stack.last().copied(),
                    &namespace,
                    &element,
                    decoder,
                    resolver,
                    &mut previous_row,
                    &mut values,
                    capture_rows,
                )?;
                stack.push(context);
            },
            Event::Empty(element) => {
                reference_start(
                    stack.last().copied(),
                    &namespace,
                    &element,
                    decoder,
                    resolver,
                    &mut previous_row,
                    &mut values,
                    capture_rows,
                )?;
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| invalid("worksheet extension XML has an unexpected end"))?;
            },
            Event::Eof if !stack.is_empty() => {
                return Err(invalid(
                    "worksheet extension XML has an unterminated element",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(values)
}

#[allow(clippy::too_many_arguments)]
fn reference_start(
    parent: Option<Context>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    previous_row: &mut u32,
    values: &mut Values,
    capture_rows: bool,
) -> Result<Context> {
    if parent.is_none() && super::is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
        return Ok(Context::Worksheet);
    }
    if parent == Some(Context::Worksheet)
        && super::is_spreadsheetml_name(namespace, element.name(), b"sheetFormatPr")
    {
        if let Some(value) = reference_descent(element, decoder, resolver)?
            && values.defaults.replace(value).is_some()
        {
            return Err(invalid("duplicate worksheet default dyDescent"));
        }
        return Ok(Context::Other);
    }
    if parent == Some(Context::Worksheet)
        && super::is_spreadsheetml_name(namespace, element.name(), b"sheetData")
    {
        return Ok(Context::SheetData);
    }
    if parent == Some(Context::SheetData)
        && super::is_spreadsheetml_name(namespace, element.name(), b"row")
    {
        if !capture_rows {
            return Ok(Context::Row);
        }
        let number = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(value) => crate::raw::worksheet::parse_one_based_row(&value)?,
            None => previous_row
                .checked_add(1)
                .filter(|value| *value <= ROWS)
                .ok_or_else(|| invalid("inferred extension row exceeds the spreadsheet grid"))?,
        };
        *previous_row = number;
        if let Some(value) = reference_descent(element, decoder, resolver)?
            && values.rows.insert(number, value).is_some()
        {
            return Err(invalid(format!(
                "duplicate worksheet row {number} dyDescent"
            )));
        }
        return Ok(Context::Row);
    }
    Ok(Context::Other)
}

fn reference_descent(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<Descent>> {
    let mut result = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let is_target = local.as_ref() == b"dyDescent"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == super::NAMESPACE);
        if !is_target {
            continue;
        }
        let lexical = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        let parsed = lexical
            .parse::<f64>()
            .map_err(|_source| invalid(format!("invalid x14ac:dyDescent value '{lexical}'")))?;
        let parsed = Descent::new(parsed)?;
        if result.replace(parsed).is_some() {
            return Err(invalid("duplicate x14ac:dyDescent attribute"));
        }
    }
    Ok(result)
}
