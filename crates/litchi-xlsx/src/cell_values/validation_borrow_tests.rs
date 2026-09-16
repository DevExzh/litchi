//! Differential coverage for the event-borrowing XML validation path.
//!
//! `validate_xml_owned_reference` is intentionally a small frozen copy of the
//! pre-borrowing loop.  It shares the validator helpers with the production
//! implementation, so these tests compare event lifetime/resolver handling
//! without copying the policy tables into a second implementation.

use quick_xml::events::Event;

use quick_xml::reader::NsReader;

use crate::error::Result;

use super::{
    Admission, STRICT_SML, TRANSITIONAL_SML, XmlOwner, text_allowed, text_context_allowed,
    validate_close, validate_element,
};

#[derive(Debug, Eq, PartialEq)]
struct Observation {
    accepted: bool,
    error: Option<String>,
}

fn observe(result: Result<()>) -> Observation {
    match result {
        Ok(()) => Observation {
            accepted: true,
            error: None,
        },
        Err(error) => Observation {
            accepted: false,
            error: Some(error.to_string()),
        },
    }
}

fn owner_name(owner: XmlOwner) -> &'static str {
    match owner {
        XmlOwner::Workbook => "workbook",
        XmlOwner::Worksheet => "worksheet",
    }
}

fn production_validate(content: &[u8], owner: XmlOwner) -> Result<()> {
    match owner {
        XmlOwner::Workbook => super::workbook_xml(content),
        XmlOwner::Worksheet => super::worksheet_xml(content),
    }
}

fn assert_differential(label: &str, content: &[u8], owner: XmlOwner, accepted: bool) {
    let reference = observe(validate_xml_owned_reference(content, owner));
    let actual = observe(production_validate(content, owner));

    assert_eq!(
        reference.accepted,
        accepted,
        "frozen reference status for {label} ({})",
        owner_name(owner)
    );
    assert_eq!(
        actual.accepted,
        accepted,
        "production status for {label} ({})",
        owner_name(owner)
    );
    assert_eq!(
        actual,
        reference,
        "borrowed validation diverged from owned reference for {label} ({})",
        owner_name(owner)
    );
}

fn assert_differential_error(
    label: &str,
    content: &[u8],
    owner: XmlOwner,
    expected_error: &'static str,
) {
    let reference = observe(validate_xml_owned_reference(content, owner));
    let actual = observe(production_validate(content, owner));
    let expected_error = format!("invalid XLSX structure: {expected_error}");

    assert!(
        !reference.accepted,
        "reference unexpectedly accepted {label}"
    );
    assert_eq!(
        reference.error.as_deref(),
        Some(expected_error.as_str()),
        "frozen error changed for {label} ({})",
        owner_name(owner)
    );
    assert_eq!(
        actual,
        reference,
        "borrowed validation changed the exact error for {label} ({})",
        owner_name(owner)
    );
}

fn namespace(value: &[u8]) -> &'static str {
    if value == TRANSITIONAL_SML {
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    } else if value == STRICT_SML {
        "http://purl.oclc.org/ooxml/spreadsheetml/main"
    } else {
        panic!("fixture requested an unknown SpreadsheetML namespace")
    }
}

fn valid_default_worksheet(sml: &[u8]) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C2"/><sheetViews><sheetView showGridLines="1"><selection activeCell="A1"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">Café 😀 &amp; <![CDATA[Δ &amp; &lt;literal&gt;]]> &#x1F600;</t></is></c><c r="B1"><f>SUM(A1)</f><v>42</v></c></row></sheetData></worksheet>"#,
        namespace(sml)
    )
    .into_bytes()
}

fn valid_prefixed_workbook(sml: &[u8]) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><x:workbook xmlns:x="{}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:fileVersion appName="xl"/><x:bookViews><x:workbookView activeTab="0"/></x:bookViews><x:sheets><x:sheet name="Sheet1" sheetId="1" state="visible" r:id="rId1"/></x:sheets><x:calcPr calcId="0" fullCalcOnLoad="1"/></x:workbook>"#,
        namespace(sml)
    )
    .into_bytes()
}

fn valid_mixed_alias_and_default_worksheet(sml: &[u8]) -> Vec<u8> {
    format!(
        r#"<s:worksheet xmlns:s="{}"><sheetData xmlns="{}"><row r="1"><s:c r="A1"><s:v>7</s:v></s:c></row></sheetData></s:worksheet>"#,
        namespace(sml),
        namespace(sml)
    )
    .into_bytes()
}

/// This fixture has a foreign `a` binding in scope, then temporarily rebinds
/// it on an empty element.  The following sibling and all closing tags must
/// see the scope restored after that `Empty` event.
fn namespace_rich_worksheet() -> Vec<u8> {
    let sml = namespace(TRANSITIONAL_SML);
    format!(
        r#"<?xml version="1.0"?><o:worksheet xmlns:o="{sml}" xmlns:a="urn:fixture:foreign" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><o:dimension ref="A1:C2"/><o:sheetData><o:row r="1"><o:c r="A1" t="inlineStr"><o:is><o:t xml:space="preserve">Café 😀 &amp; <![CDATA[Δ &amp; &lt;literal&gt;]]> &#x1F600;</o:t></o:is></o:c><a:c xmlns:a="{sml}" r="B1"/><o:c r="C1"><o:f>SUM(A1)</o:f><o:v>42</o:v></o:c></o:row></o:sheetData></o:worksheet>"#,
        sml = sml
    )
    .into_bytes()
}

fn empty_rebind_foreign_sibling() -> Vec<u8> {
    let sml = namespace(TRANSITIONAL_SML);
    format!(
        r#"<s:worksheet xmlns:s="{sml}" xmlns:a="urn:fixture:foreign"><s:sheetData><s:row r="1"><a:c xmlns:a="{sml}" r="A1"/><a:c r="A2"/></s:row></s:sheetData></s:worksheet>"#,
        sml = sml
    )
    .into_bytes()
}

fn mixed_dialect() -> Vec<u8> {
    format!(
        r#"<worksheet xmlns="{transitional}"><sheetData xmlns="{strict}"/></worksheet>"#,
        transitional = namespace(TRANSITIONAL_SML),
        strict = namespace(STRICT_SML)
    )
    .into_bytes()
}

fn foreign_element() -> Vec<u8> {
    format!(
        r#"<worksheet xmlns="{sml}" xmlns:f="urn:fixture:foreign"><sheetData><f:row/></sheetData></worksheet>"#,
        sml = namespace(TRANSITIONAL_SML)
    )
    .into_bytes()
}

fn mismatched_namespace_closing() -> Vec<u8> {
    format!(
        r#"<s:worksheet xmlns:s="{sml}" xmlns:a="{strict}"><s:sheetData></a:sheetData></s:worksheet>"#,
        sml = namespace(TRANSITIONAL_SML),
        strict = namespace(STRICT_SML)
    )
    .into_bytes()
}

#[test]
fn transitional_and_strict_worksheet_and_workbook_inputs_are_accepted() {
    for (label, sml) in [("transitional", TRANSITIONAL_SML), ("strict", STRICT_SML)] {
        let worksheet = valid_default_worksheet(sml);
        assert_differential(label, &worksheet, XmlOwner::Worksheet, true);

        let workbook = valid_prefixed_workbook(sml);
        assert_differential(
            &format!("{label} prefixed workbook"),
            &workbook,
            XmlOwner::Workbook,
            true,
        );
    }
}

#[test]
fn aliases_and_default_namespace_scopes_are_resolved_identically() {
    for (label, sml) in [("transitional", TRANSITIONAL_SML), ("strict", STRICT_SML)] {
        let worksheet = valid_mixed_alias_and_default_worksheet(sml);
        assert_differential(
            &format!("{label} alias/default worksheet"),
            &worksheet,
            XmlOwner::Worksheet,
            true,
        );
    }
}

#[test]
fn empty_namespace_rebinding_restores_scope_for_siblings_and_closing_tags() {
    let valid = namespace_rich_worksheet();
    assert_differential(
        "empty namespace rebind with valid sibling",
        &valid,
        XmlOwner::Worksheet,
        true,
    );

    let invalid = empty_rebind_foreign_sibling();
    assert_differential_error(
        "empty namespace rebind followed by foreign sibling",
        &invalid,
        XmlOwner::Worksheet,
        "value-only XML has a foreign element namespace",
    );
}

#[test]
fn namespace_failures_have_the_frozen_exact_errors() {
    let mixed = mixed_dialect();
    assert_differential_error(
        "mixed SpreadsheetML dialects",
        &mixed,
        XmlOwner::Worksheet,
        "value-only XML mixes SpreadsheetML dialects",
    );

    let foreign = foreign_element();
    assert_differential_error(
        "foreign element namespace",
        &foreign,
        XmlOwner::Worksheet,
        "value-only XML has a foreign element namespace",
    );

    let unbound = b"<worksheet><sheetData/></worksheet>";
    assert_differential_error(
        "unbound default namespace",
        unbound,
        XmlOwner::Worksheet,
        "value-only XML has an unbound element namespace",
    );

    let unknown_prefix = b"<x:worksheet><x:sheetData/></x:worksheet>";
    assert_differential_error(
        "unknown element prefix",
        unknown_prefix,
        XmlOwner::Worksheet,
        "value-only XML has an unbound element namespace",
    );

    let reset_default = format!(
        r#"<worksheet xmlns="{sml}"><sheetData xmlns=""/></worksheet>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential_error(
        "reset default namespace",
        reset_default.as_bytes(),
        XmlOwner::Worksheet,
        "value-only XML has an unbound element namespace",
    );
}

#[test]
fn mismatched_and_unbalanced_closing_tags_have_exact_errors() {
    let mismatched = format!(
        r#"<worksheet xmlns="{sml}"><sheetData></worksheet>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential(
        "mismatched local closing name",
        mismatched.as_bytes(),
        XmlOwner::Worksheet,
        false,
    );

    let mismatched_namespace = mismatched_namespace_closing();
    assert_differential(
        "mismatched closing namespace",
        &mismatched_namespace,
        XmlOwner::Worksheet,
        false,
    );

    let unmatched = format!(
        r#"</worksheet><worksheet xmlns="{sml}"/>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential(
        "unmatched closing element",
        unmatched.as_bytes(),
        XmlOwner::Worksheet,
        false,
    );

    let incomplete = format!(
        r#"<worksheet xmlns="{sml}"><sheetData/>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential_error(
        "incomplete root",
        incomplete.as_bytes(),
        XmlOwner::Worksheet,
        "value-only XML has no complete root element",
    );
}

#[test]
fn malformed_and_duplicate_attributes_match_parser_errors() {
    let duplicate = format!(
        r#"<worksheet xmlns="{sml}"><dimension ref="A1" ref="B2"/></worksheet>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential(
        "duplicate attribute",
        duplicate.as_bytes(),
        XmlOwner::Worksheet,
        false,
    );

    let duplicate_namespace = format!(
        r#"<worksheet xmlns="{sml}" xmlns="{sml}"/>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential(
        "duplicate namespace attribute",
        duplicate_namespace.as_bytes(),
        XmlOwner::Worksheet,
        false,
    );

    let unterminated_quote = format!(
        r#"<worksheet xmlns="{sml}"><dimension ref='A1></dimension></worksheet>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential(
        "unterminated attribute quote",
        unterminated_quote.as_bytes(),
        XmlOwner::Worksheet,
        false,
    );

    let missing_attribute_value = format!(
        r#"<worksheet xmlns="{sml}"><dimension ref=A1/></worksheet>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential(
        "unquoted attribute value",
        missing_attribute_value.as_bytes(),
        XmlOwner::Worksheet,
        false,
    );
}

#[test]
fn text_cdata_and_general_references_are_checked_in_context() {
    let valid = namespace_rich_worksheet();
    assert_differential(
        "unicode text cdata and refs",
        &valid,
        XmlOwner::Worksheet,
        true,
    );

    let sml = namespace(TRANSITIONAL_SML);
    let text_outside_scalar = format!(r#"<worksheet xmlns="{sml}">text<sheetData/></worksheet>"#);
    assert_differential_error(
        "text outside scalar value",
        text_outside_scalar.as_bytes(),
        XmlOwner::Worksheet,
        "value-only XML has text outside a scalar value element",
    );

    let cdata_outside_scalar =
        format!(r#"<worksheet xmlns="{sml}"><![CDATA[text]]><sheetData/></worksheet>"#);
    assert_differential_error(
        "CDATA outside scalar value",
        cdata_outside_scalar.as_bytes(),
        XmlOwner::Worksheet,
        "value-only XML has text outside a scalar value element",
    );

    let reference_outside_scalar =
        format!(r#"<worksheet xmlns="{sml}">&unknown;<sheetData/></worksheet>"#);
    assert_differential_error(
        "general reference outside scalar value",
        reference_outside_scalar.as_bytes(),
        XmlOwner::Worksheet,
        "value-only XML has a reference outside a scalar value element",
    );

    let malformed_reference = format!(
        r#"<worksheet xmlns="{sml}"><sheetData><row r="1"><c r="A1"><v>&broken</v></c></row></sheetData></worksheet>"#
    );
    assert_differential(
        "malformed general reference",
        malformed_reference.as_bytes(),
        XmlOwner::Worksheet,
        false,
    );
}

#[test]
fn document_type_declarations_are_refused_with_exact_errors() {
    let dtd = format!(
        r#"<!DOCTYPE worksheet><worksheet xmlns="{sml}"/>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential_error(
        "worksheet DTD",
        dtd.as_bytes(),
        XmlOwner::Worksheet,
        "value-only edits refuse XML document type declarations",
    );

    let dtd_with_subset = format!(
        r#"<!DOCTYPE workbook [<!ELEMENT workbook ANY>]><workbook xmlns="{sml}"/>"#,
        sml = namespace(TRANSITIONAL_SML)
    );
    assert_differential_error(
        "workbook DTD with internal subset",
        dtd_with_subset.as_bytes(),
        XmlOwner::Workbook,
        "value-only edits refuse XML document type declarations",
    );
}

#[test]
fn malformed_and_truncated_inputs_keep_parser_errors_identical() {
    let sml = namespace(TRANSITIONAL_SML);
    let malformed = [
        format!(r#"<worksheet xmlns="{sml}""#),
        format!(r#"<worksheet xmlns="{sml}"><sheetData>"#),
        format!(r#"<worksheet xmlns="{sml}"><dimension ref="A1""#),
        format!(r#"<worksheet xmlns="{sml}"><sheetData>&"#),
    ];
    for (index, input) in malformed.iter().enumerate() {
        assert_differential(
            &format!("malformed input {index}"),
            input.as_bytes(),
            XmlOwner::Worksheet,
            false,
        );
    }

    let mut invalid_utf8 =
        format!(r#"<worksheet xmlns="{sml}"><sheetData><row r="1"><c r="A1"><v>"#).into_bytes();
    invalid_utf8.push(0xff);
    invalid_utf8.extend_from_slice(b"</v></c></row></sheetData></worksheet>");
    assert_differential(
        "invalid UTF-8 in scalar text",
        &invalid_utf8,
        XmlOwner::Worksheet,
        false,
    );
}

#[test]
fn every_truncation_of_namespace_rich_valid_fixture_is_differentially_checked() {
    let valid = namespace_rich_worksheet();
    assert_differential(
        "complete namespace-rich fixture",
        &valid,
        XmlOwner::Worksheet,
        true,
    );

    let mut refused_prefixes = 0usize;
    for cut in 0..valid.len() {
        assert_differential(
            &format!("namespace-rich prefix cut at byte {cut}"),
            &valid[..cut],
            XmlOwner::Worksheet,
            false,
        );
        refused_prefixes += 1;
    }
    assert!(
        refused_prefixes > 256,
        "truncation sweep must exercise a broad bounded input set"
    );
}

#[test]
fn workbook_invalid_inputs_are_checked_against_the_same_owned_loop() {
    let sml = namespace(TRANSITIONAL_SML);
    let wrong_root = format!(r#"<worksheet xmlns="{sml}"/>"#);
    assert_differential(
        "worksheet supplied to workbook validator",
        wrong_root.as_bytes(),
        XmlOwner::Workbook,
        false,
    );

    // Change 0657: `<fileVersion>` is not a span the rewrite composes, so it
    // and every attribute on it are copied through unread.
    let unfamiliar_attribute =
        format!(r#"<workbook xmlns="{sml}"><fileVersion unknown="1"/></workbook>"#);
    assert_differential(
        "unfamiliar workbook attribute",
        unfamiliar_attribute.as_bytes(),
        XmlOwner::Workbook,
        true,
    );

    // A child of the catalog is composed, and is still held to the modelled
    // vocabulary.
    let unknown_catalog_child =
        format!(r#"<workbook xmlns="{sml}"><sheets><future/></sheets></workbook>"#);
    assert_differential_error(
        "unknown catalog child",
        unknown_catalog_child.as_bytes(),
        XmlOwner::Workbook,
        "value-only edits refuse dependency-bearing or unknown element 'future'",
    );

    let invalid_text = format!(r#"<workbook xmlns="{sml}">text</workbook>"#);
    assert_differential(
        "workbook text outside scalar value",
        invalid_text.as_bytes(),
        XmlOwner::Workbook,
        false,
    );
}

/// Frozen pre-borrowing validator loop.  Keep this limited to event ownership
/// and namespace resolver sequencing; policy remains in the parent helpers.
fn validate_xml_owned_reference(content: &[u8], owner: XmlOwner) -> Result<()> {
    let mut reader = NsReader::from_reader(content);
    let mut depth = 0usize;
    let mut elements = Vec::<Box<[u8]>>::new();
    let mut dialect = None::<Box<[u8]>>;
    let mut copied_from = None::<usize>;
    let mut saw_root = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| crate::error::invalid(format!("value-only XML scan failed: {error}")))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let local = element
                    .name()
                    .local_name()
                    .as_ref()
                    .to_vec()
                    .into_boxed_slice();
                let admission = validate_element(
                    owner,
                    &namespace,
                    &element,
                    &local,
                    elements.last().map(AsRef::as_ref),
                    depth,
                    &mut dialect,
                    copied_from.is_some(),
                )?;
                saw_root = true;
                if admission == Admission::Copied && copied_from.is_none() {
                    copied_from = Some(depth);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| crate::error::invalid("value-only XML depth overflow"))?;
                elements.push(local);
            },
            Event::Empty(element) => {
                let local = element.name().local_name().as_ref().to_vec();
                validate_element(
                    owner,
                    &namespace,
                    &element,
                    &local,
                    elements.last().map(AsRef::as_ref),
                    depth,
                    &mut dialect,
                    copied_from.is_some(),
                )?;
                saw_root = true;
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::error::invalid("value-only XML has an unmatched closing element")
                })?;
                let expected = elements.pop().ok_or_else(|| {
                    crate::error::invalid("value-only XML has no open element to close")
                })?;
                let modeled = copied_from.is_none();
                if copied_from == Some(depth) {
                    copied_from = None;
                }
                validate_close(&namespace, &element, &expected, dialect.as_deref(), modeled)?;
            },
            Event::DocType(_) => {
                return Err(crate::error::invalid(
                    "value-only edits refuse XML document type declarations",
                ));
            },
            Event::Eof => break,
            Event::Text(value) => {
                let decoded = value.decode().map_err(|error| {
                    crate::error::invalid(format!("invalid value-only XML text: {error}"))
                })?;
                if !text_allowed(
                    owner,
                    elements.last().map(AsRef::as_ref),
                    copied_from.is_some(),
                    &decoded,
                ) {
                    return Err(crate::error::invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::CData(value) => {
                let decoded = value.decode().map_err(|error| {
                    crate::error::invalid(format!("invalid value-only XML text: {error}"))
                })?;
                if !text_allowed(
                    owner,
                    elements.last().map(AsRef::as_ref),
                    copied_from.is_some(),
                    &decoded,
                ) {
                    return Err(crate::error::invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::GeneralRef(_) => {
                if !text_context_allowed(
                    owner,
                    elements.last().map(AsRef::as_ref),
                    copied_from.is_some(),
                ) {
                    return Err(crate::error::invalid(
                        "value-only XML has a reference outside a scalar value element",
                    ));
                }
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
    }
    if !saw_root || depth != 0 {
        return Err(crate::error::invalid(
            "value-only XML has no complete root element",
        ));
    }
    Ok(())
}
