//! Byte-preserving worksheet publication for core conditional formatting.

use std::ops::Range as ByteRange;

use quick_xml::escape::escape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::{parse_conditional_formattings, validate_and_associate};
use super::model::{
    Association, Color, ColorRole, DifferentialRef, Formatting, Payload, Rule, Source, Value,
};
use crate::error::{Error, Result, allocation, invalid};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_WORKSHEET_BYTES: usize = 32 * 1024 * 1024;
const MAX_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;
const MAX_COLLECTIONS: usize = 65_536;

const SUCCESSORS: &[&[u8]] = &[
    b"dataValidations",
    b"hyperlinks",
    b"printOptions",
    b"pageMargins",
    b"pageSetup",
    b"headerFooter",
    b"rowBreaks",
    b"colBreaks",
    b"customProperties",
    b"cellWatches",
    b"ignoredErrors",
    b"smartTags",
    b"drawing",
    b"legacyDrawing",
    b"legacyDrawingHF",
    b"picture",
    b"oleObjects",
    b"controls",
    b"webPublishItems",
    b"tableParts",
    b"extLst",
];

const WORKSHEET_CHILDREN: &[&[u8]] = &[
    b"sheetPr",
    b"dimension",
    b"sheetViews",
    b"sheetFormatPr",
    b"cols",
    b"sheetData",
    b"sheetCalcPr",
    b"sheetProtection",
    b"protectedRanges",
    b"scenarios",
    b"autoFilter",
    b"sortState",
    b"dataConsolidate",
    b"customSheetViews",
    b"mergeCells",
    b"phoneticPr",
    b"conditionalFormatting",
    b"dataValidations",
    b"hyperlinks",
    b"printOptions",
    b"pageMargins",
    b"pageSetup",
    b"headerFooter",
    b"rowBreaks",
    b"colBreaks",
    b"customProperties",
    b"cellWatches",
    b"ignoredErrors",
    b"smartTags",
    b"drawing",
    b"legacyDrawing",
    b"legacyDrawingHF",
    b"picture",
    b"oleObjects",
    b"controls",
    b"webPublishItems",
    b"tableParts",
    b"extLst",
];
const SHEET_DATA_CHILD: usize = 5;

/// Replace the complete ordered core `conditionalFormatting` collection.
///
/// Bytes outside the direct owners are copied exactly. Office 2010, MCE-selected,
/// and opaque selected-owner content is deliberately refused rather than lost.
pub fn replace_conditional_formattings(
    worksheet_xml: &[u8],
    values: &[Formatting],
    differential_format_count: usize,
) -> Result<Vec<u8>> {
    let layout = scan_layout(worksheet_xml)?;
    let before = parse_conditional_formattings(worksheet_xml, differential_format_count)?;
    if before.len() != layout.spans.len() {
        return Err(invalid(
            "conditional formatting selected outside direct worksheet owners",
        ));
    }
    validate_authored(values, differential_format_count)?;
    if before == values {
        return Ok(worksheet_xml.to_vec());
    }
    let replacement = write_collections(values, layout.namespace)?;
    let removed = layout.spans.iter().try_fold(0usize, |sum, span| {
        sum.checked_add(span.len())
            .ok_or_else(|| invalid("conditional-formatting removed size overflow"))
    })?;
    let capacity = worksheet_xml
        .len()
        .checked_sub(removed)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| invalid("conditional-formatting worksheet output size overflow"))?;
    if capacity > MAX_WORKSHEET_BYTES {
        return Err(invalid(
            "conditional-formatting worksheet output is too large",
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("conditional-formatting worksheet output", source))?;
    if layout.spans.is_empty() {
        output.extend_from_slice(&worksheet_xml[..layout.insertion]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&worksheet_xml[layout.insertion..]);
    } else {
        let mut cursor = 0usize;
        for (index, span) in layout.spans.iter().enumerate() {
            if span.start < cursor || span.end > worksheet_xml.len() {
                return Err(invalid("overlapping conditional-formatting owners"));
            }
            output.extend_from_slice(&worksheet_xml[cursor..span.start]);
            if index == 0 {
                output.extend_from_slice(&replacement);
            }
            cursor = span.end;
        }
        output.extend_from_slice(&worksheet_xml[cursor..]);
    }
    scan_layout(&output)?;
    let readback = parse_conditional_formattings(&output, differential_format_count)?;
    if readback != values {
        return Err(invalid(
            "conditional-formatting worksheet write verification failed",
        ));
    }
    Ok(output)
}

pub(crate) fn parse_editable_conditional_formattings(
    worksheet_xml: &[u8],
    differential_format_count: usize,
) -> Result<Vec<Formatting>> {
    let layout = scan_layout(worksheet_xml)?;
    let values = parse_conditional_formattings(worksheet_xml, differential_format_count)?;
    if values.len() != layout.spans.len() {
        return Err(invalid(
            "conditional formatting selected outside direct worksheet owners",
        ));
    }
    validate_authored(&values, differential_format_count)?;
    Ok(values)
}

pub(crate) fn validate_authored(
    values: &[Formatting],
    differential_format_count: usize,
) -> Result<()> {
    if values.len() > MAX_COLLECTIONS {
        return Err(invalid("too many conditional-formatting collections"));
    }
    let total_rules = values.iter().try_fold(0usize, |total, formatting| {
        total
            .checked_add(formatting.rules.len())
            .ok_or_else(|| invalid("conditional-formatting rule count overflow"))
    })?;
    if total_rules > super::codec::MAX_RULES {
        return Err(invalid("too many conditional-formatting rules"));
    }
    let mut checked = values.to_vec();
    validate_and_associate(&mut checked, differential_format_count)?;
    if checked != values {
        return Err(invalid(
            "conditional-formatting extension association is not authorable by the core owner",
        ));
    }
    for formatting in values {
        if formatting.ranges.is_empty() || formatting.rules.is_empty() {
            return Err(invalid(
                "conditionalFormatting requires ranges and at least one rule",
            ));
        }
        for range in &formatting.ranges {
            let parsed = super::codec::parse_sqref(range.as_str())?;
            if parsed.len() != 1 {
                return Err(invalid("conditional-formatting range is not one area"));
            }
        }
        for rule in &formatting.rules {
            validate_core_rule(rule)?;
        }
    }
    let transitional = write_collections(values, CORE)?;
    let mut worksheet = Vec::with_capacity(transitional.len() + CORE.len() + 32);
    worksheet.extend_from_slice(b"<worksheet xmlns=\"");
    worksheet.extend_from_slice(CORE);
    worksheet.extend_from_slice(b"\">");
    worksheet.extend_from_slice(&transitional);
    worksheet.extend_from_slice(b"</worksheet>");
    let reparsed = parse_conditional_formattings(&worksheet, differential_format_count)?;
    if reparsed != values {
        return Err(invalid(
            "conditional-formatting authored state does not round-trip",
        ));
    }
    Ok(())
}

fn validate_core_rule(rule: &Rule) -> Result<()> {
    if rule.source != Source::Core
        || rule.rule_type.is_none()
        || rule.priority.is_none_or(|priority| priority <= 0)
        || rule.extension_id.is_some()
        || rule.extension_association != Association::Independent
        || matches!(rule.differential_format, Some(DifferentialRef::Inline(_)))
        || matches!(rule.payload, Some(Payload::IconSet14(_)))
    {
        return Err(invalid(
            "conditional-formatting publication supports complete core rules only",
        ));
    }
    if rule
        .formulas
        .iter()
        .any(|value| value.len() > super::codec::MAX_FORMULA_BYTES)
    {
        return Err(invalid("conditional-formatting formula is too large"));
    }
    Ok(())
}

fn write_collections(values: &[Formatting], namespace: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for formatting in values {
        output.extend_from_slice(b"<x:conditionalFormatting xmlns:x=\"");
        output.extend_from_slice(namespace);
        output.extend_from_slice(b"\" sqref=\"");
        for (index, range) in formatting.ranges.iter().enumerate() {
            if index != 0 {
                output.push(b' ');
            }
            escaped(&mut output, range.as_str());
        }
        output.push(b'\"');
        if formatting.pivot {
            output.extend_from_slice(b" pivot=\"1\"");
        }
        output.push(b'>');
        for rule in &formatting.rules {
            write_rule(&mut output, rule)?;
        }
        output.extend_from_slice(b"</x:conditionalFormatting>");
        if output.len() > MAX_FRAGMENT_BYTES {
            return Err(invalid("conditional-formatting replacement is too large"));
        }
    }
    Ok(output)
}

fn write_rule(output: &mut Vec<u8>, rule: &Rule) -> Result<()> {
    validate_core_rule(rule)?;
    output.extend_from_slice(b"<x:cfRule type=\"");
    output.extend_from_slice(
        rule.rule_type
            .ok_or_else(|| invalid("core cfRule has no type"))?
            .as_str()
            .as_bytes(),
    );
    output.extend_from_slice(b"\" priority=\"");
    output.extend_from_slice(
        rule.priority
            .ok_or_else(|| invalid("core cfRule has no priority"))?
            .to_string()
            .as_bytes(),
    );
    output.push(b'\"');
    if let Some(DifferentialRef::StylesIndex(value)) = rule.differential_format {
        attribute(output, "dxfId", &value.to_string());
    }
    boolean_attribute(output, "stopIfTrue", rule.stop_if_true, false);
    boolean_attribute(output, "aboveAverage", rule.above_average, true);
    boolean_attribute(output, "equalAverage", rule.equal_average, false);
    boolean_attribute(output, "percent", rule.percent, false);
    boolean_attribute(output, "bottom", rule.bottom, false);
    if let Some(value) = rule.operator {
        attribute(output, "operator", value.as_str());
    }
    if let Some(value) = rule.text.as_deref() {
        attribute(output, "text", value);
    }
    if let Some(value) = rule.time_period {
        attribute(output, "timePeriod", value.as_str());
    }
    if let Some(value) = rule.rank {
        attribute(output, "rank", &value.to_string());
    }
    if let Some(value) = rule.standard_deviations {
        attribute(output, "stdDev", &value.to_string());
    }
    if rule.formulas.is_empty() && rule.payload.is_none() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    for formula in &rule.formulas {
        output.extend_from_slice(b"<x:formula>");
        escaped(output, formula);
        output.extend_from_slice(b"</x:formula>");
    }
    if let Some(payload) = rule.payload.as_ref() {
        write_payload(output, payload)?;
    }
    output.extend_from_slice(b"</x:cfRule>");
    Ok(())
}

fn write_payload(output: &mut Vec<u8>, payload: &Payload) -> Result<()> {
    match payload {
        Payload::ColorScale(value) => {
            output.extend_from_slice(b"<x:colorScale>");
            for threshold in &value.thresholds {
                write_threshold(output, threshold)?;
            }
            for color in &value.colors {
                write_color(output, "color", color)?;
            }
            output.extend_from_slice(b"</x:colorScale>");
        },
        Payload::DataBar(value) => {
            if value.border
                || !value.gradient
                || value.direction != super::Direction::Context
                || value.axis_position != super::Axis::Automatic
                || value
                    .colors
                    .iter()
                    .any(|color| color.role != ColorRole::Color)
            {
                return Err(invalid("data bar contains Office 2010-only state"));
            }
            output.extend_from_slice(b"<x:dataBar");
            attribute(output, "minLength", &value.min_length.to_string());
            attribute(output, "maxLength", &value.max_length.to_string());
            boolean_attribute(output, "showValue", value.show_value, true);
            output.push(b'>');
            for threshold in &value.thresholds {
                write_threshold(output, threshold)?;
            }
            for color in &value.colors {
                write_color(output, "color", &color.color)?;
            }
            output.extend_from_slice(b"</x:dataBar>");
        },
        Payload::IconSet(value) => {
            output.extend_from_slice(b"<x:iconSet");
            attribute(output, "iconSet", value.set.as_str());
            boolean_attribute(output, "showValue", value.show_value, true);
            boolean_attribute(output, "percent", value.percent, true);
            boolean_attribute(output, "reverse", value.reverse, false);
            output.push(b'>');
            for threshold in &value.thresholds {
                write_threshold(output, threshold)?;
            }
            output.extend_from_slice(b"</x:iconSet>");
        },
        Payload::IconSet14(_) => {
            return Err(invalid(
                "Office 2010 icon sets are not editable by the core owner",
            ));
        },
    }
    Ok(())
}

fn write_threshold(output: &mut Vec<u8>, value: &Value) -> Result<()> {
    if !value.kind.is_core() || value.formula.is_some() {
        return Err(invalid("Office 2010 threshold state is not a core cfvo"));
    }
    output.extend_from_slice(b"<x:cfvo type=\"");
    output.extend_from_slice(value.kind.as_str().as_bytes());
    output.push(b'\"');
    if let Some(value) = value.value.as_deref() {
        attribute(output, "val", value);
    }
    boolean_attribute(output, "gte", value.greater_than_or_equal, true);
    output.extend_from_slice(b"/>");
    Ok(())
}

fn write_color(output: &mut Vec<u8>, role: &str, value: &Color) -> Result<()> {
    if value.tint.is_some_and(|value| !value.is_finite()) {
        return Err(invalid("conditional-formatting color tint is not finite"));
    }
    output.extend_from_slice(b"<x:");
    output.extend_from_slice(role.as_bytes());
    if let Some(value) = value.rgb {
        attribute(output, "rgb", &value.to_string());
    }
    if let Some(value) = value.indexed {
        attribute(output, "indexed", &value.to_string());
    }
    if let Some(value) = value.theme {
        attribute(output, "theme", &value.to_string());
    }
    if let Some(value) = value.tint {
        attribute(output, "tint", &value.to_string());
    }
    if let Some(value) = value.automatic {
        attribute(output, "auto", if value { "1" } else { "0" });
    }
    output.extend_from_slice(b"/>");
    Ok(())
}

fn attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escaped(output, value);
    output.push(b'\"');
}

fn boolean_attribute(output: &mut Vec<u8>, name: &str, value: bool, default: bool) {
    if value != default {
        attribute(output, name, if value { "1" } else { "0" });
    }
}

fn escaped(output: &mut Vec<u8>, value: &str) {
    output.extend_from_slice(escape(value).as_bytes());
}

struct Layout {
    namespace: &'static [u8],
    spans: Vec<ByteRange<usize>>,
    insertion: usize,
}

fn scan_layout(xml: &[u8]) -> Result<Layout> {
    if xml.len() > MAX_WORKSHEET_BYTES {
        return Err(invalid("conditional-formatting worksheet XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut namespace = None;
    let mut root_close = None;
    let mut insertion = None;
    let mut owner: Option<(usize, usize)> = None;
    let mut owner_stack: Vec<Vec<u8>> = Vec::new();
    let mut spans = Vec::new();
    let mut previous_child = None;
    let mut seen_children = [false; WORKSHEET_CHILDREN.len()];
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("conditional-formatting event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid(
                "conditional-formatting worksheet exceeds event limit",
            ));
        }
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (resolved, event) = resolver.resolve_event(event);
        if exact(&resolved, X14) {
            return Err(invalid(
                "Office 2010 conditional formatting is not editable",
            ));
        }
        match event {
            Event::Start(element) => {
                reject_compatibility_attributes(&element, &resolver)?;
                if exact(&resolved, MCE) {
                    return Err(invalid(
                        "MCE-selected conditional formatting is not editable byte-exactly",
                    ));
                }
                if depth == 0 {
                    if element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid(
                            "conditional-formatting publication requires a worksheet root",
                        ));
                    }
                    namespace = spreadsheet_namespace(&resolved);
                    if namespace.is_none() {
                        return Err(invalid("invalid SpreadsheetML worksheet namespace"));
                    }
                } else if depth == 1 {
                    if !spreadsheet(&resolved) {
                        return Err(invalid(
                            "foreign direct worksheet children are not editable",
                        ));
                    }
                    let local = element.local_name();
                    previous_child = Some(check_child_order(
                        local.as_ref(),
                        previous_child,
                        &mut seen_children,
                    )?);
                    if local.as_ref() == b"conditionalFormatting" {
                        if owner.is_some() || spans.len() >= MAX_COLLECTIONS {
                            return Err(invalid("invalid conditionalFormatting owner count"));
                        }
                        validate_owner_element(&element, local.as_ref(), None)?;
                        owner = Some((depth + 1, start));
                        owner_stack.push(local.as_ref().to_vec());
                    } else if SUCCESSORS.contains(&local.as_ref()) && insertion.is_none() {
                        insertion = Some(start);
                    }
                } else if let Some((owner_depth, _)) = owner {
                    validate_nested_owner_element(
                        &resolved,
                        &element,
                        depth - owner_depth,
                        owner_stack.last().map(Vec::as_slice),
                    )?;
                    owner_stack.push(element.local_name().as_ref().to_vec());
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("conditional-formatting nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid(
                        "conditional-formatting worksheet nesting is too deep",
                    ));
                }
            },
            Event::Empty(element) => {
                reject_compatibility_attributes(&element, &resolver)?;
                if exact(&resolved, MCE) {
                    return Err(invalid(
                        "MCE-selected conditional formatting is not editable byte-exactly",
                    ));
                }
                if depth == 1 {
                    if !spreadsheet(&resolved) {
                        return Err(invalid(
                            "foreign direct worksheet children are not editable",
                        ));
                    }
                    let local = element.local_name();
                    previous_child = Some(check_child_order(
                        local.as_ref(),
                        previous_child,
                        &mut seen_children,
                    )?);
                    if local.as_ref() == b"conditionalFormatting" {
                        validate_owner_element(&element, local.as_ref(), None)?;
                        if spans.len() >= MAX_COLLECTIONS {
                            return Err(invalid("too many conditionalFormatting owners"));
                        }
                        spans.push(start..end);
                    } else if SUCCESSORS.contains(&local.as_ref()) && insertion.is_none() {
                        insertion = Some(start);
                    }
                } else if let Some((owner_depth, _)) = owner {
                    if spreadsheet(&resolved) && element.local_name().as_ref() == b"formula" {
                        return Err(invalid(
                            "empty conditional-formatting formulas are not editable",
                        ));
                    }
                    validate_nested_owner_element(
                        &resolved,
                        &element,
                        depth - owner_depth,
                        owner_stack.last().map(Vec::as_slice),
                    )?;
                }
            },
            Event::End(element) => {
                if let Some((owner_depth, owner_start)) = owner {
                    if depth == owner_depth
                        && spreadsheet(&resolved)
                        && element.local_name().as_ref() == b"conditionalFormatting"
                    {
                        spans.push(owner_start..end);
                        owner = None;
                        owner_stack.clear();
                    } else if depth > owner_depth {
                        owner_stack.pop().ok_or_else(|| {
                            invalid("conditional-formatting owner stack underflow")
                        })?;
                    }
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("conditional-formatting XML nesting underflow"))?;
            },
            Event::Comment(_) if owner.is_some() => {
                return Err(invalid(
                    "comments inside conditionalFormatting cannot be preserved",
                ));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "conditional-formatting publication rejects DTD and processing instructions",
                ));
            },
            Event::Text(text) if owner.is_some() => {
                if owner_stack.last().is_none_or(|value| value != b"formula")
                    && !text.decode().map_err(xml_error)?.trim().is_empty()
                {
                    return Err(invalid(
                        "opaque text inside conditionalFormatting cannot be preserved",
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if owner.is_some() => {
                if owner_stack.last().is_none_or(|value| value != b"formula") {
                    return Err(invalid(
                        "opaque character data inside conditionalFormatting cannot be preserved",
                    ));
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 || owner.is_some() || !owner_stack.is_empty() {
        return Err(invalid("incomplete conditional-formatting worksheet XML"));
    }
    if !seen_children[SHEET_DATA_CHILD] {
        return Err(invalid(
            "conditional-formatting worksheet has no required sheetData",
        ));
    }
    Ok(Layout {
        namespace: namespace.ok_or_else(|| invalid("worksheet XML has no root"))?,
        spans,
        insertion: insertion
            .or(root_close)
            .ok_or_else(|| invalid("worksheet has no conditionalFormatting insertion point"))?,
    })
}

fn reject_compatibility_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if exact(&namespace, MCE) || exact(&namespace, X14) {
            return Err(invalid(
                "MCE and Office 2010 worksheet markup is not editable by the core owner",
            ));
        }
    }
    Ok(())
}

fn check_child_order(
    local: &[u8],
    previous: Option<usize>,
    seen: &mut [bool; WORKSHEET_CHILDREN.len()],
) -> Result<usize> {
    let index = WORKSHEET_CHILDREN
        .iter()
        .position(|known| *known == local)
        .ok_or_else(|| {
            invalid(format!(
                "unknown direct worksheet child '{}'",
                String::from_utf8_lossy(local)
            ))
        })?;
    if previous.is_some_and(|previous| index < previous) {
        return Err(invalid(format!(
            "worksheet child '{}' is out of schema order",
            String::from_utf8_lossy(local)
        )));
    }
    if local != b"conditionalFormatting" && seen[index] {
        return Err(invalid(format!(
            "duplicate direct worksheet child '{}'",
            String::from_utf8_lossy(local)
        )));
    }
    seen[index] = true;
    Ok(index)
}

fn validate_nested_owner_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    relative_depth: usize,
    parent: Option<&[u8]>,
) -> Result<()> {
    if !spreadsheet(namespace) {
        return Err(invalid(
            "foreign content inside conditionalFormatting cannot be preserved",
        ));
    }
    let local = element.local_name();
    let allowed = matches!(
        (relative_depth, parent, local.as_ref()),
        (0, Some(b"conditionalFormatting"), b"cfRule")
            | (
                1,
                Some(b"cfRule"),
                b"formula" | b"colorScale" | b"dataBar" | b"iconSet"
            )
            | (2, Some(b"colorScale" | b"dataBar"), b"cfvo" | b"color")
            | (2, Some(b"iconSet"), b"cfvo")
    );
    if !allowed {
        return Err(invalid(format!(
            "unsupported element '{}' inside conditionalFormatting",
            String::from_utf8_lossy(local.as_ref())
        )));
    }
    validate_owner_element(element, local.as_ref(), parent)
}

fn validate_owner_element(
    element: &BytesStart<'_>,
    local: &[u8],
    _parent: Option<&[u8]>,
) -> Result<()> {
    let allowed: &[&[u8]] = match local {
        b"conditionalFormatting" => &[b"sqref", b"pivot"],
        b"cfRule" => &[
            b"type",
            b"dxfId",
            b"priority",
            b"stopIfTrue",
            b"aboveAverage",
            b"equalAverage",
            b"percent",
            b"bottom",
            b"operator",
            b"text",
            b"timePeriod",
            b"rank",
            b"stdDev",
        ],
        b"formula" | b"colorScale" => &[],
        b"dataBar" => &[b"minLength", b"maxLength", b"showValue"],
        b"iconSet" => &[b"iconSet", b"showValue", b"percent", b"reverse"],
        b"cfvo" => &[b"type", b"val", b"gte"],
        b"color" => &[b"rgb", b"indexed", b"theme", b"tint", b"auto"],
        _ => return Err(invalid("unknown conditional-formatting owner element")),
    };
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if name.contains(&b':') || !allowed.contains(&name) {
            return Err(invalid(format!(
                "unsupported attribute '{}' inside conditionalFormatting",
                String::from_utf8_lossy(name)
            )));
        }
    }
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source| invalid("conditional-formatting XML position does not fit usize"))
}

fn spreadsheet_namespace(namespace: &ResolveResult<'_>) -> Option<&'static [u8]> {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == CORE => Some(CORE),
        ResolveResult::Bound(value) if value.as_ref() == STRICT => Some(STRICT),
        _ => None,
    }
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool {
    spreadsheet_namespace(namespace).is_some()
}

fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
