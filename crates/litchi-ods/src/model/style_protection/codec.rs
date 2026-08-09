//! Lossless XML codecs for table-cell protection styles.

use super::{
    OFFICE_NAMESPACE, STYLE_NAMESPACE,
    model::{ConditionalStyle, PreservedXmlFragment, Protection, Rule, TableStyle},
    semantic::{is_namespace, optional_attribute},
    validation::validate_protection_style_collection,
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    reader::NsReader,
};
use std::collections::{BTreeMap, HashSet};
use std::ops::Range;

/// # Errors
///
/// Returns an error when the value cannot be serialized.
#[cfg_attr(not(test), allow(dead_code))]
pub fn rewrite_conditional_styles(
    fragment: Option<&PreservedXmlFragment>,
    styles: &[ConditionalStyle],
) -> Result<PreservedXmlFragment> {
    let canonical = write_conditional_styles(styles);
    let Some(fragment) = fragment else {
        return Ok(PreservedXmlFragment {
            xml: format!(
                "<office:automatic-styles xmlns:office=\"{}\">{canonical}</office:automatic-styles>",
                String::from_utf8_lossy(OFFICE_NAMESPACE)
            ),
            namespaces: BTreeMap::new(),
        });
    };
    let (ranges, insertion) = conditional_style_ranges(&fragment.xml)?;
    let xml = match insertion {
        AutomaticStylesInsertion::BeforeEnd(position) => {
            let mut out = String::with_capacity(fragment.xml.len() + canonical.len());
            let mut cursor = 0usize;
            for range in ranges {
                if range.end > position {
                    return Err(Error::InvalidFormat(
                        "conditional style range exceeds automatic styles container".to_string(),
                    ));
                }
                out.push_str(&fragment.xml[cursor..range.start]);
                cursor = range.end;
            }
            out.push_str(&fragment.xml[cursor..position]);
            out.push_str(&canonical);
            out.push_str(&fragment.xml[position..]);
            out
        },
        AutomaticStylesInsertion::ExpandEmpty { slash, name } => {
            let mut out = String::with_capacity(fragment.xml.len() + canonical.len() + name.len());
            out.push_str(&fragment.xml[..slash]);
            out.push('>');
            out.push_str(&canonical);
            out.push_str("</");
            out.push_str(&name);
            out.push('>');
            out
        },
    };
    Ok(PreservedXmlFragment {
        xml,
        namespaces: fragment.namespaces.clone(),
    })
}

/// # Errors
///
/// Returns an error when the value cannot be serialized.
pub fn rewrite_managed_cell_styles(
    fragment: Option<&PreservedXmlFragment>,
    conditional_styles: &[ConditionalStyle],
    protection_styles: &[TableStyle],
) -> Result<PreservedXmlFragment> {
    validate_protection_style_collection(protection_styles)?;
    for conditional in conditional_styles {
        if let Some(protection) = protection_styles
            .iter()
            .find(|style| style.style_name == conditional.style_name)
            && conditional.parent_style_name != protection.parent_style_name
        {
            return Err(Error::InvalidFormat(format!(
                "conditional and protection definitions for '{}' have different parent styles",
                conditional.style_name
            )));
        }
    }
    let canonical = write_managed_styles(conditional_styles, protection_styles);
    let Some(fragment) = fragment else {
        return Ok(PreservedXmlFragment {
            xml: format!(
                "<office:automatic-styles xmlns:office=\"{}\">{canonical}</office:automatic-styles>",
                String::from_utf8_lossy(OFFICE_NAMESPACE)
            ),
            namespaces: BTreeMap::new(),
        });
    };
    let (ranges, insertion) = managed_style_ranges(&fragment.xml)?;
    let xml = rewrite_ranges(&fragment.xml, ranges, insertion, &canonical)?;
    Ok(PreservedXmlFragment {
        xml,
        namespaces: fragment.namespaces.clone(),
    })
}

fn rewrite_ranges(
    xml: &str,
    ranges: Vec<Range<usize>>,
    insertion: AutomaticStylesInsertion,
    canonical: &str,
) -> Result<String> {
    match insertion {
        AutomaticStylesInsertion::BeforeEnd(position) => {
            let mut out = String::with_capacity(xml.len() + canonical.len());
            let mut cursor = 0usize;
            for range in ranges {
                if range.end > position {
                    return Err(Error::InvalidFormat(
                        "managed style range exceeds automatic styles container".to_string(),
                    ));
                }
                out.push_str(&xml[cursor..range.start]);
                cursor = range.end;
            }
            out.push_str(&xml[cursor..position]);
            out.push_str(canonical);
            out.push_str(&xml[position..]);
            Ok(out)
        },
        AutomaticStylesInsertion::ExpandEmpty { slash, name } => {
            Ok(format!("{}>{canonical}</{name}>", &xml[..slash]))
        },
    }
}

fn managed_style_ranges(xml: &str) -> Result<(Vec<Range<usize>>, AutomaticStylesInsertion)> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut candidate: Option<(usize, bool, bool)> = None;
    let mut ranges = Vec::new();
    let mut insertion = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_style_namespace = is_namespace(&namespace, STYLE_NAMESPACE);
        let is_office_namespace = is_namespace(&namespace, OFFICE_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not allowed in automatic styles XML".to_string(),
                ));
            },
            Event::Start(element) => {
                if depth == 1 && is_style_namespace && element.local_name().as_ref() == b"style" {
                    let is_cell = optional_attribute(
                        reader.resolver(),
                        reader.decoder(),
                        &element,
                        b"family",
                    )?
                    .as_deref()
                        == Some("table-cell");
                    candidate = Some((event_start, is_cell, false));
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && (element.local_name().as_ref() == b"map"
                        || element.local_name().as_ref() == b"table-cell-properties"
                            && optional_attribute(
                                reader.resolver(),
                                reader.decoder(),
                                &element,
                                b"cell-protect",
                            )?
                            .is_some())
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    let slash = xml[..event_end].rfind("/>").ok_or_else(|| {
                        Error::InvalidFormat("malformed empty automatic styles".to_string())
                    })?;
                    let name =
                        String::from_utf8(element.name().as_ref().to_vec()).map_err(|_error| {
                            Error::InvalidFormat("automatic styles name is not UTF-8".to_string())
                        })?;
                    insertion = Some(AutomaticStylesInsertion::ExpandEmpty { slash, name });
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && (element.local_name().as_ref() == b"map"
                        || element.local_name().as_ref() == b"table-cell-properties"
                            && optional_attribute(
                                reader.resolver(),
                                reader.decoder(),
                                &element,
                                b"cell-protect",
                            )?
                            .is_some())
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid automatic styles depth".to_string())
                })?;
                if depth == 1
                    && is_style_namespace
                    && element.local_name().as_ref() == b"style"
                    && let Some((start, is_cell, managed)) = candidate.take()
                    && is_cell
                    && managed
                {
                    ranges.push(start..event_end);
                }
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    insertion = Some(AutomaticStylesInsertion::BeforeEnd(event_start));
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    Ok((
        ranges,
        insertion.ok_or_else(|| {
            Error::InvalidFormat("missing office:automatic-styles container".to_string())
        })?,
    ))
}

fn write_managed_styles(conditionals: &[ConditionalStyle], protections: &[TableStyle]) -> String {
    let formula_prefixes = conditionals
        .iter()
        .flat_map(|style| &style.rules)
        .filter_map(|rule| rule.formula_namespace.as_ref())
        .map(|namespace| namespace.prefix.as_str())
        .collect::<HashSet<_>>();
    let mut style_prefix = "style".to_string();
    let mut suffix = 0usize;
    while formula_prefixes.contains(style_prefix.as_str()) {
        suffix += 1;
        style_prefix = format!("style{suffix}");
    }
    let mut out = String::new();
    for conditional in conditionals {
        let protection = protections
            .iter()
            .find(|style| style.style_name == conditional.style_name)
            .map(|style| style.protection);
        write_managed_style(
            &mut out,
            &style_prefix,
            &conditional.style_name,
            conditional.parent_style_name.as_deref(),
            protection,
            &conditional.rules,
        );
    }
    for protection in protections {
        if conditionals
            .iter()
            .any(|style| style.style_name == protection.style_name)
        {
            continue;
        }
        write_managed_style(
            &mut out,
            &style_prefix,
            &protection.style_name,
            protection.parent_style_name.as_deref(),
            Some(protection.protection),
            &[],
        );
    }
    out
}

fn write_managed_style(
    out: &mut String,
    prefix: &str,
    name: &str,
    parent: Option<&str>,
    protection: Option<Protection>,
    rules: &[Rule],
) {
    out.push('<');
    out.push_str(prefix);
    out.push_str(":style xmlns:");
    out.push_str(prefix);
    out.push_str("=\"");
    out.push_str(&escape_xml(&String::from_utf8_lossy(STYLE_NAMESPACE)));
    out.push_str("\" ");
    out.push_str(prefix);
    out.push_str(":name=\"");
    out.push_str(&escape_xml(name));
    out.push_str("\" ");
    out.push_str(prefix);
    out.push_str(":family=\"table-cell\"");
    if let Some(parent) = parent {
        out.push(' ');
        out.push_str(prefix);
        out.push_str(":parent-style-name=\"");
        out.push_str(&escape_xml(parent));
        out.push('"');
    }
    out.push('>');
    if let Some(protection) = protection {
        out.push('<');
        out.push_str(prefix);
        out.push_str(":table-cell-properties ");
        out.push_str(prefix);
        out.push_str(":cell-protect=\"");
        out.push_str(protection.as_str());
        out.push_str("\"/>");
    }
    for rule in rules {
        out.push('<');
        out.push_str(prefix);
        out.push_str(":map");
        if let Some(namespace) = &rule.formula_namespace {
            out.push_str(" xmlns:");
            out.push_str(&namespace.prefix);
            out.push_str("=\"");
            out.push_str(&escape_xml(&namespace.uri));
            out.push('"');
        }
        out.push(' ');
        out.push_str(prefix);
        out.push_str(":condition=\"");
        out.push_str(&escape_xml(&rule.condition));
        out.push_str("\" ");
        out.push_str(prefix);
        out.push_str(":apply-style-name=\"");
        out.push_str(&escape_xml(&rule.apply_style_name));
        out.push('"');
        if let Some(base) = &rule.base_cell_address {
            out.push(' ');
            out.push_str(prefix);
            out.push_str(":base-cell-address=\"");
            out.push_str(&escape_xml(base));
            out.push('"');
        }
        out.push_str("/>");
    }
    out.push_str("</");
    out.push_str(prefix);
    out.push_str(":style>");
}

enum AutomaticStylesInsertion {
    BeforeEnd(usize),
    ExpandEmpty { slash: usize, name: String },
}

#[cfg_attr(not(test), allow(dead_code))]
fn conditional_style_ranges(xml: &str) -> Result<(Vec<Range<usize>>, AutomaticStylesInsertion)> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut candidate: Option<(usize, bool, bool)> = None;
    let mut ranges = Vec::new();
    let mut insertion = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_style_namespace = is_namespace(&namespace, STYLE_NAMESPACE);
        let is_office_namespace = is_namespace(&namespace, OFFICE_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not allowed in automatic styles XML".to_string(),
                ));
            },
            Event::Start(element) => {
                if depth == 1 && is_style_namespace && element.local_name().as_ref() == b"style" {
                    let is_cell = optional_attribute(
                        reader.resolver(),
                        reader.decoder(),
                        &element,
                        b"family",
                    )?
                    .as_deref()
                        == Some("table-cell");
                    candidate = Some((event_start, is_cell, false));
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && element.local_name().as_ref() == b"map"
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    let slash = xml[..event_end].rfind("/>").ok_or_else(|| {
                        Error::InvalidFormat("malformed empty automatic styles".to_string())
                    })?;
                    let name =
                        String::from_utf8(element.name().as_ref().to_vec()).map_err(|_error| {
                            Error::InvalidFormat("automatic styles name is not UTF-8".to_string())
                        })?;
                    insertion = Some(AutomaticStylesInsertion::ExpandEmpty { slash, name });
                } else if depth == 2
                    && candidate.is_some()
                    && is_style_namespace
                    && element.local_name().as_ref() == b"map"
                {
                    candidate.as_mut().expect("checked candidate").2 = true;
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid automatic styles depth".to_string())
                })?;
                if depth == 1
                    && is_style_namespace
                    && element.local_name().as_ref() == b"style"
                    && let Some((start, is_cell, has_map)) = candidate.take()
                    && is_cell
                    && has_map
                {
                    ranges.push(start..event_end);
                }
                if depth == 0
                    && is_office_namespace
                    && element.local_name().as_ref() == b"automatic-styles"
                {
                    insertion = Some(AutomaticStylesInsertion::BeforeEnd(event_start));
                }
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    let insertion = insertion.ok_or_else(|| {
        Error::InvalidFormat("missing office:automatic-styles container".to_string())
    })?;
    Ok((ranges, insertion))
}

#[cfg_attr(not(test), allow(dead_code))]
fn write_conditional_styles(styles: &[ConditionalStyle]) -> String {
    let formula_prefixes = styles
        .iter()
        .flat_map(|style| &style.rules)
        .filter_map(|rule| rule.formula_namespace.as_ref())
        .map(|namespace| namespace.prefix.as_str())
        .collect::<HashSet<_>>();
    let mut style_prefix = "style".to_string();
    let mut suffix = 0usize;
    while formula_prefixes.contains(style_prefix.as_str()) {
        suffix += 1;
        style_prefix = format!("style{suffix}");
    }
    let mut out = String::new();
    for style in styles {
        out.push('<');
        out.push_str(&style_prefix);
        out.push_str(":style xmlns:");
        out.push_str(&style_prefix);
        out.push_str("=\"");
        out.push_str(&escape_xml(&String::from_utf8_lossy(STYLE_NAMESPACE)));
        out.push_str("\" ");
        out.push_str(&style_prefix);
        out.push_str(":name=\"");
        out.push_str(&escape_xml(&style.style_name));
        out.push_str("\" ");
        out.push_str(&style_prefix);
        out.push_str(":family=\"table-cell\"");
        if let Some(parent) = &style.parent_style_name {
            out.push(' ');
            out.push_str(&style_prefix);
            out.push_str(":parent-style-name=\"");
            out.push_str(&escape_xml(parent));
            out.push('"');
        }
        out.push('>');
        for rule in &style.rules {
            out.push('<');
            out.push_str(&style_prefix);
            out.push_str(":map");
            if let Some(namespace) = &rule.formula_namespace {
                out.push_str(" xmlns:");
                out.push_str(&namespace.prefix);
                out.push_str("=\"");
                out.push_str(&escape_xml(&namespace.uri));
                out.push('"');
            }
            out.push(' ');
            out.push_str(&style_prefix);
            out.push_str(":condition=\"");
            out.push_str(&escape_xml(&rule.condition));
            out.push_str("\" ");
            out.push_str(&style_prefix);
            out.push_str(":apply-style-name=\"");
            out.push_str(&escape_xml(&rule.apply_style_name));
            out.push('"');
            if let Some(base) = &rule.base_cell_address {
                out.push(' ');
                out.push_str(&style_prefix);
                out.push_str(":base-cell-address=\"");
                out.push_str(&escape_xml(base));
                out.push('"');
            }
            out.push_str("/>");
        }
        out.push_str("</");
        out.push_str(&style_prefix);
        out.push_str(":style>");
    }
    out
}

/// # Errors
///
/// Returns an error when the source XML is malformed.
pub fn extract_automatic_styles(xml: &str) -> Result<Option<PreservedXmlFragment>> {
    extract_office_fragment(xml, b"automatic-styles")
}

/// # Errors
///
/// Returns an error when the source XML is malformed.
pub fn extract_font_face_decls(xml: &str) -> Result<Option<PreservedXmlFragment>> {
    extract_office_fragment(xml, b"font-face-decls")
}

fn extract_office_fragment(
    xml: &str,
    expected_local_name: &[u8],
) -> Result<Option<PreservedXmlFragment>> {
    let fragment_name = String::from_utf8_lossy(expected_local_name);
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut namespaces = BTreeMap::new();
    let mut range_start = None;
    let mut depth = 0usize;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_office_namespace = is_namespace(&namespace, OFFICE_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element)
                if is_office_namespace && element.local_name().as_ref() == b"document-content" =>
            {
                collect_namespaces(&reader, &element, &mut namespaces)?;
            },
            Event::Start(element)
                if is_office_namespace && element.local_name().as_ref() == expected_local_name =>
            {
                if range_start.is_some() {
                    return Err(Error::InvalidFormat(
                        "nested preserved office fragment".to_string(),
                    ));
                }
                range_start = Some(event_start);
                depth = 1;
            },
            Event::Empty(element)
                if is_office_namespace && element.local_name().as_ref() == expected_local_name =>
            {
                if range_start.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate preserved office fragment".to_string(),
                    ));
                }
                return Ok(Some(PreservedXmlFragment {
                    xml: xml[event_start..event_end].to_string(),
                    namespaces,
                }));
            },
            Event::Start(_) if range_start.is_some() => depth += 1,
            Event::End(element) if range_start.is_some() => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid office:{fragment_name} depth"))
                })?;
                if depth == 0 {
                    if !is_office_namespace || element.local_name().as_ref() != expected_local_name
                    {
                        return Err(Error::InvalidFormat(format!(
                            "malformed office:{fragment_name} element"
                        )));
                    }
                    let start = range_start.take().expect("checked range");
                    return Ok(Some(PreservedXmlFragment {
                        xml: xml[start..event_end].to_string(),
                        namespaces,
                    }));
                }
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if range_start.is_some() {
        return Err(Error::InvalidFormat(format!(
            "unterminated office:{fragment_name} element"
        )));
    }
    Ok(None)
}

fn collect_namespaces(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespaces: &mut BTreeMap<String, String>,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        let prefix = String::from_utf8(prefix.to_vec()).map_err(|_error| {
            Error::InvalidFormat("namespace prefix is not valid UTF-8".to_string())
        })?;
        let uri = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid namespace URI: {error}")))?
            .into_owned();
        namespaces.insert(prefix, uri);
    }
    Ok(())
}
