//! Bounded SpreadsheetML data-consolidation parsing and serialization.

use std::fmt::Write;

use litchi_ooxml_common::mce::process_str;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::Result;

use super::model::{
    Conformance, DataConsolidation, Function, RangeReference, Reference, ReferenceSource,
    References, checked_relationship_id, checked_xstring, validate_reference_count,
};
use super::{
    MAX_DATA_REFERENCES, STRICT_MAIN, STRICT_REL, TRANSITIONAL_MAIN, TRANSITIONAL_REL, invalid,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope {
    Worksheet,
    Consolidate,
    DataRefs,
    DataRef,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Unbound,
    Main,
    Relationship,
    Other,
}

#[derive(Default)]
struct ConsolidationBuilder {
    function: Option<Function>,
    left_labels: Option<bool>,
    start_labels: Option<bool>,
    top_labels: Option<bool>,
    link: Option<bool>,
    data_references: Option<References>,
}

/// Parses the direct worksheet `dataConsolidate` child after applying shared MCE processing.
pub fn parse_worksheet_data_consolidation(xml: &[u8]) -> Result<Option<DataConsolidation>> {
    let source = std::str::from_utf8(xml)
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    let processed = process_str(source)?;
    let mut reader = NsReader::from_reader(processed.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut scopes = Vec::new();
    let mut builder: Option<ConsolidationBuilder> = None;
    let mut declared_count: Option<u32> = None;
    let mut references = Vec::new();
    let mut seen_consolidation = false;
    let mut passed_consolidation_position = false;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid worksheet XML: {error}")))?;
        let namespace = namespace_kind(resolved)?;
        match event {
            Event::Start(element) => {
                let scope = begin_element(
                    &reader,
                    &element,
                    namespace,
                    scopes.last().copied(),
                    &mut builder,
                    &mut declared_count,
                    &mut references,
                    &mut seen_consolidation,
                    &mut passed_consolidation_position,
                )?;
                scopes.push(scope);
            },
            Event::Empty(element) => {
                let scope = begin_element(
                    &reader,
                    &element,
                    namespace,
                    scopes.last().copied(),
                    &mut builder,
                    &mut declared_count,
                    &mut references,
                    &mut seen_consolidation,
                    &mut passed_consolidation_position,
                )?;
                end_scope(scope, &mut builder, &mut declared_count, &mut references)?;
            },
            Event::End(_) => {
                let scope = scopes
                    .pop()
                    .ok_or_else(|| invalid("unexpected worksheet end element"))?;
                end_scope(scope, &mut builder, &mut declared_count, &mut references)?;
            },
            Event::Text(text)
                if matches!(
                    scopes.last(),
                    Some(Scope::Consolidate | Scope::DataRefs | Scope::DataRef)
                ) && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("dataConsolidate family cannot contain text"));
            },
            Event::CData(text)
                if matches!(
                    scopes.last(),
                    Some(Scope::Consolidate | Scope::DataRefs | Scope::DataRef)
                ) && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("dataConsolidate family cannot contain CDATA"));
            },
            Event::DocType(_) => {
                return Err(invalid("worksheet XML cannot contain a document type"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !scopes.is_empty() {
        return Err(invalid("unterminated worksheet XML"));
    }
    builder.map(finish_builder).transpose()
}

#[allow(clippy::too_many_arguments)]
fn begin_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    parent: Option<Scope>,
    builder: &mut Option<ConsolidationBuilder>,
    declared_count: &mut Option<u32>,
    references: &mut Vec<Reference>,
    seen_consolidation: &mut bool,
    passed_consolidation_position: &mut bool,
) -> Result<Scope> {
    let local = element.local_name();
    let local = local.as_ref();
    let main = namespace == NamespaceKind::Main;
    match parent {
        None => {
            if !main || local != b"worksheet" {
                return Err(invalid("expected SpreadsheetML worksheet root"));
            }
            Ok(Scope::Worksheet)
        },
        Some(Scope::Worksheet) => {
            if local == b"dataConsolidate" {
                if !main {
                    return Err(invalid("spoofed dataConsolidate element namespace"));
                }
                if *seen_consolidation {
                    return Err(invalid("duplicate worksheet dataConsolidate element"));
                }
                if *passed_consolidation_position {
                    return Err(invalid("dataConsolidate is out of worksheet schema order"));
                }
                *seen_consolidation = true;
                *builder = Some(parse_consolidation_attributes(reader, element)?);
                Ok(Scope::Consolidate)
            } else {
                if main {
                    let position = worksheet_child_position(local);
                    if *seen_consolidation && position == Some(false) {
                        return Err(invalid(
                            "worksheet child precedes dataConsolidate in schema order",
                        ));
                    }
                    if !*seen_consolidation && position == Some(true) {
                        *passed_consolidation_position = true;
                    }
                }
                Ok(Scope::Other)
            }
        },
        Some(Scope::Consolidate) => {
            if local != b"dataRefs" || !main {
                return Err(invalid(if local == b"dataRefs" {
                    "spoofed dataRefs element namespace"
                } else {
                    "unknown dataConsolidate child element"
                }));
            }
            let current = builder
                .as_ref()
                .ok_or_else(|| invalid("missing dataConsolidate state"))?;
            if current.data_references.is_some()
                || declared_count.is_some()
                || !references.is_empty()
            {
                return Err(invalid("duplicate dataRefs element"));
            }
            *declared_count = parse_data_refs_attributes(reader, element)?;
            Ok(Scope::DataRefs)
        },
        Some(Scope::DataRefs) => {
            if local != b"dataRef" || !main {
                return Err(invalid(if local == b"dataRef" {
                    "spoofed dataRef element namespace"
                } else {
                    "unknown dataRefs child element"
                }));
            }
            if references.len() >= MAX_DATA_REFERENCES {
                return Err(invalid(format!(
                    "dataRefs exceeds safety limit {MAX_DATA_REFERENCES}"
                )));
            }
            references.push(parse_data_ref_attributes(reader, element)?);
            Ok(Scope::DataRef)
        },
        Some(Scope::DataRef) => Err(invalid("dataRef must be a leaf element")),
        Some(Scope::Other) => Ok(Scope::Other),
    }
}

fn end_scope(
    scope: Scope,
    builder: &mut Option<ConsolidationBuilder>,
    declared_count: &mut Option<u32>,
    references: &mut Vec<Reference>,
) -> Result<()> {
    if scope == Scope::DataRefs {
        validate_reference_count(references.len())?;
        if let Some(count) = *declared_count
            && count as usize != references.len()
        {
            return Err(invalid(format!(
                "dataRefs count {count} does not match {} dataRef children",
                references.len()
            )));
        }
        let collection = References::from_parts(std::mem::take(references), declared_count.take());
        let target = builder
            .as_mut()
            .ok_or_else(|| invalid("missing dataConsolidate state"))?;
        if target.data_references.replace(collection).is_some() {
            return Err(invalid("duplicate dataRefs element"));
        }
    }
    Ok(())
}

fn finish_builder(builder: ConsolidationBuilder) -> Result<DataConsolidation> {
    Ok(DataConsolidation::from_parts(
        builder.function.unwrap_or_default(),
        builder.left_labels.unwrap_or(false),
        builder.start_labels.unwrap_or(false),
        builder.top_labels.unwrap_or(false),
        builder.link.unwrap_or(false),
        builder.data_references,
    ))
}

fn parse_consolidation_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ConsolidationBuilder> {
    let mut value = ConsolidationBuilder::default();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid dataConsolidate attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(namespace)?;
        if namespace != NamespaceKind::Unbound {
            return Err(invalid("unknown namespaced dataConsolidate attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                invalid(format!("invalid dataConsolidate attribute value: {error}"))
            })?;
        match local.as_ref() {
            b"function" => set_once(&mut value.function, Function::parse(&text)?, "function")?,
            b"leftLabels" => set_once(
                &mut value.left_labels,
                parse_bool(&text, "leftLabels")?,
                "leftLabels",
            )?,
            b"startLabels" => set_once(
                &mut value.start_labels,
                parse_bool(&text, "startLabels")?,
                "startLabels",
            )?,
            b"topLabels" => set_once(
                &mut value.top_labels,
                parse_bool(&text, "topLabels")?,
                "topLabels",
            )?,
            b"link" => set_once(&mut value.link, parse_bool(&text, "link")?, "link")?,
            _ => return Err(invalid("unknown dataConsolidate attribute")),
        }
    }
    Ok(value)
}

fn parse_data_refs_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<u32>> {
    let mut count = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid dataRefs attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(namespace)? != NamespaceKind::Unbound || local.as_ref() != b"count" {
            return Err(invalid("unknown dataRefs attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid dataRefs count: {error}")))?;
        let parsed = text
            .parse::<u32>()
            .map_err(|_| invalid("dataRefs count must be unsignedInt"))?;
        if parsed as usize > MAX_DATA_REFERENCES {
            return Err(invalid(format!(
                "dataRefs count exceeds safety limit {MAX_DATA_REFERENCES}"
            )));
        }
        set_once(&mut count, parsed, "count")?;
    }
    Ok(count)
}

fn parse_data_ref_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Reference> {
    let mut name = None;
    let mut sheet = None;
    let mut reference = None;
    let mut relationship_id = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid dataRef attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(namespace)?;
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid dataRef attribute value: {error}")))?
            .into_owned();
        match (namespace, local.as_ref()) {
            (NamespaceKind::Unbound, b"name") => {
                set_once(&mut name, checked_xstring(text, "dataRef name")?, "name")?
            },
            (NamespaceKind::Unbound, b"sheet") => {
                set_once(&mut sheet, checked_xstring(text, "dataRef sheet")?, "sheet")?
            },
            (NamespaceKind::Unbound, b"ref") => {
                set_once(&mut reference, RangeReference::new(text)?, "ref")?
            },
            (NamespaceKind::Relationship, b"id") => {
                set_once(&mut relationship_id, checked_relationship_id(text)?, "r:id")?
            },
            _ => return Err(invalid("unknown or spoofed dataRef attribute")),
        }
    }
    let source = match (name, sheet, reference) {
        (Some(name), None, None) => ReferenceSource::DefinedName(name),
        (None, Some(sheet), Some(reference)) => ReferenceSource::Range { sheet, reference },
        _ => {
            return Err(invalid(
                "dataRef requires exactly name or the sheet and ref pair",
            ));
        },
    };
    Ok(Reference::from_parts(source, relationship_id))
}

/// Serializes one canonical, namespace-complete `dataConsolidate` fragment.
pub fn write_worksheet_data_consolidation(
    value: &DataConsolidation,
    conformance: Conformance,
) -> Result<String> {
    if let Some(data_refs) = value.data_references() {
        validate_reference_count(data_refs.references().len())?;
    }
    let has_relationships = value.data_references().is_some_and(|refs| {
        refs.references()
            .iter()
            .any(|reference| reference.relationship_id().is_some())
    });
    let mut xml = String::new();
    write!(
        xml,
        "<dataConsolidate xmlns=\"{}\"",
        conformance.main_namespace()
    )
    .unwrap();
    if has_relationships {
        write!(xml, " xmlns:r=\"{}\"", conformance.relationship_namespace()).unwrap();
    }
    if value.function() != Function::Sum {
        write!(xml, " function=\"{}\"", value.function().as_str()).unwrap();
    }
    write_true_attribute(&mut xml, "leftLabels", value.left_labels());
    write_true_attribute(&mut xml, "startLabels", value.start_labels());
    write_true_attribute(&mut xml, "topLabels", value.top_labels());
    write_true_attribute(&mut xml, "link", value.link());
    let Some(data_refs) = value.data_references() else {
        xml.push_str("/>");
        return Ok(xml);
    };
    xml.push('>');
    write!(xml, "<dataRefs count=\"{}\">", data_refs.references().len()).unwrap();
    for reference in data_refs.references() {
        xml.push_str("<dataRef");
        match reference.source() {
            ReferenceSource::DefinedName(name) => write_attribute(&mut xml, "name", name),
            ReferenceSource::Range { sheet, reference } => {
                write_attribute(&mut xml, "ref", reference.as_str());
                write_attribute(&mut xml, "sheet", sheet);
            },
        }
        if let Some(id) = reference.relationship_id() {
            write_attribute(&mut xml, "r:id", id);
        }
        xml.push_str("/>");
    }
    xml.push_str("</dataRefs></dataConsolidate>");
    Ok(xml)
}

fn write_true_attribute(xml: &mut String, name: &str, value: bool) {
    if value {
        write!(xml, " {name}=\"1\"").unwrap();
    }
}

fn write_attribute(xml: &mut String, name: &str, value: &str) {
    write!(xml, " {name}=\"").unwrap();
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
    xml.push('"');
}

fn namespace_kind(result: ResolveResult<'_>) -> Result<NamespaceKind> {
    match result {
        ResolveResult::Unbound => Ok(NamespaceKind::Unbound),
        ResolveResult::Bound(namespace) if is_main_namespace(namespace.as_ref()) => {
            Ok(NamespaceKind::Main)
        },
        ResolveResult::Bound(namespace) if is_relationship_namespace(namespace.as_ref()) => {
            Ok(NamespaceKind::Relationship)
        },
        ResolveResult::Bound(_) => Ok(NamespaceKind::Other),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix {}",
            String::from_utf8_lossy(&prefix)
        ))),
    }
}

fn is_main_namespace(namespace: &[u8]) -> bool {
    namespace == TRANSITIONAL_MAIN.as_bytes() || namespace == STRICT_MAIN.as_bytes()
}

fn is_relationship_namespace(namespace: &[u8]) -> bool {
    namespace == TRANSITIONAL_REL.as_bytes() || namespace == STRICT_REL.as_bytes()
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn worksheet_child_position(local: &[u8]) -> Option<bool> {
    const BEFORE: &[&[u8]] = &[
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
    ];
    const AFTER: &[&[u8]] = &[
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
    if BEFORE.contains(&local) {
        Some(false)
    } else if AFTER.contains(&local) {
        Some(true)
    } else {
        None
    }
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{name} must be an XML boolean"))),
    }
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        Err(invalid(format!("duplicate {name} attribute")))
    } else {
        Ok(())
    }
}
