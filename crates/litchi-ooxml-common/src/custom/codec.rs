//! Namespace-aware XML codec for the custom-properties part.

use chrono::{DateTime, SecondsFormat, Utc};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;
use std::borrow::Cow;
use std::collections::{BTreeSet, HashSet};

use super::model::{
    Props, Value, WireKind, fold_name, validate_name, validate_value, validate_xml_text,
    value_text_bytes,
};
use super::reserved;
use super::schema::{
    CUSTOM_NS, DOCUMENT_SUMMARY_FORMAT_ID, MAX_ATTRIBUTE_BYTES, MAX_ATTRIBUTES, MAX_PROPERTIES,
    MAX_TEXT_BYTES, MAX_TOTAL_NAME_BYTES, MAX_TOTAL_TEXT_BYTES, MAX_XML_BYTES, MAX_XML_DEPTH,
    SUMMARY_FORMAT_ID, VT_NS, checked_depth, checked_increment, checked_total, count_node, invalid,
    limit,
};
use crate::leniency::{CustomPropertiesDefect, ToleranceReport};
use crate::xml::attributes::BytesStartExt as _;
use crate::{Error, Result};

struct PendingProperty {
    name: String,
    pid: i32,
    format_id: String,
    value: Option<(WireKind, Value)>,
}

/// Decode custom-properties XML strictly: every deviation a package read
/// repairs is an error here.
#[cfg(test)]
pub(crate) fn decode(xml: &[u8]) -> Result<Props> {
    decode_with(xml, &mut Sink::strict())
}

/// Decode custom-properties XML, repairing per property and recording each
/// repair into the returned collection's report (LENIENCY_AUDIT 8.3).
pub(crate) fn decode_tolerant(xml: &[u8], part: &str) -> Result<Props> {
    let mut report = ToleranceReport::default();
    let mut props = decode_with(xml, &mut Sink::tolerant(&mut report, part))?;
    props.tolerance_report = report;
    Ok(props)
}

/// Where per-property repairs go. A strict sink turns every repair back into
/// the error it replaces.
pub(super) struct Sink<'a> {
    report: Option<&'a mut ToleranceReport>,
    part: &'a str,
}

impl<'a> Sink<'a> {
    #[cfg(test)]
    pub(super) fn strict() -> Self {
        Self {
            report: None,
            part: "",
        }
    }

    pub(super) fn tolerant(report: &'a mut ToleranceReport, part: &'a str) -> Self {
        Self {
            report: Some(report),
            part,
        }
    }

    /// Repair `error` by recording `defect`, or return it when strict. Only
    /// [`Error::Invalid`] content errors are repairable: XML, limit and
    /// allocation errors always propagate.
    pub(super) fn tolerate(
        &mut self,
        error: Error,
        defect: CustomPropertiesDefect,
        item: Option<&str>,
        value: Option<&str>,
    ) -> Result<()> {
        match (&mut self.report, &error) {
            (Some(report), Error::Invalid(_)) => {
                report.record(defect, self.part, item, value);
                Ok(())
            },
            _ => Err(error),
        }
    }
}

fn decode_with(xml: &[u8], sink: &mut Sink<'_>) -> Result<Props> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit(
            "custom-properties XML bytes",
            MAX_XML_BYTES,
            xml.len(),
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut props = Props::new();
    let mut pids = HashSet::new();
    let mut pending: Option<PendingProperty> = None;
    let mut value_kind: Option<WireKind> = None;
    let mut value_text = String::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    // The property (or unknown root child) at depth 2 is being skipped until
    // its end tag, after a recorded per-property repair.
    let mut abandoned = false;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event()?;
        if abandoned {
            match event {
                Event::Start(_) => {
                    count_node(&mut nodes)?;
                    depth = checked_depth(depth)?;
                },
                Event::End(_) => {
                    count_node(&mut nodes)?;
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("custom-properties XML depth underflow"))?;
                    if depth <= 1 {
                        abandoned = false;
                        pending = None;
                        value_kind = None;
                        value_text.clear();
                    }
                },
                Event::DocType(_) => {
                    return Err(invalid(
                        "DTD declarations are forbidden in custom-properties XML",
                    ));
                },
                Event::Eof => break,
                Event::PI(_) => {},
                _ => count_node(&mut nodes)?,
            }
            continue;
        }
        match event {
            Event::Decl(_) => {
                count_node(&mut nodes)?;
                if declaration_seen || root_seen {
                    return Err(invalid("XML declaration must occur once before the root"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "DTD declarations are forbidden in custom-properties XML",
                ));
            },
            Event::PI(_) => sink.tolerate(
                invalid("processing instructions are forbidden in custom-properties XML"),
                CustomPropertiesDefect::ProcessingInstruction,
                None,
                None,
            )?,
            Event::Start(element) => {
                count_node(&mut nodes)?;
                let child_depth = checked_depth(depth)?;
                match child_depth {
                    1 => {
                        if root_seen || root_closed {
                            return Err(invalid(
                                "custom-properties XML must contain exactly one root",
                            ));
                        }
                        validate_root(&namespace, &element, decoder)?;
                        root_seen = true;
                    },
                    2 => {
                        if !is_name(&namespace, &element, CUSTOM_NS, b"property") {
                            let name = display_name(element.name().as_ref());
                            sink.tolerate(
                                invalid(format!(
                                    "unexpected element '{name}' below custom-properties root"
                                )),
                                CustomPropertiesDefect::UnknownElement,
                                Some(&name),
                                None,
                            )?;
                            abandoned = true;
                        } else if pending.is_some() {
                            return Err(invalid("custom properties cannot be nested"));
                        } else {
                            match admit_property(&element, decoder, &mut pids, sink)? {
                                Some(parsed) => pending = Some(parsed),
                                None => abandoned = true,
                            }
                        }
                    },
                    3 => {
                        let property = pending.as_ref().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        match start_value(property, value_kind.is_some(), &namespace, &element) {
                            Ok(kind) => {
                                value_kind = Some(kind);
                                value_text.clear();
                            },
                            Err((error, defect, value)) => {
                                let name = property.name.clone();
                                sink.tolerate(error, defect, Some(&name), value.as_deref())?;
                                abandoned = true;
                            },
                        }
                    },
                    _ => {
                        return Err(limit(
                            "custom-properties XML depth",
                            MAX_XML_DEPTH,
                            child_depth,
                        ));
                    },
                }
                depth = child_depth;
            },
            Event::Empty(element) => {
                count_node(&mut nodes)?;
                let child_depth = checked_depth(depth)?;
                match child_depth {
                    1 => {
                        if root_seen || root_closed {
                            return Err(invalid(
                                "custom-properties XML must contain exactly one root",
                            ));
                        }
                        validate_root(&namespace, &element, decoder)?;
                        root_seen = true;
                        root_closed = true;
                    },
                    2 => {
                        let is_property = is_name(&namespace, &element, CUSTOM_NS, b"property");
                        let name = if is_property {
                            property_name_attribute(&element, decoder)
                        } else {
                            Some(display_name(element.name().as_ref()))
                        };
                        sink.tolerate(
                            invalid("custom property must contain exactly one typed value"),
                            if is_property {
                                CustomPropertiesDefect::MalformedProperty
                            } else {
                                CustomPropertiesDefect::UnknownElement
                            },
                            name.as_deref(),
                            None,
                        )?;
                    },
                    3 => {
                        let property = pending.as_mut().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        let parsed =
                            start_value(property, value_kind.is_some(), &namespace, &element)
                                .and_then(|kind| {
                                    parse_value(kind, "").map(|value| (kind, value)).map_err(
                                        |error| (error, CustomPropertiesDefect::InvalidValue, None),
                                    )
                                });
                        match parsed {
                            Ok(value) => property.value = Some(value),
                            Err((error, defect, value)) => {
                                let name = property.name.clone();
                                sink.tolerate(error, defect, Some(&name), value.as_deref())?;
                                // The property's own end tag closes the skip.
                                abandoned = true;
                                depth = 2;
                            },
                        }
                    },
                    _ => {
                        return Err(limit(
                            "custom-properties XML depth",
                            MAX_XML_DEPTH,
                            child_depth,
                        ));
                    },
                }
            },
            Event::End(_) => {
                count_node(&mut nodes)?;
                match depth {
                    3 => {
                        let kind = value_kind
                            .take()
                            .ok_or_else(|| invalid("custom-property value state is incomplete"))?;
                        let property = pending.as_mut().ok_or_else(|| {
                            invalid("custom-property value has no owning property")
                        })?;
                        match parse_value(kind, &value_text) {
                            Ok(value) => property.value = Some((kind, value)),
                            Err(error) => {
                                let name = property.name.clone();
                                sink.tolerate(
                                    error,
                                    CustomPropertiesDefect::InvalidValue,
                                    Some(&name),
                                    Some(&value_text),
                                )?;
                                abandoned = true;
                            },
                        }
                        value_text.clear();
                    },
                    2 => {
                        let property = pending
                            .take()
                            .ok_or_else(|| invalid("custom-property record is incomplete"))?;
                        insert_property(&mut props, property, sink)?;
                    },
                    1 => {
                        root_closed = true;
                    },
                    _ => return Err(invalid("unexpected XML end element")),
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("custom-properties XML depth underflow"))?;
            },
            Event::Text(text) => {
                count_node(&mut nodes)?;
                let text = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(format!("invalid XML text: {error}")))?;
                if depth == 3 {
                    append_value_text(&mut value_text, &text)?;
                } else if !text.trim().is_empty() {
                    let error =
                        invalid("non-whitespace text is not allowed outside a property value");
                    abandon_pending(&mut pending, &mut abandoned, depth, error, sink)?;
                }
            },
            Event::CData(text) => {
                count_node(&mut nodes)?;
                if depth != 3 {
                    let error = invalid("CDATA is only allowed inside a property value");
                    abandon_pending(&mut pending, &mut abandoned, depth, error, sink)?;
                    continue;
                }
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(format!("invalid CDATA text: {error}")))?;
                append_value_text(&mut value_text, &text)?;
            },
            Event::GeneralRef(reference) => {
                count_node(&mut nodes)?;
                if depth != 3 {
                    let error =
                        invalid("entity references are only allowed inside a property value");
                    abandon_pending(&mut pending, &mut abandoned, depth, error, sink)?;
                    continue;
                }
                match decode_reference(&reference) {
                    Ok(decoded) => append_value_text(&mut value_text, &decoded)?,
                    Err(error) => {
                        let name = pending.as_ref().map(|property| property.name.clone());
                        sink.tolerate(
                            error,
                            CustomPropertiesDefect::InvalidValue,
                            name.as_deref(),
                            None,
                        )?;
                        abandoned = true;
                    },
                }
            },
            Event::Comment(_) => count_node(&mut nodes)?,
            Event::Eof => break,
        }
    }

    if !root_seen || !root_closed || depth != 0 || pending.is_some() || value_kind.is_some() {
        return Err(invalid(
            "custom-properties XML must contain one complete Properties root",
        ));
    }
    repair_shape_fragments(&mut props, sink)?;
    validate_collection(&props)?;
    Ok(props)
}

/// Content outside a value inside a property costs that property; at the
/// root level it stays an error.
fn abandon_pending(
    pending: &mut Option<PendingProperty>,
    abandoned: &mut bool,
    depth: usize,
    error: Error,
    sink: &mut Sink<'_>,
) -> Result<()> {
    if depth != 2 {
        return Err(error);
    }
    let name = pending.as_ref().map(|property| property.name.clone());
    sink.tolerate(
        error,
        CustomPropertiesDefect::MalformedProperty,
        name.as_deref(),
        None,
    )?;
    *abandoned = true;
    Ok(())
}

/// Admit a property start tag, or `None` after a recorded repair.
fn admit_property(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    pids: &mut HashSet<i32>,
    sink: &mut Sink<'_>,
) -> Result<Option<PendingProperty>> {
    let parsed = match parse_property_attributes(element, decoder) {
        Ok(parsed) => parsed,
        Err(error) => {
            let name = property_name_attribute(element, decoder);
            sink.tolerate(
                error,
                CustomPropertiesDefect::InvalidIdentity,
                name.as_deref(),
                None,
            )?;
            return Ok(None);
        },
    };
    if !pids.insert(parsed.pid) {
        let pid = parsed.pid.to_string();
        sink.tolerate(
            invalid(format!("duplicate custom property PID {pid}")),
            CustomPropertiesDefect::DuplicateProperty,
            Some(&parsed.name),
            Some(&pid),
        )?;
        return Ok(None);
    }
    Ok(Some(parsed))
}

/// Best-effort `name` attribute of a property element, for repair context.
fn property_name_attribute(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Option<String> {
    element
        .checked_attributes()
        .take(MAX_ATTRIBUTES)
        .filter_map(std::result::Result::ok)
        .find(|attribute| attribute.key.as_ref() == b"name")
        .filter(|attribute| attribute.value.len() <= MAX_ATTRIBUTE_BYTES)
        .and_then(|attribute| {
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .ok()
                .map(Cow::into_owned)
        })
}

type ValueFailure = (Error, CustomPropertiesDefect, Option<String>);

/// Validate a value element's position, type, and attributes.
fn start_value(
    property: &PendingProperty,
    value_open: bool,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> std::result::Result<WireKind, ValueFailure> {
    if property.value.is_some() || value_open {
        return Err((
            invalid(format!(
                "custom property '{}' must contain exactly one value",
                property.name
            )),
            CustomPropertiesDefect::MalformedProperty,
            None,
        ));
    }
    let kind = parse_value_element(namespace, element).map_err(|error| {
        (
            error,
            CustomPropertiesDefect::UnsupportedValueType,
            Some(display_name(element.name().as_ref())),
        )
    })?;
    validate_value_attributes(element)
        .map_err(|error| (error, CustomPropertiesDefect::MalformedProperty, None))?;
    Ok(kind)
}

/// Insert a completed property, recording the per-property failures a read
/// repairs; collection-level limits stay errors.
fn insert_property(
    props: &mut Props,
    property: PendingProperty,
    sink: &mut Sink<'_>,
) -> Result<()> {
    let Some((wire, value)) = property.value else {
        return sink.tolerate(
            invalid(format!(
                "custom property '{}' must contain exactly one value",
                property.name
            )),
            CustomPropertiesDefect::MalformedProperty,
            Some(&property.name),
            None,
        );
    };
    if let Err(error) = validate_name(&property.name) {
        return sink.tolerate(
            error,
            CustomPropertiesDefect::InvalidIdentity,
            Some(&property.name),
            None,
        );
    }
    if let Err(error) = validate_value(&value) {
        return sink.tolerate(
            error,
            CustomPropertiesDefect::InvalidValue,
            Some(&property.name),
            None,
        );
    }
    if let Err(error) = reserved::validate_property(&property.name, &value, None) {
        return sink.tolerate(
            error,
            CustomPropertiesDefect::ReservedProperty,
            Some(&property.name),
            None,
        );
    }
    if props.folded_names.contains_key(&fold_name(&property.name)) {
        return sink.tolerate(
            invalid(format!(
                "duplicate custom property name '{}' (names are case-insensitive)",
                property.name
            )),
            CustomPropertiesDefect::DuplicateProperty,
            Some(&property.name),
            None,
        );
    }
    props.insert_parsed(property.name, property.pid, property.format_id, wire, value)
}

/// Reserved shape-ID fragment sets (MS-OI29500 3.11) are validated as a
/// group; a read drops an inconsistent group's fragments, recording each.
fn repair_shape_fragments(props: &mut Props, sink: &mut Sink<'_>) -> Result<()> {
    for (base, error) in reserved::invalid_shape_fragment_sets(props) {
        let fragments: Vec<String> = props
            .names()
            .filter(|name| reserved::is_shape_fragment_of(name, base))
            .map(str::to_owned)
            .collect();
        for name in fragments {
            sink.tolerate(
                invalid(error.to_string()),
                CustomPropertiesDefect::ReservedProperty,
                Some(&name),
                None,
            )?;
            props.remove(&name);
        }
    }
    Ok(())
}

pub(crate) fn encode(props: &Props) -> Result<Vec<u8>> {
    validate_collection(props)?;
    let estimated = estimated_xml_size(props)?;
    let mut writer = Writer::new(Vec::with_capacity(estimated));
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    let mut root = BytesStart::new("Properties");
    root.push_attribute((
        "xmlns",
        std::str::from_utf8(CUSTOM_NS).map_err(|error| {
            Error::Xml(format!(
                "invalid built-in custom-properties namespace: {error}"
            ))
        })?,
    ));
    root.push_attribute((
        "xmlns:vt",
        std::str::from_utf8(VT_NS).map_err(|error| {
            Error::Xml(format!("invalid built-in variant-types namespace: {error}"))
        })?,
    ));
    writer.write_event(Event::Start(root))?;

    let mut ordered: Vec<_> = props.properties.iter().collect();
    ordered.sort_unstable_by_key(|(_, property)| property.pid);
    for (name, property) in ordered {
        let pid = property.pid.to_string();
        let mut element = BytesStart::new("property");
        element.push_attribute(("fmtid", property.format_id.as_str()));
        element.push_attribute(("pid", pid.as_str()));
        element.push_attribute(("name", name.as_str()));
        writer.write_event(Event::Start(element))?;

        let value_name = property.wire.qualified_name();
        if property.wire == WireKind::Empty {
            writer.write_event(Event::Empty(BytesStart::new(value_name)))?;
        } else {
            writer.write_event(Event::Start(BytesStart::new(value_name)))?;
            let value = value_lexical(&property.value)?;
            writer.write_event(Event::Text(BytesText::new(&value)))?;
            writer.write_event(Event::End(BytesEnd::new(value_name)))?;
        }
        writer.write_event(Event::End(BytesEnd::new("property")))?;
    }
    writer.write_event(Event::End(BytesEnd::new("Properties")))?;
    let xml = writer.into_inner();
    if xml.len() > MAX_XML_BYTES {
        return Err(limit(
            "custom-properties XML bytes",
            MAX_XML_BYTES,
            xml.len(),
        ));
    }
    Ok(xml)
}

fn validate_collection(props: &Props) -> Result<()> {
    reserved::validate(props, None)?;
    if props.properties.len() > MAX_PROPERTIES {
        return Err(limit(
            "custom properties",
            MAX_PROPERTIES,
            props.properties.len(),
        ));
    }
    let mut pids = HashSet::with_capacity(props.properties.len());
    let mut folded = BTreeSet::new();
    let mut name_bytes = 0usize;
    let mut text_bytes = 0usize;
    for (name, property) in &props.properties {
        validate_name(name)?;
        validate_value(&property.value)?;
        validate_wire_value(property.wire, &property.value)?;
        if property.pid < 2 || !pids.insert(property.pid) {
            return Err(invalid(format!(
                "invalid or duplicate custom property PID {}",
                property.pid
            )));
        }
        validate_format_id(&property.format_id)?;
        if !folded.insert(fold_name(name)) {
            return Err(invalid(format!(
                "duplicate custom property name '{name}' (names are case-insensitive)"
            )));
        }
        name_bytes = checked_total(
            name_bytes,
            name.len(),
            MAX_TOTAL_NAME_BYTES,
            "custom-property name bytes",
        )?;
        text_bytes = checked_total(
            text_bytes,
            value_text_bytes(&property.value),
            MAX_TOTAL_TEXT_BYTES,
            "custom-property text bytes",
        )?;
    }
    Ok(())
}

fn estimated_xml_size(props: &Props) -> Result<usize> {
    let mut size = 256usize;
    for (name, property) in &props.properties {
        size = size
            .checked_add(256)
            .and_then(|size| {
                name.len()
                    .checked_mul(6)
                    .and_then(|name| size.checked_add(name))
            })
            .and_then(|size| {
                value_text_bytes(&property.value)
                    .checked_mul(6)
                    .and_then(|text| size.checked_add(text))
            })
            .ok_or_else(|| limit("custom-properties XML bytes", MAX_XML_BYTES, usize::MAX))?;
        if size > MAX_XML_BYTES {
            return Err(limit("custom-properties XML bytes", MAX_XML_BYTES, size));
        }
    }
    Ok(size)
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    if !is_name(namespace, element, CUSTOM_NS, b"Properties") {
        return Err(invalid(format!(
            "custom-properties root must be Properties in namespace '{}'",
            String::from_utf8_lossy(CUSTOM_NS)
        )));
    }
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if !is_namespace_declaration(attribute.key.as_ref()) {
            let name = attribute.key.local_name().as_ref().to_vec();
            let _ = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(format!("invalid root attribute: {error}")))?;
            return Err(invalid(format!(
                "unexpected custom-properties root attribute '{}'",
                display_name(&name)
            )));
        }
    }
    Ok(())
}

fn parse_property_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<PendingProperty> {
    let mut name = None;
    let mut pid = None;
    let mut format_id = None;
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        if attribute.key.prefix().is_some() {
            return Err(invalid(format!(
                "custom property has unexpected qualified attribute '{}'",
                display_name(attribute.key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(format!("invalid property attribute: {error}")))?
            .into_owned();
        match attribute.key.local_name().as_ref() {
            b"name" if name.is_none() => name = Some(value),
            b"pid" if pid.is_none() => {
                let parsed = value.parse::<i32>().map_err(|error| {
                    invalid(format!("invalid custom property PID '{value}': {error}"))
                })?;
                pid = Some(parsed);
            },
            b"fmtid" if format_id.is_none() => format_id = Some(normalize_format_id(&value)?),
            local => {
                return Err(invalid(format!(
                    "duplicate or unexpected custom property attribute '{}'",
                    display_name(local)
                )));
            },
        }
    }
    let name = name.ok_or_else(|| invalid("custom property is missing its name attribute"))?;
    validate_name(&name)?;
    let pid = pid.ok_or_else(|| invalid(format!("custom property '{name}' is missing its PID")))?;
    if pid < 2 {
        return Err(invalid(format!(
            "custom property '{name}' has PID {pid}; PIDs must be at least 2"
        )));
    }
    let format_id = format_id
        .ok_or_else(|| invalid(format!("custom property '{name}' is missing its format ID")))?;
    Ok(PendingProperty {
        name,
        pid,
        format_id,
        value: None,
    })
}

fn validate_value_attributes(element: &BytesStart<'_>) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        count = checked_increment(count, "custom-properties XML attributes")?;
        if count > MAX_ATTRIBUTES {
            return Err(limit(
                "custom-properties XML attributes",
                MAX_ATTRIBUTES,
                count,
            ));
        }
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(limit(
                "custom-properties XML attribute bytes",
                MAX_ATTRIBUTE_BYTES,
                attribute.value.len(),
            ));
        }
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err(invalid(format!(
                "custom-property value has unexpected attribute '{}'",
                display_name(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn parse_value_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Result<WireKind> {
    if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == VT_NS) {
        return Err(invalid(format!(
            "custom-property value '{}' is not in the variant-types namespace",
            display_name(element.name().as_ref())
        )));
    }
    WireKind::from_local_name(element.local_name().as_ref()).ok_or_else(|| {
        invalid(format!(
            "unsupported custom-property value type '{}'",
            display_name(element.local_name().as_ref())
        ))
    })
}

fn parse_value(kind: WireKind, text: &str) -> Result<Value> {
    match kind {
        WireKind::Empty => {
            if text.is_empty() {
                Ok(Value::Empty)
            } else {
                Err(invalid("vt:empty cannot contain text"))
            }
        },
        WireKind::Lpstr | WireKind::Lpwstr => {
            validate_xml_text(text, "custom-property text")?;
            Ok(Value::Text(text.to_owned()))
        },
        WireKind::I4 => text
            .trim()
            .parse::<i32>()
            .map(Value::I32)
            .map_err(|error| invalid(format!("invalid vt:i4 value '{text}': {error}"))),
        WireKind::I8 => text
            .trim()
            .parse::<i64>()
            .map(Value::I64)
            .map_err(|error| invalid(format!("invalid vt:i8 value '{text}': {error}"))),
        WireKind::R4 => {
            let value = text
                .trim()
                .parse::<f32>()
                .map_err(|error| invalid(format!("invalid vt:r4 value '{text}': {error}")))?;
            if !value.is_finite() {
                return Err(invalid("vt:r4 custom property must be finite"));
            }
            Ok(Value::F32(value))
        },
        WireKind::R8 => {
            let value = text
                .trim()
                .parse::<f64>()
                .map_err(|error| invalid(format!("invalid vt:r8 value '{text}': {error}")))?;
            if !value.is_finite() {
                return Err(invalid("vt:r8 custom property must be finite"));
            }
            Ok(Value::F64(value))
        },
        WireKind::Bool => match text.trim() {
            "true" | "1" => Ok(Value::Bool(true)),
            "false" | "0" => Ok(Value::Bool(false)),
            value => Err(invalid(format!("invalid vt:bool value '{value}'"))),
        },
        WireKind::Filetime => {
            let value = DateTime::parse_from_rfc3339(text.trim()).map_err(|error| {
                invalid(format!(
                    "invalid vt:filetime RFC3339 date-time '{text}': {error}"
                ))
            })?;
            Ok(Value::Time(value.with_timezone(&Utc)))
        },
    }
}

fn value_lexical(value: &Value) -> Result<Cow<'_, str>> {
    validate_value(value)?;
    Ok(match value {
        Value::Empty => Cow::Borrowed(""),
        Value::Text(text) => Cow::Borrowed(text),
        Value::I32(value) => Cow::Owned(value.to_string()),
        Value::I64(value) => Cow::Owned(value.to_string()),
        Value::F32(value) => Cow::Owned(value.to_string()),
        Value::F64(value) => Cow::Owned(value.to_string()),
        Value::Bool(true) => Cow::Borrowed("true"),
        Value::Bool(false) => Cow::Borrowed("false"),
        Value::Time(value) => Cow::Owned(value.to_rfc3339_opts(SecondsFormat::AutoSi, true)),
    })
}

fn validate_wire_value(wire: WireKind, value: &Value) -> Result<()> {
    let matches = matches!(
        (wire, value),
        (WireKind::Empty, Value::Empty)
            | (WireKind::Lpstr | WireKind::Lpwstr, Value::Text(_))
            | (WireKind::I4, Value::I32(_))
            | (WireKind::I8, Value::I64(_))
            | (WireKind::R4, Value::F32(_))
            | (WireKind::R8, Value::F64(_))
            | (WireKind::Bool, Value::Bool(_))
            | (WireKind::Filetime, Value::Time(_))
    );
    if matches {
        Ok(())
    } else {
        Err(invalid(
            "custom-property wire type does not match its value",
        ))
    }
}

fn validate_format_id(format_id: &str) -> Result<()> {
    let normalized = normalize_format_id(format_id)?;
    if normalized == SUMMARY_FORMAT_ID || normalized == DOCUMENT_SUMMARY_FORMAT_ID {
        return Err(invalid(format!(
            "format ID {normalized} is forbidden for custom properties"
        )));
    }
    Ok(())
}

fn normalize_format_id(format_id: &str) -> Result<String> {
    let bytes = format_id.as_bytes();
    let valid = bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && [9, 14, 19, 24]
            .iter()
            .all(|position| bytes.get(*position) == Some(&b'-'))
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || byte.is_ascii_hexdigit()
        });
    if !valid {
        return Err(invalid(format!(
            "invalid custom property format ID '{format_id}'"
        )));
    }
    let normalized = format_id.to_ascii_uppercase();
    if normalized == SUMMARY_FORMAT_ID || normalized == DOCUMENT_SUMMARY_FORMAT_ID {
        return Err(invalid(format!(
            "format ID {normalized} is forbidden for custom properties"
        )));
    }
    Ok(normalized)
}

fn append_value_text(buffer: &mut String, text: &str) -> Result<()> {
    let actual = buffer
        .len()
        .checked_add(text.len())
        .ok_or_else(|| limit("custom-property text bytes", MAX_TEXT_BYTES, usize::MAX))?;
    if actual > MAX_TEXT_BYTES {
        return Err(limit("custom-property text bytes", MAX_TEXT_BYTES, actual));
    }
    buffer.push_str(text);
    Ok(())
}

fn decode_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| Error::Xml(format!("invalid character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::Xml(format!("invalid entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_owned()),
        "lt" => Ok("<".to_owned()),
        "gt" => Ok(">".to_owned()),
        "quot" => Ok("\"".to_owned()),
        "apos" => Ok("'".to_owned()),
        _ => Err(invalid(format!("unsupported entity reference '&{name};'"))),
    }
}

fn is_name(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    element.local_name().as_ref() == expected_local_name
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected_namespace)
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn display_name(name: &[u8]) -> String {
    String::from_utf8_lossy(name).into_owned()
}
