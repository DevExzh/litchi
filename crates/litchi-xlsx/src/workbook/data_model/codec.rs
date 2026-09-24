//! Bounded XML codec for the inline MS-XLDM workbook descriptor.

use std::collections::HashSet;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};

use super::model::{
    CalculatedTimeColumn, Definition, ModelTimeGrouping, ModelTimeGroupingContentType,
    ModelTimeGroupings, OpaqueXml, Relationship, Table,
};
use super::{
    DATA_MODEL_EXTENSION_URI, MAX_CALCULATED_TIME_COLUMNS, MAX_DEPTH, MAX_EXTENSION_BYTES,
    MAX_NODES, MAX_RELATIONSHIPS, MAX_REWRITE_BYTES, MAX_STRING_BYTES, MAX_TABLES,
    MAX_TIME_GROUPINGS, MAX_TOTAL_STRING_BYTES, MAX_XML_BYTES, MODEL_TIME_GROUPINGS_EXTENSION_URI,
    MODEL_TIME_GROUPINGS_NAMESPACE, SML, STRICT_SML, X15, invalid, limit, xml_error,
};

const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

#[derive(Clone)]
pub(crate) struct Attribute {
    pub namespace: String,
    pub name: String,
    pub value: String,
}

#[derive(Clone)]
pub(crate) struct Node {
    pub namespace: String,
    pub name: String,
    pub attributes: Vec<Attribute>,
    pub children: Vec<Node>,
    pub text: String,
    opaque: Option<Box<OpaqueCapture>>,
}

#[derive(Clone)]
struct OpaqueCapture {
    start: usize,
    insertion: usize,
    declarations: Vec<u8>,
    context_attributes: Vec<u8>,
    replacements: Vec<SourceReplacement>,
    xml: Vec<u8>,
}

#[derive(Clone)]
struct SourceReplacement {
    range: Range<usize>,
    value: Vec<u8>,
}

#[derive(Clone, Default)]
struct XmlContext {
    base: Option<Arc<str>>,
    lang: Option<Arc<str>>,
    space: Option<Arc<str>>,
}

/// Parse an inline `x15:dataModel` descriptor.
pub fn parse_data_model(xml: &[u8]) -> Result<Definition> {
    let root = parse_document(xml)?;
    parse_data_model_node(&root)
}

/// Parse a standalone MS-XLSX `modelTimeGroupings` element.
pub fn parse_model_time_groupings(xml: &[u8]) -> Result<ModelTimeGroupings> {
    let root = parse_document(xml)?;
    parse_model_time_groupings_node(&root)
}

/// Serialize a standalone MS-XLSX `modelTimeGroupings` element.
pub fn write_model_time_groupings(value: &ModelTimeGroupings) -> Result<Vec<u8>> {
    validate_model_time_groupings(value)?;
    let encoded_size = serialized_model_time_groupings_size(value)?;
    let mut output = Vec::new();
    reserve_bytes(&mut output, encoded_size, "modelTimeGroupings bytes")?;
    write_model_time_groupings_node(&mut output, value);
    if output.len() != encoded_size {
        return Err(invalid(
            "serialized modelTimeGroupings size preflight mismatch",
        ));
    }
    Ok(output)
}

/// Project the known `modelTimeGroupings` child from an opaque Data Model
/// extension list. The outer list and unknown siblings remain untouched.
pub(crate) fn parse_model_time_groupings_extension(
    xml: &[u8],
) -> Result<Option<ModelTimeGroupings>> {
    let spans = locate_model_time_groupings(xml)?;
    let Some(child) = spans.child else {
        return Ok(None);
    };
    parse_model_time_groupings_fragment(&xml[child.start..child.end], &child.qname, &child.bindings)
        .map(Some)
}

/// Replace only the known `modelTimeGroupings` child in an opaque extension
/// list. All unrelated bytes, comments, namespaces, and extension siblings
/// are retained. `None` means that no extension list should be created.
pub(crate) fn rewrite_model_time_groupings_extension(
    source: Option<&[u8]>,
    value: Option<&ModelTimeGroupings>,
) -> Result<Option<Vec<u8>>> {
    let value_size = value
        .map(|value| {
            validate_model_time_groupings(value)?;
            serialized_model_time_groupings_size(value)
        })
        .transpose()?;
    let Some(source) = source else {
        let Some(value_size) = value_size else {
            return Ok(None);
        };
        let output_size = new_model_time_groupings_document_size(value_size)?;
        let node = write_model_time_groupings(value.ok_or_else(|| {
            invalid("modelTimeGroupings value disappeared during serialization")
        })?)?;
        let mut output = Vec::new();
        reserve_bytes(&mut output, output_size, "modelTimeGroupings bytes")?;
        output.extend_from_slice(b"<x15:extLst xmlns:x15=\"");
        escape(&mut output, X15);
        output.extend_from_slice(b"\" xmlns=\"\"><x15:ext uri=\"");
        escape(&mut output, MODEL_TIME_GROUPINGS_EXTENSION_URI);
        output.extend_from_slice(b"\">");
        output.extend_from_slice(&node);
        output.extend_from_slice(b"</x15:ext></x15:extLst>");
        if output.len() != output_size {
            return Err(invalid(
                "serialized modelTimeGroupings extension size preflight mismatch",
            ));
        }
        return Ok(Some(output));
    };
    let spans = locate_model_time_groupings(source)?;
    let Some(value_size) = value_size else {
        let Some(child) = spans.child else {
            return copy_bounded(source, "modelTimeGroupings bytes").map(Some);
        };
        if let Some(target) = &spans.target_extension
            && !target.empty
            && target.open_end <= child.start
            && child.end <= target.close_start.unwrap_or(target.end)
        {
            let close_start = target
                .close_start
                .ok_or_else(|| invalid("missing modelTimeGroupings extension close tag"))?;
            if is_xml_whitespace(&source[target.open_end..child.start])
                && is_xml_whitespace(&source[child.end..close_start])
            {
                return splice(source, target.start..target.end, &[]).map(Some);
            }
        }
        return splice(source, child.start..child.end, &[]).map(Some);
    };
    if let Some(child) = spans.child {
        let _ = splice_length(source, child.start..child.end, value_size)?;
        let node = write_model_time_groupings(value.ok_or_else(|| {
            invalid("modelTimeGroupings value disappeared during serialization")
        })?)?;
        return splice(source, child.start..child.end, &node).map(Some);
    }
    let list = spans
        .root
        .ok_or_else(|| invalid("missing Data Model extension list"))?;
    let extension = if let Some(target) = spans.target_extension {
        if target.empty {
            preflight_expand_empty(source, &target, value_size)?;
            let node = write_model_time_groupings(value.ok_or_else(|| {
                invalid("modelTimeGroupings value disappeared during serialization")
            })?)?;
            expand_empty(source, &target, &node)
        } else {
            let at = target
                .close_start
                .ok_or_else(|| invalid("missing modelTimeGroupings extension close tag"))?;
            let _ = splice_length(source, at..at, value_size)?;
            let node = write_model_time_groupings(value.ok_or_else(|| {
                invalid("modelTimeGroupings value disappeared during serialization")
            })?)?;
            splice(source, at..at, &node)
        }
    } else {
        let extension_size = model_time_groupings_extension_size(&list.qname, value_size)?;
        if list.empty {
            preflight_expand_empty(source, &list, extension_size)?;
            let node = write_model_time_groupings(value.ok_or_else(|| {
                invalid("modelTimeGroupings value disappeared during serialization")
            })?)?;
            let extension =
                write_model_time_groupings_extension(&list.qname, &node, extension_size)?;
            expand_empty(source, &list, &extension)
        } else {
            let at = list
                .close_start
                .ok_or_else(|| invalid("missing extension list close tag"))?;
            let _ = splice_length(source, at..at, extension_size)?;
            let node = write_model_time_groupings(value.ok_or_else(|| {
                invalid("modelTimeGroupings value disappeared during serialization")
            })?)?;
            let extension =
                write_model_time_groupings_extension(&list.qname, &node, extension_size)?;
            splice(source, at..at, &extension)
        }
    };
    extension.map(Some)
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn parse_model_time_groupings_node(root: &Node) -> Result<ModelTimeGroupings> {
    require(root, MODEL_TIME_GROUPINGS_NAMESPACE, "modelTimeGroupings")?;
    no_attributes(root, &[])?;
    whitespace(root)?;
    if root.children.is_empty() {
        return Err(invalid(
            "modelTimeGroupings must contain at least one modelTimeGrouping",
        ));
    }
    if root.children.len() > MAX_TIME_GROUPINGS {
        return Err(limit("modelTimeGrouping count"));
    }
    let groupings = root
        .children
        .iter()
        .map(parse_model_time_grouping_node)
        .collect::<Result<Vec<_>>>()?;
    let value = ModelTimeGroupings { groupings };
    validate_model_time_groupings_for_read(&value)?;
    Ok(value)
}

fn parse_model_time_groupings_fragment(
    xml: &[u8],
    qname: &[u8],
    bindings: &[(Vec<u8>, Vec<u8>)],
) -> Result<ModelTimeGroupings> {
    let prefix = qname
        .iter()
        .position(|byte| *byte == b':')
        .map(|index| &qname[..index]);
    let wrapper_prefix = unique_wrapper_prefix(bindings);
    const WRAPPER_NAMESPACE: &str = "urn:litchi-xlsx-model-time-groupings-wrapper";
    let mut wrapped = Vec::with_capacity(xml.len().saturating_add(256));
    wrapped.extend_from_slice(b"<");
    wrapped.extend_from_slice(&wrapper_prefix);
    wrapped.extend_from_slice(b":wrapper xmlns:");
    wrapped.extend_from_slice(&wrapper_prefix);
    wrapped.extend_from_slice(b"=\"");
    wrapped.extend_from_slice(WRAPPER_NAMESPACE.as_bytes());
    wrapped.push(b'"');
    for (binding_prefix, namespace) in bindings {
        if binding_prefix.as_slice() == b"xml" || binding_prefix.as_slice() == b"xmlns" {
            continue;
        }
        wrapped.extend_from_slice(b" xmlns");
        if !binding_prefix.is_empty() {
            wrapped.push(b':');
            wrapped.extend_from_slice(binding_prefix);
        }
        wrapped.extend_from_slice(b"=\"");
        wrapped.extend_from_slice(namespace);
        wrapped.push(b'"');
    }
    if prefix.is_some()
        && !bindings
            .iter()
            .any(|(binding_prefix, _)| Some(binding_prefix.as_slice()) == prefix)
    {
        return Err(invalid("modelTimeGroupings namespace binding is absent"));
    }
    wrapped.extend_from_slice(b">");
    wrapped.extend_from_slice(xml);
    wrapped.extend_from_slice(b"</");
    wrapped.extend_from_slice(&wrapper_prefix);
    wrapped.extend_from_slice(b":wrapper>");
    if wrapped.len() > MAX_XML_BYTES {
        return Err(limit("XML bytes"));
    }
    let wrapper = parse_document(&wrapped)?;
    require(&wrapper, WRAPPER_NAMESPACE, "wrapper")?;
    if wrapper.children.len() != 1 {
        return Err(invalid(
            "modelTimeGroupings wrapper must contain one element",
        ));
    }
    parse_model_time_groupings_node(&wrapper.children[0])
}

fn unique_wrapper_prefix(bindings: &[(Vec<u8>, Vec<u8>)]) -> Vec<u8> {
    let base = b"__litchi_wrapper";
    let mut suffix = 0usize;
    loop {
        let mut candidate = base.to_vec();
        if suffix != 0 {
            candidate.extend_from_slice(suffix.to_string().as_bytes());
        }
        if bindings
            .iter()
            .all(|(prefix, _)| prefix.as_slice() != candidate.as_slice())
        {
            return candidate;
        }
        suffix = suffix.saturating_add(1);
    }
}

fn qualified_name_len(owner_qname: &[u8], local_name: &[u8]) -> Result<usize> {
    let prefix_len = owner_qname
        .iter()
        .position(|byte| *byte == b':')
        .unwrap_or(0);
    prefix_len
        .checked_add((prefix_len != 0) as usize)
        .and_then(|length| length.checked_add(local_name.len()))
        .ok_or_else(|| limit("modelTimeGroupings bytes"))
}

fn append_qualified_name(output: &mut Vec<u8>, owner_qname: &[u8], local_name: &[u8]) {
    if let Some(colon) = owner_qname.iter().position(|byte| *byte == b':') {
        output.extend_from_slice(&owner_qname[..colon]);
        output.push(b':');
    }
    output.extend_from_slice(local_name);
}

fn model_time_groupings_extension_size(owner_qname: &[u8], node_size: usize) -> Result<usize> {
    let qname_size = qualified_name_len(owner_qname, b"ext")?;
    let mut size = 0usize;
    add_bytes_size(&mut size, b"<")?;
    add_size(&mut size, qname_size)?;
    add_bytes_size(&mut size, b" uri=\"")?;
    add_escaped_size(&mut size, MODEL_TIME_GROUPINGS_EXTENSION_URI)?;
    add_bytes_size(&mut size, b"\">")?;
    add_size(&mut size, node_size)?;
    add_bytes_size(&mut size, b"</")?;
    add_size(&mut size, qname_size)?;
    add_bytes_size(&mut size, b">")?;
    if size > MAX_EXTENSION_BYTES {
        return Err(limit("modelTimeGroupings bytes"));
    }
    Ok(size)
}

fn new_model_time_groupings_document_size(node_size: usize) -> Result<usize> {
    let mut size = 0usize;
    add_bytes_size(&mut size, b"<x15:extLst xmlns:x15=\"")?;
    add_escaped_size(&mut size, X15)?;
    add_bytes_size(&mut size, b"\" xmlns=\"\"><x15:ext uri=\"")?;
    add_escaped_size(&mut size, MODEL_TIME_GROUPINGS_EXTENSION_URI)?;
    add_bytes_size(&mut size, b"\">")?;
    add_size(&mut size, node_size)?;
    add_bytes_size(&mut size, b"</x15:ext></x15:extLst>")?;
    if size > MAX_EXTENSION_BYTES {
        return Err(limit("modelTimeGroupings bytes"));
    }
    Ok(size)
}

fn write_model_time_groupings_extension(
    owner_qname: &[u8],
    node: &[u8],
    expected_size: usize,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    reserve_bytes(&mut output, expected_size, "modelTimeGroupings bytes")?;
    output.push(b'<');
    append_qualified_name(&mut output, owner_qname, b"ext");
    output.extend_from_slice(b" uri=\"");
    escape(&mut output, MODEL_TIME_GROUPINGS_EXTENSION_URI);
    output.extend_from_slice(b"\">");
    output.extend_from_slice(node);
    output.extend_from_slice(b"</");
    append_qualified_name(&mut output, owner_qname, b"ext");
    output.push(b'>');
    if output.len() != expected_size {
        return Err(invalid(
            "serialized modelTimeGroupings extension size preflight mismatch",
        ));
    }
    Ok(output)
}

fn parse_model_time_grouping_node(node: &Node) -> Result<ModelTimeGrouping> {
    require(node, MODEL_TIME_GROUPINGS_NAMESPACE, "modelTimeGrouping")?;
    no_attributes(
        node,
        &[("", "tableName"), ("", "columnName"), ("", "columnId")],
    )?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid(
            "modelTimeGrouping must contain at least one calculatedTimeColumn",
        ));
    }
    if node.children.len() > MAX_CALCULATED_TIME_COLUMNS {
        return Err(limit("calculatedTimeColumn count"));
    }
    let calculated_time_columns = node
        .children
        .iter()
        .map(parse_calculated_time_column_node)
        .collect::<Result<Vec<_>>>()?;
    Ok(ModelTimeGrouping {
        table_name: required(node, "", "tableName")?.to_owned(),
        column_name: required(node, "", "columnName")?.to_owned(),
        column_id: required(node, "", "columnId")?.to_owned(),
        calculated_time_columns,
    })
}

fn parse_calculated_time_column_node(node: &Node) -> Result<CalculatedTimeColumn> {
    require(node, MODEL_TIME_GROUPINGS_NAMESPACE, "calculatedTimeColumn")?;
    no_attributes(
        node,
        &[
            ("", "columnName"),
            ("", "columnId"),
            ("", "contentType"),
            ("", "isSelected"),
        ],
    )?;
    leaf(node)?;
    let is_selected = parse_xml_boolean(required(node, "", "isSelected")?, "isSelected")?;
    Ok(CalculatedTimeColumn {
        column_name: required(node, "", "columnName")?.to_owned(),
        column_id: required(node, "", "columnId")?.to_owned(),
        content_type: parse_model_time_grouping_content_type(required(node, "", "contentType")?),
        is_selected,
    })
}

fn parse_model_time_grouping_content_type(value: &str) -> ModelTimeGroupingContentType {
    match value {
        "years" => ModelTimeGroupingContentType::Years,
        "quarters" => ModelTimeGroupingContentType::Quarters,
        "monthsindex" => ModelTimeGroupingContentType::MonthsIndex,
        "months" => ModelTimeGroupingContentType::Months,
        "daysindex" => ModelTimeGroupingContentType::DaysIndex,
        "days" => ModelTimeGroupingContentType::Days,
        "hours" => ModelTimeGroupingContentType::Hours,
        "minutes" => ModelTimeGroupingContentType::Minutes,
        "seconds" => ModelTimeGroupingContentType::Seconds,
        value => ModelTimeGroupingContentType::Other(value.to_owned()),
    }
}

fn model_time_grouping_content_type_value(value: &ModelTimeGroupingContentType) -> &str {
    match value {
        ModelTimeGroupingContentType::Years => "years",
        ModelTimeGroupingContentType::Quarters => "quarters",
        ModelTimeGroupingContentType::MonthsIndex => "monthsindex",
        ModelTimeGroupingContentType::Months => "months",
        ModelTimeGroupingContentType::DaysIndex => "daysindex",
        ModelTimeGroupingContentType::Days => "days",
        ModelTimeGroupingContentType::Hours => "hours",
        ModelTimeGroupingContentType::Minutes => "minutes",
        ModelTimeGroupingContentType::Seconds => "seconds",
        ModelTimeGroupingContentType::Other(value) => value,
    }
}

fn write_model_time_groupings_node(output: &mut Vec<u8>, value: &ModelTimeGroupings) {
    output.extend_from_slice(b"<x14:modelTimeGroupings xmlns:x14=\"");
    escape(output, MODEL_TIME_GROUPINGS_NAMESPACE);
    output.extend_from_slice(b"\">");
    for grouping in &value.groupings {
        output.extend_from_slice(b"<x14:modelTimeGrouping");
        attr(output, "tableName", &grouping.table_name);
        attr(output, "columnName", &grouping.column_name);
        attr(output, "columnId", &grouping.column_id);
        output.push(b'>');
        for column in &grouping.calculated_time_columns {
            output.extend_from_slice(b"<x14:calculatedTimeColumn");
            attr(output, "columnName", &column.column_name);
            attr(output, "columnId", &column.column_id);
            attr(
                output,
                "contentType",
                model_time_grouping_content_type_value(&column.content_type),
            );
            attr(
                output,
                "isSelected",
                if column.is_selected { "1" } else { "0" },
            );
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x14:modelTimeGrouping>");
    }
    output.extend_from_slice(b"</x14:modelTimeGroupings>");
}

fn serialized_model_time_groupings_size(value: &ModelTimeGroupings) -> Result<usize> {
    let mut size = 0usize;
    add_bytes_size(&mut size, b"<x14:modelTimeGroupings xmlns:x14=\"")?;
    add_escaped_size(&mut size, MODEL_TIME_GROUPINGS_NAMESPACE)?;
    add_bytes_size(&mut size, b"\">")?;
    for grouping in &value.groupings {
        add_bytes_size(&mut size, b"<x14:modelTimeGrouping")?;
        add_attribute_size(&mut size, "tableName", escaped_size(&grouping.table_name)?)?;
        add_attribute_size(
            &mut size,
            "columnName",
            escaped_size(&grouping.column_name)?,
        )?;
        add_attribute_size(&mut size, "columnId", escaped_size(&grouping.column_id)?)?;
        add_bytes_size(&mut size, b">")?;
        for column in &grouping.calculated_time_columns {
            add_bytes_size(&mut size, b"<x14:calculatedTimeColumn")?;
            add_attribute_size(&mut size, "columnName", escaped_size(&column.column_name)?)?;
            add_attribute_size(&mut size, "columnId", escaped_size(&column.column_id)?)?;
            add_attribute_size(
                &mut size,
                "contentType",
                escaped_size(model_time_grouping_content_type_value(&column.content_type))?,
            )?;
            add_attribute_size(&mut size, "isSelected", 1)?;
            add_bytes_size(&mut size, b"/>")?;
        }
        add_bytes_size(&mut size, b"</x14:modelTimeGrouping>")?;
    }
    add_bytes_size(&mut size, b"</x14:modelTimeGroupings>")?;
    if size > MAX_EXTENSION_BYTES {
        return Err(limit("modelTimeGroupings bytes"));
    }
    Ok(size)
}

fn validate_model_time_groupings(value: &ModelTimeGroupings) -> Result<()> {
    validate_model_time_groupings_with_policy(value, false)
}

fn validate_model_time_groupings_for_read(value: &ModelTimeGroupings) -> Result<()> {
    validate_model_time_groupings_with_policy(value, true)
}

fn validate_model_time_groupings_with_policy(
    value: &ModelTimeGroupings,
    allow_unknown_content_type: bool,
) -> Result<()> {
    if value.groupings.is_empty() {
        return Err(invalid(
            "modelTimeGroupings must contain at least one modelTimeGrouping",
        ));
    }
    if value.groupings.len() > MAX_TIME_GROUPINGS {
        return Err(limit("modelTimeGrouping count"));
    }
    for grouping in &value.groupings {
        for (field, label) in [
            (&grouping.table_name, "tableName"),
            (&grouping.column_name, "columnName"),
            (&grouping.column_id, "columnId"),
        ] {
            bounded_nonempty(field, label)?;
        }
        if grouping.calculated_time_columns.is_empty() {
            return Err(invalid(
                "modelTimeGrouping must contain at least one calculatedTimeColumn",
            ));
        }
        if grouping.calculated_time_columns.len() > MAX_CALCULATED_TIME_COLUMNS {
            return Err(limit("calculatedTimeColumn count"));
        }
        for column in &grouping.calculated_time_columns {
            bounded_nonempty(&column.column_name, "columnName")?;
            bounded_nonempty(&column.column_id, "columnId")?;
            if !allow_unknown_content_type
                && matches!(&column.content_type, ModelTimeGroupingContentType::Other(_))
            {
                return Err(Error::Unsupported {
                    feature: "unknown modelTimeGrouping contentType is read-only until its schema semantics are implemented",
                });
            }
            bounded_nonempty(
                model_time_grouping_content_type_value(&column.content_type),
                "contentType",
            )?;
        }
    }
    Ok(())
}

fn parse_xml_boolean(value: &str, label: &str) -> Result<bool> {
    match value.trim() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{label} must be an XML boolean"))),
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TimeGroupingRole {
    Root,
    TargetExtension,
    Child,
    Other,
}

#[derive(Clone, Debug)]
struct ElementSpan {
    start: usize,
    open_end: usize,
    close_start: Option<usize>,
    end: usize,
    qname: Vec<u8>,
    bindings: Vec<(Vec<u8>, Vec<u8>)>,
    empty: bool,
}

#[derive(Default)]
struct TimeGroupingSpans {
    root: Option<ElementSpan>,
    target_extension: Option<ElementSpan>,
    child: Option<ElementSpan>,
}

struct ElementFrame {
    role: TimeGroupingRole,
    span: ElementSpan,
}

fn locate_model_time_groupings(xml: &[u8]) -> Result<TimeGroupingSpans> {
    if xml.len() > MAX_EXTENSION_BYTES {
        return Err(limit("extension bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut stack: Vec<ElementFrame> = Vec::new();
    let mut found = TimeGroupingSpans::default();
    let mut nodes = 0usize;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("modelTimeGroupings source position"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("modelTimeGroupings source position"))?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML structure"))?;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let role = time_grouping_role(&reader, &stack, &element)?;
                let span = ElementSpan {
                    start,
                    open_end: end,
                    close_start: None,
                    end: 0,
                    qname: element.name().as_ref().to_vec(),
                    bindings: (role == TimeGroupingRole::Child)
                        .then(|| namespace_bindings(&reader))
                        .transpose()?
                        .unwrap_or_default(),
                    empty: false,
                };
                stack.push(ElementFrame { role, span });
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML structure"))?;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let role = time_grouping_role(&reader, &stack, &element)?;
                let span = ElementSpan {
                    start,
                    open_end: end,
                    close_start: None,
                    end,
                    qname: element.name().as_ref().to_vec(),
                    bindings: (role == TimeGroupingRole::Child)
                        .then(|| namespace_bindings(&reader))
                        .transpose()?
                        .unwrap_or_default(),
                    empty: true,
                };
                record_time_grouping_span(&mut found, role, &span)?;
            },
            Event::End(_) => {
                let mut frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unbalanced modelTimeGroupings XML"))?;
                frame.span.close_start = Some(start);
                frame.span.end = end;
                if frame.role != TimeGroupingRole::Other {
                    record_time_grouping_span(&mut found, frame.role, &frame.span)?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated modelTimeGroupings XML"));
    }
    if found.root.is_none() {
        return Err(invalid("model Data Model extension list is absent"));
    }
    Ok(found)
}

fn namespace_bindings(reader: &NsReader<&[u8]>) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut bindings = Vec::new();
    for (prefix, namespace) in reader.resolver().bindings() {
        let prefix = match prefix {
            PrefixDeclaration::Default => &b""[..],
            PrefixDeclaration::Named(prefix) => prefix,
        };
        if prefix == b"xml" || prefix == b"xmlns" {
            continue;
        }
        bindings.push((prefix.to_vec(), namespace.as_ref().to_vec()));
    }
    Ok(bindings)
}

fn time_grouping_role(
    reader: &NsReader<&[u8]>,
    stack: &[ElementFrame],
    element: &BytesStart<'_>,
) -> Result<TimeGroupingRole> {
    let depth = stack.len();
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = element.local_name();
    if depth == 0 && namespace == X15 && name.as_ref() == b"extLst" {
        return Ok(TimeGroupingRole::Root);
    }
    if depth == 1
        && stack
            .last()
            .is_some_and(|frame| frame.role == TimeGroupingRole::Root)
        && namespace == X15
        && name.as_ref() == b"ext"
        && extension_uri(element, reader.decoder())?.as_deref()
            == Some(MODEL_TIME_GROUPINGS_EXTENSION_URI)
    {
        return Ok(TimeGroupingRole::TargetExtension);
    }
    if depth == 2
        && stack
            .last()
            .is_some_and(|frame| frame.role == TimeGroupingRole::TargetExtension)
        && namespace == MODEL_TIME_GROUPINGS_NAMESPACE
        && name.as_ref() == b"modelTimeGroupings"
    {
        return Ok(TimeGroupingRole::Child);
    }
    Ok(TimeGroupingRole::Other)
}

fn record_time_grouping_span(
    found: &mut TimeGroupingSpans,
    role: TimeGroupingRole,
    span: &ElementSpan,
) -> Result<()> {
    let destination = match role {
        TimeGroupingRole::Root => &mut found.root,
        TimeGroupingRole::TargetExtension => &mut found.target_extension,
        TimeGroupingRole::Child => &mut found.child,
        TimeGroupingRole::Other => return Ok(()),
    };
    if destination.replace(span.clone()).is_some() {
        return Err(invalid("duplicate modelTimeGroupings owner"));
    }
    Ok(())
}

fn expand_empty(source: &[u8], span: &ElementSpan, child: &[u8]) -> Result<Vec<u8>> {
    if !span.empty
        || span.open_end < span.start + 2
        || &source[span.open_end - 2..span.open_end] != b"/>"
    {
        return Err(invalid("invalid empty modelTimeGroupings owner"));
    }
    let replacement_size = span
        .open_end
        .checked_sub(span.start + 2)
        .and_then(|size| size.checked_add(1))
        .and_then(|size| size.checked_add(child.len()))
        .and_then(|size| size.checked_add(2))
        .and_then(|size| size.checked_add(span.qname.len()))
        .and_then(|size| size.checked_add(1))
        .ok_or_else(|| limit("modelTimeGroupings bytes"))?;
    preflight_splice(
        source,
        span.start..span.end,
        replacement_size,
        "modelTimeGroupings bytes",
    )?;
    let mut replacement = Vec::new();
    reserve_bytes(
        &mut replacement,
        replacement_size,
        "modelTimeGroupings bytes",
    )?;
    replacement.extend_from_slice(&source[span.start..span.open_end - 2]);
    replacement.push(b'>');
    replacement.extend_from_slice(child);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(&span.qname);
    replacement.extend_from_slice(b">");
    splice(source, span.start..span.end, &replacement)
}

fn splice(source: &[u8], range: Range<usize>, replacement: &[u8]) -> Result<Vec<u8>> {
    let length = splice_length(source, range.clone(), replacement.len())?;
    let mut output = Vec::new();
    reserve_bytes(&mut output, length, "modelTimeGroupings bytes")?;
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[range.end..]);
    Ok(output)
}

fn preflight_expand_empty(source: &[u8], span: &ElementSpan, child_size: usize) -> Result<usize> {
    if !span.empty
        || span.open_end < span.start + 2
        || &source[span.open_end - 2..span.open_end] != b"/>"
    {
        return Err(invalid("invalid empty modelTimeGroupings owner"));
    }
    let replacement_size = span
        .open_end
        .checked_sub(span.start + 2)
        .and_then(|size| size.checked_add(1))
        .and_then(|size| size.checked_add(child_size))
        .and_then(|size| size.checked_add(2))
        .and_then(|size| size.checked_add(span.qname.len()))
        .and_then(|size| size.checked_add(1))
        .ok_or_else(|| limit("modelTimeGroupings bytes"))?;
    preflight_splice(
        source,
        span.start..span.end,
        replacement_size,
        "modelTimeGroupings bytes",
    )
}

fn splice_length(source: &[u8], range: Range<usize>, replacement_len: usize) -> Result<usize> {
    if range.start > range.end || range.end > source.len() {
        return Err(invalid("invalid modelTimeGroupings source range"));
    }
    preflight_splice(source, range, replacement_len, "modelTimeGroupings bytes")
}

fn preflight_splice(
    source: &[u8],
    range: Range<usize>,
    replacement_len: usize,
    resource: &'static str,
) -> Result<usize> {
    if range.start > range.end || range.end > source.len() {
        return Err(invalid("invalid modelTimeGroupings source range"));
    }
    let length = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| limit("modelTimeGroupings bytes"))?;
    if length > MAX_EXTENSION_BYTES {
        return Err(limit(resource));
    }
    Ok(length)
}

fn reserve_bytes(output: &mut Vec<u8>, size: usize, resource: &'static str) -> Result<()> {
    output
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation { resource, source })
}

fn copy_bounded(source: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    if source.len() > MAX_EXTENSION_BYTES {
        return Err(limit(resource));
    }
    let mut output = Vec::new();
    reserve_bytes(&mut output, source.len(), resource)?;
    output.extend_from_slice(source);
    Ok(output)
}

/// Close a detached extension's namespace context without rewriting its markup.
pub(crate) fn close_extension(xml: &[u8]) -> Result<Vec<u8>> {
    let root = parse_document(xml)?;
    require(&root, X15, "extLst")?;
    Ok(root
        .opaque
        .ok_or_else(|| invalid("missing extension source"))?
        .xml)
}

/// Deterministically serialize an inline `x15:dataModel` descriptor.
pub fn write_data_model(value: &Definition) -> Result<Vec<u8>> {
    // Compute the complete encoded length before validating or materializing
    // the output buffer.  The descriptor limit applies to escaped wire bytes,
    // not to the caller's UTF-8 string lengths, so counting after `push` would
    // permit an oversized allocation.
    let encoded_size = serialized_data_model_size(value)?;
    validate_definition(value, false)?;
    let mut output = Vec::with_capacity(encoded_size);
    output.extend_from_slice(b"<x15:dataModel xmlns:x15=\"");
    escape(&mut output, X15);
    output.push(b'"');
    if value.min_version_load != 5 {
        attr(
            &mut output,
            "minVersionLoad",
            &value.min_version_load.to_string(),
        );
    }
    if value.tables.is_empty() && value.relationships.is_empty() && value.extension_list.is_none() {
        output.extend_from_slice(b"/>");
        if output.len() != encoded_size {
            return Err(invalid("serialized descriptor size preflight mismatch"));
        }
        return Ok(output);
    }
    output.push(b'>');
    if !value.tables.is_empty() {
        output.extend_from_slice(b"<x15:modelTables>");
        for table in &value.tables {
            output.extend_from_slice(b"<x15:modelTable");
            attr(&mut output, "id", &table.id);
            attr(&mut output, "name", &table.name);
            attr(&mut output, "connection", &table.connection);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:modelTables>");
    }
    if !value.relationships.is_empty() {
        output.extend_from_slice(b"<x15:modelRelationships>");
        for relationship in &value.relationships {
            output.extend_from_slice(b"<x15:modelRelationship");
            attr(&mut output, "fromTable", &relationship.from_table);
            attr(&mut output, "fromColumn", &relationship.from_column);
            attr(&mut output, "toTable", &relationship.to_table);
            attr(&mut output, "toColumn", &relationship.to_column);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:modelRelationships>");
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    output.extend_from_slice(b"</x15:dataModel>");
    if output.len() != encoded_size {
        return Err(invalid("serialized descriptor size preflight mismatch"));
    }
    Ok(output)
}

fn serialized_data_model_size(value: &Definition) -> Result<usize> {
    let size = serialized_data_model_size_unbounded(value)?;
    if size > MAX_XML_BYTES {
        return Err(limit("serialized descriptor bytes"));
    }
    Ok(size)
}

fn serialized_data_model_size_unbounded(value: &Definition) -> Result<usize> {
    let mut size = 0usize;
    add_bytes_size(&mut size, b"<x15:dataModel xmlns:x15=\"")?;
    add_escaped_size(&mut size, X15)?;
    add_bytes_size(&mut size, b"\"")?;
    if value.min_version_load != 5 {
        add_attribute_size(
            &mut size,
            "minVersionLoad",
            decimal_len(value.min_version_load),
        )?;
    }
    if value.tables.is_empty() && value.relationships.is_empty() && value.extension_list.is_none() {
        add_bytes_size(&mut size, b"/>")?;
    } else {
        add_bytes_size(&mut size, b">")?;
        if !value.tables.is_empty() {
            add_bytes_size(&mut size, b"<x15:modelTables>")?;
            for table in &value.tables {
                add_bytes_size(&mut size, b"<x15:modelTable")?;
                add_attribute_size(&mut size, "id", escaped_size(&table.id)?)?;
                add_attribute_size(&mut size, "name", escaped_size(&table.name)?)?;
                add_attribute_size(&mut size, "connection", escaped_size(&table.connection)?)?;
                add_bytes_size(&mut size, b"/>")?;
            }
            add_bytes_size(&mut size, b"</x15:modelTables>")?;
        }
        if !value.relationships.is_empty() {
            add_bytes_size(&mut size, b"<x15:modelRelationships>")?;
            for relationship in &value.relationships {
                add_bytes_size(&mut size, b"<x15:modelRelationship")?;
                add_attribute_size(
                    &mut size,
                    "fromTable",
                    escaped_size(&relationship.from_table)?,
                )?;
                add_attribute_size(
                    &mut size,
                    "fromColumn",
                    escaped_size(&relationship.from_column)?,
                )?;
                add_attribute_size(&mut size, "toTable", escaped_size(&relationship.to_table)?)?;
                add_attribute_size(
                    &mut size,
                    "toColumn",
                    escaped_size(&relationship.to_column)?,
                )?;
                add_bytes_size(&mut size, b"/>")?;
            }
            add_bytes_size(&mut size, b"</x15:modelRelationships>")?;
        }
        if let Some(extension) = &value.extension_list {
            add_size(&mut size, extension.xml.len())?;
        }
        add_bytes_size(&mut size, b"</x15:dataModel>")?;
    }
    Ok(size)
}

fn add_size(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("serialized descriptor bytes"))?;
    Ok(())
}

fn add_bytes_size(total: &mut usize, value: &[u8]) -> Result<()> {
    add_size(total, value.len())
}

fn add_attribute_size(total: &mut usize, name: &str, value_size: usize) -> Result<()> {
    add_size(total, name.len())?;
    add_size(total, value_size)?;
    add_size(total, 4)
}

fn escaped_size(value: &str) -> Result<usize> {
    let mut size = 0usize;
    add_escaped_size(&mut size, value)?;
    Ok(size)
}

fn add_escaped_size(total: &mut usize, value: &str) -> Result<()> {
    for character in value.chars() {
        let size = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        add_size(total, size)?;
    }
    Ok(())
}

fn decimal_len(value: u8) -> usize {
    match value {
        0..=9 => 1,
        10..=99 => 2,
        _ => 3,
    }
}

pub(crate) fn parse_data_model_node(root: &Node) -> Result<Definition> {
    require(root, X15, "dataModel")?;
    no_attributes(root, &[("", "minVersionLoad")])?;
    whitespace(root)?;
    let min_version_load = optional(root, "", "minVersionLoad")
        .map(|value| {
            value
                .trim_matches([' ', '\t', '\n', '\r'])
                .parse::<u8>()
                .map_err(|_source| invalid("minVersionLoad must be an unsigned byte"))
        })
        .transpose()?
        .unwrap_or(5);
    let mut tables = Vec::new();
    let mut relationships = Vec::new();
    let mut extension_list = None;
    let mut stage = 0u8;
    for child in &root.children {
        match child.name.as_str() {
            "modelTables" if child.namespace == X15 && stage == 0 => {
                stage = 1;
                tables = parse_tables(child)?;
            },
            "modelRelationships" if child.namespace == X15 && stage <= 1 => {
                stage = 2;
                relationships = parse_relationships(child)?;
            },
            "extLst" if child.namespace == X15 && stage <= 2 => {
                stage = 3;
                let xml = child
                    .opaque
                    .as_ref()
                    .ok_or_else(|| invalid("missing opaque extension source"))?
                    .xml
                    .clone();
                if xml.len() > MAX_EXTENSION_BYTES {
                    return Err(limit("extension bytes"));
                }
                extension_list = Some(OpaqueXml { xml });
            },
            _ => return Err(invalid("unexpected or out-of-order dataModel child")),
        }
    }
    let value = Definition {
        min_version_load,
        tables,
        relationships,
        extension_list,
    };
    validate_definition(&value, true)?;
    Ok(value)
}

fn parse_tables(node: &Node) -> Result<Vec<Table>> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid("modelTables must contain at least one modelTable"));
    }
    if node.children.len() > MAX_TABLES {
        return Err(limit("table count"));
    }
    node.children
        .iter()
        .map(|child| {
            require(child, X15, "modelTable")?;
            no_attributes(child, &[("", "id"), ("", "name"), ("", "connection")])?;
            leaf(child)?;
            Ok(Table {
                id: required(child, "", "id")?.to_owned(),
                name: required(child, "", "name")?.to_owned(),
                connection: required(child, "", "connection")?.to_owned(),
            })
        })
        .collect()
}

fn parse_relationships(node: &Node) -> Result<Vec<Relationship>> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid(
            "modelRelationships must contain at least one modelRelationship",
        ));
    }
    if node.children.len() > MAX_RELATIONSHIPS {
        return Err(limit("relationship count"));
    }
    node.children
        .iter()
        .map(|child| {
            require(child, X15, "modelRelationship")?;
            no_attributes(
                child,
                &[
                    ("", "fromTable"),
                    ("", "fromColumn"),
                    ("", "toTable"),
                    ("", "toColumn"),
                ],
            )?;
            leaf(child)?;
            Ok(Relationship {
                from_table: required(child, "", "fromTable")?.to_owned(),
                from_column: required(child, "", "fromColumn")?.to_owned(),
                to_table: required(child, "", "toTable")?.to_owned(),
                to_column: required(child, "", "toColumn")?.to_owned(),
            })
        })
        .collect()
}

pub(crate) fn validate_definition(
    value: &Definition,
    extension_already_parsed: bool,
) -> Result<()> {
    if value.min_version_load < 5 {
        return Err(invalid("minVersionLoad must be at least 5"));
    }
    if value.tables.len() > MAX_TABLES {
        return Err(limit("table count"));
    }
    if value.relationships.len() > MAX_RELATIONSHIPS {
        return Err(limit("relationship count"));
    }
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for table in &value.tables {
        for (field, label) in [
            (&table.id, "table id"),
            (&table.name, "table name"),
            (&table.connection, "connection name"),
        ] {
            bounded_nonempty(field, label)?;
        }
        if !ids.insert(table.id.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive Data Model table id '{}'",
                table.id
            )));
        }
        if !names.insert(table.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive Data Model table name '{}'",
                table.name
            )));
        }
    }
    let mut relationships = HashSet::new();
    for relationship in &value.relationships {
        for (field, label) in [
            (&relationship.from_table, "fromTable"),
            (&relationship.from_column, "fromColumn"),
            (&relationship.to_table, "toTable"),
            (&relationship.to_column, "toColumn"),
        ] {
            bounded_nonempty(field, label)?;
        }
        if !names.contains(&relationship.from_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown fromTable '{}'",
                relationship.from_table
            )));
        }
        if !names.contains(&relationship.to_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown toTable '{}'",
                relationship.to_table
            )));
        }
        let key = (
            relationship.from_table.to_lowercase(),
            relationship.from_column.to_lowercase(),
            relationship.to_table.to_lowercase(),
            relationship.to_column.to_lowercase(),
        );
        if !relationships.insert(key) {
            return Err(invalid(
                "duplicate case-insensitive Data Model relationship",
            ));
        }
    }
    if let Some(extension) = &value.extension_list {
        if extension.xml.len() > MAX_EXTENSION_BYTES {
            return Err(limit("extension bytes"));
        }
        if !extension_already_parsed {
            let root = parse_document(&extension.xml)?;
            require(&root, X15, "extLst")?;
        }
        // Validate the one typed extension owner while retaining every
        // unrecognized sibling as opaque source bytes.
        parse_model_time_groupings_extension(&extension.xml)?;
    }
    Ok(())
}

pub(crate) fn workbook_definition(root: &Node) -> Result<(&str, Option<Definition>)> {
    if root.name != "workbook" || !(root.namespace == SML || root.namespace == STRICT_SML) {
        return Err(invalid("expected SpreadsheetML workbook root"));
    }
    let core = root.namespace.as_str();
    let lists: Vec<_> = root
        .children
        .iter()
        .filter(|child| child.namespace == core && child.name == "extLst")
        .collect();
    if lists.len() > 1 {
        return Err(invalid("workbook has multiple direct extLst elements"));
    }
    let mut found = None;
    if let Some(list) = lists.first() {
        for extension in &list.children {
            if extension.namespace == core
                && extension.name == "ext"
                && optional(extension, "", "uri") == Some(DATA_MODEL_EXTENSION_URI)
            {
                if found.is_some() {
                    return Err(invalid("workbook has multiple dataModel extensions"));
                }
                no_attributes(extension, &[("", "uri")])?;
                whitespace(extension)?;
                if extension.children.len() != 1 {
                    return Err(invalid(
                        "dataModel extension must contain exactly one dataModel element",
                    ));
                }
                found = Some(parse_data_model_node(&extension.children[0])?);
            }
        }
    }
    Ok((core, found))
}

/// Update the singleton descriptor owner using checked source element edits.
/// Existing list/root spelling and unrelated extension bytes are retained.
pub(crate) fn rewrite_data_model_extension(
    source: &litchi_opc::OwnedXmlPart,
    core: &str,
    fragment: Option<&[u8]>,
) -> Result<litchi_opc::OwnedXmlPart> {
    let mut reader = NsReader::from_reader(source.bytes());
    let origin = ReaderOrigin::of(source.bytes());
    let mut depth = 0usize;
    let mut root_tag = None;
    let mut list_tag = None;
    let mut model_tag = None;
    let mut in_list = false;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("rewrite position"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("rewrite position"))?;
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                if depth == 0 {
                    root_tag = Some(start..end);
                }
                if depth == 1 && namespace == core && element.local_name().as_ref() == b"extLst" {
                    if list_tag.replace(start..end).is_some() {
                        return Err(invalid("multiple workbook extension lists"));
                    }
                    in_list = is_start;
                }
                if depth == 2
                    && in_list
                    && namespace == core
                    && element.local_name().as_ref() == b"ext"
                    && extension_uri(&element, reader.decoder())?.as_deref()
                        == Some(DATA_MODEL_EXTENSION_URI)
                    && model_tag.replace(start..end).is_some()
                {
                    return Err(invalid("multiple Data Model extensions"));
                }
                if is_start {
                    depth += 1;
                }
            },
            Event::End(_) => {
                if depth == 2 {
                    in_list = false;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced workbook XML"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let updated = match (model_tag, fragment) {
        (Some(tag), Some(fragment)) => source.replace_element(tag, fragment)?,
        (Some(tag), None) => source.remove_element(tag)?,
        (None, None) => return Ok(source.clone()),
        (None, Some(fragment)) => {
            if let Some(list) = list_tag {
                source.append_element(list, fragment)?
            } else {
                let size = fragment
                    .len()
                    .checked_add(core.len())
                    .and_then(|size| size.checked_add(64))
                    .ok_or_else(|| limit("rewrite bytes"))?;
                if size > MAX_REWRITE_BYTES {
                    return Err(limit("rewrite bytes"));
                }
                let mut wrapper = Vec::with_capacity(size);
                wrapper.extend_from_slice(b"<extLst xmlns=\"");
                escape(&mut wrapper, core);
                wrapper.extend_from_slice(b"\">");
                wrapper.extend_from_slice(fragment);
                wrapper.extend_from_slice(b"</extLst>");
                source.append_element(
                    root_tag.ok_or_else(|| invalid("missing workbook root"))?,
                    &wrapper,
                )?
            }
        },
    };
    if updated.bytes().len() > MAX_REWRITE_BYTES {
        return Err(limit("rewrite bytes"));
    }
    Ok(updated)
}

/// Change only the load-version attribute in an already validated workbook.
/// Other descriptor markup, including comments and extension lexical forms,
/// remains source material.
pub(crate) fn rewrite_load_version(
    source: &litchi_opc::OwnedXmlPart,
    core: &str,
    version: u8,
) -> Result<litchi_opc::OwnedXmlPart> {
    let xml = source.bytes();
    let (owner_start, owner_end) = find_data_model_extension(xml, core)?
        .ok_or_else(|| invalid("workbook has no Data Model extension"))?;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("rewrite position"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("rewrite position"))?;
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if depth == 3
                    && start >= owner_start
                    && end <= owner_end
                    && element.local_name().as_ref() == b"dataModel"
                    && resolved(reader.resolver().resolve_element(element.name()).0)? == X15
                {
                    let lexical = version.to_string();
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(xml_error)?;
                        if attribute.key.as_ref() == b"minVersionLoad" {
                            let range = crate::source_attributes::value_span(
                                xml,
                                attribute.value.as_ref(),
                            )?;
                            return Ok(source.replace_attributes(&[(range, lexical.into_bytes())])?);
                        }
                    }
                    let result = source.insert_unqualified_attribute(
                        start..end,
                        "minVersionLoad",
                        lexical.as_bytes(),
                    )?;
                    if result.bytes().len() > MAX_REWRITE_BYTES {
                        return Err(limit("rewrite bytes"));
                    }
                    return Ok(result);
                }
                if is_start {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced workbook XML"))?;
            },
            Event::Eof => return Err(invalid("Data Model element is absent from its extension")),
            _ => {},
        }
    }
}

fn find_data_model_extension(xml: &[u8], core: &str) -> Result<Option<(usize, usize)>> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut open = None;
    let mut found = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("rewrite position"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| limit("rewrite position"))?;
        match event {
            Event::Start(element) => {
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                if depth == 2
                    && namespace == core
                    && element.local_name().as_ref() == b"ext"
                    && extension_uri(&element, reader.decoder())?.as_deref()
                        == Some(DATA_MODEL_EXTENSION_URI)
                {
                    if open.is_some() || found.is_some() {
                        return Err(invalid("workbook has multiple dataModel extensions"));
                    }
                    open = Some(start);
                }
                depth = depth.checked_add(1).ok_or_else(|| limit("rewrite depth"))?;
            },
            Event::Empty(element) => {
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                if depth == 2
                    && namespace == core
                    && element.local_name().as_ref() == b"ext"
                    && extension_uri(&element, reader.decoder())?.as_deref()
                        == Some(DATA_MODEL_EXTENSION_URI)
                {
                    if open.is_some() || found.is_some() {
                        return Err(invalid("workbook has multiple dataModel extensions"));
                    }
                    found = Some((start, end));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook closing element"));
                }
                if let Some(open_start) = open {
                    if depth == 3 && element.local_name().as_ref() == b"ext" {
                        found = Some((open_start, end));
                        open = None;
                    }
                }
                depth -= 1;
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
    if open.is_some() || depth != 0 {
        return Err(invalid("unterminated workbook XML"));
    }
    Ok(found)
}

fn extension_uri(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.local_name().as_ref() == b"uri" {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn namespace_matches(value: ResolveResult<'_>, expected: &str) -> Result<bool> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Ok(value == expected.as_bytes()),
        ResolveResult::Unbound => Ok(expected.is_empty()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn validate_namespace(value: ResolveResult<'_>) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = value {
        return Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )));
    }
    Ok(())
}

fn is_element_name(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    name: &str,
) -> Result<bool> {
    Ok(namespace_matches(
        reader.resolver().resolve_element(element.name()).0,
        namespace,
    )? && element.local_name().as_ref() == name.as_bytes())
}

fn validate_opaque_element(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    validate_namespace(reader.resolver().resolve_element(element.name()).0)?;
    let attribute_count = element.attributes().with_checks(true).count();
    let mut expanded = HashSet::new();
    expanded
        .try_reserve(attribute_count)
        .map_err(|_| limit("opaque attribute allocation"))?;
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        // Opaque markup is retained byte-for-byte, so do not decode or
        // normalize the value. We still need to enforce XML's entity rules;
        // the bounded scanner validates one reference at a time without
        // materializing an attribute-sized replacement string.
        validate_opaque_attribute_value(item.value.as_ref())?;
        if item.key.as_ref() == b"xmlns" || item.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = match namespace {
            ResolveResult::Bound(namespace) => Some(namespace),
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "unbound XML prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        if !expanded.insert((namespace, local)) {
            return Err(invalid("duplicate expanded XML attribute"));
        }
    }
    Ok(())
}

fn validate_opaque_attribute_value(value: &[u8]) -> Result<()> {
    let mut cursor = 0usize;
    while let Some(offset) = value[cursor..].iter().position(|byte| *byte == b'&') {
        let start = cursor + offset;
        let end = value[start + 1..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|offset| start + 1 + offset)
            .ok_or_else(|| xml_error("unterminated XML entity in attribute value"))?;
        validate_opaque_entity(&value[start + 1..end])?;
        cursor = end + 1;
    }
    Ok(())
}

fn validate_opaque_entity(entity: &[u8]) -> Result<()> {
    let entity = std::str::from_utf8(entity).map_err(xml_error)?;
    if matches!(entity, "amp" | "lt" | "gt" | "apos" | "quot") {
        return Ok(());
    }
    if entity.starts_with('#') {
        // BytesRef delegates numeric checking to the same XML codepoint
        // parser used for GeneralRef events, without allocating the value.
        BytesRef::new(entity)
            .resolve_char_ref()
            .map_err(xml_error)?;
        return Ok(());
    }
    Err(invalid("custom XML entity is rejected"))
}

pub(crate) fn parse_document(xml: &[u8]) -> Result<Node> {
    parse_document_with_base(xml, None)
}

/// Parse an XML owner with its OPC part URI as the initial XML Base context.
/// Standalone fragments retain the historical context-free entry point above;
/// package-owned workbook parts can provide their concrete part URI so a
/// relative root `xml:base` remains portable when detached.
pub(crate) fn parse_document_with_base(xml: &[u8], initial_base: Option<&str>) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("XML bytes"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut contexts: Vec<XmlContext> = Vec::new();
    let initial_context = XmlContext {
        base: initial_base.map(Arc::from),
        ..XmlContext::default()
    };
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    let mut opaque_depth = None;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Data Model XML offset exceeds usize"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Data Model XML offset exceeds usize"))?;
        if opaque_depth.is_some() {
            let capture = stack
                .last()
                .and_then(|node| node.opaque.as_deref())
                .ok_or_else(|| invalid("missing opaque extension capture"))?;
            opaque_size(capture, end)?;
        }
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                if opaque_depth.is_some() {
                    validate_opaque_element(&reader, element)?;
                    if !empty {
                        let depth = opaque_depth
                            .ok_or_else(|| invalid("missing opaque extension depth"))?;
                        let outer_depth = stack
                            .len()
                            .checked_sub(1)
                            .ok_or_else(|| invalid("missing opaque extension root"))?;
                        if outer_depth
                            .checked_add(depth)
                            .is_none_or(|depth| depth >= MAX_DEPTH)
                        {
                            return Err(limit("XML structure"));
                        }
                        *opaque_depth
                            .as_mut()
                            .ok_or_else(|| invalid("missing opaque extension depth"))? += 1;
                    }
                    continue;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let is_opaque_root = is_element_name(&reader, element, X15, "extLst")?
                    && (stack.is_empty()
                        || stack.last().is_some_and(|parent| {
                            parent.namespace == X15 && parent.name == "dataModel"
                        }));
                if is_opaque_root
                    && end
                        .checked_sub(start)
                        .is_none_or(|size| size > MAX_EXTENSION_BYTES)
                {
                    return Err(limit("extension bytes"));
                }
                let mut node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                let parent_context = contexts
                    .last()
                    .cloned()
                    .unwrap_or_else(|| initial_context.clone());
                let context = element_context(
                    &parent_context,
                    &node,
                    if is_opaque_root {
                        MAX_EXTENSION_BYTES
                    } else {
                        MAX_TOTAL_STRING_BYTES
                    },
                )?;
                if is_opaque_root {
                    let capture = opaque_capture(
                        &reader,
                        element,
                        xml,
                        start,
                        end,
                        end - if empty { 2 } else { 1 },
                        &parent_context,
                        &context,
                        &node,
                    )?;
                    opaque_size(&capture, end)?;
                    node.opaque = Some(Box::new(capture));
                }
                if empty {
                    finish_opaque(&mut node, xml, end, &mut strings)?;
                    attach(node, &mut stack, &mut root)?;
                } else {
                    if is_opaque_root {
                        opaque_depth = Some(1);
                    }
                    stack.push(node);
                    contexts.push(context);
                }
            },
            Event::End(_) if opaque_depth.is_some() => {
                if opaque_depth == Some(1) {
                    opaque_depth = None;
                    let mut node = stack
                        .pop()
                        .ok_or_else(|| invalid("missing opaque extension root"))?;
                    contexts
                        .pop()
                        .ok_or_else(|| invalid("missing XML context"))?;
                    finish_opaque(&mut node, xml, end, &mut strings)?;
                    attach(node, &mut stack, &mut root)?;
                } else {
                    *opaque_depth
                        .as_mut()
                        .ok_or_else(|| invalid("missing opaque extension depth"))? -= 1;
                }
            },
            Event::Text(_) if opaque_depth.is_some() => {},
            Event::CData(_) if opaque_depth.is_some() => {},
            Event::PI(_) if opaque_depth.is_some() => {},
            Event::End(_) => {
                let mut node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                contexts
                    .pop()
                    .ok_or_else(|| invalid("missing XML context"))?;
                finish_opaque(&mut node, xml, end, &mut strings)?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let known = reference.resolve_char_ref().map_err(xml_error)?.is_some()
                    || matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot");
                if opaque_depth.is_some() {
                    if !known {
                        return Err(invalid("custom XML entity is rejected"));
                    }
                    continue;
                }
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::CData(text) if stack.iter().any(|node| node.opaque.is_some()) => {
                add_strings(&mut strings, text.len())?;
            },
            Event::PI(_) if stack.iter().any(|node| node.opaque.is_some()) => {},
            Event::CData(_) => return Err(invalid("CDATA is rejected in modeled Data Model XML")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
        opaque: None,
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

// Retain every in-scope binding: unknown attribute/text values may contain
// QNames or expressions even when the prefix is absent from markup names.
fn opaque_capture(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    source: &[u8],
    start: usize,
    event_end: usize,
    insertion: usize,
    parent_context: &XmlContext,
    context: &XmlContext,
    node: &Node,
) -> Result<OpaqueCapture> {
    // The resolver's effective-bindings iterator scans later bindings for
    // shadowing. Cap its input before iterating to bound that quadratic work.
    let mut binding_count = 0usize;
    for level in 0..=reader.resolver().level() {
        for _ in reader.resolver().bindings_of(level) {
            binding_count += 1;
            if binding_count > 4096 {
                return Err(limit("in-scope namespace count"));
            }
        }
    }
    let mut local: HashSet<&[u8]> = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if let Some(prefix) = attribute.key.as_namespace_binding() {
            local.insert(match prefix {
                PrefixDeclaration::Default => &[],
                PrefixDeclaration::Named(prefix) => prefix,
            });
        }
    }
    let mut has_default = local.contains(&b""[..]);
    let mut declaration_len = 0usize;
    for (prefix, namespace) in reader.resolver().bindings() {
        let prefix = match prefix {
            PrefixDeclaration::Default => {
                has_default = true;
                &b""[..]
            },
            PrefixDeclaration::Named(prefix) => prefix,
        };
        if prefix == b"xml" || prefix == b"xmlns" || local.contains(prefix) {
            continue;
        }
        // NamespaceResolver exposes the original escaped URI spelling.
        let raw = std::str::from_utf8(namespace.as_ref()).map_err(xml_error)?;
        let escaped_len = namespace_escaped_len(raw)?;
        let charge = prefix
            .len()
            .checked_add(escaped_len)
            .and_then(|size| size.checked_add(10))
            .ok_or_else(|| limit("namespace bytes"))?;
        if declaration_len
            .checked_add(charge)
            .is_none_or(|size| size > MAX_EXTENSION_BYTES)
        {
            return Err(limit("namespace bytes"));
        }
        declaration_len = declaration_len
            .checked_add(charge)
            .ok_or_else(|| limit("namespace bytes"))?;
    }
    if !has_default {
        declaration_len = declaration_len
            .checked_add(b" xmlns=\"\"".len())
            .ok_or_else(|| limit("namespace bytes"))?;
    }
    if declaration_len > MAX_EXTENSION_BYTES {
        return Err(limit("namespace bytes"));
    }
    let local_base = xml_attribute(node, XML_NAMESPACE, "base");
    let local_lang = xml_attribute(node, XML_NAMESPACE, "lang");
    let local_space = xml_attribute(node, XML_NAMESPACE, "space");

    // The extension is detached from its original ancestors when it is
    // imported into another workbook. Materialize inherited XML context at
    // the detached root so QName-like values, language-sensitive content, and
    // relative links keep their source meaning.
    let mut context_len = 0usize;
    if local_base.is_none() {
        if let Some(base) = &parent_context.base {
            context_len = context_len
                .checked_add(context_attribute_len("base", base)?)
                .ok_or_else(|| limit("XML context bytes"))?;
        }
    }
    if local_lang.is_none()
        && let Some(lang) = &parent_context.lang
    {
        context_len = context_len
            .checked_add(context_attribute_len("lang", lang)?)
            .ok_or_else(|| limit("XML context bytes"))?;
    }
    if local_space.is_none()
        && let Some(space) = &parent_context.space
    {
        context_len = context_len
            .checked_add(context_attribute_len("space", space)?)
            .ok_or_else(|| limit("XML context bytes"))?;
    }
    if context_len > MAX_EXTENSION_BYTES {
        return Err(limit("XML context bytes"));
    }

    let replacement = if let (Some(parent_base), Some(local)) = (&parent_context.base, local_base)
        && !is_absolute_uri(local)
        && !parent_base.is_empty()
    {
        let effective = context
            .base
            .as_deref()
            .ok_or_else(|| invalid("cannot preserve inherited xml:base"))?;
        if effective != local {
            let range = xml_attribute_range(reader, element, source, XML_NAMESPACE, "base")?
                .ok_or_else(|| invalid("xml:base source range is absent"))?;
            Some((range, effective))
        } else {
            None
        }
    } else {
        None
    };

    let mut size = event_end
        .checked_sub(start)
        .and_then(|size| size.checked_add(declaration_len))
        .and_then(|size| size.checked_add(context_len))
        .ok_or_else(|| limit("extension bytes"))?;
    if let Some((range, effective)) = &replacement {
        let escaped_len = escaped_attribute_len(effective)?;
        size = size
            .checked_sub(range.len())
            .and_then(|size| size.checked_add(escaped_len))
            .ok_or_else(|| limit("extension bytes"))?;
    }
    if size > MAX_EXTENSION_BYTES {
        return Err(limit("extension bytes"));
    }

    let mut declarations = Vec::new();
    declarations
        .try_reserve_exact(declaration_len)
        .map_err(|_| limit("namespace allocation"))?;
    for (prefix, namespace) in reader.resolver().bindings() {
        let prefix = match prefix {
            PrefixDeclaration::Default => &b""[..],
            PrefixDeclaration::Named(prefix) => prefix,
        };
        if prefix == b"xml" || prefix == b"xmlns" || local.contains(prefix) {
            continue;
        }
        let raw = std::str::from_utf8(namespace.as_ref()).map_err(xml_error)?;
        declarations.extend_from_slice(b" xmlns");
        if !prefix.is_empty() {
            declarations.push(b':');
            declarations.extend_from_slice(prefix);
        }
        declarations.extend_from_slice(b"=\"");
        append_escaped_namespace(&mut declarations, raw)?;
        declarations.push(b'"');
    }
    // An unqualified child must stay unqualified when moved into a workbook
    // whose default namespace is SpreadsheetML.
    if !has_default {
        declarations.extend_from_slice(b" xmlns=\"\"");
    }

    let mut context_attributes = Vec::new();
    context_attributes
        .try_reserve_exact(context_len)
        .map_err(|_| limit("XML context allocation"))?;
    if local_base.is_none() {
        if let Some(base) = &parent_context.base {
            append_context_attribute(&mut context_attributes, "base", base)?;
        }
    }
    if local_lang.is_none()
        && let Some(lang) = &parent_context.lang
    {
        append_context_attribute(&mut context_attributes, "lang", lang)?;
    }
    if local_space.is_none()
        && let Some(space) = &parent_context.space
    {
        append_context_attribute(&mut context_attributes, "space", space)?;
    }

    let mut replacements = Vec::new();
    if let Some((range, effective)) = replacement {
        replacements
            .try_reserve_exact(1)
            .map_err(|_| limit("XML context allocation"))?;
        replacements.push(SourceReplacement {
            range,
            value: escaped_attribute_value(effective)?,
        });
    }

    // The context attributes are assembled separately from namespace
    // declarations so callers can charge and splice each bounded source
    // replacement without reparsing or reserializing opaque children.
    Ok(OpaqueCapture {
        start,
        insertion,
        declarations,
        context_attributes,
        replacements,
        xml: Vec::new(),
    })
}

fn finish_opaque(node: &mut Node, source: &[u8], end: usize, strings: &mut usize) -> Result<()> {
    let Some(capture) = node.opaque.as_mut() else {
        return Ok(());
    };
    let size = opaque_size(capture, end)?;
    add_strings(strings, size)?;
    capture
        .xml
        .try_reserve_exact(size)
        .map_err(|_| limit("extension allocation"))?;
    append_replaced_source(
        &mut capture.xml,
        source,
        capture.start..capture.insertion,
        &capture.replacements,
    )?;
    capture.xml.extend_from_slice(&capture.declarations);
    capture.xml.extend_from_slice(&capture.context_attributes);
    capture
        .xml
        .extend_from_slice(&source[capture.insertion..end]);
    capture.declarations = Vec::new();
    capture.context_attributes = Vec::new();
    capture.replacements = Vec::new();
    Ok(())
}

fn opaque_size(capture: &OpaqueCapture, end: usize) -> Result<usize> {
    let mut size = end
        .checked_sub(capture.start)
        .and_then(|size| size.checked_add(capture.declarations.len()))
        .and_then(|size| size.checked_add(capture.context_attributes.len()))
        .ok_or_else(|| limit("extension bytes"))?;
    for replacement in &capture.replacements {
        size = size
            .checked_sub(replacement.range.len())
            .and_then(|size| size.checked_add(replacement.value.len()))
            .ok_or_else(|| limit("extension bytes"))?;
    }
    if size > MAX_EXTENSION_BYTES {
        return Err(limit("extension bytes"));
    }
    Ok(size)
}

fn append_replaced_source(
    output: &mut Vec<u8>,
    source: &[u8],
    span: Range<usize>,
    replacements: &[SourceReplacement],
) -> Result<()> {
    let mut cursor = span.start;
    for replacement in replacements {
        if replacement.range.start < cursor
            || replacement.range.end > span.end
            || replacement.range.start > replacement.range.end
        {
            return Err(invalid("overlapping opaque XML source replacement"));
        }
        output.extend_from_slice(&source[cursor..replacement.range.start]);
        output.extend_from_slice(&replacement.value);
        cursor = replacement.range.end;
    }
    output.extend_from_slice(&source[cursor..span.end]);
    Ok(())
}

fn xml_attribute<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn element_context(parent: &XmlContext, node: &Node, max_length: usize) -> Result<XmlContext> {
    let base = match xml_attribute(node, XML_NAMESPACE, "base") {
        Some(value) => Some(Arc::from(resolve_xml_base_with_limit(
            parent.base.as_deref(),
            value,
            max_length,
        )?)),
        None => parent.base.clone(),
    };
    Ok(XmlContext {
        base,
        lang: xml_attribute(node, XML_NAMESPACE, "lang")
            .map(Arc::from)
            .or_else(|| parent.lang.clone()),
        space: xml_attribute(node, XML_NAMESPACE, "space")
            .map(Arc::from)
            .or_else(|| parent.space.clone()),
    })
}

fn append_context_attribute(output: &mut Vec<u8>, name: &str, value: &str) -> Result<()> {
    let additional = context_attribute_len(name, value)?;
    let total = output
        .len()
        .checked_add(additional)
        .ok_or_else(|| limit("XML context bytes"))?;
    if total > MAX_EXTENSION_BYTES {
        return Err(limit("XML context bytes"));
    }
    output
        .try_reserve(additional)
        .map_err(|_| limit("XML context allocation"))?;
    output.extend_from_slice(b" xml:");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'"');
    Ok(())
}

fn context_attribute_len(name: &str, value: &str) -> Result<usize> {
    let escaped_len = escaped_attribute_len(value)?;
    b" xml:"
        .len()
        .checked_add(name.len())
        .and_then(|size| size.checked_add(b"=\"".len()))
        .and_then(|size| size.checked_add(escaped_len))
        .and_then(|size| size.checked_add(1))
        .ok_or_else(|| limit("XML context bytes"))
}

fn escaped_attribute_value(value: &str) -> Result<Vec<u8>> {
    let length = escaped_attribute_len(value)?;
    if length > MAX_EXTENSION_BYTES {
        return Err(limit("XML context bytes"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| limit("XML context allocation"))?;
    escape(&mut output, value);
    Ok(output)
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let escaped = escaped_character_len(character);
        length
            .checked_add(escaped)
            .ok_or_else(|| limit("XML context bytes"))
    })
}

fn escaped_character_len(character: char) -> usize {
    match character {
        '&' => 5,
        '<' => 4,
        '"' => 6,
        '\t' => 5,
        '\n' | '\r' => 6,
        _ => character.len_utf8(),
    }
}

fn namespace_escaped_len(raw: &str) -> Result<usize> {
    let mut length = 0usize;
    let mut cursor = 0usize;
    while cursor < raw.len() {
        let Some(relative_ampersand) = raw[cursor..].find('&') else {
            length = checked_namespace_literal_len(length, &raw[cursor..])?;
            break;
        };
        let ampersand = cursor + relative_ampersand;
        length = checked_namespace_literal_len(length, &raw[cursor..ampersand])?;
        let relative_semicolon = raw[ampersand..]
            .find(';')
            .ok_or_else(|| xml_error("unterminated namespace entity"))?;
        let end = ampersand + relative_semicolon + 1;
        let decoded = quick_xml::escape::unescape(&raw[ampersand..end]).map_err(xml_error)?;
        length = length
            .checked_add(escaped_attribute_len(&decoded)?)
            .ok_or_else(|| limit("namespace bytes"))?;
        cursor = end;
    }
    Ok(length)
}

fn checked_namespace_literal_len(mut length: usize, literal: &str) -> Result<usize> {
    for character in literal.chars() {
        let character = match character {
            '\t' | '\n' | '\r' => ' ',
            character => character,
        };
        length = length
            .checked_add(escaped_character_len(character))
            .ok_or_else(|| limit("namespace bytes"))?;
    }
    Ok(length)
}

fn append_escaped_namespace(output: &mut Vec<u8>, raw: &str) -> Result<()> {
    let mut cursor = 0usize;
    while cursor < raw.len() {
        let Some(relative_ampersand) = raw[cursor..].find('&') else {
            append_escaped_namespace_literal(output, &raw[cursor..]);
            break;
        };
        let ampersand = cursor + relative_ampersand;
        append_escaped_namespace_literal(output, &raw[cursor..ampersand]);
        let relative_semicolon = raw[ampersand..]
            .find(';')
            .ok_or_else(|| xml_error("unterminated namespace entity"))?;
        let end = ampersand + relative_semicolon + 1;
        let decoded = quick_xml::escape::unescape(&raw[ampersand..end]).map_err(xml_error)?;
        escape(output, &decoded);
        cursor = end;
    }
    Ok(())
}

fn append_escaped_namespace_literal(output: &mut Vec<u8>, literal: &str) {
    for character in literal.chars() {
        let character = match character {
            '\t' | '\n' | '\r' => ' ',
            character => character,
        };
        escape_character(output, character);
    }
}

fn xml_attribute_range(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    source: &[u8],
    namespace: &str,
    name: &str,
) -> Result<Option<Range<usize>>> {
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let (attribute_namespace, local) = reader.resolver().resolve_attribute(item.key);
        if resolved(attribute_namespace)? == namespace && local.as_ref() == name.as_bytes() {
            return Ok(Some(crate::source_attributes::value_span(
                source,
                item.value.as_ref(),
            )?));
        }
    }
    Ok(None)
}

fn is_absolute_uri(value: &str) -> bool {
    let Some(colon) = value.find(':') else {
        return false;
    };
    let Some(first) = value.as_bytes().first().copied() else {
        return false;
    };
    if !first.is_ascii_alphabetic()
        || value[..colon]
            .bytes()
            .any(|byte| !(byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.')))
    {
        return false;
    }
    value
        .as_bytes()
        .iter()
        .position(|byte| matches!(*byte, b'/' | b'?' | b'#'))
        .is_none_or(|delimiter| colon < delimiter)
}

fn resolve_xml_base_with_limit(
    parent: Option<&str>,
    local: &str,
    max_length: usize,
) -> Result<String> {
    let (local_main, local_fragment) = split_fragment(local);
    let (local_main, local_query) = split_query(local_main);
    if let Some(local_scheme) = scheme_prefix(local_main) {
        let (authority, path) = authority_and_path(local_main, Some(local_scheme));
        return compose_uri(
            Some(local_scheme),
            authority,
            &remove_dot_segments(path, max_length)?,
            local_query,
            local_fragment,
            max_length,
        );
    }
    if parent.is_none() {
        return clone_bounded(local, max_length, "XML context bytes");
    }
    let parent = parent.unwrap_or_default();
    if parent.is_empty() {
        return clone_bounded(local, max_length, "XML context bytes");
    }

    let (parent_main, _parent_fragment) = split_fragment(parent);
    let (parent_main, parent_query) = split_query(parent_main);
    let parent_scheme = scheme_prefix(parent_main);
    let (parent_authority, parent_path) = authority_and_path(parent_main, parent_scheme);

    if local_main.starts_with("//") {
        let (authority, path) = authority_and_path(local_main, None);
        return compose_uri(
            parent_scheme,
            authority,
            &remove_dot_segments(path, max_length)?,
            local_query,
            local_fragment,
            max_length,
        );
    }

    let (authority, path, query) = if local_main.is_empty() {
        (
            parent_authority,
            clone_bounded(parent_path, max_length, "XML context bytes")?,
            local_query.or(parent_query),
        )
    } else if local_main.starts_with('/') {
        (
            parent_authority,
            remove_dot_segments(local_main, max_length)?,
            local_query,
        )
    } else {
        let merged = merge_uri_paths(
            parent_path,
            local_main,
            parent_authority.is_some(),
            max_length,
        )?;
        (
            parent_authority,
            remove_dot_segments(&merged, max_length)?,
            local_query,
        )
    };
    compose_uri(
        parent_scheme,
        authority,
        &path,
        query,
        local_fragment,
        max_length,
    )
}

fn compose_uri(
    scheme: Option<&str>,
    authority: Option<&str>,
    path: &str,
    query: Option<&str>,
    fragment: Option<&str>,
    max_length: usize,
) -> Result<String> {
    let authority_length = authority.map_or(Ok(0), |authority| {
        authority
            .len()
            .checked_add(2)
            .ok_or_else(|| limit("XML context bytes"))
    })?;
    let query_length = query.map_or(Ok(0), |query| {
        query
            .len()
            .checked_add(1)
            .ok_or_else(|| limit("XML context bytes"))
    })?;
    let fragment_length = fragment.map_or(Ok(0), |fragment| {
        fragment
            .len()
            .checked_add(1)
            .ok_or_else(|| limit("XML context bytes"))
    })?;
    let length = scheme
        .map_or(0, str::len)
        .checked_add(authority_length)
        .and_then(|length| length.checked_add(path.len()))
        .and_then(|length| length.checked_add(query_length))
        .and_then(|length| length.checked_add(fragment_length))
        .ok_or_else(|| limit("XML context bytes"))?;
    if length > max_length {
        return Err(limit("XML context bytes"));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| limit("XML context allocation"))?;
    if let Some(scheme) = scheme {
        output.push_str(scheme);
    }
    if let Some(authority) = authority {
        output.push_str("//");
        output.push_str(authority);
    }
    output.push_str(path);
    if let Some(query) = query {
        output.push('?');
        output.push_str(query);
    }
    if let Some(fragment) = fragment {
        output.push('#');
        output.push_str(fragment);
    }
    Ok(output)
}

fn scheme_prefix(value: &str) -> Option<&str> {
    let colon = value.find(':')?;
    let first = value.as_bytes().first().copied()?;
    if !first.is_ascii_alphabetic()
        || value[..colon]
            .bytes()
            .any(|byte| !(byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.')))
    {
        None
    } else {
        Some(&value[..=colon])
    }
}

fn authority_and_path<'a>(value: &'a str, scheme: Option<&str>) -> (Option<&'a str>, &'a str) {
    let offset = scheme.map_or(0, str::len);
    let rest = value.get(offset..).unwrap_or_default();
    let Some(rest) = rest.strip_prefix("//") else {
        return (None, rest);
    };
    match rest.find('/') {
        Some(index) => (Some(&rest[..index]), &rest[index..]),
        None => (Some(rest), ""),
    }
}

fn split_fragment(value: &str) -> (&str, Option<&str>) {
    value
        .split_once('#')
        .map_or((value, None), |(main, fragment)| (main, Some(fragment)))
}

fn split_query(value: &str) -> (&str, Option<&str>) {
    value
        .split_once('?')
        .map_or((value, None), |(path, query)| (path, Some(query)))
}

fn merge_uri_paths(
    parent: &str,
    local: &str,
    has_authority: bool,
    max_length: usize,
) -> Result<String> {
    if parent.is_empty() {
        return if has_authority {
            let length = local
                .len()
                .checked_add(1)
                .ok_or_else(|| limit("XML context bytes"))?;
            if length > max_length {
                return Err(limit("XML context bytes"));
            }
            let mut output = String::new();
            output
                .try_reserve_exact(length)
                .map_err(|_| limit("XML context allocation"))?;
            output.push('/');
            output.push_str(local);
            Ok(output)
        } else {
            clone_bounded(local, max_length, "XML context bytes")
        };
    }
    let prefix_len = parent.rfind('/').map_or(0, |index| index + 1);
    let length = prefix_len
        .checked_add(local.len())
        .ok_or_else(|| limit("XML context bytes"))?;
    if length > max_length {
        return Err(limit("XML context bytes"));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| limit("XML context allocation"))?;
    output.push_str(&parent[..prefix_len]);
    output.push_str(local);
    Ok(output)
}

fn remove_dot_segments(value: &str, max_length: usize) -> Result<String> {
    if value.len() > max_length {
        return Err(limit("XML context bytes"));
    }
    let mut input = value.as_bytes();
    let mut output = Vec::new();
    output
        .try_reserve(value.len())
        .map_err(|_| limit("XML context allocation"))?;
    while !input.is_empty() {
        if input.starts_with(b"../") {
            input = &input[3..];
        } else if input.starts_with(b"./") {
            input = &input[2..];
        } else if input == b"." || input == b".." {
            input = &[];
        } else if input.starts_with(b"/./") {
            input = &input[2..];
        } else if input == b"/." {
            input = &input[..1];
        } else if input.starts_with(b"/../") {
            input = &input[3..];
            remove_last_path_segment(&mut output);
        } else if input == b"/.." {
            input = &input[..1];
            remove_last_path_segment(&mut output);
        } else {
            let end = if input.starts_with(b"/") {
                input[1..]
                    .iter()
                    .position(|byte| *byte == b'/')
                    .map_or(input.len(), |index| index + 1)
            } else {
                input
                    .iter()
                    .position(|byte| *byte == b'/')
                    .unwrap_or(input.len())
            };
            output.extend_from_slice(&input[..end]);
            input = &input[end..];
        }
    }
    Ok(String::from_utf8(output).expect("dot-segment input is valid UTF-8"))
}

fn clone_bounded(value: &str, max_length: usize, label: &str) -> Result<String> {
    if value.len() > max_length {
        return Err(limit(label));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| limit("XML context allocation"))?;
    output.push_str(value);
    Ok(output)
}

fn remove_last_path_segment(output: &mut Vec<u8>) {
    if let Some(index) = output.iter().rposition(|byte| *byte == b'/') {
        output.truncate(index);
    } else {
        output.clear();
    }
}

fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!("expected {{{namespace}}}{name}")))
    }
}

fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for attribute in &node.attributes {
        if !allowed
            .iter()
            .any(|(namespace, name)| attribute.namespace == *namespace && attribute.name == *name)
        {
            return Err(invalid(format!(
                "unexpected attribute '{}' on {}",
                attribute.name, node.name
            )));
        }
    }
    Ok(())
}

fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .ok_or_else(|| invalid(format!("missing required {name} attribute")))
}

fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}

fn leaf(node: &Node) -> Result<()> {
    whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} cannot have children", node.name)))
    }
}

fn bounded_nonempty(value: &str, label: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{label} cannot be empty")));
    }
    if value.len() > MAX_STRING_BYTES {
        return Err(limit(label));
    }
    if let Some(character) = value.chars().find(|character| !is_xml_10_char(*character)) {
        return Err(invalid(format!(
            "{label} contains XML 1.0-forbidden character U+{:04X}",
            u32::from(character)
        )));
    }
    Ok(())
}

fn is_xml_10_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(
            u32::from(character),
            0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
}

fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_TOTAL_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}

fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'"');
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        escape_character(output, character);
    }
}

fn escape_character(output: &mut Vec<u8>, character: char) {
    match character {
        '&' => output.extend_from_slice(b"&amp;"),
        '<' => output.extend_from_slice(b"&lt;"),
        '"' => output.extend_from_slice(b"&quot;"),
        '\t' => output.extend_from_slice(b"&#x9;"),
        '\n' => output.extend_from_slice(b"&#xA;"),
        '\r' => output.extend_from_slice(b"&#xD;"),
        _ => {
            let mut bytes = [0; 4];
            output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
        },
    }
}

#[cfg(test)]
mod tests {
    use crate::error::Error;

    use super::*;

    #[test]
    fn xml_base_resolution_follows_rfc3986_edge_cases() {
        let cases = [
            (
                "https://example.test/a/b?parent=1#parent",
                "",
                "https://example.test/a/b?parent=1",
            ),
            (
                "https://example.test/a/b?parent=1#parent",
                "?local=2",
                "https://example.test/a/b?local=2",
            ),
            (
                "https://example.test/a/b?parent=1#parent",
                "#fragment",
                "https://example.test/a/b?parent=1#fragment",
            ),
            (
                "https://example.test/a/b?parent=1#parent",
                "child",
                "https://example.test/a/child",
            ),
            (
                "https://example.test/a/b?parent=1#parent",
                "child?local=2#fragment",
                "https://example.test/a/child?local=2#fragment",
            ),
            (
                "https://example.test/a/b",
                "//other.test/a/../b#fragment",
                "https://other.test/b#fragment",
            ),
            (
                "https://example.test/a/b",
                "//other.test/a//b#fragment",
                "https://other.test/a//b#fragment",
            ),
            (
                "https://example.test/a/b",
                "https://other.test/a/../b?local#fragment",
                "https://other.test/b?local#fragment",
            ),
            (
                "https://example.test/a//b/",
                "./c",
                "https://example.test/a//b/c",
            ),
            (
                "https://example.test/a/b",
                "/x//y",
                "https://example.test/x//y",
            ),
            ("https://example.test/a/b", ".", "https://example.test/a/"),
            ("https://example.test/a/b", "..", "https://example.test/"),
            (
                "https://example.test/a/b/c",
                "../d/",
                "https://example.test/a/d/",
            ),
            ("urn:example:item", "child", "urn:child"),
        ];

        for (parent, local, expected) in cases {
            assert_eq!(
                resolve_xml_base_with_limit(Some(parent), local, MAX_TOTAL_STRING_BYTES).unwrap(),
                expected,
                "parent={parent:?}, local={local:?}"
            );
        }
    }

    #[test]
    fn write_preflights_exact_escaped_descriptor_size_at_the_limit() {
        let escaped = Definition {
            tables: vec![Table {
                id: "a&<>\"\t\n\r".into(),
                name: "table".into(),
                connection: "connection".into(),
            }],
            ..Definition::default()
        };
        let escaped_size = serialized_data_model_size(&escaped).unwrap();
        let escaped_output = write_data_model(&escaped).unwrap();
        assert_eq!(escaped_output.len(), escaped_size);
        assert!(
            String::from_utf8_lossy(&escaped_output).contains("a&amp;&lt;>&quot;&#x9;&#xA;&#xD;")
        );

        let max_ascii = |prefix: &str| {
            let mut value = prefix.to_owned();
            value.push_str(&"x".repeat(MAX_STRING_BYTES - prefix.len()));
            value
        };
        let mut value = Definition {
            tables: (0..4)
                .map(|index| Table {
                    id: max_ascii(&format!("id{index}")),
                    name: format!("table{index}"),
                    connection: max_ascii(&format!("connection{index}")),
                })
                .collect(),
            relationships: (0..4)
                .map(|index| Relationship {
                    from_table: format!("table{index}"),
                    from_column: max_ascii(&format!("from{index}")),
                    to_table: format!("table{}", (index + 1) % 4),
                    to_column: max_ascii(&format!("to{index}")),
                })
                .collect(),
            extension_list: None,
            ..Definition::default()
        };
        let unbounded = serialized_data_model_size_unbounded(&value).unwrap();
        assert!(unbounded > MAX_XML_BYTES);
        let excess = unbounded - MAX_XML_BYTES;
        assert!(excess < value.tables[0].id.len());
        let id = &mut value.tables[0].id;
        id.truncate(id.len() - excess);

        assert_eq!(serialized_data_model_size(&value).unwrap(), MAX_XML_BYTES);
        let exact = write_data_model(&value).unwrap();
        assert_eq!(exact.len(), MAX_XML_BYTES);

        value.tables[0].id.pop();
        assert_eq!(write_data_model(&value).unwrap().len(), MAX_XML_BYTES - 1);

        value.tables[0].id.push('x');
        value.tables[0].id.push('y');
        assert!(matches!(
            write_data_model(&value),
            Err(Error::Invalid(message)) if message.contains("serialized descriptor bytes")
        ));
    }

    #[test]
    fn context_attribute_size_is_checked_before_output_allocation() {
        let value = "&".repeat(MAX_EXTENSION_BYTES / 5 + 1);
        let mut output = Vec::new();
        assert!(append_context_attribute(&mut output, "base", &value).is_err());
        assert!(output.is_empty());
        assert!(escaped_attribute_value(&value).is_err());
    }

    #[test]
    fn oversized_opaque_source_is_rejected_before_child_materialization() {
        let text = "x".repeat(MAX_EXTENSION_BYTES);
        let xml = format!("<m:dataModel xmlns:m='{X15}'><m:extLst>{text}</m:extLst></m:dataModel>");
        let result = parse_data_model(xml.as_bytes());
        assert!(matches!(
            result,
            Err(Error::Invalid(message)) if message.contains("extension bytes")
        ));
    }

    #[test]
    fn oversized_in_scope_namespace_is_rejected_before_unescaping() {
        let namespace = "x".repeat(MAX_EXTENSION_BYTES);
        let xml =
            format!("<m:dataModel xmlns:m='{X15}' xmlns:q='{namespace}'><m:extLst/></m:dataModel>");
        let result = parse_data_model(xml.as_bytes());
        assert!(matches!(
            result,
            Err(Error::Invalid(message)) if message.contains("namespace bytes")
        ));
    }

    #[test]
    fn dot_segment_buffer_rejects_over_limit_before_reserving() {
        assert!(remove_dot_segments("a/b", 2).is_err());
    }

    fn sample_model_time_groupings() -> ModelTimeGroupings {
        ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Date & Sales".into(),
                column_name: "Order <Date>".into(),
                column_id: "date-1".into(),
                calculated_time_columns: vec![
                    CalculatedTimeColumn {
                        column_name: "Year".into(),
                        column_id: "year-1".into(),
                        content_type: ModelTimeGroupingContentType::Years,
                        is_selected: true,
                    },
                    CalculatedTimeColumn {
                        column_name: "Future".into(),
                        column_id: "future-1".into(),
                        content_type: ModelTimeGroupingContentType::Months,
                        is_selected: false,
                    },
                ],
            }],
        }
    }

    #[test]
    fn model_time_groupings_round_trip_and_escape_values() {
        let expected = sample_model_time_groupings();
        let xml = write_model_time_groupings(&expected).unwrap();
        let text = String::from_utf8_lossy(&xml);
        assert!(text.contains("Date &amp; Sales"));
        assert!(text.contains("Order &lt;Date>"));
        assert!(text.contains("isSelected=\"1\""));
        assert!(text.contains("isSelected=\"0\""));
        assert_eq!(parse_model_time_groupings(&xml).unwrap(), expected);
    }

    #[test]
    fn model_time_groupings_unknown_content_type_is_read_only() {
        let xml = format!(
            r#"<x14:modelTimeGroupings xmlns:x14="{MODEL_TIME_GROUPINGS_NAMESPACE}">
                <x14:modelTimeGrouping tableName="Date" columnName="OrderDate" columnId="d1">
                    <x14:calculatedTimeColumn columnName="Future" columnId="f1" contentType="futureUnit" isSelected="0"/>
                </x14:modelTimeGrouping>
            </x14:modelTimeGroupings>"#
        );
        let parsed = parse_model_time_groupings(xml.as_bytes()).unwrap();
        assert!(matches!(
            &parsed.groupings[0].calculated_time_columns[0].content_type,
            ModelTimeGroupingContentType::Other(value) if value == "futureUnit"
        ));
        assert!(matches!(
            write_model_time_groupings(&parsed),
            Err(Error::Unsupported { feature })
                if feature.contains("unknown modelTimeGrouping contentType")
        ));
    }

    #[test]
    fn model_time_groupings_projection_splices_only_its_known_child() {
        let source = format!(
            r#"<x15:extLst xmlns:x15="{X15}" xmlns:t="{MODEL_TIME_GROUPINGS_NAMESPACE}" xmlns:u="{MODEL_TIME_GROUPINGS_NAMESPACE}">
                <x15:ext uri="urn:vendor"><v:keep xmlns:v="urn:vendor" value="q:source"/></x15:ext>
                <x15:ext uri="{MODEL_TIME_GROUPINGS_EXTENSION_URI}"><t:modelTimeGroupings>
                    <u:modelTimeGrouping tableName="Date" columnName="OrderDate" columnId="d1">
                        <u:calculatedTimeColumn columnName="Year" columnId="y1" contentType="years" isSelected="true"/>
                    </u:modelTimeGrouping>
                </t:modelTimeGroupings></x15:ext>
            </x15:extLst>"#
        );
        let mut definition = parse_data_model(
            format!(r#"<x15:dataModel xmlns:x15="{X15}">{source}</x15:dataModel>"#).as_bytes(),
        )
        .unwrap();
        let expected = ModelTimeGroupings {
            groupings: vec![ModelTimeGrouping {
                table_name: "Date".into(),
                column_name: "OrderDate".into(),
                column_id: "d1".into(),
                calculated_time_columns: vec![CalculatedTimeColumn {
                    column_name: "Year".into(),
                    column_id: "y1".into(),
                    content_type: ModelTimeGroupingContentType::Years,
                    is_selected: true,
                }],
            }],
        };
        assert_eq!(
            definition.model_time_groupings().unwrap(),
            Some(expected.clone())
        );
        let before = definition.extension_list.as_ref().unwrap().xml.clone();
        definition
            .set_model_time_groupings(Some(sample_model_time_groupings()))
            .unwrap();
        let updated = String::from_utf8_lossy(&definition.extension_list.as_ref().unwrap().xml);
        assert!(updated.contains("urn:vendor"));
        assert!(updated.contains("Date &amp; Sales"));
        let no_op = definition.extension_list.as_ref().unwrap().xml.clone();
        let current = definition.model_time_groupings().unwrap();
        definition.set_model_time_groupings(current).unwrap();
        assert_eq!(definition.extension_list.as_ref().unwrap().xml, no_op);
        definition.set_model_time_groupings(None).unwrap();
        let removed = String::from_utf8_lossy(&definition.extension_list.as_ref().unwrap().xml);
        assert!(removed.contains("urn:vendor"));
        assert!(!removed.contains(MODEL_TIME_GROUPINGS_EXTENSION_URI));
        assert_ne!(before, definition.extension_list.as_ref().unwrap().xml);
    }

    #[test]
    fn model_time_groupings_handles_inherited_default_namespace_and_owner_prefix() {
        let source = format!(
            r#"<x15:extLst xmlns:x15="{X15}" xmlns="urn:foreign" xmlns:t="{MODEL_TIME_GROUPINGS_NAMESPACE}">
                <x15:ext uri="urn:vendor"><foreign:keep xmlns:foreign="urn:foreign"/></x15:ext>
                <x15:ext uri="{MODEL_TIME_GROUPINGS_EXTENSION_URI}"><t:modelTimeGroupings>
                    <t:modelTimeGrouping tableName="Date" columnName="OrderDate" columnId="d1">
                        <t:calculatedTimeColumn columnName="Year" columnId="y1" contentType="years" isSelected="1"/>
                    </t:modelTimeGrouping>
                </t:modelTimeGroupings></x15:ext>
            </x15:extLst>"#
        );
        let mut definition = parse_data_model(
            format!(r#"<x15:dataModel xmlns:x15="{X15}">{source}</x15:dataModel>"#).as_bytes(),
        )
        .unwrap();
        assert!(definition.model_time_groupings().unwrap().is_some());

        let value = sample_model_time_groupings();
        definition
            .set_model_time_groupings(Some(value.clone()))
            .unwrap();
        let updated = String::from_utf8_lossy(&definition.extension_list.as_ref().unwrap().xml);
        assert!(updated.contains("foreign:keep"));
        assert_eq!(definition.model_time_groupings().unwrap(), Some(value));

        definition.set_model_time_groupings(None).unwrap();
        let removed = String::from_utf8_lossy(&definition.extension_list.as_ref().unwrap().xml);
        assert!(removed.contains("urn:vendor"));
        assert!(definition.model_time_groupings().unwrap().is_none());
    }
}
