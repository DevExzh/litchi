//! Namespace-aware XML parsing and serialization for spreadsheet tracked changes.

#![allow(
    clippy::expect_used,
    reason = "expectations follow checked fixed-arity records and an active XML owner stack"
)]

use super::model::{
    Acceptance, Cell, CellAddress, CellValue, Change, Changes, ContentChange, CutOff, Deletion,
    Dimension, Info, Insertion, Integer, Metadata, Movement, NestedDeletion, PositiveInteger,
    RangeAddress, Resources,
};
use super::{MAX_DEPTH, MAX_VALUE_BYTES, limits::Limits};
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{NamespaceResolver, ResolveResult},
    reader::NsReader,
};

const MAX_EVENTS: usize = 2_000_000;
const MAX_ATTRIBUTES_PER_ELEMENT: usize = 256;
const MAX_COMMENT_BYTES: usize = 65_536;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) end: usize,
}

#[derive(Clone, Debug)]
pub(crate) struct ElementSpan {
    pub(crate) whole: Span,
    pub(crate) open: Span,
    pub(crate) close: Option<Span>,
    pub(crate) qname: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct AttributeSpan {
    pub(crate) whole: Span,
    pub(crate) value: Span,
}

#[derive(Clone, Debug)]
pub(crate) struct AttributeInsertion {
    pub(crate) at: usize,
    pub(crate) qualified_name: String,
    pub(crate) namespace_declaration: Option<String>,
}

#[derive(Clone, Debug)]
pub(crate) struct OwnerSpan {
    pub(crate) element: ElementSpan,
    pub(crate) tracking: Option<AttributeSpan>,
    pub(crate) tracking_insert: AttributeInsertion,
    pub(crate) has_unsupported_content: bool,
}

#[derive(Clone, Debug)]
pub(crate) struct RecordSpan {
    pub(crate) id: String,
    pub(crate) element: ElementSpan,
    pub(crate) acceptance: Option<AttributeSpan>,
    pub(crate) acceptance_insert: AttributeInsertion,
    pub(crate) has_unsupported_content: bool,
    pub(crate) has_rich_content: bool,
    pub(crate) regenerable: bool,
    pub(crate) resources: Resources,
}

#[derive(Clone, Debug)]
pub(crate) struct SourceMap {
    pub(crate) changes: Option<Changes>,
    pub(crate) spreadsheet: ElementSpan,
    pub(crate) owner: Option<OwnerSpan>,
    pub(crate) schema_insert: usize,
    pub(crate) records: Vec<RecordSpan>,
    pub(crate) resources: Option<Resources>,
    pub(crate) validated: Option<std::sync::Arc<super::model::Validated>>,
    pub(crate) acceptance: Vec<Option<Acceptance>>,
    pub(crate) record_resources: Vec<Resources>,
}

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DC_NS: &[u8] = b"http://purl.org/dc/elements/1.1/";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Namespace {
    None,
    Office,
    Table,
    Text,
    Dc,
    Reserved,
    Other,
}

#[derive(Debug)]
struct Attribute {
    namespace: Namespace,
    local: String,
    value: String,
}

#[derive(Debug)]
enum Content {
    Text(String),
    Node(Node),
}

#[derive(Debug)]
struct Node {
    namespace: Namespace,
    local: String,
    attributes: Vec<Attribute>,
    content: Vec<Content>,
    max_value_bytes: usize,
    opaque_foreign: bool,
    limits: Limits,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum HostKind {
    Document,
    Body,
    Spreadsheet,
    Owner,
    Record,
    Paragraph,
    Other,
}

#[derive(Debug)]
struct SourceOpen {
    kind: HostKind,
    start: usize,
    open: Span,
    qname: String,
    tracking: Option<AttributeSpan>,
    insertion: AttributeInsertion,
    unsupported: bool,
    opaque_foreign: bool,
}

#[derive(Debug)]
struct PendingRecord {
    id: String,
    start: usize,
    open: Span,
    qname: String,
    acceptance: Option<AttributeSpan>,
    insertion: AttributeInsertion,
    unsupported: bool,
    rich: bool,
    resource_start_nodes: usize,
    resource_start_aggregate_bytes: usize,
}

#[derive(Debug, Default)]
struct StartFacts {
    id: Option<String>,
    tracking: Option<AttributeSpan>,
    acceptance: Option<AttributeSpan>,
    unsupported: bool,
    retained_aggregate_bytes: usize,
}

pub(crate) fn inspect_tracked_changes_source(xml: &str, limits: &Limits) -> Result<SourceMap> {
    let mut source = inspect_source(xml, limits)?;
    let changes = if source.owner.is_some() {
        parse_tracked_changes_semantic(xml, limits)?
    } else {
        None
    };
    if source.owner.is_some() != changes.is_some() {
        return invalid("tracked-change source and semantic owner disagree");
    }
    if let Some(changes) = &changes {
        if changes.changes.len() != source.records.len() {
            return invalid("tracked-change source record count disagrees with semantic records");
        }
        for (change, record) in changes.changes.iter().zip(&source.records) {
            if change.metadata().id != record.id {
                return invalid("tracked-change source record identity mismatch");
            }
        }
    }
    let validated = changes
        .as_ref()
        .map(|changes| changes.validate_indexed_with_limits(limits))
        .transpose()?;
    let mut acceptance = Vec::new();
    acceptance
        .try_reserve_exact(source.records.len())
        .map_err(allocation_error)?;
    if let Some(changes) = &changes {
        for (change, record) in changes.changes.iter().zip(&source.records) {
            acceptance.push(record.acceptance.map(|_| change.metadata().acceptance));
        }
    }
    if let Some(validated) = &validated {
        let retained = source.resources.unwrap_or_default();
        source.resources = Some(combine_resources(validated.resources, retained, limits)?);
        if validated.record_resources.len() != source.record_resources.len() {
            return invalid("tracked-change source resource count disagrees with semantic records");
        }
        for (resource, retained) in validated
            .record_resources
            .iter()
            .zip(&mut source.record_resources)
        {
            *retained = combine_resources(*resource, *retained, limits)?;
        }
    }
    source.changes = changes;
    source.validated = validated.map(std::sync::Arc::new);
    source.acceptance = acceptance;
    Ok(source)
}

fn inspect_source(xml: &str, limits: &Limits) -> Result<SourceMap> {
    if xml.len() > limits.max_input_bytes() {
        return invalid("spreadsheet tracked-change XML exceeds the configured input limit");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<SourceOpen>::new();
    stack.try_reserve(MAX_DEPTH).map_err(allocation_error)?;
    let mut records = Vec::new();
    records
        .try_reserve(limits.max_changes().min(4096))
        .map_err(allocation_error)?;
    let mut body = None;
    let mut spreadsheet = None;
    let mut owner = None;
    let mut active_record = None::<PendingRecord>;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut spreadsheet_child_seen = false;
    let mut events = 0usize;
    let mut nodes = 0usize;
    let mut retained_nodes = 0usize;
    let mut semantic_bytes = 0usize;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("tracked-change XML event count overflow"))?;
        if events > MAX_EVENTS {
            return invalid("spreadsheet tracked-change XML exceeds the event limit");
        }
        let event_start = reader_position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&resolved)?;
        let event = event.into_owned();
        let event_is_empty = matches!(&event, Event::Empty(_));
        let event_end = reader_position(&reader)?;

        match event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = event_is_empty;
                let retained_start_nodes = retained_nodes;
                let retained_start_aggregate_bytes = semantic_bytes;
                if stack.len() >= MAX_DEPTH {
                    return invalid("spreadsheet tracked-change XML exceeds the depth limit");
                }
                if stack.is_empty() {
                    if root_seen || root_closed {
                        return invalid("ODS content.xml has more than one root element");
                    }
                    root_seen = true;
                }
                let local = decode_name(element.local_name().as_ref(), "element")?;
                let parent = stack.last().map(|open| open.kind);
                let kind = classify_host(parent, namespace, &local);
                reject_illegal_host_position(kind, parent, namespace, &local)?;
                let inside_opaque_foreign = stack.last().is_some_and(|open| open.opaque_foreign);
                let opaque_foreign = inside_opaque_foreign
                    || is_foreign(namespace)
                        && (parent == Some(HostKind::Owner) || active_record.is_some());
                let tracked_scope = kind == HostKind::Owner
                    || kind == HostKind::Record
                    || parent == Some(HostKind::Spreadsheet) && local == "tracked-changes"
                    || active_record.is_some()
                    || stack.iter().any(|open| open.kind == HostKind::Owner);
                if tracked_scope {
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| invalid_error("tracked-change XML node count overflow"))?;
                    if nodes > limits.max_nodes() {
                        return invalid("spreadsheet tracked-change XML exceeds the node limit");
                    }
                    if !inside_opaque_foreign {
                        reject_spoofed_name(namespace, &local)?;
                    }
                }
                if parent == Some(HostKind::Owner)
                    && is_foreign(namespace)
                    && let Some(open) = stack.last_mut()
                {
                    open.unsupported = true;
                }
                if parent == Some(HostKind::Spreadsheet) {
                    if kind == HostKind::Owner {
                        if spreadsheet_child_seen {
                            return invalid(
                                "table:tracked-changes must be the first office:spreadsheet element child",
                            );
                        }
                    } else {
                        spreadsheet_child_seen = true;
                    }
                }
                let facts = if tracked_scope && !opaque_foreign {
                    inspect_start_attributes(
                        xml,
                        event_start,
                        event_end,
                        &element,
                        reader.resolver(),
                        reader.decoder(),
                        true,
                    )?
                } else {
                    StartFacts::default()
                };
                let rich_markup = active_record.is_some()
                    && stack.iter().any(|open| open.kind == HostKind::Paragraph);
                if opaque_foreign || rich_markup {
                    retained_nodes = retained_nodes.checked_add(1).ok_or_else(|| {
                        invalid_error("retained tracked-change node count overflow")
                    })?;
                    let retained_attributes =
                        retained_attribute_bytes(&element, reader.resolver(), reader.decoder())?;
                    append_limited_size(
                        &mut semantic_bytes,
                        local
                            .len()
                            .checked_add(retained_attributes)
                            .ok_or_else(|| {
                                invalid_error("retained tracked-change attribute size overflow")
                            })?,
                        limits,
                    )?;
                } else {
                    append_limited_size(
                        &mut semantic_bytes,
                        facts.retained_aggregate_bytes,
                        limits,
                    )?;
                }
                if let Some(record) = active_record.as_mut() {
                    record.unsupported |= facts.unsupported;
                    if stack.iter().any(|open| open.kind == HostKind::Paragraph) {
                        record.rich = true;
                    }
                    if is_foreign(namespace) {
                        record.unsupported = true;
                    }
                }
                let qname = decode_name(element.name().as_ref(), "qualified element")?;
                let open_span = Span {
                    start: event_start,
                    end: event_end,
                };
                let mut insertion = table_attribute_insertion(&qname, event_start, event_end, xml)?;
                if kind == HostKind::Owner {
                    insertion.qualified_name = insertion
                        .qualified_name
                        .replace("acceptance-state", "track-changes");
                }
                if kind == HostKind::Record {
                    if active_record.is_some() {
                        return invalid("nested top-level tracked-change record");
                    }
                    if records.len() >= limits.max_changes() {
                        return invalid(
                            "spreadsheet tracked-change count exceeds the configured limit",
                        );
                    }
                    active_record = Some(PendingRecord {
                        id: facts.id.clone().unwrap_or_default(),
                        start: event_start,
                        open: open_span,
                        qname: qname.clone(),
                        acceptance: facts.acceptance,
                        insertion: insertion.clone(),
                        unsupported: facts.unsupported,
                        rich: false,
                        resource_start_nodes: retained_start_nodes,
                        resource_start_aggregate_bytes: retained_start_aggregate_bytes,
                    });
                }
                let opened = SourceOpen {
                    kind,
                    start: event_start,
                    open: open_span,
                    qname,
                    tracking: facts.tracking,
                    insertion,
                    unsupported: facts.unsupported,
                    opaque_foreign,
                };
                if empty {
                    complete_source_element(
                        opened,
                        None,
                        event_end,
                        &mut body,
                        &mut spreadsheet,
                        &mut owner,
                        &mut active_record,
                        &mut records,
                        retained_nodes,
                        semantic_bytes,
                    )?;
                    if stack.is_empty() {
                        root_closed = true;
                    }
                } else {
                    stack.push(opened);
                }
            },
            Event::End(_) => {
                let open = stack
                    .pop()
                    .ok_or_else(|| invalid_error("unbalanced ODS content.xml elements"))?;
                complete_source_element(
                    open,
                    Some(Span {
                        start: event_start,
                        end: event_end,
                    }),
                    event_end,
                    &mut body,
                    &mut spreadsheet,
                    &mut owner,
                    &mut active_record,
                    &mut records,
                    retained_nodes,
                    semantic_bytes,
                )?;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    invalid_error(format!("invalid tracked-change text: {error}"))
                })?;
                if stack.is_empty() && !value.trim().is_empty() {
                    return invalid("character data outside the ODS root element");
                }
                if stack.last().is_some_and(|open| open.opaque_foreign) {
                    ensure_limited_value(&value, "text", limits)?;
                    append_limited_size(&mut semantic_bytes, value.len(), limits)?;
                }
            },
            Event::CData(_) => {
                if stack.is_empty() {
                    return invalid("CDATA outside the ODS root element");
                }
                if let Some(record) = active_record.as_mut() {
                    record.unsupported = true;
                }
                if stack.iter().any(|open| open.kind == HostKind::Owner) {
                    let length = event_end
                        .checked_sub(event_start)
                        .ok_or_else(|| invalid_error("CDATA event span underflow"))?;
                    if length > limits.max_value_bytes() {
                        return invalid("tracked-change CDATA exceeds the configured value limit");
                    }
                    append_limited_size(&mut semantic_bytes, length, limits)?;
                }
            },
            Event::Comment(comment) => {
                if comment.len() > MAX_COMMENT_BYTES {
                    return invalid("ODS XML comment exceeds the 64 KiB limit");
                }
                if stack.iter().any(|open| open.kind == HostKind::Owner) {
                    append_limited_size(&mut semantic_bytes, comment.len(), limits)?;
                }
                if let Some(record) = active_record.as_mut() {
                    record.unsupported = true;
                } else if stack
                    .last()
                    .is_some_and(|open| open.kind == HostKind::Owner)
                    && let Some(open) = stack.last_mut()
                {
                    open.unsupported = true;
                }
            },
            Event::GeneralRef(reference) => {
                let name = std::str::from_utf8(reference.as_ref())
                    .map_err(|_error| invalid_error("invalid XML entity reference"))?;
                resolve_reference(name)?;
                if stack.is_empty() {
                    return invalid("entity reference outside the ODS root element");
                }
                let _ = reference;
            },
            Event::Decl(_) => {
                if declaration_seen || root_seen || root_closed {
                    return invalid("misplaced or duplicate XML declaration");
                }
                declaration_seen = true;
            },
            Event::PI(_) => {
                if stack.iter().any(|open| open.kind == HostKind::Owner) {
                    append_limited_size(
                        &mut semantic_bytes,
                        event_end
                            .checked_sub(event_start)
                            .ok_or_else(|| invalid_error("PI event span underflow"))?,
                        limits,
                    )?;
                }
                if let Some(record) = active_record.as_mut() {
                    record.unsupported = true;
                } else if stack
                    .last()
                    .is_some_and(|open| open.kind == HostKind::Owner)
                    && let Some(open) = stack.last_mut()
                {
                    open.unsupported = true;
                }
            },
            Event::DocType(_) => {
                return invalid("DTDs are prohibited in ODS tracked-change XML");
            },
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return invalid("incomplete ODS content.xml document");
    }
    if body.is_none() {
        return Err(invalid_error("ODS content.xml has no direct office:body"));
    }
    let spreadsheet = spreadsheet
        .ok_or_else(|| invalid_error("ODS content.xml has no direct office:spreadsheet"))?;
    let schema_insert = spreadsheet.open.end;
    let resources = owner.as_ref().map(|_| Resources {
        changes: 0,
        nodes: retained_nodes,
        aggregate_bytes: semantic_bytes,
    });
    let mut record_resources = Vec::new();
    record_resources
        .try_reserve_exact(records.len())
        .map_err(allocation_error)?;
    record_resources.extend(records.iter().map(|record| record.resources));
    Ok(SourceMap {
        changes: None,
        spreadsheet,
        owner,
        schema_insert,
        records,
        resources,
        validated: None,
        acceptance: Vec::new(),
        record_resources,
    })
}

fn classify_host(parent: Option<HostKind>, namespace: Namespace, local: &str) -> HostKind {
    if parent.is_none() && namespace == Namespace::Office && local == "document-content" {
        HostKind::Document
    } else if parent == Some(HostKind::Document)
        && namespace == Namespace::Office
        && local == "body"
    {
        HostKind::Body
    } else if parent == Some(HostKind::Body)
        && namespace == Namespace::Office
        && local == "spreadsheet"
    {
        HostKind::Spreadsheet
    } else if parent == Some(HostKind::Spreadsheet)
        && namespace == Namespace::Table
        && local == "tracked-changes"
    {
        HostKind::Owner
    } else if parent == Some(HostKind::Owner)
        && namespace == Namespace::Table
        && matches!(
            local,
            "insertion" | "deletion" | "movement" | "cell-content-change"
        )
    {
        HostKind::Record
    } else if namespace == Namespace::Text && local == "p" {
        HostKind::Paragraph
    } else {
        HostKind::Other
    }
}

fn reject_illegal_host_position(
    kind: HostKind,
    parent: Option<HostKind>,
    namespace: Namespace,
    local: &str,
) -> Result<()> {
    if namespace == Namespace::Office && local == "document-content" && kind != HostKind::Document
        || namespace == Namespace::Office && local == "body" && kind != HostKind::Body
        || namespace == Namespace::Office && local == "spreadsheet" && kind != HostKind::Spreadsheet
        || namespace == Namespace::Table && local == "tracked-changes" && kind != HostKind::Owner
    {
        return invalid(format!("illegal ODS host placement for {local}"));
    }
    if parent.is_none() && kind != HostKind::Document {
        return invalid("ODS content.xml root must be office:document-content");
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn complete_source_element(
    open: SourceOpen,
    close: Option<Span>,
    end: usize,
    body: &mut Option<ElementSpan>,
    spreadsheet: &mut Option<ElementSpan>,
    owner: &mut Option<OwnerSpan>,
    active_record: &mut Option<PendingRecord>,
    records: &mut Vec<RecordSpan>,
    nodes: usize,
    aggregate_bytes: usize,
) -> Result<()> {
    let element = ElementSpan {
        whole: Span {
            start: open.start,
            end,
        },
        open: open.open,
        close,
        qname: open.qname,
    };
    match open.kind {
        HostKind::Body => set_unique(body, element, "office:body")?,
        HostKind::Spreadsheet => set_unique(spreadsheet, element, "office:spreadsheet")?,
        HostKind::Owner => {
            if owner.is_some() {
                return invalid("duplicate table:tracked-changes");
            }
            *owner = Some(OwnerSpan {
                element,
                tracking: open.tracking,
                tracking_insert: open.insertion,
                has_unsupported_content: open.unsupported,
            });
        },
        HostKind::Record => {
            let pending = active_record
                .take()
                .ok_or_else(|| invalid_error("tracked-change record source state was lost"))?;
            let unsupported = pending.unsupported;
            let rich = pending.rich;
            records.try_reserve(1).map_err(allocation_error)?;
            records.push(RecordSpan {
                id: pending.id,
                element: ElementSpan {
                    whole: Span {
                        start: pending.start,
                        end,
                    },
                    open: pending.open,
                    close: element.close,
                    qname: pending.qname,
                },
                acceptance: pending.acceptance,
                acceptance_insert: pending.insertion,
                has_unsupported_content: unsupported,
                has_rich_content: rich,
                regenerable: !unsupported && !rich,
                resources: Resources {
                    changes: 0,
                    nodes: nodes
                        .checked_sub(pending.resource_start_nodes)
                        .ok_or_else(|| {
                            invalid_error("tracked-change record node counter underflow")
                        })?,
                    aggregate_bytes: aggregate_bytes
                        .checked_sub(pending.resource_start_aggregate_bytes)
                        .ok_or_else(|| {
                            invalid_error("tracked-change record byte counter underflow")
                        })?,
                },
            });
        },
        HostKind::Document | HostKind::Paragraph | HostKind::Other => {},
    }
    Ok(())
}

fn set_unique(slot: &mut Option<ElementSpan>, value: ElementSpan, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return invalid(format!("ODS content.xml has duplicate {name}"));
    }
    Ok(())
}

fn inspect_start_attributes(
    source: &str,
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    mark_foreign: bool,
) -> Result<StartFacts> {
    let lexical = lexical_attributes(source, start, end)?;
    let mut facts = StartFacts::default();
    let mut count = 0usize;
    for attribute in element.attributes() {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid_error("XML attribute count overflow"))?;
        if count > MAX_ATTRIBUTES_PER_ELEMENT {
            return invalid("XML element exceeds the attribute-count limit");
        }
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid tracked-change attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let span = lexical
            .iter()
            .find(|(name, _)| name.as_slice() == key)
            .map(|(_, span)| *span)
            .ok_or_else(|| invalid_error("could not locate XML attribute source span"))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid_error(format!("invalid XML attribute value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("XML attribute value exceeds the 64 KiB limit");
        }
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = resolver.resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode_name(local.as_ref(), "attribute")?;
        reject_spoofed_attribute(namespace, &local)?;
        if namespace == Namespace::Table && local == "id" {
            facts.id = Some(value);
        } else if namespace == Namespace::Table && local == "track-changes" {
            facts.tracking = Some(span);
        } else if namespace == Namespace::Table && local == "acceptance-state" {
            facts.acceptance = Some(span);
        } else if mark_foreign && matches!(namespace, Namespace::Other | Namespace::None) {
            facts.unsupported = true;
            facts.retained_aggregate_bytes = facts
                .retained_aggregate_bytes
                .checked_add(local.len())
                .and_then(|size| size.checked_add(value.len()))
                .ok_or_else(|| invalid_error("retained foreign attribute size overflow"))?;
        }
    }
    Ok(facts)
}

fn retained_attribute_bytes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<usize> {
    let mut total = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid retained attribute: {error}")))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = resolver.resolve_attribute(attribute.key);
        namespace_kind(&resolved)?;
        let local = decode_name(local.as_ref(), "attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid_error(format!("invalid retained attribute value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("XML attribute value exceeds the 64 KiB limit");
        }
        total = total
            .checked_add(local.len())
            .and_then(|size| size.checked_add(value.len()))
            .ok_or_else(|| invalid_error("retained attribute size overflow"))?;
    }
    Ok(total)
}

fn lexical_attributes(
    source: &str,
    start: usize,
    end: usize,
) -> Result<Vec<(Vec<u8>, AttributeSpan)>> {
    let bytes = source.as_bytes();
    if start >= end || end > bytes.len() || bytes[start] != b'<' {
        return invalid("invalid XML start-tag span");
    }
    let mut cursor = start + 1;
    while cursor < end
        && !bytes[cursor].is_ascii_whitespace()
        && !matches!(bytes[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    let mut result = Vec::new();
    result.try_reserve(8).map_err(allocation_error)?;
    loop {
        let whitespace = cursor;
        while cursor < end && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end || matches!(bytes[cursor], b'>' | b'/') {
            break;
        }
        let name_start = cursor;
        while cursor < end
            && !bytes[cursor].is_ascii_whitespace()
            && !matches!(bytes[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < end && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end || bytes[cursor] != b'=' {
            return invalid("malformed XML attribute assignment");
        }
        cursor += 1;
        while cursor < end && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= end || !matches!(bytes[cursor], b'\'' | b'\"') {
            return invalid("XML attribute value must be quoted");
        }
        let quote = bytes[cursor];
        cursor += 1;
        let value_start = cursor;
        while cursor < end && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor >= end {
            return invalid("unterminated XML attribute value");
        }
        let value_end = cursor;
        cursor += 1;
        result.try_reserve(1).map_err(allocation_error)?;
        result.push((
            bytes[name_start..name_end].to_vec(),
            AttributeSpan {
                whole: Span {
                    start: whitespace,
                    end: cursor,
                },
                value: Span {
                    start: value_start,
                    end: value_end,
                },
            },
        ));
    }
    Ok(result)
}

fn table_attribute_insertion(
    qname: &str,
    start: usize,
    end: usize,
    source: &str,
) -> Result<AttributeInsertion> {
    let bytes = source.as_bytes();
    if end > bytes.len() || end <= start || bytes[end - 1] != b'>' {
        return invalid("invalid XML start tag for attribute insertion");
    }
    let at = if end >= 2 && bytes[end - 2] == b'/' {
        end - 2
    } else {
        end - 1
    };
    let prefix = qname.split_once(':').map(|(prefix, _)| prefix);
    Ok(match prefix {
        Some(prefix) if !prefix.is_empty() => AttributeInsertion {
            at,
            qualified_name: format!("{prefix}:acceptance-state"),
            namespace_declaration: None,
        },
        _ => AttributeInsertion {
            at,
            qualified_name: "litchi_table:acceptance-state".to_string(),
            namespace_declaration: Some(format!(
                "xmlns:litchi_table=\"{}\"",
                String::from_utf8_lossy(TABLE_NS)
            )),
        },
    })
}

fn reader_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid_error("XML source position does not fit usize"))
}

fn allocation_error(error: std::collections::TryReserveError) -> Error {
    invalid_error(format!("tracked-change allocation failed: {error}"))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

/// # Errors
///
/// Returns an error when the input is malformed or exceeds the parser's resource limits.
pub fn parse_tracked_changes(xml: &str) -> Result<Option<Changes>> {
    Ok(inspect_tracked_changes_source(xml, &Limits::default())?.changes)
}

fn parse_tracked_changes_semantic(xml: &str, limits: &Limits) -> Result<Option<Changes>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut result = None;
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut aggregate = 0usize;
    let mut nodes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace)?;
        let event = event.into_owned();
        let is_start = matches!(&event, Event::Start(_));
        let is_end = matches!(&event, Event::End(_));

        if let Event::Start(element) = &event
            && namespace == Namespace::Office
            && element.local_name().as_ref() == b"spreadsheet"
        {
            spreadsheet_depth = Some(depth);
        }
        let direct_spreadsheet_child = spreadsheet_depth.is_some_and(|value| depth == value + 1);

        match event {
            Event::Start(element)
                if namespace == Namespace::Table
                    && element.local_name().as_ref() == b"tracked-changes" =>
            {
                ensure_tracked_location(direct_spreadsheet_child, result.is_some())?;
                let root = build_subtree(
                    &mut reader,
                    namespace,
                    &element,
                    &mut aggregate,
                    &mut nodes,
                    limits,
                )?;
                result = Some(parse_root(&root)?);
                break;
            },
            Event::Empty(element)
                if namespace == Namespace::Table
                    && element.local_name().as_ref() == b"tracked-changes" =>
            {
                ensure_tracked_location(direct_spreadsheet_child, result.is_some())?;
                let root = create_node(
                    reader.resolver(),
                    reader.decoder(),
                    namespace,
                    &element,
                    &mut aggregate,
                    &mut nodes,
                    false,
                    limits,
                )?;
                result = Some(parse_root(&root)?);
                break;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs are prohibited in tracked changes".to_string(),
                ));
            },
            Event::PI(_) => {},
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }

        if is_start {
            depth = depth.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("spreadsheet XML depth overflow".to_string())
            })?;
        } else if is_end {
            depth = depth.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat("spreadsheet XML depth underflow".to_string())
            })?;
        }
        buffer.clear();
    }
    Ok(result)
}

fn ensure_tracked_location(direct_child: bool, seen: bool) -> Result<()> {
    if !direct_child {
        return Err(Error::InvalidFormat(
            "table:tracked-changes must be a direct office:spreadsheet child".to_string(),
        ));
    }
    if seen {
        return Err(Error::InvalidFormat(
            "duplicate table:tracked-changes".to_string(),
        ));
    }
    Ok(())
}

fn build_subtree(
    reader: &mut NsReader<&[u8]>,
    namespace: Namespace,
    start: &BytesStart<'_>,
    aggregate: &mut usize,
    nodes: &mut usize,
    limits: &Limits,
) -> Result<Node> {
    let root = create_node(
        reader.resolver(),
        reader.decoder(),
        namespace,
        start,
        aggregate,
        nodes,
        false,
        limits,
    )?;
    let mut stack = Vec::new();
    stack.try_reserve(MAX_DEPTH).map_err(allocation_error)?;
    stack.push(root);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&namespace)?;
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "tracked-change XML exceeds depth {MAX_DEPTH}"
                    )));
                }
                let node = create_node(
                    reader.resolver(),
                    reader.decoder(),
                    namespace,
                    &element,
                    aggregate,
                    nodes,
                    stack.last().is_some_and(|parent| parent.opaque_foreign),
                    limits,
                )?;
                stack.try_reserve(1).map_err(allocation_error)?;
                stack.push(node);
            },
            Event::Empty(element) => {
                let mut node = create_node(
                    reader.resolver(),
                    reader.decoder(),
                    namespace,
                    &element,
                    aggregate,
                    nodes,
                    stack.last().is_some_and(|parent| parent.opaque_foreign),
                    limits,
                )?;
                add_semantic_leaf_text(&mut node, aggregate, limits)?;
                stack
                    .last_mut()
                    .expect("tracked-change root remains active")
                    .content
                    .try_reserve(1)
                    .map_err(allocation_error)?;
                stack
                    .last_mut()
                    .expect("tracked-change root remains active")
                    .content
                    .push(Content::Node(node));
            },
            Event::Text(text) => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid tracked-change text: {error}"))
                    })?
                    .into_owned();
                append_text(
                    stack.last_mut().expect("active node"),
                    value,
                    aggregate,
                    limits,
                )?;
            },
            Event::CData(text) => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid tracked-change CDATA: {error}"))
                })?;
                append_text(
                    stack.last_mut().expect("active node"),
                    value.into_owned(),
                    aggregate,
                    limits,
                )?;
            },
            Event::GeneralRef(reference) => {
                let name = std::str::from_utf8(reference.as_ref()).map_err(|_error| {
                    Error::InvalidFormat("invalid tracked-change entity reference".to_string())
                })?;
                append_text(
                    stack.last_mut().expect("active node"),
                    resolve_reference(name)?,
                    aggregate,
                    limits,
                )?;
            },
            Event::End(_) => {
                let mut node = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("tracked-change XML stack underflow".to_string())
                })?;
                add_semantic_leaf_text(&mut node, aggregate, limits)?;
                if let Some(parent) = stack.last_mut() {
                    parent.content.try_reserve(1).map_err(allocation_error)?;
                    parent.content.push(Content::Node(node));
                } else {
                    return Ok(node);
                }
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs are prohibited in tracked changes".to_string(),
                ));
            },
            Event::PI(_) => {},
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated table:tracked-changes".to_string(),
                ));
            },
            Event::Comment(_) | Event::Decl(_) => {},
        }
        buffer.clear();
    }
}

fn create_node(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    namespace: Namespace,
    start: &BytesStart<'_>,
    aggregate: &mut usize,
    nodes: &mut usize,
    parent_opaque_foreign: bool,
    limits: &Limits,
) -> Result<Node> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("tracked-change node count overflow".to_string()))?;
    if *nodes > limits.max_nodes() {
        return Err(Error::InvalidFormat(format!(
            "tracked changes exceed {} XML nodes",
            limits.max_nodes()
        )));
    }
    let local = decode_name(start.local_name().as_ref(), "element")?;
    let opaque_foreign = parent_opaque_foreign || is_foreign(namespace);
    if !opaque_foreign {
        reject_spoofed_name(namespace, &local)?;
    }
    append_limited_size(aggregate, local.len(), limits)?;
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid tracked-change attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local_name) = resolver.resolve_attribute(attribute.key);
        let attribute_namespace = namespace_kind(&resolved)?;
        let local_name = decode_name(local_name.as_ref(), "attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid tracked-change attribute value: {error}"))
            })?
            .into_owned();
        ensure_limited_value(&value, "attribute", limits)?;
        let size = local_name
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid_error("tracked-change attribute size overflow"))?;
        append_limited_size(aggregate, size, limits)?;
        if attributes.iter().any(|existing: &Attribute| {
            existing.namespace == attribute_namespace && existing.local == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded tracked-change attribute {local_name}"
            )));
        }
        attributes.try_reserve(1).map_err(allocation_error)?;
        attributes.push(Attribute {
            namespace: attribute_namespace,
            local: local_name,
            value,
        });
    }
    Ok(Node {
        namespace,
        local,
        attributes,
        content: Vec::new(),
        max_value_bytes: limits.max_value_bytes(),
        opaque_foreign,
        limits: *limits,
    })
}

fn append_text(
    node: &mut Node,
    value: String,
    aggregate: &mut usize,
    limits: &Limits,
) -> Result<()> {
    append_limited_size(aggregate, value.len(), limits)?;
    if let Some(Content::Text(existing)) = node.content.last_mut() {
        let length = existing
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid_error("tracked-change text size overflow"))?;
        if length > limits.max_value_bytes() {
            return Err(Error::InvalidFormat(
                "tracked-change text exceeds the configured value limit".to_string(),
            ));
        }
        existing
            .try_reserve(value.len())
            .map_err(allocation_error)?;
        existing.push_str(&value);
    } else {
        ensure_limited_value(&value, "text", limits)?;
        node.content.try_reserve(1).map_err(allocation_error)?;
        node.content.push(Content::Text(value));
    }
    Ok(())
}

fn add_semantic_leaf_text(node: &mut Node, aggregate: &mut usize, limits: &Limits) -> Result<()> {
    if node.namespace == Namespace::Text && node.local == "s" {
        if !node.content.is_empty() {
            return Err(Error::InvalidFormat("text:s must be empty".to_string()));
        }
        let count = attribute(node, Namespace::Text, "c")
            .map(|value| parse_positive(value, "text:c"))
            .transpose()?
            .unwrap_or(1);
        if count > limits.max_value_bytes() {
            return Err(Error::InvalidFormat(
                "text:s count exceeds 64 KiB".to_string(),
            ));
        }
        append_limited_size(aggregate, count, limits)?;
        let mut spaces = String::new();
        spaces.try_reserve_exact(count).map_err(allocation_error)?;
        spaces.extend(std::iter::repeat_n(' ', count));
        node.content.try_reserve(1).map_err(allocation_error)?;
        node.content.push(Content::Text(spaces));
    } else if node.namespace == Namespace::Text && node.local == "tab" {
        if !node.content.is_empty() {
            return Err(Error::InvalidFormat("text:tab must be empty".to_string()));
        }
        append_limited_size(aggregate, 1, limits)?;
        node.content.try_reserve(1).map_err(allocation_error)?;
        node.content.push(Content::Text("\t".to_string()));
    } else if node.namespace == Namespace::Text && node.local == "line-break" {
        if !node.content.is_empty() {
            return Err(Error::InvalidFormat(
                "text:line-break must be empty".to_string(),
            ));
        }
        append_limited_size(aggregate, 1, limits)?;
        node.content.try_reserve(1).map_err(allocation_error)?;
        node.content.push(Content::Text("\n".to_string()));
    }
    Ok(())
}

fn parse_root(root: &Node) -> Result<Changes> {
    reject_attributes(root, &[(Namespace::Table, "track-changes")])?;
    require_whitespace(root)?;
    let enabled = attribute(root, Namespace::Table, "track-changes")
        .map(|value| parse_bool(value, "table:track-changes"))
        .transpose()?
        .unwrap_or(false);
    let mut changes = Vec::new();
    for child in children(root) {
        if child.namespace != Namespace::Table {
            if is_foreign(child.namespace) {
                continue;
            }
            return unexpected_child(child, "table:tracked-changes");
        }
        let change = match child.local.as_str() {
            "insertion" => Change::Insertion(parse_insertion(child)?),
            "deletion" => Change::Deletion(parse_deletion(child)?),
            "movement" => Change::Movement(parse_movement(child)?),
            "cell-content-change" => Change::CellContent(parse_cell_content_change(child)?),
            _ => return unexpected_child(child, "table:tracked-changes"),
        };
        changes.try_reserve(1).map_err(allocation_error)?;
        changes.push(change);
    }
    Ok(Changes { enabled, changes })
}

fn parse_insertion(node: &Node) -> Result<Insertion> {
    reject_attributes(
        node,
        &common_attributes(&["type", "position", "count", "table"]),
    )?;
    require_child_sequence(
        node,
        &[
            ChildRule::one(Namespace::Office, "change-info"),
            ChildRule::optional(Namespace::Table, "dependencies"),
            ChildRule::optional(Namespace::Table, "deletions"),
        ],
    )?;
    Ok(Insertion {
        metadata: parse_metadata(node)?,
        dimension: parse_dimension(required_attribute(node, Namespace::Table, "type")?)?,
        position: parse_integer(
            node,
            required_attribute(node, Namespace::Table, "position")?,
            "table:position",
        )?,
        count: PositiveInteger::parse_with_limits(
            attribute(node, Namespace::Table, "count").unwrap_or("1"),
            &node.limits,
        )?,
        table: attribute(node, Namespace::Table, "table")
            .map(|value| parse_integer(node, value, "table:table"))
            .transpose()?,
    })
}

fn parse_deletion(node: &Node) -> Result<Deletion> {
    reject_attributes(
        node,
        &common_attributes(&["type", "position", "table", "multi-deletion-spanned"]),
    )?;
    require_child_sequence(
        node,
        &[
            ChildRule::one(Namespace::Office, "change-info"),
            ChildRule::optional(Namespace::Table, "dependencies"),
            ChildRule::optional(Namespace::Table, "deletions"),
            ChildRule::optional(Namespace::Table, "cut-offs"),
        ],
    )?;
    let cut_offs = optional_child(node, Namespace::Table, "cut-offs")?
        .map(parse_cut_offs)
        .transpose()?
        .unwrap_or_default();
    Ok(Deletion {
        metadata: parse_metadata(node)?,
        dimension: parse_dimension(required_attribute(node, Namespace::Table, "type")?)?,
        position: parse_integer(
            node,
            required_attribute(node, Namespace::Table, "position")?,
            "table:position",
        )?,
        table: attribute(node, Namespace::Table, "table")
            .map(|value| parse_integer(node, value, "table:table"))
            .transpose()?,
        multi_deletion_spanned: attribute(node, Namespace::Table, "multi-deletion-spanned")
            .map(|value| parse_integer(node, value, "table:multi-deletion-spanned"))
            .transpose()?,
        cut_offs,
    })
}

fn parse_movement(node: &Node) -> Result<Movement> {
    reject_attributes(node, &common_attributes(&[]))?;
    require_child_sequence(
        node,
        &[
            ChildRule::one(Namespace::Table, "source-range-address"),
            ChildRule::one(Namespace::Table, "target-range-address"),
            ChildRule::one(Namespace::Office, "change-info"),
            ChildRule::optional(Namespace::Table, "dependencies"),
            ChildRule::optional(Namespace::Table, "deletions"),
        ],
    )?;
    Ok(Movement {
        metadata: parse_metadata(node)?,
        source: parse_range(required_child(
            node,
            Namespace::Table,
            "source-range-address",
        )?)?,
        target: parse_range(required_child(
            node,
            Namespace::Table,
            "target-range-address",
        )?)?,
    })
}

fn parse_cell_content_change(node: &Node) -> Result<ContentChange> {
    reject_attributes(node, &common_attributes(&[]))?;
    require_child_sequence(
        node,
        &[
            ChildRule::one(Namespace::Table, "cell-address"),
            ChildRule::one(Namespace::Office, "change-info"),
            ChildRule::optional(Namespace::Table, "dependencies"),
            ChildRule::optional(Namespace::Table, "deletions"),
            ChildRule::one(Namespace::Table, "previous"),
        ],
    )?;
    let previous = required_child(node, Namespace::Table, "previous")?;
    reject_attributes(previous, &[(Namespace::Table, "id")])?;
    require_child_sequence(
        previous,
        &[ChildRule::one(Namespace::Table, "change-track-table-cell")],
    )?;
    let previous_cell = required_child(previous, Namespace::Table, "change-track-table-cell")?;
    Ok(ContentChange {
        metadata: parse_metadata(node)?,
        address: parse_cell_address(required_child(node, Namespace::Table, "cell-address")?)?,
        previous_change_id: attribute(previous, Namespace::Table, "id").map(str::to_string),
        previous: parse_tracked_cell(previous_cell)?,
    })
}

fn parse_metadata(node: &Node) -> Result<Metadata> {
    let id = required_attribute(node, Namespace::Table, "id")?.to_string();
    ensure_nonempty(&id, "table:id")?;
    let info = parse_change_info(required_child(node, Namespace::Office, "change-info")?)?;
    let dependencies = optional_child(node, Namespace::Table, "dependencies")?
        .map(parse_dependencies)
        .transpose()?
        .unwrap_or_default();
    let deletions = optional_child(node, Namespace::Table, "deletions")?
        .map(parse_nested_deletions)
        .transpose()?
        .unwrap_or_default();
    Ok(Metadata {
        id,
        acceptance: attribute(node, Namespace::Table, "acceptance-state")
            .map(parse_acceptance)
            .transpose()?
            .unwrap_or_default(),
        rejecting_change_id: attribute(node, Namespace::Table, "rejecting-change-id")
            .map(str::to_string),
        info,
        dependencies,
        deletions,
    })
}

fn parse_change_info(node: &Node) -> Result<Info> {
    reject_attributes(node, &[])?;
    require_child_sequence(
        node,
        &[
            ChildRule::one(Namespace::Dc, "creator"),
            ChildRule::one(Namespace::Dc, "date"),
            ChildRule::many(Namespace::Text, "p"),
        ],
    )?;
    let creator = Some(text_content(required_child(
        node,
        Namespace::Dc,
        "creator",
    )?)?);
    let date_text = text_content(required_child(node, Namespace::Dc, "date")?)?;
    let date = Some(collapse_atomic(&date_text).to_string());
    let mut comments = Vec::new();
    for child in named_children(node, Namespace::Text, "p") {
        let value = text_content(child)?;
        comments.try_reserve(1).map_err(allocation_error)?;
        comments.push(value);
    }
    for child in children(node) {
        if !matches!(
            (child.namespace, child.local.as_str()),
            (Namespace::Dc, "creator" | "date") | (Namespace::Text, "p")
        ) {
            reject_known_child(child, "office:change-info")?;
        }
    }
    Ok(Info {
        creator,
        date,
        comments,
    })
}

fn parse_dependencies(node: &Node) -> Result<Vec<String>> {
    reject_attributes(node, &[])?;
    require_whitespace(node)?;
    let mut dependencies = Vec::new();
    for child in named_children(node, Namespace::Table, "dependency") {
        reject_attributes(child, &[(Namespace::Table, "id")])?;
        reject_children(child, &[], false)?;
        let id = required_attribute(child, Namespace::Table, "id")?.to_string();
        ensure_nonempty(&id, "table:dependency table:id")?;
        dependencies.try_reserve(1).map_err(allocation_error)?;
        dependencies.push(id);
    }
    if dependencies.is_empty() {
        return Err(Error::InvalidFormat(
            "table:dependencies requires at least one table:dependency".to_string(),
        ));
    }
    for child in children(node) {
        if child.namespace != Namespace::Table || child.local != "dependency" {
            reject_known_child(child, "table:dependencies")?;
        }
    }
    Ok(dependencies)
}

fn parse_nested_deletions(node: &Node) -> Result<Vec<NestedDeletion>> {
    reject_attributes(node, &[])?;
    require_whitespace(node)?;
    let mut result = Vec::new();
    for child in children(node) {
        if child.namespace != Namespace::Table {
            reject_known_child(child, "table:deletions")?;
            continue;
        }
        match child.local.as_str() {
            "cell-content-deletion" => {
                reject_attributes(child, &[(Namespace::Table, "id")])?;
                require_child_sequence(
                    child,
                    &[
                        ChildRule::optional(Namespace::Table, "cell-address"),
                        ChildRule::optional(Namespace::Table, "change-track-table-cell"),
                    ],
                )?;
                result.try_reserve(1).map_err(allocation_error)?;
                result.push(NestedDeletion::CellContent {
                    change_id: attribute(child, Namespace::Table, "id").map(str::to_string),
                    address: optional_child(child, Namespace::Table, "cell-address")?
                        .map(parse_cell_address)
                        .transpose()?,
                    cell: optional_child(child, Namespace::Table, "change-track-table-cell")?
                        .map(parse_tracked_cell)
                        .transpose()?,
                });
            },
            "change-deletion" => {
                reject_attributes(child, &[(Namespace::Table, "id")])?;
                reject_children(child, &[], false)?;
                result.try_reserve(1).map_err(allocation_error)?;
                result.push(NestedDeletion::Change {
                    change_id: attribute(child, Namespace::Table, "id").map(str::to_string),
                });
            },
            _ => return unexpected_child(child, "table:deletions"),
        }
    }
    if result.is_empty() {
        return Err(Error::InvalidFormat(
            "table:deletions requires at least one deletion".to_string(),
        ));
    }
    Ok(result)
}

fn parse_cut_offs(node: &Node) -> Result<Vec<CutOff>> {
    reject_attributes(node, &[])?;
    require_whitespace(node)?;
    let mut result = Vec::new();
    let mut insertion_seen = false;
    for child in children(node) {
        if child.namespace != Namespace::Table {
            reject_known_child(child, "table:cut-offs")?;
            continue;
        }
        match child.local.as_str() {
            "insertion-cut-off" => {
                if insertion_seen || !result.is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:insertion-cut-off must occur at most once and first".to_string(),
                    ));
                }
                insertion_seen = true;
                reject_attributes(
                    child,
                    &[(Namespace::Table, "id"), (Namespace::Table, "position")],
                )?;
                reject_children(child, &[], false)?;
                result.try_reserve(1).map_err(allocation_error)?;
                result.push(CutOff::Insertion {
                    change_id: required_attribute(child, Namespace::Table, "id")?.to_string(),
                    position: parse_integer(
                        node,
                        required_attribute(child, Namespace::Table, "position")?,
                        "table:position",
                    )?,
                });
            },
            "movement-cut-off" => {
                reject_attributes(
                    child,
                    &[
                        (Namespace::Table, "position"),
                        (Namespace::Table, "start-position"),
                        (Namespace::Table, "end-position"),
                    ],
                )?;
                reject_children(child, &[], false)?;
                let position = attribute(child, Namespace::Table, "position");
                let start = attribute(child, Namespace::Table, "start-position");
                let end = attribute(child, Namespace::Table, "end-position");
                let cut_off = match (position, start, end) {
                    (Some(value), None, None) => CutOff::MovementPoint {
                        position: parse_integer(node, value, "table:position")?,
                    },
                    (None, Some(start), Some(end)) => {
                        let start = parse_integer(node, start, "table:start-position")?;
                        let end = parse_integer(node, end, "table:end-position")?;
                        if start >= end {
                            return Err(Error::InvalidFormat(
                                "movement cut-off start must precede end".to_string(),
                            ));
                        }
                        CutOff::MovementRange { start, end }
                    },
                    _ => {
                        return Err(Error::InvalidFormat(
                            "movement cut-off requires position or start/end positions".to_string(),
                        ));
                    },
                };
                result.try_reserve(1).map_err(allocation_error)?;
                result.push(cut_off);
            },
            _ => return unexpected_child(child, "table:cut-offs"),
        }
    }
    if result.is_empty() {
        return Err(Error::InvalidFormat(
            "table:cut-offs requires at least one cut-off".to_string(),
        ));
    }
    Ok(result)
}

fn parse_range(node: &Node) -> Result<RangeAddress> {
    reject_attributes(
        node,
        &[
            (Namespace::Table, "column"),
            (Namespace::Table, "row"),
            (Namespace::Table, "table"),
            (Namespace::Table, "start-column"),
            (Namespace::Table, "start-row"),
            (Namespace::Table, "start-table"),
            (Namespace::Table, "end-column"),
            (Namespace::Table, "end-row"),
            (Namespace::Table, "end-table"),
        ],
    )?;
    reject_children(node, &[], false)?;
    let cell = ["table", "column", "row"].map(|name| attribute(node, Namespace::Table, name));
    let range = [
        "start-table",
        "start-column",
        "start-row",
        "end-table",
        "end-column",
        "end-row",
    ]
    .map(|name| attribute(node, Namespace::Table, name));
    if cell.iter().all(Option::is_some) && range.iter().all(Option::is_none) {
        return Ok(RangeAddress::Cell(CellAddress {
            table: parse_integer(node, cell[0].expect("present"), "table:table")?,
            column: parse_integer(node, cell[1].expect("present"), "table:column")?,
            row: parse_integer(node, cell[2].expect("present"), "table:row")?,
        }));
    }
    if cell.iter().all(Option::is_none) && range.iter().all(Option::is_some) {
        return Ok(RangeAddress::Range {
            start: CellAddress {
                table: parse_integer(node, range[0].expect("present"), "table:start-table")?,
                column: parse_integer(node, range[1].expect("present"), "table:start-column")?,
                row: parse_integer(node, range[2].expect("present"), "table:start-row")?,
            },
            end: CellAddress {
                table: parse_integer(node, range[3].expect("present"), "table:end-table")?,
                column: parse_integer(node, range[4].expect("present"), "table:end-column")?,
                row: parse_integer(node, range[5].expect("present"), "table:end-row")?,
            },
        });
    }
    Err(Error::InvalidFormat(
        "tracked range requires either cell or complete start/end coordinates".to_string(),
    ))
}

fn parse_cell_address(node: &Node) -> Result<CellAddress> {
    reject_attributes(
        node,
        &[
            (Namespace::Table, "table"),
            (Namespace::Table, "column"),
            (Namespace::Table, "row"),
        ],
    )?;
    reject_children(node, &[], false)?;
    Ok(CellAddress {
        table: parse_integer(
            node,
            required_attribute(node, Namespace::Table, "table")?,
            "table:table",
        )?,
        column: parse_integer(
            node,
            required_attribute(node, Namespace::Table, "column")?,
            "table:column",
        )?,
        row: parse_integer(
            node,
            required_attribute(node, Namespace::Table, "row")?,
            "table:row",
        )?,
    })
}

fn parse_tracked_cell(node: &Node) -> Result<Cell> {
    reject_attributes(
        node,
        &[
            (Namespace::Table, "cell-address"),
            (Namespace::Table, "style-name"),
            (Namespace::Table, "matrix-covered"),
            (Namespace::Table, "formula"),
            (Namespace::Table, "number-matrix-columns-spanned"),
            (Namespace::Table, "number-matrix-rows-spanned"),
            (Namespace::Office, "value-type"),
            (Namespace::Office, "value"),
            (Namespace::Office, "boolean-value"),
            (Namespace::Office, "currency"),
            (Namespace::Office, "date-value"),
            (Namespace::Office, "string-value"),
            (Namespace::Office, "time-value"),
        ],
    )?;
    require_whitespace(node)?;
    for child in children(node) {
        if child.namespace != Namespace::Text || child.local != "p" {
            reject_known_child(child, "table:change-track-table-cell")?;
        }
    }
    let mut display_text = String::new();
    for (index, child) in named_children(node, Namespace::Text, "p").enumerate() {
        let paragraph = text_content(child)?;
        let extra = paragraph
            .len()
            .checked_add(usize::from(index != 0))
            .ok_or_else(|| invalid_error("tracked cell display text size overflow"))?;
        let total = display_text
            .len()
            .checked_add(extra)
            .ok_or_else(|| invalid_error("tracked cell display text size overflow"))?;
        if total > node.max_value_bytes {
            return invalid("tracked cell display text exceeds the configured value limit");
        }
        display_text.try_reserve(extra).map_err(allocation_error)?;
        if index != 0 {
            display_text.push('\n');
        }
        display_text.push_str(&paragraph);
    }
    let value_type = attribute(node, Namespace::Office, "value-type").map(collapse_atomic);
    reject_incompatible_value_attributes(node, value_type)?;
    let has_value_attribute = node.attributes.iter().any(|attribute| {
        attribute.namespace == Namespace::Office && attribute.local != "value-type"
    });
    if value_type.is_none() && has_value_attribute {
        return Err(Error::InvalidFormat(
            "tracked cell value attributes require office:value-type".to_string(),
        ));
    }
    let value = match value_type {
        None => CellValue::Empty,
        Some("boolean") => CellValue::Boolean(parse_bool(
            required_attribute(node, Namespace::Office, "boolean-value")?,
            "office:boolean-value",
        )?),
        Some("float") => CellValue::Number(parse_f64(
            required_attribute(node, Namespace::Office, "value")?,
            "office:value",
        )?),
        Some("percentage") => CellValue::Percentage(parse_f64(
            required_attribute(node, Namespace::Office, "value")?,
            "office:value",
        )?),
        Some("currency") => CellValue::Currency {
            value: parse_f64(
                required_attribute(node, Namespace::Office, "value")?,
                "office:value",
            )?,
            code: required_attribute(node, Namespace::Office, "currency")?.to_string(),
        },
        Some("date") => CellValue::Date(
            collapse_atomic(required_attribute(node, Namespace::Office, "date-value")?).to_string(),
        ),
        Some("time") => CellValue::Time(
            collapse_atomic(required_attribute(node, Namespace::Office, "time-value")?).to_string(),
        ),
        Some("string") => CellValue::Text(
            attribute(node, Namespace::Office, "string-value")
                .unwrap_or(&display_text)
                .to_string(),
        ),
        Some("error") => {
            CellValue::Error(attribute(node, Namespace::Office, "string-value").map(str::to_string))
        },
        Some(other) => {
            return Err(Error::InvalidFormat(format!(
                "unsupported tracked cell office:value-type '{other}'"
            )));
        },
    };
    Ok(Cell {
        address: attribute(node, Namespace::Table, "cell-address")
            .map(collapse_atomic)
            .map(str::to_string),
        style_name: attribute(node, Namespace::Table, "style-name")
            .map(collapse_atomic)
            .map(str::to_string),
        matrix_covered: attribute(node, Namespace::Table, "matrix-covered")
            .map(|value| parse_bool(value, "table:matrix-covered"))
            .transpose()?
            .unwrap_or(false),
        formula: attribute(node, Namespace::Table, "formula").map(str::to_string),
        matrix_columns: optional_positive_nonzero(node, "number-matrix-columns-spanned")?,
        matrix_rows: optional_positive_nonzero(node, "number-matrix-rows-spanned")?,
        value,
        display_text,
    })
}

fn reject_incompatible_value_attributes(node: &Node, value_type: Option<&str>) -> Result<()> {
    let allowed: &[&str] = match value_type {
        None => &[],
        Some("boolean") => &["boolean-value"],
        Some("float" | "percentage") => &["value"],
        Some("currency") => &["value", "currency"],
        Some("date") => &["date-value"],
        Some("time") => &["time-value"],
        Some("string" | "error") => &["string-value"],
        Some(_) => &[],
    };
    for attribute in &node.attributes {
        if attribute.namespace == Namespace::Office
            && attribute.local != "value-type"
            && !allowed.contains(&attribute.local.as_str())
        {
            return invalid(format!(
                "office:value-type '{}' is incompatible with office:{}",
                value_type.unwrap_or("absent"),
                attribute.local
            ));
        }
    }
    Ok(())
}

impl Changes {
    /// Return canonical ODF XML for this `table:tracked-changes` fragment.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn to_xml_fragment(&self) -> Result<String> {
        write_tracked_changes_owner(self, Some(self.enabled), &Limits::default())
    }
}

/// # Errors
///
/// Returns an error when the value cannot be serialized.
pub fn write_tracked_changes(output: &mut String, changes: Option<&Changes>) -> Result<()> {
    if let Some(changes) = changes {
        let fragment =
            write_tracked_changes_owner(changes, Some(changes.enabled), &Limits::default())?;
        output
            .try_reserve(fragment.len())
            .map_err(allocation_error)?;
        output.push_str(&fragment);
    }
    Ok(())
}

pub(crate) fn write_tracked_changes_owner(
    changes: &Changes,
    tracking: Option<bool>,
    limits: &Limits,
) -> Result<String> {
    changes.validate_with_limits(limits)?;
    let estimate = estimate_changes_output(changes)?;
    if estimate > limits.max_output_bytes() {
        return invalid("tracked-change XML exceeds the configured output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(estimate)
        .map_err(allocation_error)?;
    write_tracked_changes_unchecked(&mut output, changes);
    const TRACK_PREFIX: &str = " table:track-changes=\"";
    let value_start = output
        .find(TRACK_PREFIX)
        .and_then(|start| start.checked_add(TRACK_PREFIX.len()))
        .ok_or_else(|| invalid_error("tracked-change writer omitted tracking attribute"))?;
    let value_end = output[value_start..]
        .find('\"')
        .and_then(|length| value_start.checked_add(length))
        .ok_or_else(|| {
            invalid_error("tracked-change writer emitted malformed tracking attribute")
        })?;
    match tracking {
        Some(value) => {
            output.replace_range(value_start..value_end, if value { "true" } else { "false" });
        },
        None => {
            let start = value_start
                .checked_sub(TRACK_PREFIX.len())
                .ok_or_else(|| invalid_error("tracking attribute span underflow"))?;
            let end = value_end
                .checked_add(1)
                .ok_or_else(|| invalid_error("tracking attribute span overflow"))?;
            output.replace_range(start..end, "");
        },
    }
    const OWNER: &str = "<table:tracked-changes";
    if !output.starts_with(OWNER) {
        return invalid("tracked-change writer emitted an unexpected owner element");
    }
    output.insert_str(OWNER.len(), namespace_declarations());
    ensure_output_bound(&output, limits.max_output_bytes())?;
    Ok(output)
}

pub(crate) fn write_tracked_change(
    change: &Change,
    include_acceptance: bool,
    limits: &Limits,
) -> Result<String> {
    let mut owner = String::new();
    let estimate = estimate_change_slice_output(std::slice::from_ref(change))?;
    if estimate > limits.max_output_bytes() {
        return invalid("tracked-change record exceeds the configured output limit");
    }
    owner
        .try_reserve_exact(estimate)
        .map_err(allocation_error)?;
    write_tracked_changes_slice_unchecked(&mut owner, true, std::slice::from_ref(change));
    let start = owner
        .find('>')
        .and_then(|index| index.checked_add(1))
        .ok_or_else(|| invalid_error("tracked-change writer emitted malformed owner start"))?;
    let end = owner
        .rfind("</table:tracked-changes>")
        .ok_or_else(|| invalid_error("tracked-change writer omitted owner close"))?;
    let record = owner
        .get(start..end)
        .ok_or_else(|| invalid_error("tracked-change writer emitted invalid record span"))?;
    const EXPLICIT_PENDING: &str = " table:acceptance-state=\"pending\"";
    let reserved = record
        .len()
        .checked_add(namespace_declarations().len())
        .and_then(|value| {
            value.checked_add(if include_acceptance {
                EXPLICIT_PENDING.len()
            } else {
                0
            })
        })
        .ok_or_else(|| invalid_error("tracked-change record size overflow"))?;
    if reserved > limits.max_output_bytes() {
        return invalid("tracked-change record exceeds the configured output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(reserved)
        .map_err(allocation_error)?;
    output.push_str(record);
    const ACCEPTANCE: &str = " table:acceptance-state=\"";
    let acceptance_range = output
        .find(ACCEPTANCE)
        .map(|attr_start| {
            let value_start = attr_start + ACCEPTANCE.len();
            output[value_start..]
                .find('\"')
                .map(|length| attr_start..value_start + length + 1)
                .ok_or_else(|| {
                    invalid_error("tracked-change writer emitted malformed acceptance attribute")
                })
        })
        .transpose()?;
    match (include_acceptance, acceptance_range) {
        (false, Some(range)) => output.replace_range(range, ""),
        (true, None) => {
            let insert = output
                .find(|character: char| character.is_ascii_whitespace() || character == '>')
                .ok_or_else(|| {
                    invalid_error("tracked-change writer emitted malformed record start")
                })?;
            output.insert_str(insert, EXPLICIT_PENDING);
        },
        _ => {},
    }
    let insert = output
        .find(|character: char| character.is_ascii_whitespace() || character == '>')
        .ok_or_else(|| invalid_error("tracked-change writer emitted malformed record start"))?;
    output.insert_str(insert, namespace_declarations());
    ensure_output_bound(&output, limits.max_output_bytes())?;
    Ok(output)
}

pub(crate) fn rewrite_record_acceptance(
    xml: &str,
    record: &RecordSpan,
    acceptance: Option<Acceptance>,
    limits: &Limits,
) -> Result<String> {
    rewrite_attribute_fragment(
        xml,
        &record.element,
        record.acceptance,
        &record.acceptance_insert,
        acceptance.map(acceptance_attribute_value),
        limits.max_output_bytes(),
    )
}

pub(crate) fn rewrite_owner_tracking(
    xml: &str,
    owner: &OwnerSpan,
    tracking: Option<bool>,
    limits: &Limits,
) -> Result<String> {
    rewrite_attribute_fragment(
        xml,
        &owner.element,
        owner.tracking,
        &owner.tracking_insert,
        tracking.map(|value| if value { "true" } else { "false" }),
        limits.max_output_bytes(),
    )
}

pub(crate) fn insert_tracked_change_into_owner(
    xml: &str,
    owner: &OwnerSpan,
    fragment: &str,
    limits: &Limits,
) -> Result<String> {
    let source = xml
        .get(owner.element.whole.start..owner.element.whole.end)
        .ok_or_else(|| invalid_error("invalid tracked-change owner source span"))?;
    let capacity = if owner.element.close.is_some() {
        source.len().checked_add(fragment.len())
    } else {
        source
            .len()
            .checked_sub(1)
            .and_then(|value| value.checked_add(fragment.len()))
            .and_then(|value| value.checked_add(owner.element.qname.len()))
            .and_then(|value| value.checked_add(3))
    }
    .ok_or_else(|| invalid_error("tracked-change owner output size overflow"))?;
    if capacity > limits.max_output_bytes() {
        return invalid("tracked-change owner exceeds the configured output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(allocation_error)?;
    if let Some(close) = owner.element.close {
        let insertion = close
            .start
            .checked_sub(owner.element.whole.start)
            .ok_or_else(|| invalid_error("tracked-change owner close span underflow"))?;
        let (before, after) = source
            .split_at_checked(insertion)
            .ok_or_else(|| invalid_error("tracked-change owner close span is invalid"))?;
        output.push_str(before);
        output.push_str(fragment);
        output.push_str(after);
    } else {
        let slash = source
            .len()
            .checked_sub(2)
            .filter(|index| source.as_bytes().get(*index) == Some(&b'/'))
            .ok_or_else(|| invalid_error("self-closing tracked-change owner is malformed"))?;
        output.push_str(&source[..slash]);
        output.push('>');
        output.push_str(fragment);
        output.push_str("</");
        output.push_str(&owner.element.qname);
        output.push('>');
    }
    ensure_output_bound(&output, limits.max_output_bytes())?;
    Ok(output)
}

pub(crate) fn insert_tracked_owner_into_spreadsheet(
    xml: &str,
    spreadsheet: &ElementSpan,
    owner_fragment: &str,
    limits: &Limits,
) -> Result<String> {
    let source = xml
        .get(spreadsheet.whole.start..spreadsheet.whole.end)
        .ok_or_else(|| invalid_error("invalid office:spreadsheet source span"))?;
    if spreadsheet.close.is_some() {
        return invalid("office:spreadsheet insertion helper requires a self-closing host");
    }
    let slash = source
        .len()
        .checked_sub(2)
        .filter(|index| source.as_bytes().get(*index) == Some(&b'/'))
        .ok_or_else(|| invalid_error("self-closing office:spreadsheet is malformed"))?;
    let capacity = source
        .len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(owner_fragment.len()))
        .and_then(|value| value.checked_add(spreadsheet.qname.len()))
        .and_then(|value| value.checked_add(3))
        .ok_or_else(|| invalid_error("office:spreadsheet insertion size overflow"))?;
    if capacity > limits.max_output_bytes() {
        return invalid("expanded office:spreadsheet exceeds the configured output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(allocation_error)?;
    output.push_str(&source[..slash]);
    output.push('>');
    output.push_str(owner_fragment);
    output.push_str("</");
    output.push_str(&spreadsheet.qname);
    output.push('>');
    Ok(output)
}

fn rewrite_attribute_fragment(
    xml: &str,
    element: &ElementSpan,
    attribute: Option<AttributeSpan>,
    insertion: &AttributeInsertion,
    value: Option<&str>,
    max_output_bytes: usize,
) -> Result<String> {
    let source = xml
        .get(element.whole.start..element.whole.end)
        .ok_or_else(|| invalid_error("invalid tracked-change element source span"))?;
    let mut output = String::new();
    output
        .try_reserve(
            source
                .len()
                .checked_add(256)
                .ok_or_else(|| invalid_error("rewritten XML size overflow"))?,
        )
        .map_err(allocation_error)?;
    output.push_str(source);
    let relative = |position: usize| {
        position
            .checked_sub(element.whole.start)
            .ok_or_else(|| invalid_error("attribute source span precedes its element"))
    };
    match (attribute, value) {
        (Some(attribute), Some(value)) => {
            let start = relative(attribute.value.start)?;
            let end = relative(attribute.value.end)?;
            let escaped = escaped_xml(value, true, max_output_bytes)?;
            output.replace_range(start..end, &escaped);
        },
        (Some(attribute), None) => {
            output.replace_range(
                relative(attribute.whole.start)?..relative(attribute.whole.end)?,
                "",
            );
        },
        (None, Some(value)) => {
            let at = relative(insertion.at)?;
            let escaped = escaped_xml(value, true, max_output_bytes)?;
            let mut addition = String::new();
            addition
                .try_reserve(
                    escaped
                        .len()
                        .checked_add(192)
                        .ok_or_else(|| invalid_error("attribute insertion size overflow"))?,
                )
                .map_err(allocation_error)?;
            if let Some(declaration) = &insertion.namespace_declaration {
                addition.push(' ');
                addition.push_str(declaration);
            }
            addition.push(' ');
            addition.push_str(&insertion.qualified_name);
            addition.push_str("=\"");
            addition.push_str(&escaped);
            addition.push('\"');
            output.insert_str(at, &addition);
        },
        (None, None) => {},
    }
    ensure_output_bound(&output, max_output_bytes)?;
    Ok(output)
}

pub(crate) const fn acceptance_attribute_value(value: Acceptance) -> &'static str {
    match value {
        Acceptance::Accepted => "accepted",
        Acceptance::Rejected => "rejected",
        Acceptance::Pending => "pending",
    }
}

fn namespace_declarations() -> &'static str {
    " xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\""
}

fn escaped_xml(value: &str, attribute: bool, max_output_bytes: usize) -> Result<String> {
    let capacity = value
        .len()
        .checked_mul(6)
        .ok_or_else(|| invalid_error("escaped XML size overflow"))?;
    if capacity > max_output_bytes {
        return invalid("escaped XML exceeds the output limit");
    }
    let mut output = String::new();
    output.try_reserve(capacity).map_err(allocation_error)?;
    escape_tracked_xml(&mut output, value, attribute);
    Ok(output)
}

fn ensure_output_bound(output: &str, max_output_bytes: usize) -> Result<()> {
    if output.len() > max_output_bytes {
        return invalid("tracked-change XML exceeds the configured output limit");
    }
    Ok(())
}

fn estimate_changes_output(changes: &Changes) -> Result<usize> {
    estimate_change_slice_output(&changes.changes)
}

fn estimate_change_slice_output(changes: &[Change]) -> Result<usize> {
    let mut size = 1024usize;
    for change in changes {
        checked_output_add(&mut size, 1024)?;
        estimate_metadata_output(&mut size, change.metadata())?;
        match change {
            Change::Insertion(value) => {
                estimate_integer_output(&mut size, &value.position)?;
                estimate_positive_integer_output(&mut size, &value.count)?;
                if let Some(value) = &value.table {
                    estimate_integer_output(&mut size, value)?;
                }
            },
            Change::Movement(value) => {
                estimate_range_output(&mut size, &value.source)?;
                estimate_range_output(&mut size, &value.target)?;
            },
            Change::Deletion(value) => {
                estimate_integer_output(&mut size, &value.position)?;
                if let Some(value) = &value.table {
                    estimate_integer_output(&mut size, value)?;
                }
                if let Some(value) = &value.multi_deletion_spanned {
                    estimate_integer_output(&mut size, value)?;
                }
                checked_output_add(
                    &mut size,
                    value
                        .cut_offs
                        .len()
                        .checked_mul(256)
                        .ok_or_else(|| invalid_error("tracked-change output estimate overflow"))?,
                )?;
                for cut_off in &value.cut_offs {
                    if let CutOff::Insertion { change_id, .. } = cut_off {
                        checked_output_add(&mut size, escaped_size(change_id, true)?)?;
                    }
                    match cut_off {
                        CutOff::Insertion { position, .. } | CutOff::MovementPoint { position } => {
                            estimate_integer_output(&mut size, position)?;
                        },
                        CutOff::MovementRange { start, end } => {
                            estimate_integer_output(&mut size, start)?;
                            estimate_integer_output(&mut size, end)?;
                        },
                    }
                }
            },
            Change::CellContent(value) => {
                checked_output_add(&mut size, 1024)?;
                estimate_address_output(&mut size, &value.address)?;
                if let Some(id) = &value.previous_change_id {
                    checked_output_add(&mut size, escaped_size(id, true)?)?;
                }
                estimate_cell_output(&mut size, &value.previous)?;
            },
        }
    }
    Ok(size)
}

fn estimate_metadata_output(size: &mut usize, metadata: &Metadata) -> Result<()> {
    checked_output_add(size, escaped_size(&metadata.id, true)?)?;
    if let Some(value) = &metadata.rejecting_change_id {
        checked_output_add(size, escaped_size(value, true)?)?;
    }
    for value in [
        metadata.info.creator.as_deref(),
        metadata.info.date.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        checked_output_add(size, escaped_size(value, false)?)?;
    }
    for value in &metadata.info.comments {
        checked_output_add(size, 128)?;
        checked_output_add(size, escaped_size(value, false)?)?;
    }
    for value in &metadata.dependencies {
        checked_output_add(size, 128)?;
        checked_output_add(size, escaped_size(value, true)?)?;
    }
    for deletion in &metadata.deletions {
        checked_output_add(size, 512)?;
        match deletion {
            NestedDeletion::CellContent {
                change_id,
                address,
                cell,
            } => {
                if let Some(value) = change_id {
                    checked_output_add(size, escaped_size(value, true)?)?;
                }
                if let Some(address) = address {
                    estimate_address_output(size, address)?;
                }
                if let Some(cell) = cell {
                    estimate_cell_output(size, cell)?;
                }
            },
            NestedDeletion::Change { change_id } => {
                if let Some(value) = change_id {
                    checked_output_add(size, escaped_size(value, true)?)?;
                }
            },
        }
    }
    Ok(())
}

fn estimate_cell_output(size: &mut usize, cell: &Cell) -> Result<()> {
    checked_output_add(size, 1024)?;
    for value in [
        cell.address.as_deref(),
        cell.style_name.as_deref(),
        cell.formula.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        checked_output_add(size, escaped_size(value, true)?)?;
    }
    match &cell.value {
        CellValue::Currency { code, .. } => {
            checked_output_add(size, escaped_size(code, true)?)?;
        },
        CellValue::Date(value) | CellValue::Time(value) | CellValue::Text(value) => {
            checked_output_add(size, escaped_size(value, true)?)?;
        },
        CellValue::Error(Some(value)) => {
            checked_output_add(size, escaped_size(value, true)?)?;
        },
        CellValue::Empty
        | CellValue::Boolean(_)
        | CellValue::Number(_)
        | CellValue::Percentage(_)
        | CellValue::Error(_) => {},
    }
    if let Some(value) = &cell.matrix_columns {
        estimate_positive_integer_output(size, value)?;
    }
    if let Some(value) = &cell.matrix_rows {
        estimate_positive_integer_output(size, value)?;
    }
    checked_output_add(size, escaped_size(&cell.display_text, false)?)?;
    checked_output_add(
        size,
        cell.display_text
            .bytes()
            .filter(|value| *value == b'\n')
            .count()
            .checked_mul(32)
            .ok_or_else(|| invalid_error("tracked-change paragraph estimate overflow"))?,
    )
}

fn estimate_range_output(size: &mut usize, range: &RangeAddress) -> Result<()> {
    match range {
        RangeAddress::Cell(value) => estimate_address_output(size, value),
        RangeAddress::Range { start, end } => {
            estimate_address_output(size, start)?;
            estimate_address_output(size, end)
        },
    }
}

fn estimate_address_output(size: &mut usize, address: &CellAddress) -> Result<()> {
    estimate_integer_output(size, &address.table)?;
    estimate_integer_output(size, &address.column)?;
    estimate_integer_output(size, &address.row)
}

fn estimate_integer_output(size: &mut usize, value: &Integer) -> Result<()> {
    checked_output_add(size, value.as_str().len())
}

fn estimate_positive_integer_output(size: &mut usize, value: &PositiveInteger) -> Result<()> {
    checked_output_add(size, value.as_str().len())
}

fn escaped_size(value: &str, attribute: bool) -> Result<usize> {
    let mut size = 0usize;
    for character in value.chars() {
        let amount = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' if attribute => 6,
            '\'' if attribute => 6,
            value => value.len_utf8(),
        };
        checked_output_add(&mut size, amount)?;
    }
    Ok(size)
}

fn checked_output_add(size: &mut usize, amount: usize) -> Result<()> {
    *size = size
        .checked_add(amount)
        .ok_or_else(|| invalid_error("tracked-change output size overflow"))?;
    Ok(())
}

fn write_tracked_changes_unchecked(output: &mut String, changes: &Changes) {
    write_tracked_changes_slice_unchecked(output, changes.enabled, &changes.changes);
}

fn write_tracked_changes_slice_unchecked(output: &mut String, enabled: bool, changes: &[Change]) {
    output.push_str("<table:tracked-changes table:track-changes=\"");
    output.push_str(if enabled { "true" } else { "false" });
    output.push_str("\">");
    for change in changes {
        match change {
            Change::Insertion(value) => {
                output.push_str("<table:insertion");
                write_common_tracked_attributes(output, &value.metadata);
                push_tracked_attr(output, "table:type", tracked_dimension(value.dimension));
                push_tracked_integer(output, "table:position", &value.position);
                if value.count.as_str() != "1" {
                    push_tracked_attr(output, "table:count", value.count.as_str());
                }
                if let Some(table) = &value.table {
                    push_tracked_integer(output, "table:table", table);
                }
                output.push('>');
                write_tracked_metadata(output, &value.metadata);
                output.push_str("</table:insertion>");
            },
            Change::Deletion(value) => {
                output.push_str("<table:deletion");
                write_common_tracked_attributes(output, &value.metadata);
                push_tracked_attr(output, "table:type", tracked_dimension(value.dimension));
                push_tracked_integer(output, "table:position", &value.position);
                if let Some(table) = &value.table {
                    push_tracked_integer(output, "table:table", table);
                }
                if let Some(span) = &value.multi_deletion_spanned {
                    push_tracked_integer(output, "table:multi-deletion-spanned", span);
                }
                output.push('>');
                write_tracked_metadata(output, &value.metadata);
                write_tracked_cut_offs(output, &value.cut_offs);
                output.push_str("</table:deletion>");
            },
            Change::Movement(value) => {
                output.push_str("<table:movement");
                write_common_tracked_attributes(output, &value.metadata);
                output.push('>');
                write_tracked_range(output, "source-range-address", &value.source);
                write_tracked_range(output, "target-range-address", &value.target);
                write_tracked_metadata(output, &value.metadata);
                output.push_str("</table:movement>");
            },
            Change::CellContent(value) => {
                output.push_str("<table:cell-content-change");
                write_common_tracked_attributes(output, &value.metadata);
                output.push('>');
                write_tracked_address(output, "cell-address", &value.address);
                write_tracked_metadata(output, &value.metadata);
                output.push_str("<table:previous");
                if let Some(id) = &value.previous_change_id {
                    push_tracked_attr(output, "table:id", id);
                }
                output.push('>');
                write_tracked_cell(output, &value.previous);
                output.push_str("</table:previous></table:cell-content-change>");
            },
        }
    }
    output.push_str("</table:tracked-changes>");
}

fn write_common_tracked_attributes(output: &mut String, metadata: &Metadata) {
    push_tracked_attr(output, "table:id", &metadata.id);
    if metadata.acceptance != Acceptance::Pending {
        push_tracked_attr(
            output,
            "table:acceptance-state",
            acceptance_attribute_value(metadata.acceptance),
        );
    }
    if let Some(id) = &metadata.rejecting_change_id {
        push_tracked_attr(output, "table:rejecting-change-id", id);
    }
}

fn write_tracked_metadata(output: &mut String, metadata: &Metadata) {
    output.push_str("<office:change-info>");
    if let Some(value) = &metadata.info.creator {
        write_tracked_text(output, "dc:creator", value);
    }
    if let Some(value) = &metadata.info.date {
        write_tracked_text(output, "dc:date", value);
    }
    for value in &metadata.info.comments {
        write_tracked_text(output, "text:p", value);
    }
    output.push_str("</office:change-info>");
    if !metadata.dependencies.is_empty() {
        output.push_str("<table:dependencies>");
        for id in &metadata.dependencies {
            output.push_str("<table:dependency");
            push_tracked_attr(output, "table:id", id);
            output.push_str("/>");
        }
        output.push_str("</table:dependencies>");
    }
    if !metadata.deletions.is_empty() {
        output.push_str("<table:deletions>");
        for deletion in &metadata.deletions {
            match deletion {
                NestedDeletion::CellContent {
                    change_id,
                    address,
                    cell,
                } => {
                    output.push_str("<table:cell-content-deletion");
                    if let Some(id) = change_id {
                        push_tracked_attr(output, "table:id", id);
                    }
                    output.push('>');
                    if let Some(address) = address {
                        write_tracked_address(output, "cell-address", address);
                    }
                    if let Some(cell) = cell {
                        write_tracked_cell(output, cell);
                    }
                    output.push_str("</table:cell-content-deletion>");
                },
                NestedDeletion::Change { change_id } => {
                    output.push_str("<table:change-deletion");
                    if let Some(id) = change_id {
                        push_tracked_attr(output, "table:id", id);
                    }
                    output.push_str("/>");
                },
            }
        }
        output.push_str("</table:deletions>");
    }
}

fn write_tracked_cut_offs(output: &mut String, cut_offs: &[CutOff]) {
    if cut_offs.is_empty() {
        return;
    }
    output.push_str("<table:cut-offs>");
    for value in cut_offs {
        output.push_str(match value {
            CutOff::Insertion { .. } => "<table:insertion-cut-off",
            CutOff::MovementPoint { .. } | CutOff::MovementRange { .. } => {
                "<table:movement-cut-off"
            },
        });
        match value {
            CutOff::Insertion {
                change_id,
                position,
            } => {
                push_tracked_attr(output, "table:id", change_id);
                push_tracked_integer(output, "table:position", position);
            },
            CutOff::MovementPoint { position } => {
                push_tracked_integer(output, "table:position", position);
            },
            CutOff::MovementRange { start, end } => {
                push_tracked_integer(output, "table:start-position", start);
                push_tracked_integer(output, "table:end-position", end);
            },
        }
        output.push_str("/>");
    }
    output.push_str("</table:cut-offs>");
}

fn write_tracked_range(output: &mut String, name: &str, range: &RangeAddress) {
    output.push_str("<table:");
    output.push_str(name);
    match range {
        RangeAddress::Cell(value) => write_tracked_address_attrs(output, "", value),
        RangeAddress::Range { start, end } => {
            write_tracked_address_attrs(output, "start-", start);
            write_tracked_address_attrs(output, "end-", end);
        },
    }
    output.push_str("/>");
}

fn write_tracked_address(output: &mut String, name: &str, value: &CellAddress) {
    output.push_str("<table:");
    output.push_str(name);
    write_tracked_address_attrs(output, "", value);
    output.push_str("/>");
}

fn write_tracked_address_attrs(output: &mut String, prefix: &str, value: &CellAddress) {
    push_tracked_integer(output, &format!("table:{prefix}table"), &value.table);
    push_tracked_integer(output, &format!("table:{prefix}column"), &value.column);
    push_tracked_integer(output, &format!("table:{prefix}row"), &value.row);
}

fn write_tracked_cell(output: &mut String, cell: &Cell) {
    output.push_str("<table:change-track-table-cell");
    for (name, value) in [
        ("table:cell-address", cell.address.as_deref()),
        ("table:style-name", cell.style_name.as_deref()),
        ("table:formula", cell.formula.as_deref()),
    ] {
        if let Some(value) = value {
            push_tracked_attr(output, name, value);
        }
    }
    if cell.matrix_covered {
        push_tracked_attr(output, "table:matrix-covered", "true");
    }
    if let Some(value) = &cell.matrix_columns {
        push_tracked_attr(
            output,
            "table:number-matrix-columns-spanned",
            value.as_str(),
        );
    }
    if let Some(value) = &cell.matrix_rows {
        push_tracked_attr(output, "table:number-matrix-rows-spanned", value.as_str());
    }
    match &cell.value {
        CellValue::Empty => {},
        CellValue::Boolean(value) => {
            push_tracked_attr(output, "office:value-type", "boolean");
            push_tracked_attr(
                output,
                "office:boolean-value",
                if *value { "true" } else { "false" },
            );
        },
        CellValue::Number(value) => write_tracked_number(output, "float", *value),
        CellValue::Percentage(value) => write_tracked_number(output, "percentage", *value),
        CellValue::Currency { value, code } => {
            write_tracked_number(output, "currency", *value);
            push_tracked_attr(output, "office:currency", code);
        },
        CellValue::Date(value) => write_tracked_value(output, "date", "office:date-value", value),
        CellValue::Time(value) => write_tracked_value(output, "time", "office:time-value", value),
        CellValue::Text(value) => {
            write_tracked_value(output, "string", "office:string-value", value);
        },
        CellValue::Error(value) => {
            push_tracked_attr(output, "office:value-type", "error");
            if let Some(value) = value {
                push_tracked_attr(output, "office:string-value", value);
            }
        },
    }
    if cell.display_text.is_empty() {
        output.push_str("/>");
    } else {
        output.push('>');
        for paragraph in cell.display_text.split('\n') {
            write_tracked_text(output, "text:p", paragraph);
        }
        output.push_str("</table:change-track-table-cell>");
    }
}

fn write_tracked_number(output: &mut String, value_type: &str, value: f64) {
    push_tracked_attr(output, "office:value-type", value_type);
    push_tracked_attr(output, "office:value", &canonical_double(value));
}

fn canonical_double(value: f64) -> String {
    if value.is_nan() {
        "NaN".to_string()
    } else if value == f64::INFINITY {
        "INF".to_string()
    } else if value == f64::NEG_INFINITY {
        "-INF".to_string()
    } else {
        value.to_string()
    }
}

fn write_tracked_value(output: &mut String, value_type: &str, name: &str, value: &str) {
    push_tracked_attr(output, "office:value-type", value_type);
    push_tracked_attr(output, name, value);
}

fn tracked_dimension(value: Dimension) -> &'static str {
    match value {
        Dimension::Row => "row",
        Dimension::Column => "column",
        Dimension::Table => "table",
    }
}

fn push_tracked_integer(output: &mut String, name: &str, value: &Integer) {
    push_tracked_attr(output, name, value.as_str());
}

fn push_tracked_attr(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    escape_tracked_xml(output, value, true);
    output.push('"');
}

fn write_tracked_text(output: &mut String, name: &str, value: &str) {
    output.push('<');
    output.push_str(name);
    output.push('>');
    escape_tracked_xml(output, value, false);
    output.push_str("</");
    output.push_str(name);
    output.push('>');
}

fn escape_tracked_xml(output: &mut String, value: &str, attribute: bool) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' if attribute => output.push_str("&quot;"),
            '\'' if attribute => output.push_str("&apos;"),
            value => output.push(value),
        }
    }
}

fn optional_positive_nonzero(node: &Node, name: &str) -> Result<Option<PositiveInteger>> {
    attribute(node, Namespace::Table, name)
        .map(|value| PositiveInteger::parse_with_limits(value, &node.limits))
        .transpose()
}

fn common_attributes(extra: &[&'static str]) -> Vec<(Namespace, &'static str)> {
    let mut attributes = vec![
        (Namespace::Table, "id"),
        (Namespace::Table, "acceptance-state"),
        (Namespace::Table, "rejecting-change-id"),
    ];
    attributes.extend(extra.iter().map(|name| (Namespace::Table, *name)));
    attributes
}

fn parse_acceptance(value: &str) -> Result<Acceptance> {
    match collapse_atomic(value) {
        "accepted" => Ok(Acceptance::Accepted),
        "rejected" => Ok(Acceptance::Rejected),
        "pending" => Ok(Acceptance::Pending),
        _ => Err(Error::InvalidFormat(format!(
            "invalid table:acceptance-state '{value}'"
        ))),
    }
}

fn parse_dimension(value: &str) -> Result<Dimension> {
    match collapse_atomic(value) {
        "row" => Ok(Dimension::Row),
        "column" => Ok(Dimension::Column),
        "table" => Ok(Dimension::Table),
        _ => Err(Error::InvalidFormat(format!(
            "invalid tracked-change table:type '{value}'"
        ))),
    }
}

fn attribute<'a>(node: &'a Node, namespace: Namespace, local: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.local == local)
        .map(|attribute| attribute.value.as_str())
}

fn required_attribute<'a>(node: &'a Node, namespace: Namespace, local: &str) -> Result<&'a str> {
    attribute(node, namespace, local)
        .ok_or_else(|| Error::InvalidFormat(format!("{} requires attribute {local}", node.local)))
}

fn children(node: &Node) -> impl Iterator<Item = &Node> {
    node.content.iter().filter_map(|content| match content {
        Content::Node(node) => Some(node),
        Content::Text(_) => None,
    })
}

fn named_children<'a>(
    node: &'a Node,
    namespace: Namespace,
    local: &'a str,
) -> impl Iterator<Item = &'a Node> {
    children(node).filter(move |node| node.namespace == namespace && node.local == local)
}

fn optional_child<'a>(
    node: &'a Node,
    namespace: Namespace,
    local: &str,
) -> Result<Option<&'a Node>> {
    let mut matches =
        children(node).filter(|child| child.namespace == namespace && child.local == local);
    let result = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "duplicate child {local} in {}",
            node.local
        )));
    }
    Ok(result)
}

fn required_child<'a>(node: &'a Node, namespace: Namespace, local: &str) -> Result<&'a Node> {
    optional_child(node, namespace, local)?
        .ok_or_else(|| Error::InvalidFormat(format!("{} requires child {local}", node.local)))
}

fn reject_attributes(node: &Node, allowed: &[(Namespace, &str)]) -> Result<()> {
    for attribute in &node.attributes {
        if is_known(attribute.namespace)
            && !allowed.iter().any(|(namespace, local)| {
                attribute.namespace == *namespace && attribute.local == *local
            })
        {
            return Err(Error::InvalidFormat(format!(
                "unexpected attribute {} on {}",
                attribute.local, node.local
            )));
        }
    }
    Ok(())
}

fn reject_children(node: &Node, table_children: &[&str], change_info: bool) -> Result<()> {
    require_whitespace(node)?;
    for child in children(node) {
        let allowed =
            (change_info && child.namespace == Namespace::Office && child.local == "change-info")
                || (child.namespace == Namespace::Table
                    && table_children.iter().any(|name| child.local == *name));
        if !allowed {
            reject_known_child(child, &node.local)?;
        }
    }
    if change_info {
        required_child(node, Namespace::Office, "change-info")?;
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct ChildRule {
    namespace: Namespace,
    local: &'static str,
    minimum: usize,
    maximum: usize,
}

impl ChildRule {
    const fn one(namespace: Namespace, local: &'static str) -> Self {
        Self {
            namespace,
            local,
            minimum: 1,
            maximum: 1,
        }
    }

    const fn optional(namespace: Namespace, local: &'static str) -> Self {
        Self {
            namespace,
            local,
            minimum: 0,
            maximum: 1,
        }
    }

    const fn many(namespace: Namespace, local: &'static str) -> Self {
        Self {
            namespace,
            local,
            minimum: 0,
            maximum: usize::MAX,
        }
    }
}

fn require_child_sequence(node: &Node, rules: &[ChildRule]) -> Result<()> {
    require_whitespace(node)?;
    let known = children(node)
        .filter(|child| is_known(child.namespace))
        .collect::<Vec<_>>();
    let mut index = 0usize;
    for rule in rules {
        let mut count = 0usize;
        while index < known.len()
            && known[index].namespace == rule.namespace
            && known[index].local == rule.local
            && count < rule.maximum
        {
            index += 1;
            count += 1;
        }
        if count < rule.minimum {
            return invalid(format!(
                "{} requires ordered child {}",
                node.local, rule.local
            ));
        }
    }
    if let Some(child) = known.get(index) {
        return unexpected_child(child, &node.local);
    }
    Ok(())
}

fn reject_known_child(child: &Node, parent: &str) -> Result<()> {
    if is_known(child.namespace) {
        return unexpected_child(child, parent);
    }
    Ok(())
}

fn unexpected_child<T>(child: &Node, parent: &str) -> Result<T> {
    Err(Error::InvalidFormat(format!(
        "unexpected {} child in {parent}",
        child.local
    )))
}

fn require_whitespace(node: &Node) -> Result<()> {
    if node
        .content
        .iter()
        .any(|content| matches!(content, Content::Text(value) if !value.trim().is_empty()))
    {
        return Err(Error::InvalidFormat(format!(
            "unexpected character data in {}",
            node.local
        )));
    }
    Ok(())
}

fn text_content(node: &Node) -> Result<String> {
    let mut output = String::new();
    let mut stack = vec![(node, 0usize)];
    while let Some((current, index)) = stack.pop() {
        if index >= current.content.len() {
            continue;
        }
        stack.push((current, index + 1));
        match &current.content[index] {
            Content::Text(value) => {
                let length = output
                    .len()
                    .checked_add(value.len())
                    .ok_or_else(|| invalid_error("tracked-change text value size overflow"))?;
                if length > node.max_value_bytes {
                    return Err(Error::InvalidFormat(
                        "tracked-change text value exceeds the configured limit".to_string(),
                    ));
                }
                output.try_reserve(value.len()).map_err(allocation_error)?;
                output.push_str(value);
            },
            Content::Node(child) => stack.push((child, 0)),
        }
    }
    Ok(output)
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match collapse_atomic(value) {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid {name} boolean '{value}'"
        ))),
    }
}

fn parse_integer(node: &Node, value: &str, name: &str) -> Result<Integer> {
    Integer::parse_with_limits(value, &node.limits)
        .map_err(|_error| Error::InvalidFormat(format!("invalid {name} integer '{value}'")))
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let value = collapse_atomic(value).parse::<usize>().map_err(|_error| {
        Error::InvalidFormat(format!("invalid {name} positive integer '{value}'"))
    })?;
    if value == 0 {
        return Err(Error::InvalidFormat(format!("{name} must be positive")));
    }
    Ok(value)
}

fn parse_f64(value: &str, name: &str) -> Result<f64> {
    let value = collapse_atomic(value);
    match value {
        "INF" => return Ok(f64::INFINITY),
        "-INF" => return Ok(f64::NEG_INFINITY),
        "NaN" => return Ok(f64::NAN),
        _ => {},
    }
    let parsed = value
        .parse::<f64>()
        .map_err(|_error| Error::InvalidFormat(format!("invalid {name} number '{value}'")))?;
    if !parsed.is_finite() {
        return invalid(format!("invalid {name} xsd:double lexical value '{value}'"));
    }
    Ok(parsed)
}

fn collapse_atomic(value: &str) -> &str {
    value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n'))
}

fn ensure_nonempty(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        Err(Error::InvalidFormat(format!("{name} must not be empty")))
    } else {
        Ok(())
    }
}

fn ensure_limited_value(value: &str, kind: &str, limits: &Limits) -> Result<()> {
    if value.len() > limits.max_value_bytes() {
        Err(Error::InvalidFormat(format!(
            "tracked-change {kind} exceeds the configured value limit"
        )))
    } else {
        Ok(())
    }
}

fn append_limited_size(aggregate: &mut usize, amount: usize, limits: &Limits) -> Result<()> {
    *aggregate = aggregate
        .checked_add(amount)
        .ok_or_else(|| invalid_error("tracked-change aggregate size overflow"))?;
    if *aggregate > limits.max_aggregate_bytes() {
        return invalid("tracked-change XML exceeds the configured aggregate limit");
    }
    Ok(())
}

fn combine_resources(left: Resources, right: Resources, limits: &Limits) -> Result<Resources> {
    let combined = Resources {
        changes: left
            .changes
            .checked_add(right.changes)
            .ok_or_else(|| invalid_error("tracked-change count overflow"))?,
        nodes: left
            .nodes
            .checked_add(right.nodes)
            .ok_or_else(|| invalid_error("tracked-change node count overflow"))?,
        aggregate_bytes: left
            .aggregate_bytes
            .checked_add(right.aggregate_bytes)
            .ok_or_else(|| invalid_error("tracked-change aggregate size overflow"))?,
    };
    if combined.changes > limits.max_changes() {
        return invalid("spreadsheet tracked-change count exceeds the configured limit");
    }
    if combined.nodes > limits.max_nodes() {
        return invalid("spreadsheet tracked changes exceed the configured node limit");
    }
    if combined.aggregate_bytes > limits.max_aggregate_bytes() {
        return invalid("spreadsheet tracked changes exceed the configured aggregate limit");
    }
    Ok(combined)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> Result<Namespace> {
    match namespace {
        ResolveResult::Unbound => Ok(Namespace::None),
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NS => Ok(Namespace::Office),
        ResolveResult::Bound(value) if value.as_ref() == TABLE_NS => Ok(Namespace::Table),
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NS => Ok(Namespace::Text),
        ResolveResult::Bound(value) if value.as_ref() == DC_NS => Ok(Namespace::Dc),
        ResolveResult::Bound(value) if is_odf_reserved_namespace(value.as_ref()) => {
            Ok(Namespace::Reserved)
        },
        ResolveResult::Bound(_) => Ok(Namespace::Other),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound tracked-change namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn is_known(namespace: Namespace) -> bool {
    matches!(
        namespace,
        Namespace::Office
            | Namespace::Table
            | Namespace::Text
            | Namespace::Dc
            | Namespace::Reserved
    )
}

fn is_foreign(namespace: Namespace) -> bool {
    matches!(namespace, Namespace::None | Namespace::Other)
}

fn is_odf_reserved_namespace(namespace: &[u8]) -> bool {
    matches!(
        namespace,
        b"urn:oasis:names:tc:opendocument:xmlns:animation:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:config:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:database:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:form:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:of:1.2"
            | b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:script:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:style:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
            | b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
            | b"http://docs.oasis-open.org/ns/office/1.2/meta/odf#"
            | b"http://docs.oasis-open.org/ns/office/1.2/meta/pkg#"
            | b"http://www.w3.org/1998/Math/MathML"
            | b"http://www.w3.org/1999/xhtml"
            | b"http://www.w3.org/1999/xlink"
            | b"http://www.w3.org/2002/xforms"
            | b"http://www.w3.org/2003/g/data-view#"
    )
}

fn reject_spoofed_name(namespace: Namespace, local: &str) -> Result<()> {
    const TABLE_NAMES: &[&str] = &[
        "tracked-changes",
        "insertion",
        "deletion",
        "movement",
        "cell-content-change",
        "dependencies",
        "dependency",
        "deletions",
        "cell-content-deletion",
        "change-deletion",
        "cut-offs",
        "insertion-cut-off",
        "movement-cut-off",
        "source-range-address",
        "target-range-address",
        "cell-address",
        "previous",
        "change-track-table-cell",
    ];
    if TABLE_NAMES.contains(&local) && namespace != Namespace::Table && !is_foreign(namespace)
        || local == "change-info" && namespace != Namespace::Office && !is_foreign(namespace)
        || matches!(local, "creator" | "date")
            && namespace != Namespace::Dc
            && !is_foreign(namespace)
    {
        return Err(Error::InvalidFormat(format!(
            "tracked-change vocabulary element '{local}' uses the wrong namespace"
        )));
    }
    Ok(())
}

fn reject_spoofed_attribute(namespace: Namespace, local: &str) -> Result<()> {
    const TABLE_ATTRIBUTES: &[&str] = &[
        "track-changes",
        "id",
        "acceptance-state",
        "rejecting-change-id",
        "type",
        "position",
        "count",
        "table",
        "multi-deletion-spanned",
        "start-position",
        "end-position",
        "column",
        "row",
        "start-table",
        "start-column",
        "start-row",
        "end-table",
        "end-column",
        "end-row",
        "cell-address",
        "matrix-covered",
        "formula",
        "number-matrix-columns-spanned",
        "number-matrix-rows-spanned",
    ];
    const OFFICE_ATTRIBUTES: &[&str] = &[
        "value-type",
        "value",
        "boolean-value",
        "currency",
        "date-value",
        "string-value",
        "time-value",
    ];
    let wrong_reserved = |expected| namespace != expected && !is_foreign(namespace);
    if TABLE_ATTRIBUTES.contains(&local) && wrong_reserved(Namespace::Table)
        || OFFICE_ATTRIBUTES.contains(&local) && wrong_reserved(Namespace::Office)
        || local == "c" && wrong_reserved(Namespace::Text)
    {
        return invalid(format!(
            "tracked-change vocabulary attribute '{local}' uses the wrong namespace"
        ));
    }
    Ok(())
}

fn decode_name(value: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_error| Error::InvalidFormat(format!("invalid UTF-8 tracked-change {kind} name")))
}

fn resolve_reference(name: &str) -> Result<String> {
    let builtin = match name {
        "amp" => Some('&'),
        "lt" => Some('<'),
        "gt" => Some('>'),
        "quot" => Some('"'),
        "apos" => Some('\''),
        _ => None,
    };
    let value = if let Some(value) = builtin {
        value
    } else {
        if let Some(hex) = name.strip_prefix("#x") {
            u32::from_str_radix(hex, 16).ok()
        } else if let Some(decimal) = name.strip_prefix('#') {
            decimal.parse::<u32>().ok()
        } else {
            None
        }
        .filter(|value| {
            matches!(
                value,
                0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
            )
        })
        .and_then(char::from_u32)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("unsupported tracked-change entity '&{name};'"))
        })?
    };
    Ok(value.to_string())
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("invalid spreadsheet tracked-change XML: {error}"))
}
