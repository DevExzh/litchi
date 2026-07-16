//! Typed, inert spreadsheet change-tracking metadata.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{NamespaceResolver, ResolveResult},
    reader::NsReader,
};
use std::{collections::HashSet, num::NonZeroUsize};

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DC_NS: &[u8] = b"http://purl.org/dc/elements/1.1/";
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 1_000_000;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// Whether a recorded spreadsheet change is pending, accepted, or rejected.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum SpreadsheetChangeAcceptance {
    Accepted,
    Rejected,
    #[default]
    Pending,
}

/// The structural unit affected by a row, column, or table insertion/deletion.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SpreadsheetChangeDimension {
    Row,
    Column,
    Table,
}

/// Author, date, and comments stored for one change.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SpreadsheetChangeInfo {
    pub creator: Option<String>,
    pub date: Option<String>,
    pub comments: Vec<String>,
}

/// Integer table/row/column coordinates used by the change-tracking vocabulary.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SpreadsheetTrackedCellAddress {
    pub table: i64,
    pub column: i64,
    pub row: i64,
}

/// A single cell or rectangular source/target range used by a movement.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SpreadsheetTrackedRangeAddress {
    Cell(SpreadsheetTrackedCellAddress),
    Range {
        start: SpreadsheetTrackedCellAddress,
        end: SpreadsheetTrackedCellAddress,
    },
}

/// Typed scalar value preserved by `table:change-track-table-cell`.
#[derive(Clone, Debug, PartialEq)]
pub enum SpreadsheetTrackedCellValue {
    Empty,
    Boolean(bool),
    Number(f64),
    Percentage(f64),
    Currency { value: f64, code: String },
    Date(String),
    Time(String),
    Text(String),
}

/// Former cell state embedded in a tracked change.
#[derive(Clone, Debug, PartialEq)]
pub struct SpreadsheetTrackedCell {
    pub address: Option<String>,
    pub matrix_covered: bool,
    pub formula: Option<String>,
    pub matrix_columns: Option<NonZeroUsize>,
    pub matrix_rows: Option<NonZeroUsize>,
    pub value: SpreadsheetTrackedCellValue,
    pub display_text: String,
}

/// A deletion nested inside another tracked change.
#[derive(Clone, Debug, PartialEq)]
pub enum SpreadsheetNestedDeletion {
    CellContent {
        change_id: Option<String>,
        address: Option<SpreadsheetTrackedCellAddress>,
        cell: Option<SpreadsheetTrackedCell>,
    },
    Change { change_id: Option<String> },
}

/// A location removed from a previously tracked insertion or movement.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SpreadsheetChangeCutOff {
    Insertion { change_id: String, position: i64 },
    MovementPoint { position: i64 },
    MovementRange { start: i64, end: i64 },
}

/// Metadata common to every top-level spreadsheet change.
#[derive(Clone, Debug, PartialEq)]
pub struct SpreadsheetChangeMetadata {
    pub id: String,
    pub acceptance: SpreadsheetChangeAcceptance,
    pub rejecting_change_id: Option<String>,
    pub info: SpreadsheetChangeInfo,
    pub dependencies: Vec<String>,
    pub deletions: Vec<SpreadsheetNestedDeletion>,
}

/// A tracked row, column, or table insertion.
#[derive(Clone, Debug, PartialEq)]
pub struct SpreadsheetInsertion {
    pub metadata: SpreadsheetChangeMetadata,
    pub dimension: SpreadsheetChangeDimension,
    pub position: i64,
    pub count: NonZeroUsize,
    pub table: Option<i64>,
}

/// A tracked row, column, or table deletion.
#[derive(Clone, Debug, PartialEq)]
pub struct SpreadsheetDeletion {
    pub metadata: SpreadsheetChangeMetadata,
    pub dimension: SpreadsheetChangeDimension,
    pub position: i64,
    pub table: Option<i64>,
    pub multi_deletion_spanned: Option<i64>,
    pub cut_offs: Vec<SpreadsheetChangeCutOff>,
}

/// A tracked cell or range movement.
#[derive(Clone, Debug, PartialEq)]
pub struct SpreadsheetMovement {
    pub metadata: SpreadsheetChangeMetadata,
    pub source: SpreadsheetTrackedRangeAddress,
    pub target: SpreadsheetTrackedRangeAddress,
}

/// A tracked replacement of one cell's content.
#[derive(Clone, Debug, PartialEq)]
pub struct SpreadsheetCellContentChange {
    pub metadata: SpreadsheetChangeMetadata,
    pub address: SpreadsheetTrackedCellAddress,
    pub previous_change_id: Option<String>,
    pub previous: SpreadsheetTrackedCell,
}

/// One top-level spreadsheet change in document order.
#[derive(Clone, Debug, PartialEq)]
pub enum SpreadsheetTrackedChange {
    Insertion(SpreadsheetInsertion),
    Deletion(SpreadsheetDeletion),
    Movement(SpreadsheetMovement),
    CellContent(SpreadsheetCellContentChange),
}

impl SpreadsheetTrackedChange {
    pub fn metadata(&self) -> &SpreadsheetChangeMetadata {
        match self {
            Self::Insertion(value) => &value.metadata,
            Self::Deletion(value) => &value.metadata,
            Self::Movement(value) => &value.metadata,
            Self::CellContent(value) => &value.metadata,
        }
    }
}

/// Spreadsheet-wide tracked-change state and ordered change records.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct SpreadsheetTrackedChanges {
    pub enabled: bool,
    pub changes: Vec<SpreadsheetTrackedChange>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Namespace {
    None,
    Office,
    Table,
    Text,
    Dc,
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
}

pub(crate) fn parse_tracked_changes(xml: &str) -> Result<Option<SpreadsheetTrackedChanges>> {
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
        let mut consumed = false;

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
                )?;
                result = Some(parse_root(&root)?);
                consumed = true;
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
                )?;
                result = Some(parse_root(&root)?);
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in tracked changes"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }

        if is_start && !consumed {
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
) -> Result<Node> {
    let root = create_node(
        reader.resolver(),
        reader.decoder(),
        namespace,
        start,
        aggregate,
        nodes,
    )?;
    let mut stack = vec![root];
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
                )?;
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
                )?;
                add_semantic_leaf_text(&mut node, aggregate)?;
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
                append_text(stack.last_mut().expect("active node"), value, aggregate)?;
            },
            Event::CData(text) => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid tracked-change CDATA: {error}"))
                })?;
                append_text(
                    stack.last_mut().expect("active node"),
                    value.into_owned(),
                    aggregate,
                )?;
            },
            Event::GeneralRef(reference) => {
                let name = std::str::from_utf8(reference.as_ref()).map_err(|_| {
                    Error::InvalidFormat("invalid tracked-change entity reference".to_string())
                })?;
                append_text(
                    stack.last_mut().expect("active node"),
                    resolve_reference(name)?,
                    aggregate,
                )?;
            },
            Event::End(_) => {
                let mut node = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("tracked-change XML stack underflow".to_string())
                })?;
                add_semantic_leaf_text(&mut node, aggregate)?;
                if let Some(parent) = stack.last_mut() {
                    parent.content.push(Content::Node(node));
                } else {
                    return Ok(node);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in tracked changes"
                        .to_string(),
                ));
            },
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
) -> Result<Node> {
    *nodes = nodes.checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("tracked-change node count overflow".to_string())
    })?;
    if *nodes > MAX_NODES {
        return Err(Error::InvalidFormat(format!(
            "tracked changes exceed {MAX_NODES} XML nodes"
        )));
    }
    let local = decode_name(start.local_name().as_ref(), "element")?;
    reject_spoofed_name(namespace, &local)?;
    append_size(aggregate, local.len())?;
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
        ensure_value_bound(&value, "attribute")?;
        append_size(aggregate, local_name.len().saturating_add(value.len()))?;
        if attributes.iter().any(|existing: &Attribute| {
            existing.namespace == attribute_namespace && existing.local == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded tracked-change attribute {local_name}"
            )));
        }
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
    })
}

fn append_text(node: &mut Node, value: String, aggregate: &mut usize) -> Result<()> {
    append_size(aggregate, value.len())?;
    if let Some(Content::Text(existing)) = node.content.last_mut() {
        if existing.len().saturating_add(value.len()) > MAX_VALUE_BYTES {
            return Err(Error::InvalidFormat(
                "tracked-change text exceeds 64 KiB".to_string(),
            ));
        }
        existing.push_str(&value);
    } else {
        ensure_value_bound(&value, "text")?;
        node.content.push(Content::Text(value));
    }
    Ok(())
}

fn add_semantic_leaf_text(node: &mut Node, aggregate: &mut usize) -> Result<()> {
    if node.namespace == Namespace::Text && node.local == "s" {
        if !node.content.is_empty() {
            return Err(Error::InvalidFormat("text:s must be empty".to_string()));
        }
        let count = attribute(node, Namespace::Text, "c")
            .map(|value| parse_positive(value, "text:c"))
            .transpose()?
            .unwrap_or(1);
        if count > MAX_VALUE_BYTES {
            return Err(Error::InvalidFormat("text:s count exceeds 64 KiB".to_string()));
        }
        append_size(aggregate, count)?;
        node.content.push(Content::Text(" ".repeat(count)));
    } else if node.namespace == Namespace::Text && node.local == "tab" {
        if !node.content.is_empty() {
            return Err(Error::InvalidFormat("text:tab must be empty".to_string()));
        }
        append_size(aggregate, 1)?;
        node.content.push(Content::Text("\t".to_string()));
    } else if node.namespace == Namespace::Text && node.local == "line-break" {
        if !node.content.is_empty() {
            return Err(Error::InvalidFormat(
                "text:line-break must be empty".to_string(),
            ));
        }
        append_size(aggregate, 1)?;
        node.content.push(Content::Text("\n".to_string()));
    }
    Ok(())
}

fn parse_root(root: &Node) -> Result<SpreadsheetTrackedChanges> {
    reject_attributes(root, &[(Namespace::Table, "track-changes")])?;
    require_whitespace(root)?;
    let enabled = attribute(root, Namespace::Table, "track-changes")
        .map(|value| parse_bool(value, "table:track-changes"))
        .transpose()?
        .unwrap_or(false);
    let mut changes = Vec::new();
    for child in children(root) {
        if child.namespace != Namespace::Table {
            reject_known_child(child, "table:tracked-changes")?;
            continue;
        }
        let change = match child.local.as_str() {
            "insertion" => SpreadsheetTrackedChange::Insertion(parse_insertion(child)?),
            "deletion" => SpreadsheetTrackedChange::Deletion(parse_deletion(child)?),
            "movement" => SpreadsheetTrackedChange::Movement(parse_movement(child)?),
            "cell-content-change" => {
                SpreadsheetTrackedChange::CellContent(parse_cell_content_change(child)?)
            },
            _ => return unexpected_child(child, "table:tracked-changes"),
        };
        changes.push(change);
    }
    let mut ids = HashSet::with_capacity(changes.len());
    for change in &changes {
        let id = change.metadata().id.as_str();
        if !ids.insert(id) {
            return Err(Error::InvalidFormat(format!(
                "duplicate spreadsheet tracked-change id '{id}'"
            )));
        }
    }
    Ok(SpreadsheetTrackedChanges { enabled, changes })
}

fn parse_insertion(node: &Node) -> Result<SpreadsheetInsertion> {
    reject_attributes(
        node,
        &common_attributes(&["type", "position", "count", "table"]),
    )?;
    reject_children(node, &["dependencies", "deletions"], true)?;
    Ok(SpreadsheetInsertion {
        metadata: parse_metadata(node)?,
        dimension: parse_dimension(required_attribute(node, Namespace::Table, "type")?)?,
        position: parse_i64(
            required_attribute(node, Namespace::Table, "position")?,
            "table:position",
        )?,
        count: NonZeroUsize::new(
            attribute(node, Namespace::Table, "count")
                .map(|value| parse_positive(value, "table:count"))
                .transpose()?
                .unwrap_or(1),
        )
        .expect("default and parser are positive"),
        table: attribute(node, Namespace::Table, "table")
            .map(|value| parse_i64(value, "table:table"))
            .transpose()?,
    })
}

fn parse_deletion(node: &Node) -> Result<SpreadsheetDeletion> {
    reject_attributes(
        node,
        &common_attributes(&["type", "position", "table", "multi-deletion-spanned"]),
    )?;
    reject_children(node, &["dependencies", "deletions", "cut-offs"], true)?;
    let cut_offs = optional_child(node, Namespace::Table, "cut-offs")?
        .map(parse_cut_offs)
        .transpose()?
        .unwrap_or_default();
    Ok(SpreadsheetDeletion {
        metadata: parse_metadata(node)?,
        dimension: parse_dimension(required_attribute(node, Namespace::Table, "type")?)?,
        position: parse_i64(
            required_attribute(node, Namespace::Table, "position")?,
            "table:position",
        )?,
        table: attribute(node, Namespace::Table, "table")
            .map(|value| parse_i64(value, "table:table"))
            .transpose()?,
        multi_deletion_spanned: attribute(node, Namespace::Table, "multi-deletion-spanned")
            .map(|value| parse_i64(value, "table:multi-deletion-spanned"))
            .transpose()?,
        cut_offs,
    })
}

fn parse_movement(node: &Node) -> Result<SpreadsheetMovement> {
    reject_attributes(node, &common_attributes(&[]))?;
    reject_children(
        node,
        &[
            "dependencies",
            "deletions",
            "source-range-address",
            "target-range-address",
        ],
        true,
    )?;
    Ok(SpreadsheetMovement {
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

fn parse_cell_content_change(node: &Node) -> Result<SpreadsheetCellContentChange> {
    reject_attributes(node, &common_attributes(&[]))?;
    reject_children(
        node,
        &["dependencies", "deletions", "cell-address", "previous"],
        true,
    )?;
    let previous = required_child(node, Namespace::Table, "previous")?;
    reject_attributes(previous, &[(Namespace::Table, "id")])?;
    reject_children(previous, &["change-track-table-cell"], false)?;
    let previous_cell = required_child(
        previous,
        Namespace::Table,
        "change-track-table-cell",
    )?;
    Ok(SpreadsheetCellContentChange {
        metadata: parse_metadata(node)?,
        address: parse_cell_address(required_child(
            node,
            Namespace::Table,
            "cell-address",
        )?)?,
        previous_change_id: attribute(previous, Namespace::Table, "id").map(str::to_string),
        previous: parse_tracked_cell(previous_cell)?,
    })
}

fn parse_metadata(node: &Node) -> Result<SpreadsheetChangeMetadata> {
    let id = required_attribute(node, Namespace::Table, "id")?.to_string();
    ensure_nonempty(&id, "table:id")?;
    let info = parse_change_info(required_child(
        node,
        Namespace::Office,
        "change-info",
    )?)?;
    let dependencies = optional_child(node, Namespace::Table, "dependencies")?
        .map(parse_dependencies)
        .transpose()?
        .unwrap_or_default();
    let deletions = optional_child(node, Namespace::Table, "deletions")?
        .map(parse_nested_deletions)
        .transpose()?
        .unwrap_or_default();
    Ok(SpreadsheetChangeMetadata {
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

fn parse_change_info(node: &Node) -> Result<SpreadsheetChangeInfo> {
    reject_attributes(node, &[])?;
    require_whitespace(node)?;
    let creator = optional_child(node, Namespace::Dc, "creator")?
        .map(text_content)
        .transpose()?;
    let date = optional_child(node, Namespace::Dc, "date")?
        .map(text_content)
        .transpose()?;
    let comments = named_children(node, Namespace::Text, "p")
        .map(text_content)
        .collect::<Result<Vec<_>>>()?;
    for child in children(node) {
        if !matches!(
            (child.namespace, child.local.as_str()),
            (Namespace::Dc, "creator") | (Namespace::Dc, "date") | (Namespace::Text, "p")
        ) {
            reject_known_child(child, "office:change-info")?;
        }
    }
    Ok(SpreadsheetChangeInfo {
        creator,
        date,
        comments,
    })
}

fn parse_dependencies(node: &Node) -> Result<Vec<String>> {
    reject_attributes(node, &[])?;
    require_whitespace(node)?;
    let dependencies = named_children(node, Namespace::Table, "dependency")
        .map(|child| {
            reject_attributes(child, &[(Namespace::Table, "id")])?;
            reject_children(child, &[], false)?;
            let id = required_attribute(child, Namespace::Table, "id")?.to_string();
            ensure_nonempty(&id, "table:dependency table:id")?;
            Ok(id)
        })
        .collect::<Result<Vec<_>>>()?;
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

fn parse_nested_deletions(node: &Node) -> Result<Vec<SpreadsheetNestedDeletion>> {
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
                reject_children(child, &["cell-address", "change-track-table-cell"], false)?;
                result.push(SpreadsheetNestedDeletion::CellContent {
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
                result.push(SpreadsheetNestedDeletion::Change {
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

fn parse_cut_offs(node: &Node) -> Result<Vec<SpreadsheetChangeCutOff>> {
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
                result.push(SpreadsheetChangeCutOff::Insertion {
                    change_id: required_attribute(child, Namespace::Table, "id")?.to_string(),
                    position: parse_i64(
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
                    (Some(value), None, None) => SpreadsheetChangeCutOff::MovementPoint {
                        position: parse_i64(value, "table:position")?,
                    },
                    (None, Some(start), Some(end)) => {
                        let start = parse_i64(start, "table:start-position")?;
                        let end = parse_i64(end, "table:end-position")?;
                        if start >= end {
                            return Err(Error::InvalidFormat(
                                "movement cut-off start must precede end".to_string(),
                            ));
                        }
                        SpreadsheetChangeCutOff::MovementRange { start, end }
                    },
                    _ => {
                        return Err(Error::InvalidFormat(
                            "movement cut-off requires position or start/end positions"
                                .to_string(),
                        ));
                    },
                };
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

fn parse_range(node: &Node) -> Result<SpreadsheetTrackedRangeAddress> {
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
    let cell = ["table", "column", "row"]
        .map(|name| attribute(node, Namespace::Table, name));
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
        return Ok(SpreadsheetTrackedRangeAddress::Cell(
            SpreadsheetTrackedCellAddress {
                table: parse_i64(cell[0].expect("present"), "table:table")?,
                column: parse_i64(cell[1].expect("present"), "table:column")?,
                row: parse_i64(cell[2].expect("present"), "table:row")?,
            },
        ));
    }
    if cell.iter().all(Option::is_none) && range.iter().all(Option::is_some) {
        return Ok(SpreadsheetTrackedRangeAddress::Range {
            start: SpreadsheetTrackedCellAddress {
                table: parse_i64(range[0].expect("present"), "table:start-table")?,
                column: parse_i64(range[1].expect("present"), "table:start-column")?,
                row: parse_i64(range[2].expect("present"), "table:start-row")?,
            },
            end: SpreadsheetTrackedCellAddress {
                table: parse_i64(range[3].expect("present"), "table:end-table")?,
                column: parse_i64(range[4].expect("present"), "table:end-column")?,
                row: parse_i64(range[5].expect("present"), "table:end-row")?,
            },
        });
    }
    Err(Error::InvalidFormat(
        "tracked range requires either cell or complete start/end coordinates".to_string(),
    ))
}

fn parse_cell_address(node: &Node) -> Result<SpreadsheetTrackedCellAddress> {
    reject_attributes(
        node,
        &[
            (Namespace::Table, "table"),
            (Namespace::Table, "column"),
            (Namespace::Table, "row"),
        ],
    )?;
    reject_children(node, &[], false)?;
    Ok(SpreadsheetTrackedCellAddress {
        table: parse_i64(
            required_attribute(node, Namespace::Table, "table")?,
            "table:table",
        )?,
        column: parse_i64(
            required_attribute(node, Namespace::Table, "column")?,
            "table:column",
        )?,
        row: parse_i64(
            required_attribute(node, Namespace::Table, "row")?,
            "table:row",
        )?,
    })
}

fn parse_tracked_cell(node: &Node) -> Result<SpreadsheetTrackedCell> {
    reject_attributes(
        node,
        &[
            (Namespace::Table, "cell-address"),
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
    let paragraphs = named_children(node, Namespace::Text, "p")
        .map(text_content)
        .collect::<Result<Vec<_>>>()?;
    let display_text = paragraphs.join("\n");
    let value_type = attribute(node, Namespace::Office, "value-type");
    let has_value_attribute = node.attributes.iter().any(|attribute| {
        attribute.namespace == Namespace::Office && attribute.local != "value-type"
    });
    if value_type.is_none() && has_value_attribute {
        return Err(Error::InvalidFormat(
            "tracked cell value attributes require office:value-type".to_string(),
        ));
    }
    let value = match value_type {
        None => SpreadsheetTrackedCellValue::Empty,
        Some("boolean") => SpreadsheetTrackedCellValue::Boolean(parse_bool(
            required_attribute(node, Namespace::Office, "boolean-value")?,
            "office:boolean-value",
        )?),
        Some("float") => SpreadsheetTrackedCellValue::Number(parse_f64(
            required_attribute(node, Namespace::Office, "value")?,
            "office:value",
        )?),
        Some("percentage") => SpreadsheetTrackedCellValue::Percentage(parse_f64(
            required_attribute(node, Namespace::Office, "value")?,
            "office:value",
        )?),
        Some("currency") => SpreadsheetTrackedCellValue::Currency {
            value: parse_f64(
                required_attribute(node, Namespace::Office, "value")?,
                "office:value",
            )?,
            code: required_attribute(node, Namespace::Office, "currency")?.to_string(),
        },
        Some("date") => SpreadsheetTrackedCellValue::Date(
            required_attribute(node, Namespace::Office, "date-value")?.to_string(),
        ),
        Some("time") => SpreadsheetTrackedCellValue::Time(
            required_attribute(node, Namespace::Office, "time-value")?.to_string(),
        ),
        Some("string") => SpreadsheetTrackedCellValue::Text(
            attribute(node, Namespace::Office, "string-value")
                .unwrap_or(&display_text)
                .to_string(),
        ),
        Some(other) => {
            return Err(Error::InvalidFormat(format!(
                "unsupported tracked cell office:value-type '{other}'"
            )));
        },
    };
    Ok(SpreadsheetTrackedCell {
        address: attribute(node, Namespace::Table, "cell-address").map(str::to_string),
        matrix_covered: attribute(node, Namespace::Table, "matrix-covered")
            .map(|value| parse_bool(value, "table:matrix-covered"))
            .transpose()?
            .unwrap_or(false),
        formula: attribute(node, Namespace::Table, "formula").map(str::to_string),
        matrix_columns: optional_positive_nonzero(
            node,
            "number-matrix-columns-spanned",
        )?,
        matrix_rows: optional_positive_nonzero(node, "number-matrix-rows-spanned")?,
        value,
        display_text,
    })
}

fn optional_positive_nonzero(node: &Node, name: &str) -> Result<Option<NonZeroUsize>> {
    attribute(node, Namespace::Table, name)
        .map(|value| {
            NonZeroUsize::new(parse_positive(value, &format!("table:{name}"))?).ok_or_else(|| {
                Error::InvalidFormat(format!("table:{name} must be positive"))
            })
        })
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

fn parse_acceptance(value: &str) -> Result<SpreadsheetChangeAcceptance> {
    match value {
        "accepted" => Ok(SpreadsheetChangeAcceptance::Accepted),
        "rejected" => Ok(SpreadsheetChangeAcceptance::Rejected),
        "pending" => Ok(SpreadsheetChangeAcceptance::Pending),
        _ => Err(Error::InvalidFormat(format!(
            "invalid table:acceptance-state '{value}'"
        ))),
    }
}

fn parse_dimension(value: &str) -> Result<SpreadsheetChangeDimension> {
    match value {
        "row" => Ok(SpreadsheetChangeDimension::Row),
        "column" => Ok(SpreadsheetChangeDimension::Column),
        "table" => Ok(SpreadsheetChangeDimension::Table),
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

fn required_attribute<'a>(
    node: &'a Node,
    namespace: Namespace,
    local: &str,
) -> Result<&'a str> {
    attribute(node, namespace, local).ok_or_else(|| {
        Error::InvalidFormat(format!("{} requires attribute {local}", node.local))
    })
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
    let mut matches = children(node).filter(|child| {
        child.namespace == namespace && child.local == local
    });
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
    optional_child(node, namespace, local)?.ok_or_else(|| {
        Error::InvalidFormat(format!("{} requires child {local}", node.local))
    })
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
        let allowed = (change_info
            && child.namespace == Namespace::Office
            && child.local == "change-info")
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
    if node.content.iter().any(|content| {
        matches!(content, Content::Text(value) if !value.trim().is_empty())
    }) {
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
                if output.len().saturating_add(value.len()) > MAX_VALUE_BYTES {
                    return Err(Error::InvalidFormat(
                        "tracked-change text value exceeds 64 KiB".to_string(),
                    ));
                }
                output.push_str(value);
            },
            Content::Node(child) => stack.push((child, 0)),
        }
    }
    Ok(output)
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid {name} boolean '{value}'"
        ))),
    }
}

fn parse_i64(value: &str, name: &str) -> Result<i64> {
    value
        .parse::<i64>()
        .map_err(|_| Error::InvalidFormat(format!("invalid {name} integer '{value}'")))
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let value = value
        .parse::<usize>()
        .map_err(|_| Error::InvalidFormat(format!("invalid {name} positive integer '{value}'")))?;
    if value == 0 {
        return Err(Error::InvalidFormat(format!("{name} must be positive")));
    }
    Ok(value)
}

fn parse_f64(value: &str, name: &str) -> Result<f64> {
    let value = value
        .parse::<f64>()
        .map_err(|_| Error::InvalidFormat(format!("invalid {name} number '{value}'")))?;
    if !value.is_finite() {
        return Err(Error::InvalidFormat(format!("{name} must be finite")));
    }
    Ok(value)
}

fn ensure_nonempty(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        Err(Error::InvalidFormat(format!("{name} must not be empty")))
    } else {
        Ok(())
    }
}

fn ensure_value_bound(value: &str, kind: &str) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES {
        Err(Error::InvalidFormat(format!(
            "tracked-change {kind} exceeds 64 KiB"
        )))
    } else {
        Ok(())
    }
}

fn append_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    *aggregate = aggregate.checked_add(amount).ok_or_else(|| {
        Error::InvalidFormat("tracked-change aggregate size overflow".to_string())
    })?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return Err(Error::InvalidFormat(
            "tracked-change metadata exceeds 16 MiB".to_string(),
        ));
    }
    Ok(())
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> Result<Namespace> {
    match namespace {
        ResolveResult::Unbound => Ok(Namespace::None),
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NS => Ok(Namespace::Office),
        ResolveResult::Bound(value) if value.as_ref() == TABLE_NS => Ok(Namespace::Table),
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NS => Ok(Namespace::Text),
        ResolveResult::Bound(value) if value.as_ref() == DC_NS => Ok(Namespace::Dc),
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
        Namespace::Office | Namespace::Table | Namespace::Text | Namespace::Dc
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
    if TABLE_NAMES.contains(&local) && namespace != Namespace::Table
        || local == "change-info" && namespace != Namespace::Office
        || matches!(local, "creator" | "date") && namespace != Namespace::Dc
    {
        return Err(Error::InvalidFormat(format!(
            "tracked-change vocabulary element '{local}' uses the wrong namespace"
        )));
    }
    Ok(())
}

fn decode_name(value: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 tracked-change {kind} name")))
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
        let scalar = if let Some(hex) = name.strip_prefix("#x") {
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
            Error::InvalidFormat(format!(
                "unsupported tracked-change entity '&{name};'"
            ))
        })?;
        scalar
    };
    Ok(value.to_string())
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("invalid spreadsheet tracked-change XML: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:body><office:spreadsheet>"#;
    const SUFFIX: &str = "</office:spreadsheet></office:body></office:document-content>";

    fn parse(fragment: &str) -> Result<Option<SpreadsheetTrackedChanges>> {
        parse_tracked_changes(&format!("{PREFIX}{fragment}{SUFFIX}"))
    }

    #[test]
    fn parses_complete_spreadsheet_change_graph() {
        let xml = r#"<table:tracked-changes table:track-changes="true">
          <table:insertion table:id="i1" table:type="row" table:position="2" table:count="3" table:table="0"><office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17</dc:date><text:p>insert</text:p></office:change-info><table:dependencies><table:dependency table:id="m1"/></table:dependencies></table:insertion>
          <table:deletion table:id="d1" table:type="column" table:position="4" table:acceptance-state="rejected" table:multi-deletion-spanned="2"><office:change-info/><table:deletions><table:change-deletion table:id="i1"/><table:cell-content-deletion><table:cell-address table:table="0" table:column="4" table:row="1"/><table:change-track-table-cell office:value-type="float" office:value="12.5" table:formula="of:=1+1"><text:p>12.5</text:p></table:change-track-table-cell></table:cell-content-deletion></table:deletions><table:cut-offs><table:insertion-cut-off table:id="i1" table:position="1"/><table:movement-cut-off table:start-position="2" table:end-position="5"/></table:cut-offs></table:deletion>
          <table:movement table:id="m1"><table:source-range-address table:start-table="0" table:start-column="1" table:start-row="2" table:end-table="0" table:end-column="3" table:end-row="4"/><table:target-range-address table:table="0" table:column="5" table:row="6"/><office:change-info><dc:creator>Mover</dc:creator></office:change-info></table:movement>
          <table:cell-content-change table:id="c1" table:acceptance-state="accepted"><table:cell-address table:table="0" table:column="1" table:row="2"/><office:change-info><text:p>A &amp;&#x20;B</text:p></office:change-info><table:previous table:id="old"><table:change-track-table-cell office:value-type="string" office:string-value="old" table:matrix-covered="false"><text:p>old</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
        </table:tracked-changes>"#;
        let tracked = parse(xml).unwrap().unwrap();
        assert!(tracked.enabled);
        assert_eq!(tracked.changes.len(), 4);
        let SpreadsheetTrackedChange::Insertion(insertion) = &tracked.changes[0] else {
            panic!("expected insertion")
        };
        assert_eq!(insertion.count.get(), 3);
        assert_eq!(insertion.metadata.info.comments, ["insert"]);
        let SpreadsheetTrackedChange::Deletion(deletion) = &tracked.changes[1] else {
            panic!("expected deletion")
        };
        assert_eq!(deletion.cut_offs.len(), 2);
        assert_eq!(deletion.metadata.deletions.len(), 2);
        let SpreadsheetTrackedChange::Movement(movement) = &tracked.changes[2] else {
            panic!("expected movement")
        };
        assert!(matches!(movement.source, SpreadsheetTrackedRangeAddress::Range { .. }));
        let SpreadsheetTrackedChange::CellContent(change) = &tracked.changes[3] else {
            panic!("expected cell change")
        };
        assert_eq!(change.metadata.acceptance, SpreadsheetChangeAcceptance::Accepted);
        assert_eq!(change.metadata.info.comments, ["A & B"]);
        assert_eq!(change.previous.value, SpreadsheetTrackedCellValue::Text("old".into()));
    }

    #[test]
    fn applies_defaults_and_rejects_malformed_change_graphs() {
        let empty = parse("<table:tracked-changes/>").unwrap().unwrap();
        assert!(!empty.enabled);
        assert!(empty.changes.is_empty());
        let info = "<office:change-info/>";
        for fragment in [
            r#"<table:tracked-changes table:track-changes="yes"/>"#.to_string(),
            format!(r#"<table:tracked-changes><table:insertion table:id="x" table:type="bad" table:position="0">{info}</table:insertion></table:tracked-changes>"#),
            format!(r#"<table:tracked-changes><table:insertion table:id="x" table:type="row" table:position="0" table:count="0">{info}</table:insertion></table:tracked-changes>"#),
            format!(r#"<table:tracked-changes><table:movement table:id="x"><table:source-range-address table:table="0" table:column="1"/><table:target-range-address table:table="0" table:column="1" table:row="1"/>{info}</table:movement></table:tracked-changes>"#),
            format!(r#"<table:tracked-changes><table:cell-content-change table:id="x"><table:cell-address table:table="0" table:column="1" table:row="1"/>{info}<table:previous><table:change-track-table-cell office:value="1"/></table:previous></table:cell-content-change></table:tracked-changes>"#),
            format!(r#"<table:tracked-changes><table:insertion table:id="x" table:type="row" table:position="0">{info}</table:insertion><table:deletion table:id="x" table:type="row" table:position="0">{info}</table:deletion></table:tracked-changes>"#),
            r#"<table:tracked-changes><fake:insertion xmlns:fake="urn:fake"/></table:tracked-changes>"#.to_string(),
        ] {
            assert!(parse(&fragment).is_err(), "accepted {fragment}");
        }
        assert!(parse_tracked_changes(
            r#"<table:tracked-changes xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#
        )
        .is_err());
    }

    #[test]
    fn parses_libreoffice_change_tracking_fixtures() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let tracked_path = root.join(
            "3rdparty/libreoffice-core/sc/qa/unit/data/ods/change-tracking.ods",
        );
        let tracked = crate::Spreadsheet::open(tracked_path).unwrap();
        let changes = tracked.tracked_changes().unwrap();
        assert_eq!(changes.changes.len(), 2);
        assert!(changes.changes.iter().all(|change| matches!(
            change,
            SpreadsheetTrackedChange::CellContent(_)
        )));

        let protected_path = root.join(
            "3rdparty/libreoffice-core/sc/qa/extras/testdocuments/RecordChangesProtected.ods",
        );
        let protected = crate::Spreadsheet::open(protected_path).unwrap();
        assert!(protected.tracked_changes().unwrap().changes.is_empty());
    }
}
