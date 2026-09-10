//! Bounded, read-only projection of ODF body structures that are not ordinary
//! paragraphs.  The source XML remains authoritative: this module only owns
//! semantic read values and never serializes or rewrites a structure.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
use std::ops::Range;

const MAX_STRUCTURE_TEXT: usize = 16 * 1024 * 1024;
const MAX_STRUCTURE_NODES: usize = 1_000_000;
const MAX_REPEAT: usize = 1_000_000;
const MAX_STRUCTURE_SITE_DEPTH: usize = 256;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const DC_NAMESPACE: &[u8] = b"http://purl.org/dc/elements/1.1/";
const META_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Ns {
    Office,
    Text,
    Table,
    Draw,
    Xlink,
    Svg,
    Xml,
    Dc,
    Meta,
    Other,
    None,
}

fn resolve_namespace(value: Option<&[u8]>) -> Ns {
    match value {
        Some(value) if value == OFFICE_NAMESPACE => Ns::Office,
        Some(value) if value == TEXT_NAMESPACE => Ns::Text,
        Some(value) if value == TABLE_NAMESPACE => Ns::Table,
        Some(value) if value == DRAW_NAMESPACE => Ns::Draw,
        Some(value) if value == XLINK_NAMESPACE => Ns::Xlink,
        Some(value) if value == SVG_NAMESPACE => Ns::Svg,
        Some(value) if value == XML_NAMESPACE => Ns::Xml,
        Some(value) if value == DC_NAMESPACE => Ns::Dc,
        Some(value) if value == META_NAMESPACE => Ns::Meta,
        Some(_) => Ns::Other,
        None => Ns::None,
    }
}

#[derive(Clone, Debug)]
struct Attribute {
    local: String,
    namespace: Ns,
    namespace_uri: Option<String>,
    value: String,
}

#[derive(Clone, Debug)]
enum Child {
    Element(Box<Node>),
    Text(String),
}

#[derive(Clone, Debug)]
struct Node {
    children: Vec<Child>,
    local: String,
    namespace: Ns,
    namespace_uri: Option<String>,
    attributes: Vec<Attribute>,
}

impl Node {
    fn attribute(&self, namespace: Ns, local: &[u8]) -> Option<&str> {
        let local = std::str::from_utf8(local).ok()?;
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace == namespace
                    && (namespace != Ns::Other || attribute.namespace_uri.is_some())
                    && attribute.local == local
            })
            .map(|attribute| attribute.value.as_str())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RootKind {
    Table,
    Section,
    Note,
    Annotation,
    TrackedChanges,
    ChangedRegion,
    ChangeStart,
    ChangeEnd,
    Change,
    Index,
    Frame,
    Ruby,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ContextKind {
    OfficeBody,
    OfficeText,
    TextSection,
    TextParagraph,
    TextHeading,
    TextSpan,
    TextNote,
    TextNoteBody,
    TextTrackedChanges,
    TextChangedRegion,
    TextInsertion,
    TextDeletion,
    TextFormatChange,
    TextIndex,
    TextIndexBody,
    TextIndexSource,
    Table,
    TableColumns,
    TableColumn,
    TableRows,
    TableRow,
    TableCell,
    CoveredTableCell,
    OfficeAnnotation,
    DrawFrame,
    DrawTextBox,
    DrawImage,
    DrawObject,
    DrawObjectOle,
    DrawPlugin,
    DrawFloatingFrame,
    TextRuby,
    TextRubyBase,
    TextRubyText,
    BodyContainer,
    Other,
}

impl ContextKind {
    fn permits_body_structure(self) -> bool {
        !matches!(self, Self::Other | Self::OfficeBody | Self::OfficeText)
    }
}

/// Additional semantic values projected from the OTH body.
pub(crate) struct BodyStructures {
    pub(crate) annotations: Vec<crate::annotation::Annotation>,
    pub(crate) changes: Vec<crate::change::Change>,
    pub(crate) frames: Vec<crate::frame::Frame>,
    pub(crate) indexes: Vec<crate::index::Index>,
    pub(crate) notes: Vec<crate::note::Note>,
    pub(crate) rubies: Vec<crate::ruby::Ruby>,
    pub(crate) sections: Vec<crate::section::Section>,
    pub(crate) tables: Vec<crate::table::Table>,
    retained_bytes: usize,
}

/// A source-backed text span inside one admitted body-structure family.
///
/// The scanner intentionally keeps this separate from [`BodyStructures`].
/// The latter is a semantic projection and is lazy; this value is created only
/// when an edit explicitly resolves a structural selector.
#[derive(Clone, Debug)]
pub(crate) struct StructureSite {
    pub(crate) kind: EditableStructureKind,
    pub(crate) index: usize,
    pub(crate) identity: Option<String>,
    pub(crate) full: Range<usize>,
    pub(crate) text: Option<crate::codec::ReplacementSite>,
    pub(crate) text_value: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum EditableStructureKind {
    Section,
    Note,
    Annotation,
    Table,
}

impl BodyStructures {
    fn new() -> Self {
        Self {
            annotations: Vec::new(),
            changes: Vec::new(),
            frames: Vec::new(),
            indexes: Vec::new(),
            notes: Vec::new(),
            rubies: Vec::new(),
            sections: Vec::new(),
            tables: Vec::new(),
            retained_bytes: 0,
        }
    }

    fn account(&mut self, bytes: usize) -> Result<()> {
        self.retained_bytes = self
            .retained_bytes
            .checked_add(bytes)
            .ok_or_else(|| Error::InvalidFormat("OTH projected body size overflow".to_string()))?;
        if self.retained_bytes > MAX_STRUCTURE_TEXT {
            return invalid("OTH projected body structures exceed the aggregate text limit");
        }
        Ok(())
    }
}

fn table_bytes(table: &crate::table::Table) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, table.name())?;
    add_string(&mut total, table.style_name())?;
    for column in table.columns() {
        add_string(&mut total, column.style_name())?;
        add_string(&mut total, column.default_cell_style_name())?;
    }
    for row in table.rows() {
        add_string(&mut total, row.style_name())?;
        for cell in row.cells() {
            add_string(&mut total, cell.style_name())?;
            add_string(&mut total, cell.formula())?;
            add_string(&mut total, cell.value_type())?;
            add_string(&mut total, cell.value())?;
            add_string(&mut total, Some(cell.text()))?;
            for paragraph in cell.paragraphs() {
                add_string(&mut total, Some(paragraph.text()))?;
            }
        }
    }
    Ok(total)
}

fn section_bytes(section: &crate::section::Section) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, section.name())?;
    add_string(&mut total, section.style_name())?;
    add_string(&mut total, section.condition())?;
    add_string(&mut total, Some(section.text()))?;
    for paragraph in section.paragraphs() {
        add_string(&mut total, Some(paragraph.text()))?;
    }
    Ok(total)
}

fn note_bytes(note: &crate::note::Note) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, Some(note.class().as_str()))?;
    add_string(&mut total, note.id())?;
    add_string(&mut total, Some(note.citation()))?;
    add_string(&mut total, note.label())?;
    add_string(&mut total, Some(note.body()))?;
    for paragraph in note.paragraphs() {
        add_string(&mut total, Some(paragraph.text()))?;
    }
    Ok(total)
}

fn annotation_bytes(annotation: &crate::annotation::Annotation) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, annotation.name())?;
    add_string(&mut total, annotation.creator())?;
    add_string(&mut total, annotation.date())?;
    add_string(&mut total, annotation.date_string())?;
    add_string(&mut total, annotation.initials())?;
    add_string(&mut total, Some(annotation.text()))?;
    Ok(total)
}

fn change_bytes(change: &crate::change::Change) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, Some(change.kind().as_str()))?;
    add_string(&mut total, change.id())?;
    add_string(&mut total, change.author())?;
    add_string(&mut total, change.date())?;
    add_string(&mut total, Some(change.text()))?;
    Ok(total)
}

fn index_bytes(index: &crate::index::Index) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, Some(index.kind().as_str()))?;
    add_string(&mut total, index.name())?;
    add_string(&mut total, index.source())?;
    add_string(&mut total, Some(index.body()))?;
    Ok(total)
}

fn frame_bytes(frame: &crate::frame::Frame) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, frame.name())?;
    add_string(&mut total, frame.style_name())?;
    add_string(&mut total, frame.anchor_type())?;
    add_string(&mut total, frame.x())?;
    add_string(&mut total, frame.y())?;
    add_string(&mut total, frame.width())?;
    add_string(&mut total, frame.height())?;
    add_string(&mut total, frame.href())?;
    add_string(&mut total, Some(frame.text()))?;
    Ok(total)
}

fn ruby_bytes(ruby: &crate::ruby::Ruby) -> Result<usize> {
    let mut total = 0usize;
    add_string(&mut total, ruby.style_name())?;
    add_string(&mut total, ruby.text_style_name())?;
    add_string(&mut total, Some(ruby.base()))?;
    add_string(&mut total, Some(ruby.text()))?;
    Ok(total)
}

struct Budget {
    nodes: usize,
    text_bytes: usize,
}

impl Budget {
    fn node(&mut self) -> Result<()> {
        self.nodes = self.nodes.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure node count overflow".to_string())
        })?;
        if self.nodes > MAX_STRUCTURE_NODES {
            return invalid("OTH body structure node count exceeds the limit");
        }
        Ok(())
    }

    fn text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self.text_bytes.checked_add(bytes).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure text size overflow".to_string())
        })?;
        if self.text_bytes > MAX_STRUCTURE_TEXT {
            return invalid("OTH body structure text exceeds the limit");
        }
        Ok(())
    }
}

/// Projects tables, sections, notes, annotations, review metadata, indexes,
/// frames, and ruby pairs from a validated `content.xml` source.
pub(crate) fn project_structures(xml: &str) -> Result<BodyStructures> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let tracked_changes_present = document_has_tracked_changes(xml)?;
    let mut stack = Vec::<Node>::new();
    let mut contexts = Vec::<ContextKind>::new();
    let mut body_text_depth = 0usize;
    let mut output = BodyStructures::new();
    let mut budget = Budget {
        nodes: 0,
        text_bytes: 0,
    };

    loop {
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                let (_namespace, context, root) = classify_element(&reader, start.name());
                let owns_text = context == ContextKind::OfficeText
                    && contexts.last() == Some(&ContextKind::OfficeBody);
                let capture_root = stack.is_empty()
                    && body_text_depth > 0
                    && owns_structure_root(&contexts)
                    && root.is_some_and(|root| {
                        !root_requires_tracked_changes(root) || tracked_changes_present
                    });
                if !stack.is_empty() || capture_root {
                    let node = parse_node(&reader, &start, &mut budget)?;
                    reserve(&mut stack, "OTH body structure stack")?;
                    stack.push(node);
                }
                if owns_text {
                    body_text_depth = body_text_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("OTH office:text depth overflow".to_string())
                    })?;
                }
                if contexts.len() >= MAX_STRUCTURE_NODES {
                    return invalid("OTH body structure nesting exceeds the limit");
                }
                reserve(&mut contexts, "OTH body structure context stack")?;
                contexts.push(context);
            },
            Event::Empty(start) => {
                let (_namespace, _context, root) = classify_element(&reader, start.name());
                let capture_root = stack.is_empty()
                    && body_text_depth > 0
                    && owns_structure_root(&contexts)
                    && root.is_some_and(|root| {
                        !root_requires_tracked_changes(root) || tracked_changes_present
                    });
                if !stack.is_empty() || capture_root {
                    let node = parse_node(&reader, &start, &mut budget)?;
                    if stack.is_empty() {
                        if let Some(kind) = root_kind(node.namespace, &node.local) {
                            finish_root(node, kind, &mut output)?;
                        }
                    } else {
                        append_node(&mut stack, node)?;
                    }
                }
            },
            Event::End(_end) => {
                if let Some(node) = stack.pop() {
                    if let Some(parent) = stack.last_mut() {
                        append_child(parent, Child::Element(Box::new(node)))?;
                    } else if let Some(kind) = root_kind(node.namespace, &node.local) {
                        finish_root(node, kind, &mut output)?;
                    }
                }
                let owns_text_end = contexts.len() >= 2
                    && contexts[contexts.len() - 1] == ContextKind::OfficeText
                    && contexts[contexts.len() - 2] == ContextKind::OfficeBody;
                if contexts.pop() == Some(ContextKind::OfficeText)
                    && owns_text_end
                    && body_text_depth > 0
                {
                    body_text_depth -= 1;
                }
            },
            Event::Text(text) => {
                if let Some(node) = stack.last_mut() {
                    let value = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid OTH structure text: {error}"))
                        })?
                        .into_owned();
                    budget.text(value.len())?;
                    append_child(node, Child::Text(value))?;
                }
            },
            Event::CData(text) => {
                if let Some(node) = stack.last_mut() {
                    let value = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid OTH structure CDATA: {error}"))
                        })?
                        .into_owned();
                    budget.text(value.len())?;
                    append_child(node, Child::Text(value))?;
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(node) = stack.last_mut() {
                    let value = reference_value(&reference)?;
                    budget.text(value.len())?;
                    append_child(node, Child::Text(value))?;
                }
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                if !stack.is_empty() {
                    return invalid("OTH body structure stack is not empty");
                }
                return Ok(output);
            },
        }
    }
}

fn document_has_tracked_changes(xml: &str) -> Result<bool> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut contexts = Vec::new();
    loop {
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                let (_namespace, _context, root) = classify_element(&reader, start.name());
                if root == Some(RootKind::TrackedChanges) && owns_structure_root(&contexts) {
                    return Ok(true);
                }
                if contexts.len() >= MAX_STRUCTURE_SITE_DEPTH {
                    return invalid("OTH tracked-change context scan exceeds the depth limit");
                }
                contexts
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "OTH tracked-change context scan",
                        source,
                    })?;
                contexts.push(_context);
            },
            Event::Empty(start) => {
                let (_namespace, _context, root) = classify_element(&reader, start.name());
                if root == Some(RootKind::TrackedChanges) && owns_structure_root(&contexts) {
                    return Ok(true);
                }
            },
            Event::End(end) => {
                let (_namespace, context, _root) = classify_element(&reader, end.name());
                let Some(open) = contexts.pop() else {
                    return invalid("OTH tracked-change context scan has an unmatched end tag");
                };
                if open != context {
                    return invalid("OTH tracked-change context scan has mismatched tags");
                }
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::Eof => return Ok(false),
            _ => {},
        }
    }
}

fn classify_element(
    reader: &NsReader<&[u8]>,
    name: QName<'_>,
) -> (Ns, ContextKind, Option<RootKind>) {
    let (resolved, local_name) = reader.resolver().resolve_element(name);
    let resolved = match resolved {
        ResolveResult::Bound(Namespace(value)) => resolve_namespace(Some(value)),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Ns::None,
    };
    let local = local_name.as_ref();
    (
        resolved,
        context_kind(resolved, local),
        root_kind_bytes(resolved, local),
    )
}

fn root_kind_bytes(namespace: Ns, local: &[u8]) -> Option<RootKind> {
    match (namespace, local) {
        (Ns::Table, b"table") => Some(RootKind::Table),
        (Ns::Text, b"section") => Some(RootKind::Section),
        (Ns::Text, b"note") => Some(RootKind::Note),
        (Ns::Office, b"annotation") => Some(RootKind::Annotation),
        (Ns::Text, b"tracked-changes") => Some(RootKind::TrackedChanges),
        (Ns::Text, b"changed-region") => Some(RootKind::ChangedRegion),
        (Ns::Text, b"change-start") => Some(RootKind::ChangeStart),
        (Ns::Text, b"change-end") => Some(RootKind::ChangeEnd),
        (Ns::Text, b"change") => Some(RootKind::Change),
        (
            Ns::Text,
            b"table-of-content"
            | b"illustration-index"
            | b"table-index"
            | b"object-index"
            | b"user-index"
            | b"alphabetical-index"
            | b"bibliography"
            | b"bibliography-index",
        ) => Some(RootKind::Index),
        (Ns::Draw, b"frame" | b"text-box") => Some(RootKind::Frame),
        (Ns::Text, b"ruby") => Some(RootKind::Ruby),
        _ => None,
    }
}

const fn root_requires_tracked_changes(kind: RootKind) -> bool {
    matches!(
        kind,
        RootKind::ChangedRegion | RootKind::ChangeStart | RootKind::ChangeEnd | RootKind::Change
    )
}

fn context_kind(namespace: Ns, local: &[u8]) -> ContextKind {
    match (namespace, local) {
        (Ns::Office, b"body") => ContextKind::OfficeBody,
        (Ns::Office, b"text") => ContextKind::OfficeText,
        (Ns::Text, b"section") => ContextKind::TextSection,
        (Ns::Text, b"p") => ContextKind::TextParagraph,
        (Ns::Text, b"h") => ContextKind::TextHeading,
        (Ns::Text, b"span") => ContextKind::TextSpan,
        (Ns::Text, b"note") => ContextKind::TextNote,
        (Ns::Text, b"note-body") => ContextKind::TextNoteBody,
        (Ns::Text, b"tracked-changes") => ContextKind::TextTrackedChanges,
        (Ns::Text, b"changed-region") => ContextKind::TextChangedRegion,
        (Ns::Text, b"insertion") => ContextKind::TextInsertion,
        (Ns::Text, b"deletion") => ContextKind::TextDeletion,
        (Ns::Text, b"format-change") => ContextKind::TextFormatChange,
        (
            Ns::Text,
            b"list"
            | b"list-header"
            | b"list-item"
            | b"numbered-paragraph"
            | b"bookmark-start"
            | b"bookmark-end"
            | b"bookmark"
            | b"reference-mark-start"
            | b"reference-mark-end"
            | b"reference-mark"
            | b"change"
            | b"note-citation"
            | b"index-title"
            | b"index-title-template"
            | b"index-entry"
            | b"table-of-content-entry-template"
            | b"illustration-index-entry-template"
            | b"table-index-entry-template"
            | b"object-index-entry-template"
            | b"user-index-entry-template"
            | b"alphabetical-index-entry-template"
            | b"bibliography-entry-template",
        ) => ContextKind::BodyContainer,
        (Ns::Text, b"index-body") => ContextKind::TextIndexBody,
        (
            Ns::Text,
            b"section-source"
            | b"table-of-content-source"
            | b"illustration-index-source"
            | b"table-index-source"
            | b"object-index-source"
            | b"user-index-source"
            | b"alphabetical-index-source"
            | b"bibliography-source",
        ) => ContextKind::TextIndexSource,
        (Ns::Text, local) if root_kind_bytes(namespace, local) == Some(RootKind::Index) => {
            ContextKind::TextIndex
        },
        (Ns::Table, b"table") => ContextKind::Table,
        (Ns::Table, b"table-columns" | b"table-column-group" | b"table-header-columns") => {
            ContextKind::TableColumns
        },
        (Ns::Table, b"table-column") => ContextKind::TableColumn,
        (Ns::Table, b"table-rows" | b"table-row-group" | b"table-header-rows") => {
            ContextKind::TableRows
        },
        (Ns::Table, b"table-row") => ContextKind::TableRow,
        (Ns::Table, b"table-cell") => ContextKind::TableCell,
        (Ns::Table, b"covered-table-cell") => ContextKind::CoveredTableCell,
        (Ns::Office, b"annotation") => ContextKind::OfficeAnnotation,
        (Ns::Draw, b"frame") => ContextKind::DrawFrame,
        (Ns::Draw, b"text-box") => ContextKind::DrawTextBox,
        (Ns::Draw, b"image") => ContextKind::DrawImage,
        (Ns::Draw, b"object") => ContextKind::DrawObject,
        (Ns::Draw, b"object-ole") => ContextKind::DrawObjectOle,
        (Ns::Draw, b"plugin") => ContextKind::DrawPlugin,
        (Ns::Draw, b"floating-frame") => ContextKind::DrawFloatingFrame,
        (Ns::Text, b"ruby") => ContextKind::TextRuby,
        (Ns::Text, b"ruby-base") => ContextKind::TextRubyBase,
        (Ns::Text, b"ruby-text") => ContextKind::TextRubyText,
        _ => ContextKind::Other,
    }
}

fn owns_structure_root(contexts: &[ContextKind]) -> bool {
    let Some(text_index) = contexts.iter().enumerate().rposition(|(index, context)| {
        *context == ContextKind::OfficeText
            && index > 0
            && contexts[index - 1] == ContextKind::OfficeBody
    }) else {
        return false;
    };
    contexts[text_index + 1..]
        .iter()
        .copied()
        .all(ContextKind::permits_body_structure)
}

fn parse_node(
    reader: &NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    budget: &mut Budget,
) -> Result<Node> {
    // Resolve and charge every retained scalar before constructing any owned
    // name, namespace URI, attribute, or attribute value.  This keeps a
    // rejected node from allocating an unbounded metadata fan-out first.
    budget.node()?;
    let (resolved, local) = reader.resolver().resolve_element(start.name());
    let local = std::str::from_utf8(local.as_ref()).map_err(|error| {
        Error::InvalidFormat(format!("invalid OTH structure element name: {error}"))
    })?;
    let namespace_uri = match resolved {
        ResolveResult::Bound(Namespace(value)) => {
            (resolve_namespace(Some(value)) == Ns::Other).then_some(value)
        },
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    };
    let namespace = match resolved {
        ResolveResult::Bound(Namespace(value)) => resolve_namespace(Some(value)),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Ns::None,
    };
    budget.text(local.len())?;
    if let Some(namespace_uri) = namespace_uri {
        budget.text(namespace_uri.len())?;
    }
    let mut attribute_count = 0usize;
    for raw in start.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute: {error}"))
        })?;
        if attribute.key.as_ref().starts_with(b"xmlns") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let local = std::str::from_utf8(local.as_ref()).map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute name: {error}"))
        })?;
        let namespace_uri = match resolved {
            ResolveResult::Bound(Namespace(value)) => {
                (resolve_namespace(Some(value)) == Ns::Other).then_some(value)
            },
            ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
        };
        if let Some(namespace_uri) = namespace_uri {
            budget.text(namespace_uri.len())?;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid OTH structure attribute value: {error}"))
            })?;
        budget.text(local.len())?;
        budget.text(value.len())?;
        attribute_count = attribute_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH body structure attribute count overflow".to_string())
        })?;
    }
    let mut node = Node {
        children: Vec::new(),
        local: local.to_owned(),
        namespace,
        namespace_uri: namespace_uri.map(|value| {
            std::str::from_utf8(value)
                .map(str::to_owned)
                .unwrap_or_else(|_| String::from_utf8_lossy(value).into_owned())
        }),
        attributes: Vec::new(),
    };
    node.attributes
        .try_reserve_exact(attribute_count)
        .map_err(|source| Error::Allocation {
            resource: "OTH body structure attributes",
            source,
        })?;
    for raw in start.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute: {error}"))
        })?;
        if attribute.key.as_ref().starts_with(b"xmlns") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let local = std::str::from_utf8(local.as_ref()).map_err(|error| {
            Error::InvalidFormat(format!("invalid OTH structure attribute name: {error}"))
        })?;
        let namespace_uri = match resolved {
            ResolveResult::Bound(Namespace(value)) => {
                (resolve_namespace(Some(value)) == Ns::Other).then_some(value)
            },
            ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
        };
        let namespace = match resolved {
            ResolveResult::Bound(Namespace(value)) => resolve_namespace(Some(value)),
            ResolveResult::Unbound | ResolveResult::Unknown(_) => Ns::None,
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid OTH structure attribute value: {error}"))
            })?
            .into_owned();
        node.attributes.push(Attribute {
            local: local.to_owned(),
            namespace,
            namespace_uri: namespace_uri.map(|value| {
                std::str::from_utf8(value)
                    .map(str::to_owned)
                    .unwrap_or_else(|_| String::from_utf8_lossy(value).into_owned())
            }),
            value,
        });
    }
    Ok(node)
}

fn append_node(stack: &mut [Node], node: Node) -> Result<()> {
    let parent = stack
        .last_mut()
        .ok_or_else(|| Error::InvalidFormat("OTH body structure parent is missing".to_string()))?;
    append_child(parent, Child::Element(Box::new(node)))
}

fn append_child(parent: &mut Node, child: Child) -> Result<()> {
    parent
        .children
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "OTH body structure children",
            source,
        })?;
    parent.children.push(child);
    Ok(())
}

fn root_kind(namespace: Ns, local: &str) -> Option<RootKind> {
    match (namespace, local) {
        (Ns::Table, "table") => Some(RootKind::Table),
        (Ns::Text, "section") => Some(RootKind::Section),
        (Ns::Text, "note") => Some(RootKind::Note),
        (Ns::Office, "annotation") => Some(RootKind::Annotation),
        (Ns::Text, "tracked-changes") => Some(RootKind::TrackedChanges),
        (Ns::Text, "changed-region") => Some(RootKind::ChangedRegion),
        (Ns::Text, "change-start") => Some(RootKind::ChangeStart),
        (Ns::Text, "change-end") => Some(RootKind::ChangeEnd),
        (Ns::Text, "change") => Some(RootKind::Change),
        (
            Ns::Text,
            "table-of-content" | "illustration-index" | "table-index" | "object-index"
            | "user-index" | "alphabetical-index" | "bibliography" | "bibliography-index",
        ) => Some(RootKind::Index),
        (Ns::Draw, "frame" | "text-box") => Some(RootKind::Frame),
        (Ns::Text, "ruby") => Some(RootKind::Ruby),
        _ => None,
    }
}

fn finish_root(node: Node, kind: RootKind, output: &mut BodyStructures) -> Result<()> {
    match kind {
        RootKind::Table => collect_tables(&node, output, false),
        RootKind::Section => collect_sections(&node, 1, output, false),
        RootKind::Note => collect_notes(&node, output, false),
        RootKind::Annotation => collect_annotations(&node, output, false),
        RootKind::TrackedChanges
        | RootKind::ChangedRegion
        | RootKind::ChangeStart
        | RootKind::ChangeEnd
        | RootKind::Change => collect_changes(&node, output, true),
        RootKind::Index => collect_indexes(&node, output, false),
        RootKind::Frame => collect_frames(&node, output, false),
        RootKind::Ruby => collect_rubies(&node, output, false),
    }
}

fn collect_tables(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table" {
        let table = project_table(node)?;
        let bytes = table_bytes(&table)?;
        output.account(bytes)?;
        reserve(&mut output.tables, "OTH table projection")?;
        output.tables.push(table);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn collect_descendant_roots(
    node: &Node,
    section_depth: usize,
    output: &mut BodyStructures,
    tracked_scope: bool,
) -> Result<()> {
    let tracked_scope =
        tracked_scope || (node.namespace == Ns::Text && node.local == "tracked-changes");
    for child in elements(node) {
        if !valid_structure_child(child) {
            continue;
        }
        match root_kind(child.namespace, &child.local) {
            Some(RootKind::Table) => collect_tables(child, output, tracked_scope)?,
            Some(RootKind::Section) => {
                collect_sections(child, section_depth, output, tracked_scope)?
            },
            Some(RootKind::Note) => collect_notes(child, output, tracked_scope)?,
            Some(RootKind::Annotation) => collect_annotations(child, output, tracked_scope)?,
            Some(
                RootKind::TrackedChanges
                | RootKind::ChangedRegion
                | RootKind::ChangeStart
                | RootKind::ChangeEnd
                | RootKind::Change,
            ) => collect_changes(child, output, tracked_scope)?,
            Some(RootKind::Index) => collect_indexes(child, output, tracked_scope)?,
            Some(RootKind::Frame) => collect_frames(child, output, tracked_scope)?,
            Some(RootKind::Ruby) => collect_rubies(child, output, tracked_scope)?,
            None => collect_descendant_roots(child, section_depth, output, tracked_scope)?,
        }
    }
    Ok(())
}

fn collect_table_columns(
    node: &Node,
    columns: &mut Vec<crate::table::Column>,
    declared_columns: &mut usize,
) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table-column" {
        reserve(columns, "OTH table columns")?;
        let repeated = count_attr(node, b"number-columns-repeated")?;
        *declared_columns = checked_dimension_add(*declared_columns, repeated)?;
        columns.push(crate::table::Column::projected(
            attr(node, Ns::Table, b"style-name").map(str::to_owned),
            attr(node, Ns::Table, b"default-cell-style-name").map(str::to_owned),
            repeated,
        ));
        return Ok(());
    }
    for child in elements(node) {
        match (child.namespace, child.local.as_str()) {
            (Ns::Table, "table-column") => {
                collect_table_columns(child, columns, declared_columns)?;
            },
            (Ns::Table, "table-columns" | "table-column-group" | "table-header-columns") => {
                collect_table_columns(child, columns, declared_columns)?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn collect_table_rows(
    node: &Node,
    rows: &mut Vec<crate::table::Row>,
    logical_columns: &mut usize,
) -> Result<()> {
    if node.namespace == Ns::Table && node.local == "table-row" {
        reserve(rows, "OTH table rows")?;
        let row = project_row(node)?;
        *logical_columns = (*logical_columns).max(row_width(&row)?);
        rows.push(row);
        return Ok(());
    }
    for child in elements(node) {
        match (child.namespace, child.local.as_str()) {
            (Ns::Table, "table-row") => {
                collect_table_rows(child, rows, logical_columns)?;
            },
            (Ns::Table, "table-rows" | "table-row-group" | "table-header-rows") => {
                collect_table_rows(child, rows, logical_columns)?
            },
            _ => {},
        }
    }
    Ok(())
}

fn project_table(node: &Node) -> Result<crate::table::Table> {
    let mut columns = Vec::new();
    let mut rows = Vec::new();
    let mut declared_columns = 0usize;
    let mut logical_columns = 0usize;
    for child in elements(node) {
        if child.namespace != Ns::Table {
            continue;
        }
        match child.local.as_str() {
            "table-column" => {
                collect_table_columns(child, &mut columns, &mut declared_columns)?;
            },
            "table-columns" | "table-column-group" | "table-header-columns" => {
                collect_table_columns(child, &mut columns, &mut declared_columns)?;
            },
            "table-row" => collect_table_rows(child, &mut rows, &mut logical_columns)?,
            "table-rows" | "table-row-group" | "table-header-rows" => {
                collect_table_rows(child, &mut rows, &mut logical_columns)?;
            },
            _ => {},
        }
    }
    Ok(crate::table::Table::projected(
        attr(node, Ns::Table, b"name").map(str::to_owned),
        attr(node, Ns::Table, b"style-name").map(str::to_owned),
        columns,
        rows,
        declared_columns,
        checked_dimension_max(declared_columns, logical_columns)?,
    ))
}

fn row_width(row: &crate::table::Row) -> Result<usize> {
    let mut width = 0usize;
    for cell in row.cells() {
        let span = cell
            .columns_spanned()
            .checked_mul(cell.repeat_count())
            .ok_or_else(|| Error::InvalidFormat("OTH table logical width overflow".to_string()))?;
        width = checked_dimension_add(width, span)?;
    }
    Ok(width)
}

fn checked_dimension_add(left: usize, right: usize) -> Result<usize> {
    let value = left
        .checked_add(right)
        .ok_or_else(|| Error::InvalidFormat("OTH table logical width overflow".to_string()))?;
    if value > MAX_REPEAT {
        return invalid("OTH table logical width exceeds the supported limit");
    }
    Ok(value)
}

fn checked_dimension_max(left: usize, right: usize) -> Result<usize> {
    let value = left.max(right);
    if value > MAX_REPEAT {
        return invalid("OTH table logical width exceeds the supported limit");
    }
    Ok(value)
}

fn project_row(node: &Node) -> Result<crate::table::Row> {
    let mut cells = Vec::new();
    for child in elements(node) {
        if child.namespace == Ns::Table
            && matches!(child.local.as_str(), "table-cell" | "covered-table-cell")
        {
            reserve(&mut cells, "OTH table cells")?;
            cells.push(project_cell(child)?);
        }
    }
    Ok(crate::table::Row::projected(
        attr(node, Ns::Table, b"style-name").map(str::to_owned),
        count_attr(node, b"number-rows-repeated")?,
        cells,
    ))
}

fn project_cell(node: &Node) -> Result<crate::table::Cell> {
    let mut paragraphs = Vec::new();
    for child in elements(node) {
        if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
            let text = plain_text(child)?;
            reserve(&mut paragraphs, "OTH table cell paragraphs")?;
            paragraphs.push(crate::paragraph::Paragraph::new(text));
        }
    }
    Ok(crate::table::Cell::projected(
        if node.local == "covered-table-cell" {
            crate::table::CellKind::Covered
        } else {
            crate::table::CellKind::Cell
        },
        attr(node, Ns::Table, b"style-name").map(str::to_owned),
        count_attr(node, b"number-columns-repeated")?,
        count_attr_default(node, b"number-columns-spanned", 1)?,
        count_attr_default(node, b"number-rows-spanned", 1)?,
        attr(node, Ns::Table, b"formula").map(str::to_owned),
        attr(node, Ns::Office, b"value-type").map(str::to_owned),
        attr(node, Ns::Office, b"value").map(str::to_owned),
        plain_text(node)?,
        paragraphs,
    ))
}

fn collect_sections(
    node: &Node,
    depth: usize,
    output: &mut BodyStructures,
    tracked_scope: bool,
) -> Result<()> {
    if node.namespace == Ns::Text && node.local == "section" {
        let mut paragraphs = Vec::new();
        for child in elements(node) {
            if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
                reserve(&mut paragraphs, "OTH section paragraphs")?;
                paragraphs.push(crate::paragraph::Paragraph::new(plain_text(child)?));
            }
        }
        let section = crate::section::Section::projected(
            attr(node, Ns::Text, b"name").map(str::to_owned),
            attr(node, Ns::Text, b"style-name").map(str::to_owned),
            bool_attr(node, Ns::Text, b"protected")?,
            optional_bool_attr(node, Ns::Text, b"display")?,
            attr(node, Ns::Text, b"condition").map(str::to_owned),
            depth,
            plain_text(node)?,
            paragraphs,
        );
        let bytes = section_bytes(&section)?;
        output.account(bytes)?;
        reserve(&mut output.sections, "OTH section projection")?;
        output.sections.push(section);
    }
    collect_descendant_roots(
        node,
        depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("OTH section depth overflow".to_string()))?,
        output,
        tracked_scope,
    )?;
    Ok(())
}

fn collect_notes(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Text && node.local == "note" {
        let class = match attr(node, Ns::Text, b"note-class") {
            Some("footnote") => crate::note::NoteClass::Footnote,
            Some("endnote") => crate::note::NoteClass::Endnote,
            Some(value) => crate::note::NoteClass::Other(value.to_owned()),
            None => crate::note::NoteClass::Other(String::new()),
        };
        let citation_node = direct_child(node, Ns::Text, "note-citation");
        let body_node = direct_child(node, Ns::Text, "note-body");
        let mut paragraphs = Vec::new();
        if let Some(body) = body_node {
            for child in elements(body) {
                if child.namespace == Ns::Text && matches!(child.local.as_str(), "p" | "h") {
                    reserve(&mut paragraphs, "OTH note paragraphs")?;
                    paragraphs.push(crate::paragraph::Paragraph::new(plain_text(child)?));
                }
            }
        }
        let note = crate::note::Note::projected(
            class,
            attr(node, Ns::Text, b"id").map(str::to_owned),
            citation_node.and_then(|value| attr(value, Ns::Text, b"label").map(str::to_owned)),
            citation_node
                .map(plain_text)
                .transpose()?
                .unwrap_or_default(),
            body_node.map(plain_text).transpose()?.unwrap_or_default(),
            paragraphs,
        );
        let bytes = note_bytes(&note)?;
        output.account(bytes)?;
        reserve(&mut output.notes, "OTH note projection")?;
        output.notes.push(note);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn collect_annotations(
    node: &Node,
    output: &mut BodyStructures,
    tracked_scope: bool,
) -> Result<()> {
    if node.namespace == Ns::Office && node.local == "annotation" {
        let creator = direct_child(node, Ns::Dc, "creator")
            .map(plain_text)
            .transpose()?;
        let date = direct_child(node, Ns::Dc, "date")
            .map(plain_text)
            .transpose()?;
        let date_string = direct_child(node, Ns::Meta, "date-string")
            .map(plain_text)
            .transpose()?;
        let initials = direct_child(node, Ns::Meta, "creator-initials")
            .or_else(|| direct_child(node, Ns::Text, "sender-initials"))
            .map(plain_text)
            .transpose()?;
        let annotation = crate::annotation::Annotation::projected(
            attr(node, Ns::Office, b"name").map(str::to_owned),
            creator,
            date,
            date_string,
            initials,
            optional_bool_attr(node, Ns::Office, b"display")?,
            annotation_text(node)?,
        );
        let bytes = annotation_bytes(&annotation)?;
        output.account(bytes)?;
        reserve(&mut output.annotations, "OTH annotation projection")?;
        output.annotations.push(annotation);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn annotation_text(node: &Node) -> Result<String> {
    let mut text = String::new();
    for child in elements(node) {
        if child.namespace == Ns::Dc
            || child.namespace == Ns::Meta
            || (child.namespace == Ns::Text && child.local == "sender-initials")
        {
            continue;
        }
        append_plain_text(&mut text, child)?;
    }
    Ok(text)
}

fn collect_changes(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    let tracked_scope =
        tracked_scope || (node.namespace == Ns::Text && node.local == "tracked-changes");
    let kind = match (node.namespace, node.local.as_str()) {
        (Ns::Text, "changed-region") => Some(change_kind(node)),
        (Ns::Text, "change-start") => Some(crate::change::Kind::Start),
        (Ns::Text, "change-end") => Some(crate::change::Kind::End),
        (Ns::Text, "change") => Some(crate::change::Kind::Other("change".to_string())),
        _ => None,
    };
    if tracked_scope {
        if let Some(kind) = kind {
            let info = find_descendant(node, Ns::Office, "change-info");
            let author = info
                .and_then(|value| descendant(value, Ns::Dc, "creator"))
                .map(plain_text)
                .transpose()?;
            let date = info
                .and_then(|value| descendant(value, Ns::Dc, "date"))
                .map(plain_text)
                .transpose()?;
            let change = crate::change::Change::projected(
                attr(node, Ns::Text, b"change-id")
                    .or_else(|| attr(node, Ns::Text, b"id"))
                    .or_else(|| attr(node, Ns::Xml, b"id"))
                    .map(str::to_owned),
                kind,
                author,
                date,
                change_text(node)?,
            );
            let bytes = change_bytes(&change)?;
            output.account(bytes)?;
            reserve(&mut output.changes, "OTH tracked change projection")?;
            output.changes.push(change);
        }
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn change_text(node: &Node) -> Result<String> {
    let content = direct_child(node, Ns::Text, "insertion")
        .or_else(|| direct_child(node, Ns::Text, "deletion"))
        .or_else(|| direct_child(node, Ns::Text, "format-change"));
    let Some(content) = content else {
        return plain_text(node);
    };
    let mut text = String::new();
    for child in &content.children {
        match child {
            Child::Element(child) => {
                if child.namespace == Ns::Office && child.local == "change-info" {
                    continue;
                }
                append_plain_text(&mut text, child)?;
            },
            Child::Text(value) => checked_append(&mut text, value)?,
        }
    }
    Ok(text)
}

fn change_kind(node: &Node) -> crate::change::Kind {
    for child in elements(node) {
        if child.namespace == Ns::Text {
            return match child.local.as_str() {
                "insertion" => crate::change::Kind::Insertion,
                "deletion" => crate::change::Kind::Deletion,
                "format-change" => crate::change::Kind::Format,
                "change" | "move" => crate::change::Kind::Move,
                other => crate::change::Kind::Other(other.to_owned()),
            };
        }
    }
    crate::change::Kind::Other("changed-region".to_string())
}

fn collect_indexes(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Text && index_kind(&node.local).is_some() {
        let source = elements(node)
            .find(|child| child.namespace == Ns::Text && child.local.ends_with("-source"))
            .map(plain_text)
            .transpose()?;
        let body = elements(node)
            .find(|child| child.namespace == Ns::Text && child.local == "index-body")
            .map(plain_text)
            .transpose()?
            .unwrap_or_default();
        let index = crate::index::Index::projected(
            index_kind(&node.local)
                .unwrap_or_else(|| crate::index::Kind::Other(node.local.clone())),
            attr(node, Ns::Text, b"name").map(str::to_owned),
            bool_attr(node, Ns::Text, b"protected")?,
            source,
            body,
        );
        let bytes = index_bytes(&index)?;
        output.account(bytes)?;
        reserve(&mut output.indexes, "OTH index projection")?;
        output.indexes.push(index);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn index_kind(local: &str) -> Option<crate::index::Kind> {
    Some(match local {
        "table-of-content" => crate::index::Kind::TableOfContents,
        "illustration-index" => crate::index::Kind::Illustration,
        "table-index" => crate::index::Kind::Table,
        "object-index" => crate::index::Kind::Object,
        "user-index" => crate::index::Kind::User,
        "alphabetical-index" => crate::index::Kind::Alphabetical,
        "bibliography" | "bibliography-index" => crate::index::Kind::Bibliography,
        _ => return None,
    })
}

fn collect_frames(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Draw && matches!(node.local.as_str(), "frame" | "text-box") {
        let kind = frame_kind(node);
        let href = find_descendant(node, Ns::Draw, "image")
            .or_else(|| find_descendant(node, Ns::Draw, "object"))
            .or_else(|| find_descendant(node, Ns::Draw, "object-ole"))
            .or_else(|| find_descendant(node, Ns::Draw, "plugin"))
            .or_else(|| find_descendant(node, Ns::Draw, "floating-frame"))
            .and_then(|value| attr(value, Ns::Xlink, b"href"))
            .map(str::to_owned);
        let frame = crate::frame::Frame::projected(
            kind,
            attr(node, Ns::Draw, b"name").map(str::to_owned),
            attr(node, Ns::Draw, b"style-name").map(str::to_owned),
            attr(node, Ns::Text, b"anchor-type").map(str::to_owned),
            attr(node, Ns::Svg, b"x").map(str::to_owned),
            attr(node, Ns::Svg, b"y").map(str::to_owned),
            attr(node, Ns::Svg, b"width").map(str::to_owned),
            attr(node, Ns::Svg, b"height").map(str::to_owned),
            href,
            plain_text(node)?,
        );
        let bytes = frame_bytes(&frame)?;
        output.account(bytes)?;
        reserve(&mut output.frames, "OTH frame projection")?;
        output.frames.push(frame);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn frame_kind(node: &Node) -> crate::frame::Kind {
    if node.namespace == Ns::Draw && node.local == "text-box"
        || find_descendant(node, Ns::Draw, "text-box").is_some()
    {
        return crate::frame::Kind::TextBox;
    }
    if find_descendant(node, Ns::Draw, "image").is_some() {
        crate::frame::Kind::Image
    } else if find_descendant(node, Ns::Draw, "object-ole").is_some() {
        crate::frame::Kind::OleObject
    } else if find_descendant(node, Ns::Draw, "object").is_some() {
        crate::frame::Kind::Object
    } else if find_descendant(node, Ns::Draw, "plugin").is_some() {
        crate::frame::Kind::Plugin
    } else if find_descendant(node, Ns::Draw, "floating-frame").is_some() {
        crate::frame::Kind::FloatingFrame
    } else {
        crate::frame::Kind::Generic
    }
}

fn collect_rubies(node: &Node, output: &mut BodyStructures, tracked_scope: bool) -> Result<()> {
    if node.namespace == Ns::Text && node.local == "ruby" {
        let base = direct_child(node, Ns::Text, "ruby-base")
            .map(plain_text)
            .transpose()?
            .unwrap_or_default();
        let ruby_text = direct_child(node, Ns::Text, "ruby-text");
        let ruby = crate::ruby::Ruby::projected(
            attr(node, Ns::Text, b"style-name").map(str::to_owned),
            ruby_text.and_then(|value| attr(value, Ns::Text, b"style-name").map(str::to_owned)),
            base,
            ruby_text.map(plain_text).transpose()?.unwrap_or_default(),
        );
        let bytes = ruby_bytes(&ruby)?;
        output.account(bytes)?;
        reserve(&mut output.rubies, "OTH ruby projection")?;
        output.rubies.push(ruby);
    }
    collect_descendant_roots(node, 1, output, tracked_scope)?;
    Ok(())
}

fn elements(node: &Node) -> impl Iterator<Item = &Node> {
    node.children.iter().filter_map(|child| match child {
        Child::Element(element) => Some(element.as_ref()),
        Child::Text(_) => None,
    })
}

fn valid_structure_child(node: &Node) -> bool {
    if node.namespace == Ns::Other && node.namespace_uri.is_none() {
        return false;
    }
    context_kind(node.namespace, node.local.as_bytes()).permits_body_structure()
}

fn direct_child<'a>(node: &'a Node, namespace: Ns, local: &str) -> Option<&'a Node> {
    elements(node).find(|child| child.namespace == namespace && child.local == local)
}

fn descendant<'a>(node: &'a Node, namespace: Ns, local: &str) -> Option<&'a Node> {
    if node.namespace == namespace && node.local == local {
        return Some(node);
    }
    elements(node).find_map(|child| descendant(child, namespace, local))
}

fn find_descendant<'a>(node: &'a Node, namespace: Ns, local: &str) -> Option<&'a Node> {
    elements(node).find_map(|child| descendant(child, namespace, local))
}

fn attr<'a>(node: &'a Node, namespace: Ns, local: &[u8]) -> Option<&'a str> {
    node.attribute(namespace, local)
}

fn bool_attr(node: &Node, namespace: Ns, local: &[u8]) -> Result<bool> {
    Ok(optional_bool_attr(node, namespace, local)?.unwrap_or(false))
}

fn optional_bool_attr(node: &Node, namespace: Ns, local: &[u8]) -> Result<Option<bool>> {
    let Some(value) = attr(node, namespace, local) else {
        return Ok(None);
    };
    match value {
        "true" | "1" => Ok(Some(true)),
        "false" | "0" => Ok(Some(false)),
        _ => invalid("OTH boolean body attribute is invalid"),
    }
}

fn count_attr(node: &Node, local: &[u8]) -> Result<usize> {
    count_attr_default(node, local, 1)
}

fn count_attr_default(node: &Node, local: &[u8], default: usize) -> Result<usize> {
    let Some(value) = attr(node, Ns::Table, local) else {
        return Ok(default);
    };
    let parsed = value.parse::<usize>().map_err(|_| {
        Error::InvalidFormat("OTH table repetition attribute is invalid".to_string())
    })?;
    if parsed == 0 || parsed > MAX_REPEAT {
        return invalid("OTH table repetition attribute is outside the supported range");
    }
    Ok(parsed)
}

fn append_plain_text(output: &mut String, node: &Node) -> Result<()> {
    if matches!(node.namespace, Ns::Other | Ns::None) {
        return Ok(());
    }
    if node.namespace == Ns::Text {
        match node.local.as_str() {
            "s" => {
                let count = node.attribute(Ns::Text, b"c").map_or(Ok(1), |value| {
                    value.parse::<usize>().map_err(|_| {
                        Error::InvalidFormat("OTH text:s count is invalid".to_string())
                    })
                })?;
                if count == 0 || count > MAX_STRUCTURE_TEXT {
                    return invalid("OTH text:s count is outside the supported range");
                }
                checked_append_repeated(output, " ", count)?;
                return Ok(());
            },
            "tab" => {
                checked_append(output, "\t")?;
                return Ok(());
            },
            "line-break" => {
                checked_append(output, "\n")?;
                return Ok(());
            },
            _ => {},
        }
    }
    for child in &node.children {
        match child {
            Child::Element(element) => append_plain_text(output, element)?,
            Child::Text(text) => checked_append(output, text)?,
        }
    }
    Ok(())
}

fn plain_text(node: &Node) -> Result<String> {
    let mut output = String::new();
    append_plain_text(&mut output, node)?;
    Ok(output)
}

fn checked_append(output: &mut String, value: &str) -> Result<()> {
    let target = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    if target > MAX_STRUCTURE_TEXT {
        return invalid("OTH structure text exceeds the limit");
    }
    output
        .try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "OTH body structure text",
            source,
        })?;
    output.push_str(value);
    Ok(())
}

fn checked_append_repeated(output: &mut String, value: &str, count: usize) -> Result<()> {
    let additional = value
        .len()
        .checked_mul(count)
        .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    let target = output
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("OTH structure text size overflow".to_string()))?;
    if target > MAX_STRUCTURE_TEXT {
        return invalid("OTH structure text exceeds the limit");
    }
    output
        .try_reserve(additional)
        .map_err(|source| Error::Allocation {
            resource: "OTH body structure text",
            source,
        })?;
    for _ in 0..count {
        output.push_str(value);
    }
    Ok(())
}

fn add_string(total: &mut usize, value: Option<&str>) -> Result<()> {
    let Some(value) = value else { return Ok(()) };
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("OTH projected body size overflow".to_string()))?;
    Ok(())
}

fn reserve<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn reference_value(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(value) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!(
            "invalid OTH structure character reference: {error}"
        ))
    })? {
        return Ok(value.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid OTH structure entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "apos" => Ok("'".to_string()),
        "gt" => Ok(">".to_string()),
        "lt" => Ok("<".to_string()),
        "quot" => Ok("\"".to_string()),
        _ => invalid("OTH custom entities are not allowed"),
    }
}

/// Checks references outside an excluded structure using resolved namespaces
/// and XML/URI decoding.  A source-bound removal must make this check before
/// deleting the selected span; it never dereferences or fetches an external
/// URI.
pub(crate) fn has_semantic_reference(
    xml: &str,
    excluded: Range<usize>,
    identity: &str,
) -> Result<bool> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let event_start = source_offset(reader.buffer_position())?;
        let event = reader.read_event().map_err(|error| xml_error(&error))?;
        let event_end = source_offset(reader.buffer_position())?;
        let excluded_event = event_start >= excluded.start && event_end <= excluded.end;
        match event {
            Event::Start(start) | Event::Empty(start) if !excluded_event => {
                for raw in start.attributes() {
                    let attribute = raw.map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid OTH structural reference attribute: {error}"
                        ))
                    })?;
                    if attribute.key.as_ref().starts_with(b"xmlns") {
                        continue;
                    }
                    let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
                    let namespace = match resolved {
                        ResolveResult::Bound(Namespace(value)) => resolve_namespace(Some(value)),
                        ResolveResult::Unbound | ResolveResult::Unknown(_) => Ns::None,
                    };
                    if !reference_attribute(namespace, local.as_ref()) {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid OTH structural reference value: {error}"
                            ))
                        })?;
                    if semantic_reference_value(value.as_ref(), identity)? {
                        return Ok(true);
                    }
                }
            },
            Event::DocType(_) => return invalid("OTH content.xml cannot contain a DTD"),
            Event::Eof => return Ok(false),
            _ => {},
        }
    }
}

fn reference_attribute(namespace: Ns, local: &[u8]) -> bool {
    match namespace {
        Ns::Text => matches!(
            local,
            b"section-name" | b"change-id" | b"id" | b"name" | b"ref-name"
        ),
        Ns::Xlink => local == b"href",
        _ => false,
    }
}

fn semantic_reference_value(value: &str, identity: &str) -> Result<bool> {
    let Some(value) = percent_decode(value)? else {
        return Ok(false);
    };
    if value == identity {
        return Ok(true);
    }
    Ok(value
        .rsplit_once('#')
        .is_some_and(|(_, fragment)| fragment == identity))
}

fn percent_decode(value: &str) -> Result<Option<String>> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::new();
    decoded
        .try_reserve(bytes.len())
        .map_err(|source| Error::Allocation {
            resource: "OTH structural reference value",
            source,
        })?;
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            decoded.push(bytes[index]);
            index += 1;
            continue;
        }
        if index + 2 >= bytes.len() {
            return Ok(None);
        }
        let Some(high) = hex_digit(bytes[index + 1]) else {
            return Ok(None);
        };
        let Some(low) = hex_digit(bytes[index + 2]) else {
            return Ok(None);
        };
        decoded.push((high << 4) | low);
        index += 3;
    }
    Ok(String::from_utf8(decoded).ok())
}

fn hex_digit(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

/// Finds the lossless text spans admitted by the first structural write slice.
///
/// A selector is source-order within one namespace-qualified family.  The
/// scanner records every owned section, note, and annotation so empty or rich
/// structures cannot silently change selector numbering.  Only a single plain
/// `text:p`/`text:h` child is editable; all other markup remains untouched and
/// is refused when it would make that span ambiguous.
pub(crate) fn structure_sites(xml: &str) -> Result<Vec<StructureSite>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut contexts = Vec::<ContextKind>::new();
    let mut target_ids = Vec::<Option<usize>>::new();
    let mut targets = Vec::<Option<ActiveSite>>::new();
    let mut counts = [0_usize; 4];
    let mut sites = Vec::new();
    let mut active_block = None::<ActiveBlockSite>;
    let mut budget = Budget {
        nodes: 0,
        text_bytes: 0,
    };

    loop {
        let event_start = source_offset(reader.buffer_position())?;
        match reader.read_event().map_err(|error| xml_error(&error))? {
            Event::Start(start) => {
                budget.node()?;
                if contexts.len() >= MAX_STRUCTURE_SITE_DEPTH {
                    return invalid("OTH structural edit depth exceeds the limit");
                }
                let (_namespace, context, root) = classify_element(&reader, start.name());
                if let Some(active) = active_block.as_mut() {
                    active.invalid = true;
                }
                let root_id = match editable_kind(root) {
                    Some(kind) if owns_structure_root(&contexts) => {
                        let index = kind.index(&mut counts)?;
                        targets.try_reserve(1).map_err(|source| Error::Allocation {
                            resource: "OTH structural edit sites",
                            source,
                        })?;
                        targets.push(Some(ActiveSite {
                            full_start: event_start,
                            kind,
                            index,
                            identity: structure_identity(&reader, &start, kind, &mut budget)?,
                            root_depth: contexts.len(),
                            block_count: 0,
                            invalid: false,
                            text: None,
                            text_value: None,
                        }));
                        Some(targets.len() - 1)
                    },
                    _ => None,
                };
                let block = if matches!(
                    context,
                    ContextKind::TextParagraph | ContextKind::TextHeading
                ) {
                    if let Some(target_id) = block_owner(&targets, &contexts) {
                        if let Some(target) = targets.get_mut(target_id).and_then(Option::as_mut) {
                            target.block_count =
                                target.block_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "OTH structural edit block count overflow".to_string(),
                                    )
                                })?;
                            if target.block_count > 1 {
                                target.invalid = true;
                            }
                        }
                        Some(ActiveBlockSite {
                            target_id,
                            depth: contexts.len(),
                            content_start: source_offset(reader.buffer_position())?,
                            text: String::new(),
                            invalid: false,
                        })
                    } else {
                        None
                    }
                } else {
                    None
                };
                active_block = block;
                contexts.push(context);
                target_ids.push(root_id);
            },
            Event::Empty(start) => {
                budget.node()?;
                if contexts.len() >= MAX_STRUCTURE_SITE_DEPTH {
                    return invalid("OTH structural edit depth exceeds the limit");
                }
                let (_namespace, context, root) = classify_element(&reader, start.name());
                if let Some(active) = active_block.as_mut() {
                    active.invalid = true;
                }
                if let Some(kind) = editable_kind(root)
                    && owns_structure_root(&contexts)
                {
                    let index = kind.index(&mut counts)?;
                    sites.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "OTH structural edit sites",
                        source,
                    })?;
                    sites.push(StructureSite {
                        kind,
                        index,
                        identity: structure_identity(&reader, &start, kind, &mut budget)?,
                        full: event_start..source_offset(reader.buffer_position())?,
                        text: None,
                        text_value: None,
                    });
                } else if matches!(
                    context,
                    ContextKind::TextParagraph | ContextKind::TextHeading
                ) && let Some(target_id) = block_owner(&targets, &contexts)
                    && let Some(target) = targets.get_mut(target_id).and_then(Option::as_mut)
                {
                    target.block_count = target.block_count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("OTH structural edit block count overflow".to_string())
                    })?;
                    target.invalid = true;
                    target.text = None;
                    target.text_value = None;
                }
            },
            Event::Text(text) => {
                if let Some(active) = active_block.as_mut() {
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid OTH structural edit text: {error}"))
                    })?;
                    budget.text(value.len())?;
                    checked_append(&mut active.text, &value)?;
                }
            },
            Event::GeneralRef(reference) if active_block.is_some() => {
                let value = reference_value(&reference)?;
                budget.text(value.len())?;
                checked_append(
                    &mut active_block
                        .as_mut()
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "OTH structural block state is missing".to_string(),
                            )
                        })?
                        .text,
                    &value,
                )?;
            },
            Event::End(end) => {
                let (_namespace, context, _root) = classify_element(&reader, end.name());
                if matches!(
                    context,
                    ContextKind::TextParagraph | ContextKind::TextHeading
                ) && let Some(active) = active_block.take()
                {
                    if active.depth != contexts.len().saturating_sub(1) {
                        return invalid("OTH structural edit text depth is invalid");
                    }
                    if let Some(target) = targets.get_mut(active.target_id).and_then(Option::as_mut)
                    {
                        if !active.invalid && !target.invalid {
                            target.text = Some(crate::codec::ReplacementSite {
                                prefix: String::new(),
                                range: active.content_start..event_start,
                                suffix: String::new(),
                            });
                            target.text_value = Some(active.text);
                        } else {
                            target.text = None;
                            target.text_value = None;
                        }
                    }
                }
                let Some(open_context) = contexts.pop() else {
                    return invalid("OTH structural edit end tag has no matching start tag");
                };
                if open_context != context {
                    return invalid("OTH structural edit end tag does not match its start tag");
                }
                let root_id = target_ids.pop().ok_or_else(|| {
                    Error::InvalidFormat("OTH structural edit root stack is missing".to_string())
                })?;
                if let Some(root_id) = root_id {
                    let target = targets[root_id].take().ok_or_else(|| {
                        Error::InvalidFormat(
                            "OTH structural edit root state is missing".to_string(),
                        )
                    })?;
                    sites.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "OTH structural edit sites",
                        source,
                    })?;
                    sites.push(StructureSite {
                        kind: target.kind,
                        index: target.index,
                        identity: target.identity,
                        full: target.full_start..source_offset(reader.buffer_position())?,
                        text: target.text,
                        text_value: target.text_value,
                    });
                }
            },
            Event::CData(_) | Event::Comment(_) | Event::PI(_) if active_block.is_some() => {
                if let Some(active) = active_block.as_mut() {
                    active.invalid = true;
                }
            },
            Event::DocType(_) => return invalid("OTH structural edit source cannot contain a DTD"),
            Event::Decl(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::GeneralRef(_)
            | Event::PI(_) => {},
            Event::Eof => {
                if !contexts.is_empty() || active_block.is_some() {
                    return invalid("OTH structural edit source is not closed");
                }
                sites.sort_unstable_by_key(|site| site.full.start);
                return Ok(sites);
            },
        }
    }
}

#[derive(Clone)]
struct ActiveBlockSite {
    target_id: usize,
    depth: usize,
    content_start: usize,
    // The text is moved into the target when the block closes.  Keeping this
    // allocation local bounds a rejected rich block to one selected span.
    text: String,
    invalid: bool,
}

struct ActiveSite {
    full_start: usize,
    kind: EditableStructureKind,
    index: usize,
    identity: Option<String>,
    root_depth: usize,
    block_count: usize,
    invalid: bool,
    text: Option<crate::codec::ReplacementSite>,
    text_value: Option<String>,
}

impl EditableStructureKind {
    pub(crate) const fn to_public(self) -> crate::facade::StructureKind {
        match self {
            Self::Section => crate::facade::StructureKind::Section,
            Self::Note => crate::facade::StructureKind::Note,
            Self::Annotation => crate::facade::StructureKind::Annotation,
            Self::Table => crate::facade::StructureKind::Table,
        }
    }

    fn index(self, counts: &mut [usize; 4]) -> Result<usize> {
        let slot = match self {
            Self::Section => &mut counts[0],
            Self::Note => &mut counts[1],
            Self::Annotation => &mut counts[2],
            Self::Table => &mut counts[3],
        };
        let index = *slot;
        *slot = slot.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("OTH structural edit selector overflow".to_string())
        })?;
        if *slot > MAX_STRUCTURE_NODES {
            return invalid("OTH structural edit selector count exceeds the limit");
        }
        Ok(index)
    }
}

fn editable_kind(kind: Option<RootKind>) -> Option<EditableStructureKind> {
    match kind {
        Some(RootKind::Section) => Some(EditableStructureKind::Section),
        Some(RootKind::Note) => Some(EditableStructureKind::Note),
        Some(RootKind::Annotation) => Some(EditableStructureKind::Annotation),
        Some(RootKind::Table) => Some(EditableStructureKind::Table),
        _ => None,
    }
}

fn block_owner(targets: &[Option<ActiveSite>], contexts: &[ContextKind]) -> Option<usize> {
    targets.iter().enumerate().rev().find_map(|(id, target)| {
        let target = target.as_ref()?;
        let direct = match target.kind {
            EditableStructureKind::Section | EditableStructureKind::Annotation => {
                contexts.len() == target.root_depth + 1
            },
            EditableStructureKind::Note => {
                contexts.len() == target.root_depth + 2
                    && contexts.get(target.root_depth + 1) == Some(&ContextKind::TextNoteBody)
            },
            EditableStructureKind::Table => false,
        };
        direct.then_some(id)
    })
}

fn structure_identity(
    reader: &NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    kind: EditableStructureKind,
    budget: &mut Budget,
) -> Result<Option<String>> {
    let (namespace, local) = match kind {
        EditableStructureKind::Section => (Ns::Text, b"name".as_slice()),
        EditableStructureKind::Note => (Ns::Text, b"id".as_slice()),
        EditableStructureKind::Annotation => (Ns::Office, b"name".as_slice()),
        EditableStructureKind::Table => (Ns::Table, b"name".as_slice()),
    };
    for raw in start.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid OTH structural identity attribute: {error}"
            ))
        })?;
        if attribute.key.as_ref().starts_with(b"xmlns") {
            continue;
        }
        let (resolved, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let resolved = match resolved {
            ResolveResult::Bound(Namespace(value)) => resolve_namespace(Some(value)),
            ResolveResult::Unbound | ResolveResult::Unknown(_) => Ns::None,
        };
        if resolved != namespace || local_name.as_ref() != local {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid OTH structural identity value: {error}"))
            })?;
        budget.text(value.len())?;
        return Ok((!value.is_empty()).then(|| value.into_owned()));
    }
    Ok(None)
}

fn source_offset(offset: u64) -> Result<usize> {
    usize::try_from(offset)
        .map_err(|error| Error::InvalidFormat(format!("OTH source offset overflow: {error}")))
}

fn xml_error(error: &quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("invalid OTH body structure XML: {error}"))
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_string()))
}
