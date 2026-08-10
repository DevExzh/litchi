//! Document metadata fields and validated mixed-content trees.

#![allow(
    clippy::wildcard_imports,
    reason = "semantic field owners share the stable model facade namespace"
)]
use super::*;
use std::collections::HashSet;
/// One of the eight temporal/revision ODF document-metadata fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MetadataFieldKind {
    CreationDate,
    CreationTime,
    PrintDate,
    PrintTime,
    EditingCycles,
    EditingDuration,
    ModificationDate,
    ModificationTime,
}

impl MetadataFieldKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::CreationDate => "text:creation-date",
            Self::CreationTime => "text:creation-time",
            Self::PrintDate => "text:print-date",
            Self::PrintTime => "text:print-time",
            Self::EditingCycles => "text:editing-cycles",
            Self::EditingDuration => "text:editing-duration",
            Self::ModificationDate => "text:modification-date",
            Self::ModificationTime => "text:modification-time",
        }
    }

    pub(crate) const fn permits_data_style(self) -> bool {
        !matches!(self, Self::EditingCycles)
    }
}

/// Strict typed value attribute for a temporal document-metadata field.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum MetadataFieldValue {
    Date(FieldDateValue),
    Time(FieldTimeValue),
    Duration(FieldDuration),
}

/// One of the nine fixed string/identity document-metadata fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum IdentityFieldKind {
    InitialCreator,
    Description,
    PrintedBy,
    Title,
    Subject,
    Keywords,
    Creator,
    /// The cached full name of the document author from ODF 1.2 §7.3.7.1.
    AuthorName,
    /// The cached initials of the document author from ODF 1.2 §7.3.7.2.
    AuthorInitials,
}

impl IdentityFieldKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::InitialCreator => "text:initial-creator",
            Self::Description => "text:description",
            Self::PrintedBy => "text:printed-by",
            Self::Title => "text:title",
            Self::Subject => "text:subject",
            Self::Keywords => "text:keywords",
            Self::Creator => "text:creator",
            Self::AuthorName => "text:author-name",
            Self::AuthorInitials => "text:author-initials",
        }
    }
}

/// One of the fifteen ODF 1.2 subsequent-author `text:sender-*` field categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SenderFieldKind {
    FirstName,
    LastName,
    Initials,
    Title,
    Position,
    Email,
    PrivatePhone,
    Fax,
    Company,
    WorkPhone,
    Street,
    City,
    PostalCode,
    Country,
    StateOrProvince,
}

impl SenderFieldKind {
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::FirstName => "text:sender-firstname",
            Self::LastName => "text:sender-lastname",
            Self::Initials => "text:sender-initials",
            Self::Title => "text:sender-title",
            Self::Position => "text:sender-position",
            Self::Email => "text:sender-email",
            Self::PrivatePhone => "text:sender-phone-private",
            Self::Fax => "text:sender-fax",
            Self::Company => "text:sender-company",
            Self::WorkPhone => "text:sender-phone-work",
            Self::Street => "text:sender-street",
            Self::City => "text:sender-city",
            Self::PostalCode => "text:sender-postal-code",
            Self::Country => "text:sender-country",
            Self::StateOrProvince => "text:sender-state-or-province",
        }
    }
}

/// Independently optional cached values permitted by `text:user-defined`.
///
/// Unlike variable fields, ODF 1.2 does not use `office:value-type` here and
/// its schema permits more than one of these attributes to coexist.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct UserDefinedMetadataValues {
    pub number: Option<String>,
    pub date: Option<FieldDateValue>,
    pub time: Option<FieldDuration>,
    pub boolean: Option<bool>,
    pub string: Option<String>,
}

/// A namespace-resolved attribute on inert `text:meta-field` content.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetaFieldAttribute {
    pub namespace_uri: String,
    pub local_name: String,
    pub value: String,
}

/// A namespace-resolved inert inline element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetaFieldElement {
    pub namespace_uri: String,
    pub local_name: String,
    pub attributes: Vec<MetaFieldAttribute>,
    pub children: Vec<MetaFieldNode>,
}

/// Ordered mixed content retained by `text:meta-field`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum MetaFieldNode {
    Text(String),
    Element(MetaFieldElement),
}

/// Validated, inert mixed content with a cached plain-text projection.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetaFieldContent {
    nodes: Vec<MetaFieldNode>,
    display_text: String,
}

impl MetaFieldContent {
    pub fn new(nodes: Vec<MetaFieldNode>) -> Result<Self> {
        let display_text =
            validated_meta_display_text(&nodes, MetaContentGrammar::ParagraphOrHyperlink)?;
        Ok(Self {
            nodes,
            display_text,
        })
    }

    pub fn nodes(&self) -> &[MetaFieldNode] {
        &self.nodes
    }

    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    pub(super) fn write_xml(&self, output: &mut String) {
        for node in &self.nodes {
            write_meta_node(node, output);
        }
    }
}

/// Validated, inert structured content for an ODF `text:note-body`.
///
/// The ODF 1.3 schema permits paragraph-like blocks, lists, tables, selected
/// drawing content, and related structured text descendants in a note body.
/// This models ODF 1.3 Part 3, section 6.3.4. Direct character data is
/// deliberately rejected: it belongs inside one of those schema-defined child
/// elements. Links, fields, event listeners, and macro metadata are serialized
/// only as inert XML; this type never follows, evaluates, or executes them.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NoteBodyContent {
    nodes: Vec<MetaFieldNode>,
    display_text: String,
}

impl NoteBodyContent {
    /// Construct structured note-body content from namespace-resolved nodes.
    pub fn new(nodes: Vec<MetaFieldNode>) -> Result<Self> {
        if nodes
            .iter()
            .any(|node| matches!(node, MetaFieldNode::Text(_)))
        {
            return Err(Error::InvalidFormat(
                "text:note-body cannot contain direct character data".to_string(),
            ));
        }
        validated_meta_display_text(&nodes, MetaContentGrammar::NoteBody)?;
        let display_text = note_body_display_text(&nodes)?;
        Ok(Self {
            nodes,
            display_text,
        })
    }

    /// Return the ordered, namespace-resolved note-body nodes.
    pub fn nodes(&self) -> &[MetaFieldNode] {
        &self.nodes
    }

    /// Return a cached visible-text projection of the structured note body.
    ///
    /// Paragraph and heading descendants are separated by line feeds. Nested
    /// note bodies are omitted from an enclosing note's projection, while a
    /// nested note's citation remains inline. `text:s`, `text:tab`, and
    /// `text:line-break` receive their corresponding text semantics, matching
    /// the bounded semantic note reader.
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Return whether this body has no schema child elements.
    pub fn is_empty(&self) -> bool {
        self.nodes.is_empty()
    }

    /// Revalidate this value's bounded XML and resource constraints.
    pub fn validate(&self) -> Result<()> {
        Self::new(self.nodes.clone()).map(|_| ())
    }

    pub(crate) fn write_xml(&self, output: &mut String) {
        for node in &self.nodes {
            write_meta_node(node, output);
        }
    }
}

fn validated_meta_display_text(
    nodes: &[MetaFieldNode],
    grammar: MetaContentGrammar,
) -> Result<String> {
    let mut aggregate = 0usize;
    let mut node_count = 0usize;
    let mut display_text = String::new();
    validate_meta_nodes(
        nodes,
        0,
        grammar,
        &mut aggregate,
        &mut node_count,
        &mut display_text,
    )?;
    Ok(display_text)
}

fn note_body_display_text(nodes: &[MetaFieldNode]) -> Result<String> {
    let mut display_text = String::new();
    let mut seen_block = false;
    append_note_body_display_text(nodes, &mut display_text, &mut seen_block, false)?;
    Ok(display_text)
}

fn append_note_body_display_text(
    nodes: &[MetaFieldNode],
    display_text: &mut String,
    seen_block: &mut bool,
    in_paragraph: bool,
) -> Result<()> {
    for node in nodes {
        match node {
            MetaFieldNode::Text(value) if in_paragraph => {
                append_note_body_display_value(display_text, value)?;
            },
            MetaFieldNode::Text(_) => {},
            MetaFieldNode::Element(element) => {
                if element.namespace_uri == TEXT_DATABASE_NAMESPACE && element.local_name == "note"
                {
                    if in_paragraph
                        && let Some(MetaFieldNode::Element(citation)) = element.children.first()
                        && citation.namespace_uri == TEXT_DATABASE_NAMESPACE
                        && citation.local_name == "note-citation"
                    {
                        append_note_body_display_text(
                            &citation.children,
                            display_text,
                            seen_block,
                            true,
                        )?;
                    }
                    continue;
                }
                if element.namespace_uri == TEXT_DATABASE_NAMESPACE
                    && matches!(element.local_name.as_str(), "p" | "h")
                {
                    if !in_paragraph {
                        if *seen_block {
                            append_note_body_display_value(display_text, "\n")?;
                        }
                        *seen_block = true;
                    }
                    append_note_body_display_text(
                        &element.children,
                        display_text,
                        seen_block,
                        true,
                    )?;
                    continue;
                }
                if in_paragraph && element.namespace_uri == TEXT_DATABASE_NAMESPACE {
                    match element.local_name.as_str() {
                        "s" => {
                            append_note_body_spaces(display_text, element)?;
                            continue;
                        },
                        "tab" => {
                            append_note_body_display_value(display_text, "\t")?;
                            continue;
                        },
                        "line-break" => {
                            append_note_body_display_value(display_text, "\n")?;
                            continue;
                        },
                        _ => {},
                    }
                }
                append_note_body_display_text(
                    &element.children,
                    display_text,
                    seen_block,
                    in_paragraph,
                )?;
            },
        }
    }
    Ok(())
}

fn append_note_body_display_value(output: &mut String, value: &str) -> Result<()> {
    let total = output.len().checked_add(value.len()).ok_or_else(|| {
        Error::InvalidFormat("text:note-body display text size overflow".to_string())
    })?;
    if total > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "text:note-body display text exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} bytes"
        )));
    }
    output.push_str(value);
    Ok(())
}

fn append_note_body_spaces(output: &mut String, element: &MetaFieldElement) -> Result<()> {
    let count = element
        .attributes
        .iter()
        .find(|attribute| {
            attribute.namespace_uri == TEXT_DATABASE_NAMESPACE && attribute.local_name == "c"
        })
        .map(|attribute| {
            attribute.value.parse::<usize>().map_err(|_error| {
                Error::InvalidFormat("text:s text:c must be a non-negative integer".to_string())
            })
        })
        .transpose()?
        .unwrap_or(1);
    let total = output.len().checked_add(count).ok_or_else(|| {
        Error::InvalidFormat("text:note-body display text size overflow".to_string())
    })?;
    if total > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "text:note-body display text exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} bytes"
        )));
    }
    output.extend(std::iter::repeat_n(' ', count));
    Ok(())
}

impl UserDefinedMetadataValues {
    pub(super) fn validate(&self, aggregate: &mut usize) -> Result<()> {
        if let Some(number) = &self.number {
            validate_double(number)?;
            validate_dynamic_value("office:value", Some(number), true, aggregate)?;
        }
        if let Some(date) = &self.date {
            date.validate(aggregate)?;
        }
        if let Some(time) = &self.time {
            time.validate("office:time-value", aggregate)?;
        }
        validate_dynamic_value(
            "office:string-value",
            self.string.as_deref(),
            false,
            aggregate,
        )
    }

    pub(super) fn write_attributes(&self, element: &mut Element) {
        if self.number.is_none()
            && self.date.is_none()
            && self.time.is_none()
            && self.boolean.is_none()
            && self.string.is_none()
        {
            return;
        }
        element.set_attribute("xmlns:office", OFFICE_NAMESPACE);
        if let Some(number) = &self.number {
            element.set_attribute("office:value", number);
        }
        if let Some(date) = &self.date {
            element.set_attribute("office:date-value", date.as_str());
        }
        if let Some(time) = &self.time {
            element.set_attribute("office:time-value", time.as_str());
        }
        if let Some(boolean) = self.boolean {
            element.set_attribute(
                "office:boolean-value",
                if boolean { "true" } else { "false" },
            );
        }
        if let Some(string) = &self.string {
            element.set_attribute("office:string-value", string);
        }
    }
}

pub(crate) fn validate_document_metadata_value(
    kind: MetadataFieldKind,
    value: Option<&MetadataFieldValue>,
    aggregate: &mut usize,
) -> Result<()> {
    match (kind, value) {
        (_, None) => Ok(()),
        (MetadataFieldKind::CreationDate, Some(MetadataFieldValue::Date(value))) => {
            value.validate(aggregate)
        },
        (MetadataFieldKind::CreationTime, Some(MetadataFieldValue::Time(value))) => {
            value.validate(aggregate)
        },
        (
            MetadataFieldKind::PrintDate | MetadataFieldKind::ModificationDate,
            Some(MetadataFieldValue::Date(value)),
        ) if value.kind() == DateValueKind::Date => value.validate(aggregate),
        (
            MetadataFieldKind::PrintTime | MetadataFieldKind::ModificationTime,
            Some(MetadataFieldValue::Time(value)),
        ) if value.kind() == TimeValueKind::Time => value.validate(aggregate),
        (MetadataFieldKind::EditingDuration, Some(MetadataFieldValue::Duration(value))) => {
            value.validate("text:duration", aggregate)
        },
        _ => Err(Error::InvalidFormat(format!(
            "value type is not permitted by {}",
            kind.element_name()
        ))),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum MetaContentGrammar {
    ParagraphOrHyperlink,
    Paragraph,
    TextOnly,
    DropDown,
    Empty,
    Hyperlink,
    Ruby,
    Note,
    NoteBody,
    ExecuteMacro,
    EventListeners,
    PresentationEventListener,
    Annotation,
    Structured,
    ShapeBasic,
    ShapeGroup,
    ShapeFrame,
    ShapeLink,
}

fn validate_meta_nodes(
    nodes: &[MetaFieldNode],
    depth: usize,
    grammar: MetaContentGrammar,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    if depth > MAX_META_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "text:meta-field content exceeds {MAX_META_FIELD_DEPTH} levels"
        )));
    }
    match grammar {
        MetaContentGrammar::DropDown => {
            return validate_meta_drop_down(nodes, depth, aggregate, node_count, display_text);
        },
        MetaContentGrammar::Ruby => {
            return validate_meta_exact_pair(
                nodes,
                depth,
                (
                    TEXT_DATABASE_NAMESPACE,
                    "ruby-base",
                    MetaContentGrammar::ParagraphOrHyperlink,
                ),
                (
                    TEXT_DATABASE_NAMESPACE,
                    "ruby-text",
                    MetaContentGrammar::TextOnly,
                ),
                "text:ruby",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::Note => {
            return validate_meta_exact_pair(
                nodes,
                depth,
                (
                    TEXT_DATABASE_NAMESPACE,
                    "note-citation",
                    MetaContentGrammar::TextOnly,
                ),
                (
                    TEXT_DATABASE_NAMESPACE,
                    "note-body",
                    MetaContentGrammar::NoteBody,
                ),
                "text:note",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::Hyperlink => {
            return validate_meta_optional_listener_then(
                nodes,
                depth,
                MetaContentGrammar::Paragraph,
                "text:a",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::ExecuteMacro => {
            return validate_meta_optional_listener_then(
                nodes,
                depth,
                MetaContentGrammar::TextOnly,
                "text:execute-macro",
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::EventListeners => {
            return validate_meta_event_listeners(
                nodes,
                depth,
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::PresentationEventListener => {
            if nodes.is_empty() {
                return Ok(());
            }
            if nodes.len() != 1 {
                return Err(Error::InvalidFormat(
                    "presentation:event-listener permits at most one presentation:sound"
                        .to_string(),
                ));
            }
            return validate_meta_required_element(
                &nodes[0],
                depth,
                PRESENTATION_NAMESPACE,
                "sound",
                MetaContentGrammar::Empty,
                aggregate,
                node_count,
                display_text,
            );
        },
        MetaContentGrammar::Annotation => {
            return validate_meta_annotation(nodes, depth, aggregate, node_count, display_text);
        },
        MetaContentGrammar::ShapeLink => {
            if nodes.len() != 1 {
                return Err(Error::InvalidFormat(
                    "draw:a requires exactly one drawing shape".to_string(),
                ));
            }
            let MetaFieldNode::Element(element) = &nodes[0] else {
                return Err(Error::InvalidFormat(
                    "draw:a requires a drawing shape child".to_string(),
                ));
            };
            let grammar = odf_shape_grammar(&element.namespace_uri, &element.local_name)
                .ok_or_else(|| {
                    Error::InvalidFormat("draw:a child is not a drawing shape".to_string())
                })?;
            return validate_meta_required_element(
                &nodes[0],
                depth,
                &element.namespace_uri,
                &element.local_name,
                grammar,
                aggregate,
                node_count,
                display_text,
            );
        },
        _ => {},
    }
    for node in nodes {
        *node_count = node_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("text:meta-field node count overflow".to_string())
        })?;
        if *node_count > MAX_META_FIELD_NODES {
            return Err(Error::InvalidFormat(format!(
                "text:meta-field exceeds {MAX_META_FIELD_NODES} content nodes"
            )));
        }
        match node {
            MetaFieldNode::Text(value) => {
                if matches!(grammar, MetaContentGrammar::Empty) {
                    return Err(Error::InvalidFormat(
                        "ODF empty inline element contains character data".to_string(),
                    ));
                }
                validate_dynamic_value("meta-field text", Some(value), false, aggregate)?;
                display_text.push_str(value);
            },
            MetaFieldNode::Element(element) => {
                validate_meta_element_parts(
                    &element.namespace_uri,
                    &element.local_name,
                    &element.attributes,
                    aggregate,
                )?;
                let child_grammar =
                    meta_child_grammar(grammar, &element.namespace_uri, &element.local_name)?;
                validate_meta_element_attributes_for_grammar(element, child_grammar)?;
                validate_meta_nodes(
                    &element.children,
                    depth + 1,
                    child_grammar,
                    aggregate,
                    node_count,
                    display_text,
                )?;
            },
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn validate_meta_exact_pair(
    nodes: &[MetaFieldNode],
    depth: usize,
    first: (&str, &str, MetaContentGrammar),
    second: (&str, &str, MetaContentGrammar),
    owner: &str,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    if nodes.len() != 2 {
        return Err(Error::InvalidFormat(format!(
            "{owner} requires exactly two schema-ordered child elements"
        )));
    }
    validate_meta_required_element(
        &nodes[0],
        depth,
        first.0,
        first.1,
        first.2,
        aggregate,
        node_count,
        display_text,
    )?;
    validate_meta_required_element(
        &nodes[1],
        depth,
        second.0,
        second.1,
        second.2,
        aggregate,
        node_count,
        display_text,
    )
}

#[allow(clippy::too_many_arguments)]
fn validate_meta_required_element(
    node: &MetaFieldNode,
    depth: usize,
    namespace: &str,
    local: &str,
    child_grammar: MetaContentGrammar,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let MetaFieldNode::Element(element) = node else {
        return Err(Error::InvalidFormat(format!(
            "expected {namespace}:{local} element in structured metadata content"
        )));
    };
    if element.namespace_uri != namespace || element.local_name != local {
        return Err(Error::InvalidFormat(format!(
            "expected {namespace}:{local} in structured metadata content"
        )));
    }
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("text:meta-field node count overflow".to_string()))?;
    if *node_count > MAX_META_FIELD_NODES {
        return Err(Error::InvalidFormat(format!(
            "text:meta-field exceeds {MAX_META_FIELD_NODES} content nodes"
        )));
    }
    validate_meta_element_parts(namespace, local, &element.attributes, aggregate)?;
    validate_meta_element_attributes_for_grammar(element, child_grammar)?;
    validate_meta_nodes(
        &element.children,
        depth + 1,
        child_grammar,
        aggregate,
        node_count,
        display_text,
    )
}

fn validate_meta_element_attributes_for_grammar(
    element: &MetaFieldElement,
    grammar: MetaContentGrammar,
) -> Result<()> {
    if grammar == MetaContentGrammar::DropDown {
        validate_meta_drop_down_attributes(&element.attributes)?;
    }
    Ok(())
}

fn validate_meta_drop_down(
    nodes: &[MetaFieldNode],
    depth: usize,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let mut display_started = false;
    let mut labels = 0usize;
    for node in nodes {
        *node_count = node_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("text:meta-field node count overflow".to_string())
        })?;
        if *node_count > MAX_META_FIELD_NODES {
            return Err(Error::InvalidFormat(format!(
                "text:meta-field exceeds {MAX_META_FIELD_NODES} content nodes"
            )));
        }
        match node {
            MetaFieldNode::Text(value) => {
                display_started = true;
                validate_dynamic_value("meta-field drop-down text", Some(value), false, aggregate)?;
                display_text.push_str(value);
            },
            MetaFieldNode::Element(element) => {
                if display_started
                    || element.namespace_uri != TEXT_DATABASE_NAMESPACE
                    || element.local_name != "label"
                {
                    return Err(Error::InvalidFormat(
                        "text:drop-down permits only leading text:label children".to_string(),
                    ));
                }
                if !element.children.is_empty() {
                    return Err(Error::InvalidFormat(
                        "text:label must be empty in text:drop-down".to_string(),
                    ));
                }
                labels = labels.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("text:drop-down label count overflow".to_string())
                })?;
                if labels > MAX_DROP_DOWN_LABELS {
                    return Err(Error::InvalidFormat(format!(
                        "text:drop-down exceeds {MAX_DROP_DOWN_LABELS} labels"
                    )));
                }
                validate_meta_element_parts(
                    &element.namespace_uri,
                    &element.local_name,
                    &element.attributes,
                    aggregate,
                )?;
                validate_meta_drop_down_label_attributes(&element.attributes)?;
            },
        }
    }
    if depth > MAX_META_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "text:meta-field content exceeds {MAX_META_FIELD_DEPTH} levels"
        )));
    }
    Ok(())
}

fn validate_meta_drop_down_attributes(attributes: &[MetaFieldAttribute]) -> Result<()> {
    let has_name = attributes.iter().any(|attribute| {
        attribute.namespace_uri == TEXT_DATABASE_NAMESPACE && attribute.local_name == "name"
    });
    if !has_name {
        return Err(Error::InvalidFormat(
            "text:drop-down requires text:name".to_string(),
        ));
    }
    if attributes.iter().any(|attribute| {
        attribute.namespace_uri != TEXT_DATABASE_NAMESPACE || attribute.local_name != "name"
    }) {
        return Err(Error::InvalidFormat(
            "text:drop-down only permits text:name".to_string(),
        ));
    }
    Ok(())
}

fn validate_meta_drop_down_label_attributes(attributes: &[MetaFieldAttribute]) -> Result<()> {
    for attribute in attributes {
        if attribute.namespace_uri != TEXT_DATABASE_NAMESPACE
            || !matches!(attribute.local_name.as_str(), "value" | "current-selected")
        {
            return Err(Error::InvalidFormat(
                "text:label has an unsupported attribute".to_string(),
            ));
        }
        if attribute.local_name == "current-selected" {
            parse_drop_down_boolean(&attribute.value)?;
        }
    }
    Ok(())
}

fn validate_meta_optional_listener_then(
    nodes: &[MetaFieldNode],
    depth: usize,
    remaining_grammar: MetaContentGrammar,
    owner: &str,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let listener_position = nodes.iter().position(|node| {
        matches!(node, MetaFieldNode::Element(element)
            if element.namespace_uri == OFFICE_NAMESPACE && element.local_name == "event-listeners")
    });
    let start = match listener_position {
        None => 0,
        Some(0) => {
            validate_meta_required_element(
                &nodes[0],
                depth,
                OFFICE_NAMESPACE,
                "event-listeners",
                MetaContentGrammar::EventListeners,
                aggregate,
                node_count,
                display_text,
            )?;
            1
        },
        Some(_) => {
            return Err(Error::InvalidFormat(format!(
                "office:event-listeners must be the first child of {owner}"
            )));
        },
    };
    validate_meta_nodes(
        &nodes[start..],
        depth,
        remaining_grammar,
        aggregate,
        node_count,
        display_text,
    )
}

fn validate_meta_event_listeners(
    nodes: &[MetaFieldNode],
    depth: usize,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    for node in nodes {
        let MetaFieldNode::Element(element) = node else {
            return Err(Error::InvalidFormat(
                "office:event-listeners cannot contain character data".to_string(),
            ));
        };
        let child_grammar = match (element.namespace_uri.as_str(), element.local_name.as_str()) {
            (SCRIPT_NAMESPACE, "event-listener") => MetaContentGrammar::Empty,
            (PRESENTATION_NAMESPACE, "event-listener") => {
                MetaContentGrammar::PresentationEventListener
            },
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "{}:{} is not an ODF event listener",
                    element.namespace_uri, element.local_name
                )));
            },
        };
        validate_meta_required_element(
            node,
            depth,
            &element.namespace_uri,
            &element.local_name,
            child_grammar,
            aggregate,
            node_count,
            display_text,
        )?;
    }
    Ok(())
}

fn validate_meta_annotation(
    nodes: &[MetaFieldNode],
    depth: usize,
    aggregate: &mut usize,
    node_count: &mut usize,
    display_text: &mut String,
) -> Result<()> {
    let metadata = [
        (DC_NAMESPACE, "creator"),
        (DC_NAMESPACE, "date"),
        (META_NAMESPACE, "date-string"),
    ];
    let mut position = 0usize;
    for (namespace, local) in metadata {
        if matches!(nodes.get(position), Some(MetaFieldNode::Element(element))
            if element.namespace_uri == namespace && element.local_name == local)
        {
            validate_meta_required_element(
                &nodes[position],
                depth,
                namespace,
                local,
                MetaContentGrammar::TextOnly,
                aggregate,
                node_count,
                display_text,
            )?;
            position += 1;
        }
    }
    for node in &nodes[position..] {
        let MetaFieldNode::Element(element) = node else {
            return Err(Error::InvalidFormat(
                "office:annotation only permits metadata followed by text:p or text:list"
                    .to_string(),
            ));
        };
        let grammar = match (element.namespace_uri.as_str(), element.local_name.as_str()) {
            (TEXT_DATABASE_NAMESPACE, "p") => MetaContentGrammar::ParagraphOrHyperlink,
            (TEXT_DATABASE_NAMESPACE, "list") => MetaContentGrammar::Structured,
            _ => {
                return Err(Error::InvalidFormat(
                    "office:annotation only permits metadata followed by text:p or text:list"
                        .to_string(),
                ));
            },
        };
        validate_meta_required_element(
            node,
            depth,
            &element.namespace_uri,
            &element.local_name,
            grammar,
            aggregate,
            node_count,
            display_text,
        )?;
    }
    Ok(())
}

pub(crate) fn meta_child_grammar(
    parent: MetaContentGrammar,
    namespace: &str,
    local: &str,
) -> Result<MetaContentGrammar> {
    if matches!(
        parent,
        MetaContentGrammar::TextOnly | MetaContentGrammar::Empty
    ) {
        return Err(Error::InvalidFormat(format!(
            "{namespace}:{local} is not permitted inside a text-only or empty ODF element"
        )));
    }
    if parent == MetaContentGrammar::NoteBody {
        return note_body_child_grammar(namespace, local);
    }
    if matches!(
        parent,
        MetaContentGrammar::ShapeBasic
            | MetaContentGrammar::ShapeGroup
            | MetaContentGrammar::ShapeFrame
    ) {
        return shape_child_grammar(parent, namespace, local);
    }
    if parent == MetaContentGrammar::Structured {
        return structured_meta_child_grammar(namespace, local);
    }
    let allow_hyperlink = parent == MetaContentGrammar::ParagraphOrHyperlink;
    if namespace == TEXT_DATABASE_NAMESPACE {
        if local == "a" {
            return if allow_hyperlink {
                Ok(MetaContentGrammar::Hyperlink)
            } else {
                Err(Error::InvalidFormat(
                    "text:a cannot be nested in text:a paragraph content".to_string(),
                ))
            };
        }
        return match local {
            "span" | "meta" | "meta-field" => Ok(MetaContentGrammar::ParagraphOrHyperlink),
            "ruby" => Ok(MetaContentGrammar::Ruby),
            "note" => Ok(MetaContentGrammar::Note),
            "execute-macro" => Ok(MetaContentGrammar::ExecuteMacro),
            "drop-down" => Ok(MetaContentGrammar::DropDown),
            "s"
            | "tab"
            | "line-break"
            | "soft-page-break"
            | "bookmark"
            | "bookmark-start"
            | "bookmark-end"
            | "reference-mark"
            | "reference-mark-start"
            | "reference-mark-end"
            | "change"
            | "change-start"
            | "change-end"
            | "toc-mark"
            | "toc-mark-start"
            | "toc-mark-end"
            | "user-index-mark"
            | "user-index-mark-start"
            | "user-index-mark-end"
            | "alphabetical-index-mark"
            | "alphabetical-index-mark-start"
            | "alphabetical-index-mark-end" => Ok(MetaContentGrammar::Empty),
            "bibliography-mark" => Ok(MetaContentGrammar::TextOnly),
            "database-next" | "database-row-select" => Ok(MetaContentGrammar::Empty),
            _ if Field::is_field_tag(&format!("text:{local}")) => Ok(MetaContentGrammar::TextOnly),
            _ => Err(Error::InvalidFormat(format!(
                "text:{local} is not paragraph-content in text:meta-field"
            ))),
        };
    }
    match (namespace, local) {
        (OFFICE_NAMESPACE, "annotation") => Ok(MetaContentGrammar::Annotation),
        (OFFICE_NAMESPACE, "annotation-end") => Ok(MetaContentGrammar::Empty),
        (PRESENTATION_NAMESPACE, "header" | "footer" | "date-time") => {
            Ok(MetaContentGrammar::Empty)
        },
        _ if is_odf_shape_root(namespace, local) => {
            odf_shape_grammar(namespace, local).ok_or_else(|| {
                Error::InvalidFormat(format!("unsupported drawing shape {namespace}:{local}"))
            })
        },
        _ => Err(Error::InvalidFormat(format!(
            "{namespace}:{local} is not paragraph-content in text:meta-field"
        ))),
    }
}

fn note_body_child_grammar(namespace: &str, local: &str) -> Result<MetaContentGrammar> {
    if namespace == TEXT_DATABASE_NAMESPACE {
        return match local {
            "p" | "h" => Ok(MetaContentGrammar::ParagraphOrHyperlink),
            "soft-page-break" | "change" | "change-start" | "change-end" => {
                Ok(MetaContentGrammar::Empty)
            },
            "list" | "numbered-paragraph" | "section" | "table-of-content"
            | "illustration-index" | "table-index" | "object-index" | "user-index"
            | "alphabetical-index" | "bibliography" => Ok(MetaContentGrammar::Structured),
            _ => Err(Error::InvalidFormat(format!(
                "text:{local} is not text-content in text:note-body"
            ))),
        };
    }
    if namespace == TABLE_NAMESPACE && local == "table" {
        return Ok(MetaContentGrammar::Structured);
    }
    if is_odf_shape_root(namespace, local) {
        return odf_shape_grammar(namespace, local).ok_or_else(|| {
            Error::InvalidFormat(format!("unsupported drawing shape {namespace}:{local}"))
        });
    }
    Err(Error::InvalidFormat(format!(
        "{namespace}:{local} is not text-content in text:note-body"
    )))
}

fn shape_child_grammar(
    owner: MetaContentGrammar,
    namespace: &str,
    local: &str,
) -> Result<MetaContentGrammar> {
    match (namespace, local) {
        (SVG_NAMESPACE, "title" | "desc") => Ok(MetaContentGrammar::TextOnly),
        (OFFICE_NAMESPACE, "event-listeners") => Ok(MetaContentGrammar::EventListeners),
        (TEXT_DATABASE_NAMESPACE, "p") => Ok(MetaContentGrammar::ParagraphOrHyperlink),
        (DRAW_NAMESPACE, "glue-point" | "page-thumbnail" | "control") => {
            Ok(MetaContentGrammar::Empty)
        },
        (
            DRAW_NAMESPACE,
            "text-box" | "image" | "object" | "object-ole" | "applet" | "floating-frame" | "plugin"
            | "image-map" | "enhanced-geometry",
        ) if owner == MetaContentGrammar::ShapeFrame => Ok(MetaContentGrammar::Structured),
        (TABLE_NAMESPACE, "table") if owner == MetaContentGrammar::ShapeFrame => {
            Ok(MetaContentGrammar::Structured)
        },
        _ if owner == MetaContentGrammar::ShapeGroup && is_odf_shape_root(namespace, local) => {
            odf_shape_grammar(namespace, local).ok_or_else(|| {
                Error::InvalidFormat(format!("unsupported drawing shape {namespace}:{local}"))
            })
        },
        _ => Err(Error::InvalidFormat(format!(
            "{namespace}:{local} is not valid direct drawing-shape content"
        ))),
    }
}

fn structured_meta_child_grammar(namespace: &str, local: &str) -> Result<MetaContentGrammar> {
    match (namespace, local) {
        (TEXT_DATABASE_NAMESPACE, "p" | "h") => Ok(MetaContentGrammar::ParagraphOrHyperlink),
        (TEXT_DATABASE_NAMESPACE, "soft-page-break" | "change" | "change-start" | "change-end") => {
            Ok(MetaContentGrammar::Empty)
        },
        (OFFICE_NAMESPACE, "event-listeners") => Ok(MetaContentGrammar::EventListeners),
        (SVG_NAMESPACE, "title" | "desc") => Ok(MetaContentGrammar::TextOnly),
        (SCRIPT_NAMESPACE, "event-listener") => Ok(MetaContentGrammar::Empty),
        (PRESENTATION_NAMESPACE, "event-listener") => {
            Ok(MetaContentGrammar::PresentationEventListener)
        },
        _ if is_allowed_meta_namespace(namespace) => Ok(MetaContentGrammar::Structured),
        _ => Err(Error::InvalidFormat(format!(
            "foreign structured metadata namespace '{namespace}' for {local}"
        ))),
    }
}

fn is_odf_shape_root(namespace: &str, local: &str) -> bool {
    (namespace == DRAW_NAMESPACE
        && matches!(
            local,
            "rect"
                | "line"
                | "polyline"
                | "polygon"
                | "regular-polygon"
                | "path"
                | "circle"
                | "ellipse"
                | "g"
                | "page-thumbnail"
                | "frame"
                | "measure"
                | "caption"
                | "connector"
                | "control"
                | "custom-shape"
                | "a"
        ))
        || (namespace == DR3D_NAMESPACE && local == "scene")
}

fn odf_shape_grammar(namespace: &str, local: &str) -> Option<MetaContentGrammar> {
    if !is_odf_shape_root(namespace, local) {
        return None;
    }
    Some(match (namespace, local) {
        (DRAW_NAMESPACE, "g") => MetaContentGrammar::ShapeGroup,
        (DRAW_NAMESPACE, "frame") => MetaContentGrammar::ShapeFrame,
        (DRAW_NAMESPACE, "a") => MetaContentGrammar::ShapeLink,
        _ => MetaContentGrammar::ShapeBasic,
    })
}

pub(crate) fn validate_meta_element_parts(
    namespace_uri: &str,
    local_name: &str,
    attributes: &[MetaFieldAttribute],
    aggregate: &mut usize,
) -> Result<()> {
    if !is_allowed_meta_namespace(namespace_uri) || namespace_uri == XLINK_NAMESPACE {
        return Err(Error::InvalidFormat(format!(
            "foreign meta-field element namespace '{namespace_uri}'"
        )));
    }
    validate_xml_ncname(local_name, "meta-field element name")?;
    if attributes.len() > MAX_META_FIELD_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "meta-field child exceeds {MAX_META_FIELD_ATTRIBUTES} attributes"
        )));
    }
    let mut seen = HashSet::new();
    for attribute in attributes {
        if !is_allowed_meta_namespace(&attribute.namespace_uri) {
            return Err(Error::InvalidFormat(format!(
                "foreign meta-field attribute namespace '{}'",
                attribute.namespace_uri
            )));
        }
        validate_xml_ncname(&attribute.local_name, "meta-field attribute name")?;
        if !seen.insert((&attribute.namespace_uri, &attribute.local_name)) {
            return Err(Error::InvalidFormat(
                "duplicate namespace-resolved meta-field attribute".to_string(),
            ));
        }
        validate_dynamic_value(
            "meta-field attribute value",
            Some(&attribute.value),
            false,
            aggregate,
        )?;
    }
    Ok(())
}

pub(crate) fn validate_xml_id(value: &str) -> Result<()> {
    validate_xml_ncname(value, "xml:id")
}

fn validate_xml_ncname(value: &str, name: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    };
    let start = first == '_' || first.is_alphabetic() || (first as u32) >= 0x80;
    let rest = chars.all(|ch| {
        ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric() || (ch as u32) >= 0x80
    });
    if !start || !rest || value.contains(':') || !value.chars().all(is_xml_1_0_char) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML NCName {name} '{value}'"
        )));
    }
    Ok(())
}

pub(crate) fn is_allowed_meta_namespace(namespace: &str) -> bool {
    matches!(
        namespace,
        TEXT_DATABASE_NAMESPACE
            | OFFICE_NAMESPACE
            | STYLE_NAMESPACE
            | XLINK_NAMESPACE
            | XML_NAMESPACE
            | DRAW_NAMESPACE
            | TABLE_NAMESPACE
            | PRESENTATION_NAMESPACE
            | SVG_NAMESPACE
            | FO_NAMESPACE
            | NUMBER_NAMESPACE
            | META_NAMESPACE
            | DC_NAMESPACE
            | XHTML_NAMESPACE
            | DR3D_NAMESPACE
            | FORM_NAMESPACE
            | SCRIPT_NAMESPACE
    )
}

pub(crate) fn add_meta_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    if amount > MAX_DYNAMIC_FIELD_VALUE {
        return Err(Error::InvalidFormat(format!(
            "meta-field value exceeds {MAX_DYNAMIC_FIELD_VALUE} bytes"
        )));
    }
    *aggregate = aggregate
        .checked_add(amount)
        .ok_or_else(|| Error::InvalidFormat("meta-field aggregate size overflow".to_string()))?;
    if *aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "meta-field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
        )));
    }
    Ok(())
}

fn canonical_meta_prefix(namespace: &str) -> &'static str {
    match namespace {
        TEXT_DATABASE_NAMESPACE => "text",
        OFFICE_NAMESPACE => "office",
        STYLE_NAMESPACE => "style",
        XLINK_NAMESPACE => "xlink",
        XML_NAMESPACE => "xml",
        DRAW_NAMESPACE => "draw",
        TABLE_NAMESPACE => "table",
        PRESENTATION_NAMESPACE => "presentation",
        SVG_NAMESPACE => "svg",
        FO_NAMESPACE => "fo",
        NUMBER_NAMESPACE => "number",
        META_NAMESPACE => "meta",
        DC_NAMESPACE => "dc",
        XHTML_NAMESPACE => "xhtml",
        DR3D_NAMESPACE => "dr3d",
        FORM_NAMESPACE => "form",
        SCRIPT_NAMESPACE => "script",
        _ => unreachable!("validated meta-field namespace"),
    }
}

pub(super) fn write_meta_node(node: &MetaFieldNode, output: &mut String) {
    match node {
        MetaFieldNode::Text(value) => push_xml_text(output, value),
        MetaFieldNode::Element(element) => {
            let prefix = canonical_meta_prefix(&element.namespace_uri);
            output.push('<');
            output.push_str(prefix);
            output.push(':');
            output.push_str(&element.local_name);
            output.push_str(" xmlns:");
            output.push_str(prefix);
            output.push_str("=\"");
            output.push_str(&element.namespace_uri);
            output.push('"');
            let mut declared = HashSet::new();
            declared.insert(prefix);
            for attribute in &element.attributes {
                let attribute_prefix = canonical_meta_prefix(&attribute.namespace_uri);
                if attribute_prefix != "xml" && declared.insert(attribute_prefix) {
                    output.push_str(" xmlns:");
                    output.push_str(attribute_prefix);
                    output.push_str("=\"");
                    output.push_str(&attribute.namespace_uri);
                    output.push('"');
                }
                output.push(' ');
                output.push_str(attribute_prefix);
                output.push(':');
                output.push_str(&attribute.local_name);
                output.push_str("=\"");
                push_xml_attribute(output, &attribute.value);
                output.push('"');
            }
            if element.children.is_empty() {
                output.push_str("/>");
            } else {
                output.push('>');
                for child in &element.children {
                    write_meta_node(child, output);
                }
                output.push_str("</");
                output.push_str(prefix);
                output.push(':');
                output.push_str(&element.local_name);
                output.push('>');
            }
        },
    }
}
