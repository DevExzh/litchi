//! Immutable, bounded inventory of main-document section boundaries.

use super::codec::{validate_element_qname, validate_namespace_declaration, validate_qname};
use super::{Columns, Margins, PageSize, Reference, Section, Start};
use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use litchi_core::{Position, SourceVersion};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::sync::Arc;

/// Resource policy for one main-document section inventory.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Limits {
    /// Maximum source main-document XML bytes.
    pub max_input_bytes: usize,
    /// Maximum XML events examined after MCE branch selection.
    pub max_events: usize,
    /// Maximum XML nesting depth.
    pub max_depth: usize,
    /// Maximum logical paragraphs in the main-document body.
    pub max_paragraphs: usize,
    /// Maximum explicit `w:sectPr` elements.
    pub max_sections: usize,
    /// Maximum bytes in one selected `w:sectPr` fragment.
    pub max_section_bytes: usize,
    /// Maximum aggregate bytes retained for inert header/footer relationship IDs.
    pub max_reference_bytes: usize,
    /// Maximum transient bytes produced by MCE branch selection.
    pub max_mce_output_bytes: usize,
    /// Maximum MCE namespace bindings and directive tokens.
    pub max_mce_bindings: usize,
    /// Maximum choices in one `mc:AlternateContent` container.
    pub max_mce_choices: usize,
}

impl Limits {
    fn validate(&self) -> Result<()> {
        if [
            self.max_input_bytes,
            self.max_events,
            self.max_depth,
            self.max_paragraphs,
            self.max_sections,
            self.max_section_bytes,
            self.max_reference_bytes,
            self.max_mce_output_bytes,
            self.max_mce_bindings,
            self.max_mce_choices,
        ]
        .contains(&0)
        {
            return Err(Error::InvalidFormat(
                "section inventory limits must be nonzero".into(),
            ));
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: 32 * 1024 * 1024,
            max_events: 1_000_000,
            max_depth: 128,
            max_paragraphs: 1_000_000,
            max_sections: 65_536,
            max_section_bytes: super::model::MAX_XML_BYTES,
            max_reference_bytes: 1024 * 1024,
            max_mce_output_bytes: 64 * 1024 * 1024,
            max_mce_bindings: 4096,
            max_mce_choices: 1024,
        }
    }
}

/// Where the properties for one logical section are authored.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Ownership {
    /// The section ends at a paragraph whose direct `w:pPr` owns `w:sectPr`.
    Paragraph(Position),
    /// The final direct child of `w:body` owns the last section properties.
    BodyFinal,
    /// No final `w:sectPr` is authored; Word's implicit defaults describe it.
    Implicit,
}

/// A half-open range of source-order main-story paragraphs.
///
/// Positions use the same namespace-aware paragraph order as the document's
/// paragraph inventory, including paragraphs nested in main-story tables and
/// controls. A `w:sectPr` in a table-cell paragraph is content in that
/// paragraph, not a main-story section boundary, and is therefore ignored by
/// this inventory.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphRange {
    start: Position,
    end: Position,
}

impl ParagraphRange {
    /// First paragraph position in the section.
    #[must_use]
    pub const fn start(self) -> Position {
        self.start
    }

    /// Exclusive paragraph boundary after the section.
    #[must_use]
    pub const fn end(self) -> Position {
        self.end
    }

    /// Number of paragraphs in the range.
    #[must_use]
    pub const fn len(self) -> usize {
        self.end.get() - self.start.get()
    }

    /// Whether the section contains no paragraph.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start.get() == self.end.get()
    }
}

/// A typed selector for one logical section.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Selector {
    /// Select by zero-based source-order section position.
    Position(Position),
    /// Select the section whose properties have this semantic owner.
    Owner(Ownership),
}

impl Selector {
    /// Select the body-final section.
    #[must_use]
    pub const fn body_final() -> Self {
        Self::Owner(Ownership::BodyFinal)
    }

    /// Select a section ending at one paragraph.
    #[must_use]
    pub const fn paragraph(position: Position) -> Self {
        Self::Owner(Ownership::Paragraph(position))
    }
}

impl From<usize> for Selector {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}

impl From<Position> for Selector {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

impl From<Ownership> for Selector {
    fn from(value: Ownership) -> Self {
        Self::Owner(value)
    }
}

/// One property address supported by the focused query facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Property {
    /// Section-property ownership.
    Ownership,
    /// Logical paragraph boundaries.
    ParagraphRange,
    /// Locally authored page size and orientation.
    PageSize,
    /// Locally authored page margins.
    Margins,
    /// Locally authored section-break start policy.
    Start,
    /// Locally authored newspaper-style columns.
    Columns,
    /// Inert local header relationship references.
    Headers,
    /// Inert local footer relationship references.
    Footers,
}

/// Borrowed value returned for one typed section property query.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum PropertyValue<'a> {
    /// Section-property ownership.
    Ownership(Ownership),
    /// Logical paragraph boundaries.
    ParagraphRange(ParagraphRange),
    /// Optional locally authored page geometry.
    PageSize(Option<PageSize>),
    /// Optional locally authored margins.
    Margins(Option<Margins>),
    /// Optional locally authored break policy.
    Start(Option<Start>),
    /// Optional locally authored columns.
    Columns(Option<Columns>),
    /// Borrowed inert relationship references.
    Headers(&'a [Reference]),
    /// Borrowed inert relationship references.
    Footers(&'a [Reference]),
}

/// Immutable semantic descriptor for one logical section.
#[derive(Debug, PartialEq, Eq)]
pub struct Descriptor {
    position: Position,
    ownership: Ownership,
    paragraphs: ParagraphRange,
    page_size: Option<PageSize>,
    margins: Option<Margins>,
    start: Option<Start>,
    columns: Option<Columns>,
    headers: Box<[Reference]>,
    footers: Box<[Reference]>,
}

impl Descriptor {
    /// Zero-based source-order position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Where this section's properties are authored.
    #[must_use]
    pub const fn ownership(&self) -> Ownership {
        self.ownership
    }

    /// Half-open logical paragraph range.
    #[must_use]
    pub const fn paragraphs(&self) -> ParagraphRange {
        self.paragraphs
    }

    /// Locally authored page geometry, if present.
    #[must_use]
    pub const fn page_size(&self) -> Option<PageSize> {
        self.page_size
    }

    /// Locally authored margins, if present.
    #[must_use]
    pub const fn margins(&self) -> Option<Margins> {
        self.margins
    }

    /// Locally authored section-break policy, if present.
    #[must_use]
    pub const fn start(&self) -> Option<Start> {
        self.start
    }

    /// Locally authored newspaper-style columns, if present.
    #[must_use]
    pub fn columns(&self) -> Option<Columns> {
        self.columns.clone()
    }

    /// Inert header relationship references; targets are not resolved.
    #[must_use]
    pub fn headers(&self) -> &[Reference] {
        &self.headers
    }

    /// Inert footer relationship references; targets are not resolved.
    #[must_use]
    pub fn footers(&self) -> &[Reference] {
        &self.footers
    }

    fn property(&self, property: Property) -> PropertyValue<'_> {
        match property {
            Property::Ownership => PropertyValue::Ownership(self.ownership),
            Property::ParagraphRange => PropertyValue::ParagraphRange(self.paragraphs),
            Property::PageSize => PropertyValue::PageSize(self.page_size),
            Property::Margins => PropertyValue::Margins(self.margins),
            Property::Start => PropertyValue::Start(self.start),
            Property::Columns => PropertyValue::Columns(self.columns.clone()),
            Property::Headers => PropertyValue::Headers(&self.headers),
            Property::Footers => PropertyValue::Footers(&self.footers),
        }
    }
}

/// Cheaply cloneable, immutable section inventory.
#[derive(Clone, Debug)]
pub struct Inventory {
    sections: Arc<[Descriptor]>,
    paragraph_count: usize,
    sources: Arc<[Option<SourceSpan>]>,
}

/// The private source range for one locally authored main-story `w:sectPr`.
///
/// It is deliberately not part of the public section descriptor: physical XML
/// identity remains an implementation detail of the source-backed layout
/// transaction.
#[derive(Clone, Debug)]
pub(crate) struct SourceSpan {
    start: usize,
    end: usize,
    relationship_bindings: Box<[RelationshipBinding]>,
    namespace_bindings: Box<[RelationshipBinding]>,
}

impl SourceSpan {
    pub(crate) const fn range(&self) -> (usize, usize) {
        (self.start, self.end)
    }

    pub(crate) fn namespace_bindings(&self) -> &[RelationshipBinding] {
        &self.namespace_bindings
    }
}

impl Inventory {
    /// Parse a main-document XML payload with production defaults.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        Self::parse_with_limits(xml, &Limits::default())
    }

    /// Parse a main-document XML payload with caller-provided limits.
    pub fn parse_with_limits(xml: &[u8], limits: &Limits) -> Result<Self> {
        parse_inventory(xml, limits)
    }

    /// Logical sections in source order.
    #[must_use]
    pub fn sections(&self) -> &[Descriptor] {
        &self.sections
    }

    /// Number of source-order main-story paragraphs used by section ranges.
    #[must_use]
    pub const fn paragraph_count(&self) -> usize {
        self.paragraph_count
    }

    /// Resolve a typed selector without exposing physical XML identity.
    #[must_use]
    pub fn section(&self, selector: impl Into<Selector>) -> Option<&Descriptor> {
        match selector.into() {
            Selector::Position(position) => self.sections.get(position.get()),
            Selector::Owner(owner) => self
                .sections
                .iter()
                .find(|section| section.ownership == owner),
        }
    }

    /// Query one typed property without cloning the inventory or references.
    #[must_use]
    pub fn property(
        &self,
        selector: impl Into<Selector>,
        property: Property,
    ) -> Option<PropertyValue<'_>> {
        self.section(selector)
            .map(|section| section.property(property))
    }

    /// Whether two inventories share the immutable descriptor allocation.
    #[must_use]
    pub fn shares_allocation_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.sections, &other.sections) && Arc::ptr_eq(&self.sources, &other.sources)
    }

    pub(crate) fn topology_matches(&self, other: &Self) -> bool {
        self.paragraph_count == other.paragraph_count
            && self.sections.len() == other.sections.len()
            && self.sources.len() == other.sources.len()
            && self
                .sections
                .iter()
                .zip(other.sections.iter())
                .all(|(left, right)| {
                    left.position == right.position
                        && left.ownership == right.ownership
                        && left.paragraphs == right.paragraphs
                        && left.headers == right.headers
                        && left.footers == right.footers
                })
            && self
                .sources
                .iter()
                .zip(other.sources.iter())
                .all(|(left, right)| match (left, right) {
                    (None, None) => true,
                    (Some(left), Some(right)) => relationship_bindings_match(
                        &left.relationship_bindings,
                        &right.relationship_bindings,
                    ),
                    _ => false,
                })
    }

    pub(crate) fn source_span(&self, selector: impl Into<Selector>) -> Option<&SourceSpan> {
        let position = self.section(selector)?.position.get();
        self.sources.get(position)?.as_ref()
    }

    pub(crate) fn source_fragment(
        &self,
        xml: &[u8],
        selector: impl Into<Selector>,
        max_section_bytes: usize,
    ) -> Result<Option<(SourceSpan, Vec<u8>)>> {
        let Some(source) = self.source_span(selector) else {
            return Ok(None);
        };
        let (start, end) = source.range();
        let bytes = xml.get(start..end).ok_or_else(|| {
            Error::InvalidFormat("section source span is outside document XML".into())
        })?;
        Ok(Some((
            source.clone(),
            self_contained_fragment(bytes, &source.namespace_bindings, max_section_bytes)?,
        )))
    }
}

fn relationship_bindings_match(
    left: &[RelationshipBinding],
    right: &[RelationshipBinding],
) -> bool {
    left.iter()
        .filter(|binding| is_relationship_namespace(&binding.namespace))
        .map(|binding| (&binding.prefix, &binding.namespace))
        .eq(right
            .iter()
            .filter(|binding| is_relationship_namespace(&binding.namespace))
            .map(|binding| (&binding.prefix, &binding.namespace)))
}

fn is_relationship_namespace(namespace: &[u8]) -> bool {
    matches!(
        namespace,
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            | b"http://purl.oclc.org/ooxml/officeDocument/relationships"
    )
}

/// An immutable inventory optionally bound to an exact positional source.
#[derive(Clone, Debug)]
pub struct Snapshot {
    inventory: Inventory,
    source_version: Option<SourceVersion>,
}

impl Snapshot {
    /// Capture an owned XML snapshot without a positional-source lineage.
    pub fn from_xml(xml: Vec<u8>) -> Result<Self> {
        Self::from_xml_with_limits(xml, &Limits::default())
    }

    /// Capture owned XML with caller-provided semantic limits.
    pub fn from_xml_with_limits(xml: Vec<u8>, limits: &Limits) -> Result<Self> {
        Ok(Self {
            inventory: Inventory::parse_with_limits(&xml, limits)?,
            source_version: None,
        })
    }

    pub(crate) fn from_source_xml(
        xml: &[u8],
        source_version: SourceVersion,
        limits: &Limits,
    ) -> Result<Self> {
        Ok(Self {
            inventory: Inventory::parse_with_limits(xml, limits)?,
            source_version: Some(source_version),
        })
    }

    /// Borrow the immutable semantic inventory.
    #[must_use]
    pub const fn inventory(&self) -> &Inventory {
        &self.inventory
    }

    /// Resolve one typed section selector.
    #[must_use]
    pub fn section(&self, selector: impl Into<Selector>) -> Option<&Descriptor> {
        self.inventory.section(selector)
    }

    /// Query one typed property without cloning relationship references.
    #[must_use]
    pub fn property(
        &self,
        selector: impl Into<Selector>,
        property: Property,
    ) -> Option<PropertyValue<'_>> {
        self.inventory.property(selector, property)
    }

    /// Exact source identity and revision, when captured from `ReadAt`.
    #[must_use]
    pub const fn source_version(&self) -> Option<SourceVersion> {
        self.source_version
    }

    /// Whether two snapshots share their immutable section allocation.
    #[must_use]
    pub fn shares_allocation_with(&self, other: &Self) -> bool {
        self.inventory.shares_allocation_with(&other.inventory)
    }
}

#[derive(Clone, Copy)]
struct ParagraphContext {
    depth: usize,
    main_story: bool,
    section_boundary: bool,
    position: Position,
    saw_properties: bool,
    saw_content: bool,
    properties_depth: Option<usize>,
    section_seen: bool,
}

struct Capture {
    start: usize,
    depth: usize,
    ownership: Ownership,
    reference_count: usize,
    reference_bytes: usize,
    relationship_bindings: Vec<RelationshipBinding>,
    root_namespace_bindings: Vec<RelationshipBinding>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RelationshipBinding {
    pub(crate) prefix: Vec<u8>,
    pub(crate) namespace: Vec<u8>,
}

#[derive(Default)]
struct NamespaceContext {
    active: Vec<RelationshipBinding>,
    scopes: Vec<Vec<(Vec<u8>, Option<RelationshipBinding>)>>,
}

impl NamespaceContext {
    fn enter(
        &mut self,
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) -> Result<()> {
        let declarations = namespace_declarations(element, decoder)?;
        let mut changes = Vec::new();
        changes
            .try_reserve(declarations.len())
            .map_err(|source| Error::Allocation {
                resource: "section namespace scope changes",
                source,
            })?;
        for declaration in declarations {
            let prefix = declaration.prefix.clone();
            if let Some(index) = self
                .active
                .iter()
                .position(|binding| binding.prefix == prefix)
            {
                let previous = self.active[index].clone();
                self.active[index] = declaration;
                changes.push((prefix, Some(previous)));
            } else {
                self.active.push(declaration);
                changes.push((prefix, None));
            }
        }
        self.scopes.push(changes);
        Ok(())
    }

    fn exit(&mut self) -> Result<()> {
        let changes = self
            .scopes
            .pop()
            .ok_or_else(|| Error::InvalidFormat("section namespace scope underflow".into()))?;
        for (prefix, previous) in changes.into_iter().rev() {
            if let Some(index) = self
                .active
                .iter()
                .position(|binding| binding.prefix == prefix)
            {
                if let Some(previous) = previous {
                    self.active[index] = previous;
                } else {
                    self.active.remove(index);
                }
            } else {
                return Err(Error::InvalidFormat(
                    "section namespace scope restoration failed".into(),
                ));
            }
        }
        Ok(())
    }

    fn snapshot(&self) -> Vec<RelationshipBinding> {
        self.active.clone()
    }
}

struct ScanState {
    sections: Vec<Descriptor>,
    sources: Vec<Option<SourceSpan>>,
    paragraphs: usize,
    next_section_start: usize,
    reference_bytes: usize,
    body_final: bool,
}

fn parse_inventory(xml: &[u8], limits: &Limits) -> Result<Inventory> {
    limits.validate()?;
    enforce_limit("input bytes", xml.len(), limits.max_input_bytes)?;
    let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
    capabilities.understand_namespace(crate::paragraph::extensions::WORD_2010_NAMESPACE);
    let processed = litchi_ooxml_common::mce::process_markup_compatibility(
        xml,
        &capabilities,
        &litchi_ooxml_common::mce::Limits {
            max_input_bytes: limits.max_input_bytes,
            max_output_bytes: limits.max_mce_output_bytes,
            max_depth: limits.max_depth,
            max_namespace_bindings: limits.max_mce_bindings,
            max_directive_tokens: limits.max_mce_bindings,
            max_choices_per_alternate: limits.max_mce_choices,
        },
    )?;
    scan_visible_document(processed.xml.as_ref(), limits)
}

fn scan_visible_document(xml: &[u8], limits: &Limits) -> Result<Inventory> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut element_stack = Vec::<Vec<u8>>::new();
    let mut events = 0usize;
    let mut root_seen = false;
    let mut body_depth = None;
    let mut table_depth = 0usize;
    let mut body_seen = false;
    let mut root_closed = false;
    let mut paragraphs = Vec::<ParagraphContext>::new();
    let mut capture = None::<Capture>;
    let mut namespace_context = NamespaceContext::default();
    let mut state = ScanState {
        sections: Vec::new(),
        sources: Vec::new(),
        paragraphs: 0,
        next_section_start: 0,
        reference_bytes: 0,
        body_final: false,
    };

    loop {
        let event_start = offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let event_end = offset(&reader)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("section event counter overflow".into()))?;
        enforce_limit("XML events", events, limits.max_events)?;

        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(Error::InvalidFormat(
                        "main-document XML has content after w:document".into(),
                    ));
                }
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML nesting overflow".into()))?;
                enforce_limit("XML depth", child_depth, limits.max_depth)?;
                let word = is_wordprocessing_namespace(&namespace);
                let local = element.local_name();
                validate_namespace_bindings(&element, &namespace, &resolver, reader.decoder())?;
                namespace_context.enter(&element, reader.decoder())?;
                let namespace_bindings = if word && local.as_ref() == b"sectPr" {
                    namespace_context.snapshot()
                } else {
                    Vec::new()
                };
                inspect_reference(
                    &element,
                    &namespace,
                    &resolver,
                    reader.decoder(),
                    child_depth,
                    limits,
                    state.reference_bytes,
                    capture.as_mut(),
                )?;
                inspect_element(
                    xml,
                    event_start,
                    event_end,
                    child_depth,
                    word,
                    local.as_ref(),
                    false,
                    limits,
                    &mut root_seen,
                    &mut body_seen,
                    &mut body_depth,
                    &mut table_depth,
                    &mut paragraphs,
                    &mut capture,
                    &mut state,
                    namespace_bindings,
                )?;
                element_stack
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "section inventory XML element stack",
                        source,
                    })?;
                element_stack.push(element.name().as_ref().to_vec());
                depth = child_depth;
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(Error::InvalidFormat(
                        "main-document XML has content after w:document".into(),
                    ));
                }
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("section XML nesting overflow".into()))?;
                enforce_limit("XML depth", child_depth, limits.max_depth)?;
                let word = is_wordprocessing_namespace(&namespace);
                let local = element.local_name();
                validate_namespace_bindings(&element, &namespace, &resolver, reader.decoder())?;
                namespace_context.enter(&element, reader.decoder())?;
                let namespace_bindings = if word && local.as_ref() == b"sectPr" {
                    namespace_context.snapshot()
                } else {
                    Vec::new()
                };
                inspect_reference(
                    &element,
                    &namespace,
                    &resolver,
                    reader.decoder(),
                    child_depth,
                    limits,
                    state.reference_bytes,
                    capture.as_mut(),
                )?;
                inspect_element(
                    xml,
                    event_start,
                    event_end,
                    child_depth,
                    word,
                    local.as_ref(),
                    true,
                    limits,
                    &mut root_seen,
                    &mut body_seen,
                    &mut body_depth,
                    &mut table_depth,
                    &mut paragraphs,
                    &mut capture,
                    &mut state,
                    namespace_bindings,
                )?;
                namespace_context.exit()?;
            },
            Event::End(element) => {
                let expected = element_stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("main-document XML has an unmatched end element".into())
                })?;
                if expected != element.name().as_ref() {
                    return Err(Error::InvalidFormat(
                        "main-document XML has mismatched element nesting".into(),
                    ));
                }
                if capture.as_ref().is_some_and(|value| value.depth == depth) {
                    let selected = capture.take().ok_or_else(|| {
                        Error::InvalidFormat("missing selected section fragment".into())
                    })?;
                    append_section(xml, event_end, selected, limits, &mut state)?;
                }
                if let Some(context) = paragraphs.last_mut() {
                    if context.properties_depth == Some(depth) {
                        context.properties_depth = None;
                    }
                    if context.depth == depth {
                        paragraphs.pop();
                    }
                }
                if body_depth == Some(depth)
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"body"
                {
                    body_depth = None;
                }
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"tbl"
                {
                    table_depth = table_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("main-document table nesting underflow".into())
                    })?;
                }
                if depth == 1 {
                    if !is_wordprocessing_namespace(&namespace)
                        || element.local_name().as_ref() != b"document"
                    {
                        return Err(Error::InvalidFormat(
                            "main-document XML has an invalid root close".into(),
                        ));
                    }
                    root_closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid section XML nesting".into()))?;
                namespace_context.exit()?;
            },
            Event::Eof if depth != 0 || capture.is_some() || !element_stack.is_empty() => {
                return Err(Error::InvalidFormat(
                    "unterminated main-document XML".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(text)
                if root_closed && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(Error::InvalidFormat(
                    "main-document XML has trailing character data".into(),
                ));
            },
            Event::CData(text)
                if root_closed && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(Error::InvalidFormat(
                    "main-document XML has trailing character data".into(),
                ));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !root_seen || !root_closed || !body_seen {
        return Err(Error::InvalidFormat(
            "main-document section inventory requires one w:document/w:body".into(),
        ));
    }
    if !state.body_final {
        append_implicit(&mut state, limits)?;
    }
    Ok(Inventory {
        sections: Arc::from(state.sections.into_boxed_slice()),
        paragraph_count: state.paragraphs,
        sources: Arc::from(state.sources.into_boxed_slice()),
    })
}

#[allow(clippy::too_many_arguments)]
fn inspect_element(
    xml: &[u8],
    event_start: usize,
    event_end: usize,
    child_depth: usize,
    word: bool,
    local: &[u8],
    empty: bool,
    limits: &Limits,
    root_seen: &mut bool,
    body_seen: &mut bool,
    body_depth: &mut Option<usize>,
    table_depth: &mut usize,
    paragraphs: &mut Vec<ParagraphContext>,
    capture: &mut Option<Capture>,
    state: &mut ScanState,
    namespace_bindings: Vec<RelationshipBinding>,
) -> Result<()> {
    if child_depth == 1 {
        if *root_seen || !word || local != b"document" {
            return Err(Error::InvalidFormat(
                "section inventory XML has an invalid root".into(),
            ));
        }
        *root_seen = true;
        if empty {
            return Err(Error::InvalidFormat(
                "main-document root omits w:body".into(),
            ));
        }
    }
    if word && local == b"body" {
        if *body_seen || child_depth != 2 || !*root_seen {
            return Err(Error::InvalidFormat(
                "main document has an invalid or duplicate w:body".into(),
            ));
        }
        *body_seen = true;
        *body_depth = (!empty).then_some(child_depth);
        return Ok(());
    }

    if body_depth.is_some_and(|body| child_depth == body + 1) && state.body_final {
        return Err(Error::InvalidFormat(
            "body-final section properties are not the final body child".into(),
        ));
    }

    if word && local == b"tbl" && !empty {
        *table_depth = table_depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("main-document table nesting overflow".into()))?;
    }

    if word && local == b"p" && body_depth.is_some() {
        if !paragraphs.is_empty() {
            if !empty {
                paragraphs
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "section inventory paragraph stack",
                        source,
                    })?;
                paragraphs.push(ParagraphContext {
                    depth: child_depth,
                    main_story: false,
                    section_boundary: false,
                    position: Position::new(0),
                    saw_properties: false,
                    saw_content: false,
                    properties_depth: None,
                    section_seen: false,
                });
            }
            return Ok(());
        }
        enforce_limit(
            "paragraphs",
            state.paragraphs.saturating_add(1),
            limits.max_paragraphs,
        )?;
        let position = Position::new(state.paragraphs);
        state.paragraphs = state
            .paragraphs
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("section paragraph counter overflow".into()))?;
        if !empty {
            paragraphs
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "section inventory paragraph stack",
                    source,
                })?;
            paragraphs.push(ParagraphContext {
                depth: child_depth,
                main_story: true,
                section_boundary: *table_depth == 0,
                position,
                saw_properties: false,
                saw_content: false,
                properties_depth: None,
                section_seen: false,
            });
        }
        return Ok(());
    }

    if let Some(context) = paragraphs.last_mut() {
        if !context.main_story {
            return Ok(());
        }
        if child_depth == context.depth + 1 {
            if word && local == b"pPr" {
                if context.saw_properties || context.saw_content {
                    return Err(Error::InvalidFormat(
                        "paragraph properties are duplicated or out of order".into(),
                    ));
                }
                context.saw_properties = true;
                if !empty {
                    context.properties_depth = Some(child_depth);
                }
                return Ok(());
            }
            context.saw_content = true;
        }
        if context
            .properties_depth
            .is_some_and(|properties| child_depth == properties + 1)
        {
            if context.section_seen {
                return Err(Error::InvalidFormat(
                    "paragraph section properties must be the final pPr child".into(),
                ));
            }
            if word && local == b"sectPr" {
                if !context.section_boundary {
                    // ECMA-376 permits `w:pPr/w:sectPr` in the common
                    // paragraph property model, but a table-cell paragraph
                    // cannot terminate a main-story section. Keep its
                    // content inert and leave boundary/range accounting to
                    // the direct body story.
                    return Ok(());
                }
                context.section_seen = true;
                let ownership = Ownership::Paragraph(context.position);
                if empty {
                    append_section(
                        xml,
                        event_end,
                        Capture {
                            start: event_start,
                            depth: child_depth,
                            ownership,
                            reference_count: 0,
                            reference_bytes: 0,
                            relationship_bindings: Vec::new(),
                            root_namespace_bindings: namespace_bindings,
                        },
                        limits,
                        state,
                    )?;
                } else {
                    *capture = Some(Capture {
                        start: event_start,
                        depth: child_depth,
                        ownership,
                        reference_count: 0,
                        reference_bytes: 0,
                        relationship_bindings: Vec::new(),
                        root_namespace_bindings: namespace_bindings,
                    });
                }
                return Ok(());
            }
        }
    }

    if word && local == b"sectPr" {
        if body_depth.is_some_and(|body| child_depth == body + 1) {
            state.body_final = true;
            if empty {
                append_section(
                    xml,
                    event_end,
                    Capture {
                        start: event_start,
                        depth: child_depth,
                        ownership: Ownership::BodyFinal,
                        reference_count: 0,
                        reference_bytes: 0,
                        relationship_bindings: Vec::new(),
                        root_namespace_bindings: namespace_bindings,
                    },
                    limits,
                    state,
                )?;
            } else {
                *capture = Some(Capture {
                    start: event_start,
                    depth: child_depth,
                    ownership: Ownership::BodyFinal,
                    reference_count: 0,
                    reference_bytes: 0,
                    relationship_bindings: Vec::new(),
                    root_namespace_bindings: namespace_bindings,
                });
            }
            return Ok(());
        }
        return Err(Error::InvalidFormat(
            "w:sectPr must be body-final or a direct w:pPr child".into(),
        ));
    }
    Ok(())
}

fn append_section(
    xml: &[u8],
    end: usize,
    capture: Capture,
    limits: &Limits,
    state: &mut ScanState,
) -> Result<()> {
    let length = end
        .checked_sub(capture.start)
        .ok_or_else(|| Error::InvalidFormat("section fragment range underflow".into()))?;
    enforce_limit("section bytes", length, limits.max_section_bytes)?;
    enforce_limit(
        "sections",
        state.sections.len().saturating_add(1),
        limits.max_sections,
    )?;
    let source = xml.get(capture.start..end).ok_or_else(|| {
        Error::InvalidFormat("section fragment is outside main-document XML".into())
    })?;
    let relationship_bindings = capture.relationship_bindings.clone();
    let namespace_bindings = capture.root_namespace_bindings.clone();
    let fragment = self_contained_fragment(source, &namespace_bindings, limits.max_section_bytes)?;
    let mut section = Section::from_xml_bytes(fragment)?;
    let semantic = section.local_state()?;
    let added_reference_bytes = semantic
        .headers
        .iter()
        .chain(&semantic.footers)
        .try_fold(0usize, |total, reference| {
            total.checked_add(reference.relationship_id.len())
        })
        .ok_or_else(|| Error::InvalidFormat("section reference byte count overflow".into()))?;
    if added_reference_bytes != capture.reference_bytes {
        return Err(Error::InvalidFormat(
            "section reference preflight disagrees with semantic decoding".into(),
        ));
    }
    state.reference_bytes = state
        .reference_bytes
        .checked_add(added_reference_bytes)
        .ok_or_else(|| Error::InvalidFormat("section reference byte count overflow".into()))?;
    enforce_limit(
        "header/footer reference bytes",
        state.reference_bytes,
        limits.max_reference_bytes,
    )?;
    let boundary = match capture.ownership {
        Ownership::Paragraph(position) => position
            .get()
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("section paragraph boundary overflow".into()))?,
        Ownership::BodyFinal | Ownership::Implicit => state.paragraphs,
    };
    if boundary < state.next_section_start {
        return Err(Error::InvalidFormat(
            "section paragraph boundaries are out of source order".into(),
        ));
    }
    state
        .sections
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "section inventory descriptors",
            source,
        })?;
    state
        .sources
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "section inventory source spans",
            source,
        })?;
    state.sections.push(Descriptor {
        position: Position::new(state.sections.len()),
        ownership: capture.ownership,
        paragraphs: ParagraphRange {
            start: Position::new(state.next_section_start),
            end: Position::new(boundary),
        },
        page_size: semantic.page_size,
        margins: semantic.margins,
        start: semantic.start,
        columns: semantic.columns,
        headers: semantic.headers.into_boxed_slice(),
        footers: semantic.footers.into_boxed_slice(),
    });
    state.sources.push(Some(SourceSpan {
        start: capture.start,
        end,
        relationship_bindings: relationship_bindings.into_boxed_slice(),
        namespace_bindings: namespace_bindings.into_boxed_slice(),
    }));
    state.next_section_start = boundary;
    Ok(())
}

fn validate_namespace_bindings(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    validate_element_qname(element.name().as_ref(), "main-document")?;
    if let ResolveResult::Unknown(prefix) = namespace {
        return Err(Error::InvalidFormat(format!(
            "main-document element uses unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )));
    }
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        validate_qname(attribute.key.as_ref(), "main-document", "attribute")?;
        let prefix = attribute.key.prefix();
        let is_namespace_declaration = (prefix.is_none()
            && attribute.key.local_name().as_ref() == b"xmlns")
            || prefix
                .as_ref()
                .is_some_and(|value| value.as_ref() == b"xmlns");
        if is_namespace_declaration {
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            validate_namespace_declaration(
                attribute.key.as_ref(),
                value.as_ref(),
                "main-document",
            )?;
            continue;
        }
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        if let ResolveResult::Unknown(prefix) = attribute_namespace {
            return Err(Error::InvalidFormat(format!(
                "main-document attribute uses unbound namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn inspect_reference(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    depth: usize,
    limits: &Limits,
    retained_reference_bytes: usize,
    capture: Option<&mut Capture>,
) -> Result<()> {
    let Some(capture) = capture else {
        return Ok(());
    };
    if depth != capture.depth + 1
        || !is_wordprocessing_namespace(namespace)
        || !matches!(
            element.local_name().as_ref(),
            b"headerReference" | b"footerReference"
        )
    {
        return Ok(());
    }
    capture.reference_count = capture
        .reference_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("section reference count overflow".into()))?;
    if capture.reference_count > 6 {
        return Err(Error::InvalidFormat(
            "section has too many header/footer references".into(),
        ));
    }
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        let Some(namespace) = relationship_namespace(&attribute_namespace) else {
            continue;
        };
        let raw_total = retained_reference_bytes
            .checked_add(capture.reference_bytes)
            .and_then(|value| value.checked_add(attribute.value.len()))
            .ok_or_else(|| Error::InvalidFormat("section reference byte count overflow".into()))?;
        enforce_limit(
            "header/footer reference bytes",
            raw_total,
            limits.max_reference_bytes,
        )?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        capture.reference_bytes = capture
            .reference_bytes
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("section reference byte count overflow".into()))?;
        enforce_limit(
            "header/footer reference bytes",
            retained_reference_bytes
                .checked_add(capture.reference_bytes)
                .ok_or_else(|| {
                    Error::InvalidFormat("section reference byte count overflow".into())
                })?,
            limits.max_reference_bytes,
        )?;
        let prefix = attribute.key.prefix().ok_or_else(|| {
            Error::InvalidFormat("section relationship ID has no namespace prefix".into())
        })?;
        if !capture
            .relationship_bindings
            .iter()
            .any(|binding| binding.prefix.as_slice() == prefix.as_ref())
        {
            let mut owned_prefix = Vec::new();
            owned_prefix
                .try_reserve_exact(prefix.as_ref().len())
                .map_err(|source| Error::Allocation {
                    resource: "section relationship namespace prefix",
                    source,
                })?;
            owned_prefix.extend_from_slice(prefix.as_ref());
            capture
                .relationship_bindings
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "section relationship namespace bindings",
                    source,
                })?;
            capture.relationship_bindings.push(RelationshipBinding {
                prefix: owned_prefix,
                namespace: namespace.as_bytes().to_vec(),
            });
        }
    }
    Ok(())
}

fn relationship_namespace(namespace: &ResolveResult<'_>) -> Option<&'static str> {
    const TRANSITIONAL: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
    match namespace {
        ResolveResult::Bound(quick_xml::name::Namespace(value)) if *value == TRANSITIONAL => {
            Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        },
        ResolveResult::Bound(quick_xml::name::Namespace(value)) if *value == STRICT => {
            Some("http://purl.oclc.org/ooxml/officeDocument/relationships")
        },
        ResolveResult::Bound(_) | ResolveResult::Unknown(_) | ResolveResult::Unbound => None,
    }
}

fn namespace_declarations(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<RelationshipBinding>> {
    let mut bindings = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let attribute_prefix = attribute
            .key
            .prefix()
            .map_or_else(Vec::new, |prefix| prefix.as_ref().to_vec());
        let prefix = if attribute_prefix.as_slice() == b"xmlns" {
            attribute.key.local_name().as_ref().to_vec()
        } else if attribute_prefix.is_empty() && attribute.key.local_name().as_ref() == b"xmlns" {
            Vec::new()
        } else {
            continue;
        };
        let namespace = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        bindings
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "section namespace declarations",
                source,
            })?;
        bindings.push(RelationshipBinding {
            prefix,
            namespace: namespace.as_bytes().to_vec(),
        });
    }
    Ok(bindings)
}

fn self_contained_fragment(
    source: &[u8],
    bindings: &[RelationshipBinding],
    max_section_bytes: usize,
) -> Result<Vec<u8>> {
    let open_end = root_opening_end(source)?;
    let mut added = 0usize;
    for binding in bindings {
        if !root_declares_prefix(source, &binding.prefix)? {
            added = added
                .checked_add(10)
                .and_then(|value| value.checked_add(binding.prefix.len()))
                .and_then(|value| value.checked_add(binding.namespace.len()))
                .ok_or_else(|| {
                    Error::InvalidFormat("section namespace declaration size overflow".into())
                })?;
        }
    }
    let capacity = source
        .len()
        .checked_add(added)
        .ok_or_else(|| Error::InvalidFormat("section fragment size overflow".into()))?;
    enforce_limit("section bytes", capacity, max_section_bytes)?;
    let mut fragment = Vec::new();
    fragment
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "section inventory fragment",
            source,
        })?;
    let insertion = if open_end > 0 && source[open_end - 1] == b'/' {
        open_end - 1
    } else {
        open_end
    };
    fragment.extend_from_slice(&source[..insertion]);
    for binding in bindings {
        if root_declares_prefix(source, &binding.prefix)? {
            continue;
        }
        if binding.prefix.is_empty() {
            fragment.extend_from_slice(b" xmlns=\"");
        } else {
            fragment.extend_from_slice(b" xmlns:");
            fragment.extend_from_slice(&binding.prefix);
            fragment.extend_from_slice(b"=\"");
        }
        fragment.extend_from_slice(&binding.namespace);
        fragment.push(b'"');
    }
    fragment.extend_from_slice(&source[insertion..]);
    Ok(fragment)
}

fn root_opening_end(source: &[u8]) -> Result<usize> {
    if source.first() != Some(&b'<') {
        return Err(Error::InvalidFormat(
            "section fragment does not begin with an opening element".into(),
        ));
    }
    let mut quote = None;
    for (index, byte) in source.iter().enumerate().skip(1) {
        if let Some(delimiter) = quote {
            if *byte == delimiter {
                quote = None;
            }
            continue;
        }
        match *byte {
            b'\'' | b'"' => quote = Some(*byte),
            b'>' => return Ok(index),
            _ => {},
        }
    }
    if quote.is_some() {
        return Err(Error::InvalidFormat(
            "section opening element has an unterminated quoted attribute".into(),
        ));
    }
    Err(Error::InvalidFormat(
        "section opening element is incomplete".into(),
    ))
}

fn root_declares_prefix(source: &[u8], prefix: &[u8]) -> Result<bool> {
    let mut reader = quick_xml::Reader::from_reader(source);
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let declares_prefix = if prefix.is_empty() {
                        attribute.key.prefix().is_none()
                            && attribute.key.local_name().as_ref() == b"xmlns"
                    } else {
                        attribute
                            .key
                            .prefix()
                            .is_some_and(|value| value.as_ref() == b"xmlns")
                            && attribute.key.local_name().as_ref() == prefix
                    };
                    if declares_prefix {
                        return Ok(true);
                    }
                }
                return Ok(false);
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "section fragment has no root element".into(),
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn append_implicit(state: &mut ScanState, limits: &Limits) -> Result<()> {
    enforce_limit(
        "sections",
        state.sections.len().saturating_add(1),
        limits.max_sections,
    )?;
    state
        .sections
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "section inventory descriptors",
            source,
        })?;
    state
        .sources
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "section inventory source spans",
            source,
        })?;
    state.sections.push(Descriptor {
        position: Position::new(state.sections.len()),
        ownership: Ownership::Implicit,
        paragraphs: ParagraphRange {
            start: Position::new(state.next_section_start),
            end: Position::new(state.paragraphs),
        },
        page_size: None,
        margins: None,
        start: None,
        columns: None,
        headers: Box::default(),
        footers: Box::default(),
    });
    state.sources.push(None);
    Ok(())
}

fn enforce_limit(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        return Err(Error::SectionInventoryLimit {
            resource,
            maximum,
            actual,
        });
    }
    Ok(())
}

fn offset(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))
}
