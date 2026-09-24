//! In-content RDFa and `text:meta` metadata.
//!
//! ODF has two metadata surfaces which are deliberately separate from
//! `meta.xml` and from package RDF graphs: RDFa attributes attached to text
//! content and the inline `text:meta` element.  This module owns the small,
//! inert typed view of those declarations.  RDFa CURIEs and datatypes are
//! retained as strings; they are never resolved, fetched, or evaluated.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::doc_markdown,
    clippy::format_push_string,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::missing_errors_doc,
    clippy::must_use_candidate,
    clippy::needless_pass_by_value,
    clippy::similar_names,
    reason = "the metadata codec is a bounded XML projection with compatibility-oriented naming"
)]

use crate::core::ResolvedReader;
use crate::elements::field::{
    MetaFieldAttribute, MetaFieldContent, MetaFieldElement, MetaFieldNode,
};
use crate::generic::{ChargedXml, FlatMutationBudget, MemoryLease, allocate_xml};
use crate::namespace::{
    DCNS, DR3DNS, DRAWNS, FONS, FORMNS, METANS, NUMBERNS, OFFICENS, PRESENTATIONNS, SCRIPTNS,
    STYLENS, SVGNS, TABLENS, TEXTNS, XHTMLNS, XLINKNS, XMLNS, XSDNS,
};
use litchi_core::{Error, Position, Reservation, Resource, ResourceLimit, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{PrefixDeclaration, QName, ResolveResult};
use quick_xml::reader::Reader;
use std::{
    collections::{HashMap, HashSet},
    fmt::Write as _,
    mem::{align_of, size_of},
};

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 512;
const MAX_OCCURRENCES: usize = 1_000_000;
const MAX_ATTRIBUTES: usize = 256;
const MAX_TEXT_META_NODES: usize = 1_000_000;
const MAX_TEXT_META_DEPTH: usize = 256;
const MAX_TEXT_META_BYTES: usize = 16 * 1024 * 1024;
// Every RDFa lexical field is independently bounded; this ceiling also
// bounds repeated property tokens before any prefix index is built.
const MAX_RDFA_CURIE_TOKENS: usize = MAX_TEXT_META_BYTES * 3;

/// Reservations for owned metadata parser state.
///
/// The XML reader borrows events from the source string, but the typed view
/// intentionally owns resolved names, attribute values, namespace snapshots,
/// and mixed-content nodes.  Keeping one coalesced reservation lets the
/// budgeted paths charge each requested allocation before it happens without
/// growing a second, uncharged reservation ledger.
struct MetadataMemory<'a> {
    budget: Option<&'a FlatMutationBudget>,
    reservation: Option<Reservation>,
}

impl<'a> MetadataMemory<'a> {
    fn new(budget: Option<&'a FlatMutationBudget>) -> Self {
        Self {
            budget,
            reservation: None,
        }
    }

    fn check(&self) -> Result<()> {
        self.budget.map_or(Ok(()), FlatMutationBudget::check)
    }

    fn merge_reservation(
        &mut self,
        reservation: Reservation,
        resource: &'static str,
    ) -> Result<()> {
        if let Some(existing) = &mut self.reservation {
            if existing.try_merge(reservation).is_err() {
                return Err(Error::InvalidFormat(format!(
                    "{resource} reservation chain mismatch"
                )));
            }
        } else {
            self.reservation = Some(reservation);
        }
        Ok(())
    }

    fn reserve_bytes(&mut self, amount: usize, resource: &'static str) -> Result<()> {
        self.check()?;
        if amount == 0 {
            return Ok(());
        }
        let Some(budget) = self.budget else {
            return Ok(());
        };
        let reservation = budget.reserve_bytes(amount, resource)?;
        self.merge_reservation(reservation, resource)
    }

    /// Reserve a growth destination as a short-lived scratch charge.  The
    /// existing allocation remains live until the fallible growth returns;
    /// callers drop this token before retaining only the new allocation's
    /// net capacity in the aggregate lease.
    fn reserve_growth_scratch(
        &self,
        amount: usize,
        resource: &'static str,
    ) -> Result<Option<Reservation>> {
        if amount == 0 {
            return Ok(None);
        }
        self.budget
            .map(|budget| budget.reserve_bytes(amount, resource))
            .transpose()
    }

    fn reserve_vec<T>(
        &mut self,
        items: &mut Vec<T>,
        additional: usize,
        resource: &'static str,
    ) -> Result<()> {
        if additional == 0 {
            return Ok(());
        }
        let required = items
            .len()
            .checked_add(additional)
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        if required <= items.capacity() {
            return Ok(());
        }
        let old_capacity = items.capacity();
        let requested_bytes = required
            .checked_mul(size_of::<T>())
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        let old_bytes = old_capacity
            .checked_mul(size_of::<T>())
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        let requested_delta = requested_bytes
            .checked_sub(old_bytes)
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        let mut retained = self.reserve_growth_scratch(requested_delta, resource)?;
        let scratch = self.reserve_growth_scratch(old_bytes, resource)?;
        items
            .try_reserve_exact(additional)
            .map_err(|source| Error::Allocation { resource, source })?;
        let actual_capacity = items.capacity();
        let actual_allocation_bytes = actual_capacity
            .checked_mul(size_of::<T>())
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        if actual_allocation_bytes > requested_bytes {
            if let (Some(retained), Some(budget)) = (retained.as_mut(), self.budget) {
                let extra =
                    budget.reserve_bytes(actual_allocation_bytes - requested_bytes, resource)?;
                retained.try_merge(extra).map_err(|_reservation| {
                    Error::InvalidFormat(format!("{resource} reservation chain changed"))
                })?;
            }
        }
        if let Some(retained) = retained {
            self.merge_reservation(retained, resource)?;
        }
        drop(scratch);
        Ok(())
    }

    fn clone_string(&mut self, value: &str, resource: &'static str) -> Result<String> {
        let mut output = String::new();
        reserve_string_growth(&mut output, value.len(), self, resource)?;
        output.push_str(value);
        Ok(output)
    }

    fn clone_option_string(
        &mut self,
        value: Option<&str>,
        resource: &'static str,
    ) -> Result<Option<String>> {
        value
            .map(|value| self.clone_string(value, resource))
            .transpose()
    }

    fn into_memory_lease(self) -> Option<MemoryLease> {
        self.reservation.map(MemoryLease::new)
    }

    fn namespace_declaration(
        &mut self,
        name: &str,
        value: &str,
        resource: &'static str,
    ) -> Result<String> {
        let length = 1usize
            .checked_add(name.len())
            .and_then(|length| length.checked_add(3))
            .and_then(|length| length.checked_add(value.len()))
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
        let mut output = String::new();
        reserve_string_growth(&mut output, length, self, resource)?;
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(value);
        output.push('"');
        Ok(output)
    }
}

/// The XML part in which an in-content declaration was found.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MetadataPart {
    /// `content.xml`.
    Content,
    /// `styles.xml`.
    Styles,
}

/// RDFa attributes defined by ODF's in-content metadata profile.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct RdfaAttributes {
    /// RDFa `about` resource.
    pub about: Option<String>,
    /// RDFa `property` CURIE or CURIE list.
    pub property: Option<String>,
    /// RDFa `content` lexical value.
    pub content: Option<String>,
    /// RDFa `datatype` lexical datatype.
    pub datatype: Option<String>,
}

/// An unmodeled attribute retained on a `text:meta` root.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetadataAttribute {
    /// Namespace URI, or empty for an unqualified attribute.
    pub namespace_uri: String,
    /// Local attribute name.
    pub local_name: String,
    /// Lexical value.
    pub value: String,
    /// Source prefix when the attribute was qualified.
    pub prefix: Option<String>,
}

impl RdfaAttributes {
    /// Return whether no RDFa attribute is present.
    pub fn is_empty(&self) -> bool {
        self.about.is_none()
            && self.property.is_none()
            && self.content.is_none()
            && self.datatype.is_none()
    }

    /// Validate XML lexical limits without resolving RDFa values.
    pub fn validate(&self) -> Result<()> {
        for (name, value) in [
            ("xhtml:about", self.about.as_deref()),
            ("xhtml:property", self.property.as_deref()),
            ("xhtml:content", self.content.as_deref()),
            ("xhtml:datatype", self.datatype.as_deref()),
        ] {
            if let Some(value) = value {
                validate_xml_value(name, value, MAX_TEXT_META_BYTES)?;
            }
        }
        Ok(())
    }
}

/// Semantic host for one RDFa occurrence.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum RdfaHost {
    /// A `text:p` in document order within its part.
    Paragraph { index: usize },
    /// A `text:h` in document order within its part.
    Heading { index: usize },
    /// A `text:bookmark-start`; the optional name is retained as written.
    BookmarkStart { name: Option<String> },
    /// A `text:meta` element.  Its full typed value is also available in
    /// [`ContentMetadata::text_meta`].
    TextMeta { index: usize },
    /// Another namespace-resolved XML host carrying RDFa attributes.
    Element {
        namespace_uri: String,
        local_name: String,
        occurrence: usize,
    },
}

/// One in-content RDFa declaration.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RdfaOccurrence {
    /// XML part containing the host.
    pub part: MetadataPart,
    /// Host element identity.
    pub host: RdfaHost,
    /// Retained RDFa attributes.
    pub attributes: RdfaAttributes,
}

/// A typed `text:meta` inline element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextMeta {
    /// XML part containing the element.
    pub part: MetadataPart,
    /// Zero-based `text:meta` occurrence in the part.
    pub index: usize,
    /// Optional `xml:id`.
    pub xml_id: Option<String>,
    /// RDFa attributes attached to the element.
    pub rdfa: RdfaAttributes,
    /// Unknown root attributes retained by expanded name.
    pub attributes: Vec<MetadataAttribute>,
    /// Inert, validated paragraph-content mixed content.
    pub content: MetaFieldContent,
}

impl TextMeta {
    /// Construct an inline metadata element for a new document or insertion.
    pub fn new(content: MetaFieldContent) -> Result<Self> {
        content.validate()?;
        Ok(Self {
            part: MetadataPart::Content,
            index: 0,
            xml_id: None,
            rdfa: RdfaAttributes::default(),
            attributes: Vec::new(),
            content,
        })
    }

    /// Construct a plain-text inline metadata element.
    pub fn from_text(text: impl Into<String>) -> Result<Self> {
        Self::new(MetaFieldContent::new(vec![MetaFieldNode::Text(
            text.into(),
        )])?)
    }

    /// Validate attributes, identity, and mixed content.
    pub fn validate(&self) -> Result<()> {
        self.rdfa.validate()?;
        validate_metadata_attributes(&self.attributes)?;
        if let Some(id) = &self.xml_id {
            validate_xml_id(id)?;
        }
        self.content.validate()
    }

    #[allow(dead_code)]
    fn to_xml_with_context_with_limit(
        &self,
        namespace_declarations: &[String],
        maximum: usize,
    ) -> Result<String> {
        self.to_xml_with_context_with_limit_and_budget(namespace_declarations, maximum, None)
            .map(ChargedXml::into_string)
    }

    fn to_xml_with_context_with_limit_and_budget(
        &self,
        namespace_declarations: &[String],
        maximum: usize,
        budget: Option<&FlatMutationBudget>,
    ) -> Result<ChargedXml> {
        let mut scratch = MetadataMemory::new(budget);
        self.rdfa.validate()?;
        validate_metadata_attributes(&self.attributes)?;
        if let Some(id) = &self.xml_id {
            validate_xml_id(id)?;
        }
        self.content.validate_with_reservation(|amount, resource| {
            scratch.reserve_bytes(amount, resource)
        })?;
        let output_len = metadata_output_len(self, namespace_declarations, &mut scratch)?;
        bounded_output_len_with_limit(output_len, "text:meta serialization", maximum)?;
        let (mut output, memory) = allocate_xml(budget, output_len, "ODT text:meta serialization")?;
        output.push_str("<text:meta xmlns:text=\"");
        output.push_str(TEXTNS);
        output.push_str("\" xmlns:xhtml=\"");
        output.push_str(XHTMLNS);
        output.push('"');
        if let Some(id) = &self.xml_id {
            output.push_str(" xml:id=\"");
            output.push_str(&escaped_metadata_value(
                id,
                &mut scratch,
                "ODT text:meta xml:id escape scratch",
            )?);
            output.push('"');
        }
        push_rdfa_namespace_declarations(
            &mut output,
            &self.rdfa,
            namespace_declarations,
            &mut scratch,
        )?;
        push_rdfa_attributes(&mut output, &self.rdfa, &mut scratch)?;
        push_metadata_attributes(&mut output, &self.attributes, &mut scratch)?;
        if self.content.nodes().is_empty() {
            output.push_str("/>");
        } else {
            output.push('>');
            self.content.write_xml_to(&mut output);
            output.push_str("</text:meta>");
        }
        Ok(ChargedXml {
            xml: output,
            memory,
        })
    }
}

/// Bounded typed in-content metadata from one ODF document package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ContentMetadata {
    /// RDFa declarations in document and style XML parts.
    pub rdfa: Vec<RdfaOccurrence>,
    /// Inline `text:meta` elements in document and style XML parts.
    pub text_meta: Vec<TextMeta>,
}

impl ContentMetadata {
    /// Return whether no in-content metadata was found.
    pub fn is_empty(&self) -> bool {
        self.rdfa.is_empty() && self.text_meta.is_empty()
    }

    /// Validate all retained metadata.
    pub fn validate(&self) -> Result<()> {
        for item in &self.rdfa {
            item.attributes.validate()?;
        }
        for item in &self.text_meta {
            item.validate()?
        }
        Ok(())
    }
}

/// Parse in-content metadata from the supplied XML parts.
pub(crate) fn parse_parts(parts: &[(&str, MetadataPart)]) -> Result<ContentMetadata> {
    parse_parts_with_optional_budget(parts, None).map(|(metadata, _)| metadata)
}

fn parse_parts_with_optional_budget(
    parts: &[(&str, MetadataPart)],
    budget: Option<&FlatMutationBudget>,
) -> Result<(ContentMetadata, Option<MemoryLease>)> {
    let mut memory = MetadataMemory::new(budget);
    let mut output = ContentMetadata::default();
    for (xml, part) in parts {
        let parsed = parse_part_inner(xml, *part, &mut memory)?;
        for item in parsed.rdfa {
            push_bounded_with_memory(
                &mut output.rdfa,
                item,
                "ODF aggregate in-content RDFa occurrence",
                &mut memory,
            )?;
        }
        for item in parsed.text_meta {
            push_bounded_with_memory(
                &mut output.text_meta,
                item,
                "ODF aggregate text:meta projection",
                &mut memory,
            )?;
        }
    }
    let memory = memory.into_memory_lease();
    Ok((output, memory))
}

/// Parse in-content metadata while charging parser and projection ownership.
pub(crate) fn parse_parts_with_budget(
    parts: &[(&str, MetadataPart)],
    budget: &FlatMutationBudget,
) -> Result<(ContentMetadata, MemoryLease)> {
    budget.check()?;
    let (metadata, memory) = parse_parts_with_optional_budget(parts, Some(budget))?;
    Ok((metadata, memory.unwrap_or_default()))
}

/// Parse one XML part.  This is crate-visible so mutable editors can inspect
/// their authoritative content snapshot without retaining a second copy.
pub(crate) fn parse_part(xml: &str, part: MetadataPart) -> Result<ContentMetadata> {
    parse_part_with_budget(xml, part, None).map(|(metadata, _)| metadata)
}

/// Parse one XML part while charging parser and projection ownership.
pub(crate) fn parse_part_with_budget(
    xml: &str,
    part: MetadataPart,
    budget: Option<&FlatMutationBudget>,
) -> Result<(ContentMetadata, Option<MemoryLease>)> {
    let mut memory = MetadataMemory::new(budget);
    let output = parse_part_inner(xml, part, &mut memory)?;
    let memory = memory.into_memory_lease();
    Ok((output, memory))
}

fn parse_part_inner(
    xml: &str,
    part: MetadataPart,
    memory: &mut MetadataMemory<'_>,
) -> Result<ContentMetadata> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} in-content metadata limit"
        )));
    }
    reserve_namespace_resolver_memory(xml, memory)?;
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();
    let mut active: Vec<ActiveTextMeta> = Vec::new();
    let mut output = ContentMetadata::default();
    let mut paragraph_index = 0usize;
    let mut heading_index = 0usize;
    let mut text_meta_index = 0usize;
    let mut other_occurrences: HashMap<(String, String), usize> = HashMap::new();
    let mut seen = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF metadata XML: {error}")))?;
        let namespace_uri = match &event {
            Event::Start(_) | Event::Empty(_) => {
                resolved_namespace_with_memory(&namespace, memory)?
            },
            _ => None,
        };
        match event {
            Event::Start(ref source) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("in-content metadata XML depth overflow".to_string())
                })?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "in-content metadata XML exceeds {MAX_DEPTH} levels"
                    )));
                }
                let local = utf8_with_memory(
                    source.local_name().as_ref(),
                    "metadata element name",
                    memory,
                )?;
                let attrs = parse_attributes(&reader, source, memory)?;
                let rdfa = rdfa_from_attributes(&attrs, memory)?;
                let host = host_for(
                    namespace_uri.as_deref(),
                    &local,
                    &attrs,
                    &mut paragraph_index,
                    &mut heading_index,
                    &mut text_meta_index,
                    &mut other_occurrences,
                    memory,
                )?;
                if let Some(ref host) = host {
                    if !rdfa.is_empty() {
                        push_bounded_with_memory(
                            &mut output.rdfa,
                            RdfaOccurrence {
                                part,
                                host: clone_rdfa_host(host, memory)?,
                                attributes: clone_rdfa_attributes(&rdfa, memory)?,
                            },
                            "ODF in-content RDFa occurrence",
                            memory,
                        )?;
                    }
                }
                // A metadata element contributes its complete root to every
                // active outer `text:meta` builder before becoming a new root.
                if !active.is_empty() {
                    let child_attributes = parse_meta_attributes(&reader, source, memory)?;
                    let child_namespace = namespace_uri.as_deref().ok_or_else(|| {
                        Error::InvalidFormat(
                            "unqualified element inside text:meta is not supported".to_string(),
                        )
                    })?;
                    for root in &mut active {
                        root.builder.start_element(
                            memory.clone_string(
                                child_namespace,
                                "ODT text:meta namespace projection",
                            )?,
                            memory.clone_string(&local, "ODT text:meta element name projection")?,
                            clone_meta_field_attributes(&child_attributes, memory)?,
                            memory,
                        )?;
                    }
                }
                let is_meta = namespace_uri.as_deref() == Some(TEXTNS) && local == "meta";
                if is_meta {
                    let (xml_id, root_rdfa, root_attributes) =
                        text_meta_root_attributes(&reader, source, &attrs, memory)?;
                    let order = match host {
                        Some(RdfaHost::TextMeta { index }) => index,
                        _ => {
                            return Err(Error::InvalidFormat(
                                "text:meta host index missing".to_string(),
                            ));
                        },
                    };
                    push_bounded_with_memory(
                        &mut active,
                        ActiveTextMeta {
                            depth,
                            part,
                            index: order,
                            xml_id,
                            rdfa: root_rdfa,
                            attributes: root_attributes,
                            builder: MetaBuilder::default(),
                        },
                        "ODF active text:meta projection",
                        memory,
                    )?;
                }
                reserve_vec_push(
                    &mut stack,
                    (namespace_uri, local),
                    "ODF metadata element stack",
                    memory,
                )?;
            },
            Event::Empty(ref source) => {
                let local = utf8_with_memory(
                    source.local_name().as_ref(),
                    "metadata element name",
                    memory,
                )?;
                let attrs = parse_attributes(&reader, source, memory)?;
                let rdfa = rdfa_from_attributes(&attrs, memory)?;
                let host = host_for(
                    namespace_uri.as_deref(),
                    &local,
                    &attrs,
                    &mut paragraph_index,
                    &mut heading_index,
                    &mut text_meta_index,
                    &mut other_occurrences,
                    memory,
                )?;
                if let Some(ref host) = host {
                    if !rdfa.is_empty() {
                        push_bounded_with_memory(
                            &mut output.rdfa,
                            RdfaOccurrence {
                                part,
                                host: clone_rdfa_host(host, memory)?,
                                attributes: clone_rdfa_attributes(&rdfa, memory)?,
                            },
                            "ODF in-content RDFa occurrence",
                            memory,
                        )?;
                    }
                }
                if !active.is_empty() {
                    let child_attributes = parse_meta_attributes(&reader, source, memory)?;
                    let child_namespace = namespace_uri.as_deref().ok_or_else(|| {
                        Error::InvalidFormat(
                            "unqualified element inside text:meta is not supported".to_string(),
                        )
                    })?;
                    for root in &mut active {
                        root.builder.empty_element(
                            memory.clone_string(
                                child_namespace,
                                "ODT text:meta namespace projection",
                            )?,
                            memory.clone_string(&local, "ODT text:meta element name projection")?,
                            clone_meta_field_attributes(&child_attributes, memory)?,
                            memory,
                        )?;
                    }
                }
                if namespace_uri.as_deref() == Some(TEXTNS) && local == "meta" {
                    let (xml_id, root_rdfa, root_attributes) =
                        text_meta_root_attributes(&reader, source, &attrs, memory)?;
                    let index = match host {
                        Some(RdfaHost::TextMeta { index }) => index,
                        _ => {
                            return Err(Error::InvalidFormat(
                                "text:meta host index missing".to_string(),
                            ));
                        },
                    };
                    let value = TextMeta {
                        part,
                        index,
                        xml_id,
                        rdfa: root_rdfa,
                        attributes: root_attributes,
                        content: MetaFieldContent::new(Vec::new())?,
                    };
                    push_bounded_with_memory(
                        &mut output.text_meta,
                        value,
                        "ODF completed text:meta projection",
                        memory,
                    )?;
                }
            },
            Event::Text(ref value) => {
                if metadata_text_decode_needs_owned(value.as_ref()) {
                    memory.reserve_bytes(
                        value.as_ref().len(),
                        "ODT text:meta character-data decode scratch",
                    )?;
                }
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text:meta character data: {error}"))
                    })?;
                for root in &mut active {
                    root.builder.text(&value, memory)?;
                }
            },
            Event::CData(ref value) => {
                if metadata_text_decode_needs_owned(value.as_ref()) {
                    memory.reserve_bytes(
                        value.as_ref().len(),
                        "ODT text:meta CDATA decode scratch",
                    )?;
                }
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text:meta CDATA: {error}"))
                    })?;
                for root in &mut active {
                    root.builder.text(&value, memory)?;
                }
            },
            Event::GeneralRef(ref value) => {
                let reference = std::str::from_utf8(value.as_ref()).map_err(|_| {
                    Error::InvalidFormat("invalid XML entity reference in text:meta".to_string())
                })?;
                let value = quick_xml::escape::resolve_xml_entity(reference).ok_or_else(|| {
                    Error::InvalidFormat("unknown XML entity in text:meta".to_string())
                })?;
                for root in &mut active {
                    root.builder.text(value, memory)?;
                }
            },
            Event::End(_) => {
                for root in &mut active {
                    if root.depth < depth {
                        root.builder.end_element(memory)?;
                    }
                }
                if let Some(root) = active.pop_if(|root| root.depth == depth) {
                    let content = root.builder.finish(memory)?;
                    let value = TextMeta {
                        part: root.part,
                        index: root.index,
                        xml_id: root.xml_id,
                        rdfa: root.rdfa,
                        attributes: root.attributes,
                        content,
                    };
                    push_bounded_with_memory(
                        &mut output.text_meta,
                        value,
                        "ODF completed text:meta projection",
                        memory,
                    )?;
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("in-content metadata XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("in-content metadata XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF metadata XML".to_string(),
                ));
            },
            Event::Comment(_) | Event::PI(_) if !active.is_empty() => {
                return Err(Error::InvalidFormat(
                    "comments and processing instructions inside text:meta are not supported by the typed editor".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        seen = seen.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("in-content metadata event count overflow".to_string())
        })?;
        if seen > MAX_OCCURRENCES {
            return Err(Error::InvalidFormat(format!(
                "in-content metadata exceeds {MAX_OCCURRENCES} XML events"
            )));
        }
        if seen & 0x03ff == 0 {
            memory.check()?;
        }
    }
    if depth != 0 || !stack.is_empty() || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete in-content metadata XML".to_string(),
        ));
    }
    output.text_meta.sort_by_key(|value| value.index);
    Ok(output)
}

#[derive(Debug)]
struct ActiveTextMeta {
    depth: usize,
    part: MetadataPart,
    index: usize,
    xml_id: Option<String>,
    rdfa: RdfaAttributes,
    attributes: Vec<MetadataAttribute>,
    builder: MetaBuilder,
}

#[derive(Debug, Default)]
struct MetaBuilder {
    roots: Vec<MetaFieldNode>,
    stack: Vec<MetaFieldElement>,
    nodes: usize,
    bytes: usize,
}

impl MetaBuilder {
    fn text(&mut self, value: &str, memory: &mut MetadataMemory<'_>) -> Result<()> {
        self.add_node(value.len())?;
        if let Some(MetaFieldNode::Text(existing)) = self.current_mut().last_mut() {
            reserve_string_growth(
                existing,
                value.len(),
                memory,
                "ODT text:meta text projection",
            )?;
            existing.push_str(value);
        } else {
            let current = self.current_mut();
            memory.reserve_vec(current, 1, "ODT text:meta node projection")?;
            let text = memory.clone_string(value, "ODT text:meta text projection")?;
            current.push(MetaFieldNode::Text(text));
        }
        Ok(())
    }

    fn start_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<MetaFieldAttribute>,
        memory: &mut MetadataMemory<'_>,
    ) -> Result<()> {
        self.add_node(namespace_uri.len().saturating_add(local_name.len()))?;
        if self.stack.len() >= MAX_TEXT_META_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "text:meta content exceeds {MAX_TEXT_META_DEPTH} levels"
            )));
        }
        memory.reserve_vec(&mut self.stack, 1, "ODT text:meta element stack")?;
        let children = Vec::new();
        self.stack.push(MetaFieldElement {
            namespace_uri,
            local_name,
            attributes,
            children,
        });
        Ok(())
    }

    fn empty_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<MetaFieldAttribute>,
        memory: &mut MetadataMemory<'_>,
    ) -> Result<()> {
        self.start_element(namespace_uri, local_name, attributes, memory)?;
        self.end_element(memory)
    }

    fn end_element(&mut self, memory: &mut MetadataMemory<'_>) -> Result<()> {
        let element = self
            .stack
            .pop()
            .ok_or_else(|| Error::InvalidFormat("text:meta content stack underflow".to_string()))?;
        memory.reserve_vec(self.current_mut(), 1, "ODT text:meta node projection")?;
        self.current_mut().push(MetaFieldNode::Element(element));
        Ok(())
    }

    fn finish(self, memory: &mut MetadataMemory<'_>) -> Result<MetaFieldContent> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete text:meta mixed content".to_string(),
            ));
        }
        MetaFieldContent::new_with_reservation(self.roots, |amount, resource| {
            memory.reserve_bytes(amount, resource)
        })
    }

    fn current_mut(&mut self) -> &mut Vec<MetaFieldNode> {
        if let Some(element) = self.stack.last_mut() {
            &mut element.children
        } else {
            &mut self.roots
        }
    }

    fn add_node(&mut self, bytes: usize) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("text:meta node count overflow".to_string()))?;
        if self.nodes > MAX_TEXT_META_NODES {
            return Err(Error::InvalidFormat(format!(
                "text:meta exceeds {MAX_TEXT_META_NODES} content nodes"
            )));
        }
        self.bytes = self
            .bytes
            .checked_add(bytes)
            .ok_or_else(|| Error::InvalidFormat("text:meta byte count overflow".to_string()))?;
        if self.bytes > MAX_TEXT_META_BYTES {
            return Err(Error::InvalidFormat(format!(
                "text:meta exceeds {MAX_TEXT_META_BYTES} content bytes"
            )));
        }
        Ok(())
    }
}

fn validate_text_meta_with_budget(
    value: &TextMeta,
    budget: Option<&FlatMutationBudget>,
) -> Result<()> {
    let mut memory = MetadataMemory::new(budget);
    value.rdfa.validate()?;
    validate_metadata_attributes(&value.attributes)?;
    if let Some(id) = &value.xml_id {
        validate_xml_id(id)?;
    }
    value
        .content
        .validate_with_reservation(|amount, resource| memory.reserve_bytes(amount, resource))
}

fn host_for(
    namespace: Option<&str>,
    local: &str,
    attrs: &[MetaFieldAttribute],
    paragraph_index: &mut usize,
    heading_index: &mut usize,
    text_meta_index: &mut usize,
    other_occurrences: &mut HashMap<(String, String), usize>,
    memory: &mut MetadataMemory<'_>,
) -> Result<Option<RdfaHost>> {
    if namespace == Some(TEXTNS) {
        return match local {
            "p" => {
                let index = *paragraph_index;
                *paragraph_index = paragraph_index.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::Paragraph { index }))
            },
            "h" => {
                let index = *heading_index;
                *heading_index = heading_index.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("heading occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::Heading { index }))
            },
            "bookmark-start" => Ok(Some(RdfaHost::BookmarkStart {
                name: attr_value(attrs, TEXTNS, "name", memory)?,
            })),
            "meta" => {
                let index = *text_meta_index;
                *text_meta_index = text_meta_index.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("text:meta occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::TextMeta { index }))
            },
            _ => {
                let key = (
                    memory.clone_string(TEXTNS, "ODF RDFa host namespace")?,
                    memory.clone_string(local, "ODF RDFa host local name")?,
                );
                reserve_hash_map_entry(
                    other_occurrences,
                    &key,
                    memory,
                    "ODF RDFa host occurrence map",
                )?;
                let occurrence = other_occurrences.entry(key).or_insert(0);
                let current = *occurrence;
                *occurrence = occurrence.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("RDFa host occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::Element {
                    namespace_uri: memory.clone_string(TEXTNS, "ODF RDFa host namespace")?,
                    local_name: memory.clone_string(local, "ODF RDFa host local name")?,
                    occurrence: current,
                }))
            },
        };
    }
    let Some(namespace) = namespace else {
        return Ok(None);
    };
    let key = (
        memory.clone_string(namespace, "ODF RDFa host namespace")?,
        memory.clone_string(local, "ODF RDFa host local name")?,
    );
    reserve_hash_map_entry(
        other_occurrences,
        &key,
        memory,
        "ODF RDFa host occurrence map",
    )?;
    let occurrence = other_occurrences.entry(key).or_insert(0);
    let current = *occurrence;
    *occurrence = occurrence
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("RDFa host occurrence count overflow".to_string()))?;
    Ok(Some(RdfaHost::Element {
        namespace_uri: memory.clone_string(namespace, "ODF RDFa host namespace")?,
        local_name: memory.clone_string(local, "ODF RDFa host local name")?,
        occurrence: current,
    }))
}

fn parse_attributes(
    reader: &ResolvedReader<'_>,
    source: &BytesStart<'_>,
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<MetaFieldAttribute>> {
    let mut attributes = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid in-content metadata attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if attributes.len() >= MAX_ATTRIBUTES {
            return Err(Error::InvalidFormat(format!(
                "in-content metadata element exceeds {MAX_ATTRIBUTES} attributes"
            )));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace_with_memory(&namespace, memory)?.unwrap_or_default();
        if metadata_attribute_decode_needs_owned(attribute.value.as_ref()) {
            memory.reserve_bytes(
                attribute.value.len(),
                "ODT metadata attribute decode scratch",
            )?;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid metadata attribute value: {error}"))
            })?;
        let local_name = utf8_with_memory(local.as_ref(), "metadata attribute name", memory)?;
        let value =
            memory.clone_string(value.as_ref(), "ODT metadata attribute value projection")?;
        memory.reserve_vec(&mut attributes, 1, "ODT metadata attribute projection")?;
        attributes.push(MetaFieldAttribute {
            namespace_uri,
            local_name,
            value,
        });
    }
    Ok(attributes)
}

fn parse_meta_attributes(
    reader: &ResolvedReader<'_>,
    source: &BytesStart<'_>,
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<MetaFieldAttribute>> {
    let output = parse_attributes(reader, source, memory)?;
    for attribute in &output {
        if attribute.namespace_uri.is_empty() {
            return Err(Error::InvalidFormat(
                "unqualified attributes inside text:meta are not supported".to_string(),
            ));
        }
    }
    if output.len() > 64 {
        return Err(Error::InvalidFormat(
            "text:meta child has too many attributes".to_string(),
        ));
    }
    // `parse_attributes` reserves exact capacity for each pushed item.  Do
    // not shrink here: retaining that capacity is already represented by the
    // ledger and avoids an untracked reallocation.
    Ok(output)
}

fn rdfa_from_attributes(
    attrs: &[MetaFieldAttribute],
    memory: &mut MetadataMemory<'_>,
) -> Result<RdfaAttributes> {
    let mut output = RdfaAttributes::default();
    for attr in attrs {
        if attr.namespace_uri == XHTMLNS {
            let slot = match attr.local_name.as_str() {
                "about" => &mut output.about,
                "property" => &mut output.property,
                "content" => &mut output.content,
                "datatype" => &mut output.datatype,
                _ => continue,
            };
            if slot.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate xhtml:{} RDFa attribute",
                    attr.local_name
                )));
            }
            *slot = Some(memory.clone_string(&attr.value, "ODF RDFa attribute projection")?);
        }
    }
    output.validate()?;
    Ok(output)
}

fn text_meta_root_attributes(
    reader: &ResolvedReader<'_>,
    source: &BytesStart<'_>,
    attrs: &[MetaFieldAttribute],
    memory: &mut MetadataMemory<'_>,
) -> Result<(Option<String>, RdfaAttributes, Vec<MetadataAttribute>)> {
    let mut xml_id = None;
    for attr in attrs {
        if attr.namespace_uri == XMLNS && attr.local_name == "id" {
            if xml_id.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate text:meta xml:id".to_string(),
                ));
            }
            validate_xml_id(&attr.value)?;
            xml_id = Some(memory.clone_string(&attr.value, "ODT text:meta xml:id projection")?);
        }
    }
    Ok((
        xml_id,
        rdfa_from_attributes(attrs, memory)?,
        unknown_text_meta_attributes(reader, source, memory)?,
    ))
}

fn unknown_text_meta_attributes(
    reader: &ResolvedReader<'_>,
    source: &BytesStart<'_>,
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<MetadataAttribute>> {
    let mut output = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid text:meta attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace_with_memory(&namespace, memory)?.unwrap_or_default();
        let local_name = utf8_with_memory(local.as_ref(), "text:meta attribute name", memory)?;
        if (namespace_uri == XMLNS && local_name == "id") || namespace_uri == XHTMLNS {
            continue;
        }
        if output.len() >= 64 {
            return Err(Error::InvalidFormat(
                "text:meta root has too many unknown attributes".to_string(),
            ));
        }
        if metadata_attribute_decode_needs_owned(attribute.value.as_ref()) {
            memory.reserve_bytes(
                attribute.value.len(),
                "ODT text:meta attribute decode scratch",
            )?;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid text:meta attribute value: {error}"))
            })?;
        let prefix = raw
            .split(|byte| *byte == b':')
            .next()
            .filter(|candidate| *candidate != raw)
            .map(|candidate| {
                std::str::from_utf8(candidate)
                    .map_err(|_| Error::InvalidFormat("invalid attribute prefix".to_string()))
                    .and_then(|prefix| {
                        memory.clone_string(prefix, "ODT text:meta attribute prefix")
                    })
            })
            .transpose()?;
        let value =
            memory.clone_string(value.as_ref(), "ODT text:meta attribute value projection")?;
        memory.reserve_vec(&mut output, 1, "ODT text:meta attribute projection")?;
        output.push(MetadataAttribute {
            namespace_uri,
            local_name,
            value,
            prefix,
        });
    }
    if output.len() > 64 {
        return Err(Error::InvalidFormat(
            "text:meta root has too many unknown attributes".to_string(),
        ));
    }
    Ok(output)
}

fn attr_value(
    attrs: &[MetaFieldAttribute],
    namespace: &str,
    local: &str,
    memory: &mut MetadataMemory<'_>,
) -> Result<Option<String>> {
    attrs
        .iter()
        .find(|attr| attr.namespace_uri == namespace && attr.local_name == local)
        .map(|attr| memory.clone_string(&attr.value, "ODF RDFa host attribute projection"))
        .transpose()
}

fn validate_xml_id(value: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat("xml:id must not be empty".to_string()));
    };
    let valid_start = first == '_' || first.is_alphabetic() || (first as u32) >= 0x80;
    let valid_rest = chars.all(|ch| {
        ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric() || (ch as u32) >= 0x80
    });
    if !valid_start || !valid_rest || value.contains(':') {
        return Err(Error::InvalidFormat(format!(
            "invalid XML NCName '{value}'"
        )));
    }
    validate_xml_value("xml:id", value, MAX_TEXT_META_BYTES)
}

fn validate_xml_value(name: &str, value: &str, max_bytes: usize) -> Result<()> {
    if value.len() > max_bytes || !value.chars().all(is_xml_1_0_char) {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds its bounded XML lexical value"
        )));
    }
    Ok(())
}

const fn is_xml_1_0_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || (value as u32 >= 0x20 && value as u32 <= 0xD7FF)
        || (value as u32 >= 0xE000 && value as u32 <= 0xFFFD)
        || (value as u32 >= 0x10000 && value as u32 <= 0x10FFFF)
}

fn push_rdfa_attributes(
    output: &mut String,
    attrs: &RdfaAttributes,
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    for (name, value) in [
        ("about", attrs.about.as_deref()),
        ("property", attrs.property.as_deref()),
        ("content", attrs.content.as_deref()),
        ("datatype", attrs.datatype.as_deref()),
    ] {
        if let Some(value) = value {
            output.push_str(" xhtml:");
            output.push_str(name);
            output.push_str("=\"");
            output.push_str(&escaped_metadata_value(
                value,
                memory,
                "ODT RDFa attribute escape scratch",
            )?);
            output.push('"');
        }
    }
    Ok(())
}

fn push_rdfa_namespace_declarations(
    output: &mut String,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    for prefix in rdfa_prefixes_with_memory(attrs, memory)? {
        if matches!(prefix.as_str(), "text" | "xhtml" | "xml") {
            continue;
        }
        if namespace_binding(namespace_declarations, &prefix).is_some() {
            continue;
        }
        let Some(namespace) = canonical_rdfa_namespace(&prefix) else {
            continue;
        };
        output.push_str(" xmlns:");
        output.push_str(&prefix);
        output.push_str("=\"");
        output.push_str(namespace);
        output.push('"');
    }
    Ok(())
}

fn append_required_rdfa_namespaces(
    output: &mut String,
    source: &str,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
    scratch: &mut MetadataMemory<'_>,
) -> Result<()> {
    for prefix in rdfa_prefixes_with_memory(attrs, scratch)? {
        if source_namespace_binding(source, &prefix, scratch)?.is_some()
            || namespace_binding(namespace_declarations, &prefix).is_some()
        {
            continue;
        }
        let Some(namespace) = canonical_rdfa_namespace(&prefix) else {
            return Err(Error::InvalidFormat(format!(
                "RDFa CURIE prefix '{prefix}' has no in-scope namespace binding"
            )));
        };
        output.push_str(" xmlns:");
        output.push_str(&prefix);
        output.push_str("=\"");
        output.push_str(namespace);
        output.push('"');
    }
    Ok(())
}

fn ensure_rdfa_namespace_context(
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
) -> Result<()> {
    fn check_value<'a>(
        value: &'a str,
        safe_curie_only: bool,
        namespace_declarations: &[String],
        checked: &mut [Option<&'a str>; MAX_ATTRIBUTES],
        checked_len: &mut usize,
        token_count: &mut usize,
    ) -> Result<()> {
        *token_count = token_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("RDFa CURIE token count overflow".to_string()))?;
        if *token_count > MAX_RDFA_CURIE_TOKENS {
            return Err(Error::InvalidFormat(format!(
                "RDFa CURIE token list exceeds {MAX_RDFA_CURIE_TOKENS} entries"
            )));
        }
        let value = if safe_curie_only {
            let Some(value) = value
                .strip_prefix('[')
                .and_then(|value| value.strip_suffix(']'))
            else {
                return Ok(());
            };
            value
        } else {
            value
                .strip_prefix('[')
                .and_then(|value| value.strip_suffix(']'))
                .unwrap_or(value)
        };
        let Some((prefix, _)) = value.split_once(':') else {
            return Ok(());
        };
        if prefix == "_" {
            return Ok(());
        }
        if !is_valid_metadata_prefix(prefix) {
            return Err(Error::InvalidFormat(format!(
                "invalid RDFa CURIE prefix '{prefix}'"
            )));
        }
        if checked[..*checked_len]
            .iter()
            .flatten()
            .any(|current| *current == prefix)
        {
            return Ok(());
        }
        if namespace_binding(namespace_declarations, prefix).is_none()
            && canonical_rdfa_namespace(prefix).is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "RDFa CURIE prefix '{prefix}' has no in-scope namespace binding"
            )));
        }
        if *checked_len < checked.len() {
            checked[*checked_len] = Some(prefix);
            *checked_len += 1;
        }
        Ok(())
    }

    let mut checked = [None; MAX_ATTRIBUTES];
    let mut checked_len = 0usize;
    let mut token_count = 0usize;
    if let Some(value) = &attrs.about {
        check_value(
            value,
            true,
            namespace_declarations,
            &mut checked,
            &mut checked_len,
            &mut token_count,
        )?;
    }
    if let Some(value) = &attrs.property {
        for token in value.split_whitespace() {
            check_value(
                token,
                false,
                namespace_declarations,
                &mut checked,
                &mut checked_len,
                &mut token_count,
            )?;
        }
    }
    if let Some(value) = &attrs.datatype {
        check_value(
            value,
            false,
            namespace_declarations,
            &mut checked,
            &mut checked_len,
            &mut token_count,
        )?;
    }
    Ok(())
}

fn namespace_binding<'a>(namespace_declarations: &'a [String], prefix: &str) -> Option<&'a str> {
    namespace_declarations
        .iter()
        .filter_map(|declaration| declaration_parts(declaration))
        .find(|(name, _)| name.strip_prefix("xmlns:") == Some(prefix))
        .map(|(_, value)| value)
}

fn canonical_rdfa_namespace(prefix: &str) -> Option<&'static str> {
    match prefix {
        "dc" => Some(DCNS),
        "xsd" => Some(XSDNS),
        "xhtml" => Some(XHTMLNS),
        "text" => Some(TEXTNS),
        _ => None,
    }
}

fn next_rdfa_token(token_count: &mut usize, memory: &mut MetadataMemory<'_>) -> Result<()> {
    *token_count = token_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("RDFa CURIE token count overflow".to_string()))?;
    if *token_count > MAX_RDFA_CURIE_TOKENS {
        return Err(Error::InvalidFormat(format!(
            "RDFa CURIE token list exceeds {MAX_RDFA_CURIE_TOKENS} entries"
        )));
    }
    memory.check()
}

fn rdfa_prefixes_with_memory(
    attrs: &RdfaAttributes,
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<String>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    let mut token_count = 0usize;
    if let Some(value) = &attrs.about {
        next_rdfa_token(&mut token_count, memory)?;
        add_rdfa_prefix_with_memory(&mut output, &mut seen, value, true, memory)?;
    }
    if let Some(value) = &attrs.property {
        for token in value.split_whitespace() {
            next_rdfa_token(&mut token_count, memory)?;
            add_rdfa_prefix_with_memory(&mut output, &mut seen, token, false, memory)?;
        }
    }
    if let Some(value) = &attrs.datatype {
        next_rdfa_token(&mut token_count, memory)?;
        add_rdfa_prefix_with_memory(&mut output, &mut seen, value, false, memory)?;
    }
    Ok(output)
}

fn add_rdfa_prefix_with_memory(
    output: &mut Vec<String>,
    seen: &mut HashSet<String>,
    value: &str,
    safe_curie_only: bool,
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    let value = if safe_curie_only {
        let Some(value) = value
            .strip_prefix('[')
            .and_then(|value| value.strip_suffix(']'))
        else {
            return Ok(());
        };
        value
    } else {
        value
            .strip_prefix('[')
            .and_then(|value| value.strip_suffix(']'))
            .unwrap_or(value)
    };
    let Some((prefix, _)) = value.split_once(':') else {
        return Ok(());
    };
    if prefix == "_" {
        return Ok(());
    }
    if !is_valid_metadata_prefix(prefix) {
        return Err(Error::InvalidFormat(format!(
            "invalid RDFa CURIE prefix '{prefix}'"
        )));
    }
    if seen.contains(prefix) {
        return Ok(());
    }
    reserve_hash_set_insert(seen, prefix, memory, "ODT RDFa CURIE prefix index")?;
    reserve_vec_push(
        output,
        memory.clone_string(prefix, "ODT RDFa CURIE prefix")?,
        "ODT RDFa CURIE prefixes",
        memory,
    )?;
    Ok(())
}

fn validate_metadata_attributes(attrs: &[MetadataAttribute]) -> Result<()> {
    if attrs.len() > 64 {
        return Err(Error::InvalidFormat(
            "text:meta root has too many unknown attributes".to_string(),
        ));
    }
    for attr in attrs {
        if attr.local_name.contains(':') || attr.local_name == "xmlns" {
            return Err(Error::InvalidFormat(
                "text:meta attribute local name is not serializable".to_string(),
            ));
        }
        validate_xml_name(&attr.local_name, "text:meta attribute name")?;
        validate_xml_value(
            "text:meta attribute namespace",
            &attr.namespace_uri,
            MAX_TEXT_META_BYTES,
        )?;
        validate_xml_value(
            "text:meta attribute name",
            &attr.local_name,
            MAX_TEXT_META_BYTES,
        )?;
        validate_xml_value(
            "text:meta attribute value",
            &attr.value,
            MAX_TEXT_META_BYTES,
        )?;
        if let Some(prefix) = &attr.prefix {
            if !is_valid_metadata_prefix(prefix) {
                return Err(Error::InvalidFormat(
                    "text:meta attribute prefix is not serializable".to_string(),
                ));
            }
        }
    }
    Ok(())
}

fn push_metadata_attributes(
    output: &mut String,
    attrs: &[MetadataAttribute],
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    let mut dynamic = Vec::new();
    memory.reserve_vec(&mut dynamic, 3, "ODT text:meta dynamic namespaces")?;
    for (prefix, namespace) in [("text", TEXTNS), ("xhtml", XHTMLNS), ("xml", XMLNS)] {
        reserve_vec_push(
            &mut dynamic,
            (
                memory.clone_string(prefix, "ODT text:meta namespace prefix")?,
                memory.clone_string(namespace, "ODT text:meta namespace URI")?,
            ),
            "ODT text:meta dynamic namespaces",
            memory,
        )?;
    }
    let mut generated = 0usize;
    for attr in attrs {
        let prefix = if attr.namespace_uri.is_empty() {
            String::new()
        } else if let Some(prefix) = canonical_metadata_prefix(&attr.namespace_uri) {
            memory.clone_string(prefix, "ODT text:meta namespace prefix")?
        } else if let Some(prefix) = attr.prefix.as_deref().filter(|prefix| {
            is_valid_metadata_prefix(prefix)
                && dynamic.iter().all(|(current, namespace)| {
                    current != *prefix || namespace == &attr.namespace_uri
                })
        }) {
            let prefix = memory.clone_string(prefix, "ODT text:meta namespace prefix")?;
            if !dynamic
                .iter()
                .any(|(current, namespace)| current == &prefix && namespace == &attr.namespace_uri)
            {
                output.push_str(" xmlns:");
                output.push_str(&prefix);
                output.push_str("=\"");
                output.push_str(&escaped_metadata_value(
                    &attr.namespace_uri,
                    memory,
                    "ODT text:meta namespace escape scratch",
                )?);
                output.push('"');
                reserve_vec_push(
                    &mut dynamic,
                    (
                        memory.clone_string(&prefix, "ODT text:meta namespace prefix")?,
                        memory.clone_string(&attr.namespace_uri, "ODT text:meta namespace URI")?,
                    ),
                    "ODT text:meta dynamic namespaces",
                    memory,
                )?;
            }
            prefix
        } else if let Some((prefix, _)) = dynamic
            .iter()
            .find(|(_, namespace)| namespace == &attr.namespace_uri)
        {
            memory.clone_string(prefix, "ODT text:meta namespace prefix")?
        } else {
            let prefix = loop {
                let candidate = generated_metadata_prefix(generated, memory)?;
                generated = generated.saturating_add(1);
                if dynamic.iter().all(|(current, _)| current != &candidate) {
                    break candidate;
                }
            };
            output.push_str(" xmlns:");
            output.push_str(&prefix);
            output.push_str("=\"");
            output.push_str(&escaped_metadata_value(
                &attr.namespace_uri,
                memory,
                "ODT text:meta namespace escape scratch",
            )?);
            output.push('"');
            reserve_vec_push(
                &mut dynamic,
                (
                    memory.clone_string(&prefix, "ODT text:meta namespace prefix")?,
                    memory.clone_string(&attr.namespace_uri, "ODT text:meta namespace URI")?,
                ),
                "ODT text:meta dynamic namespaces",
                memory,
            )?;
            prefix
        };
        output.push(' ');
        if !prefix.is_empty() {
            output.push_str(&prefix);
            output.push(':');
        }
        output.push_str(&attr.local_name);
        output.push_str("=\"");
        output.push_str(&escaped_metadata_value(
            &attr.value,
            memory,
            "ODT text:meta attribute escape scratch",
        )?);
        output.push('"');
    }
    Ok(())
}

fn generated_metadata_prefix(generated: usize, memory: &mut MetadataMemory<'_>) -> Result<String> {
    let mut digits = 1usize;
    let mut value = generated;
    while value >= 10 {
        value /= 10;
        digits = digits.saturating_add(1);
    }
    let length = 4usize
        .checked_add(digits)
        .ok_or_else(|| Error::InvalidFormat("metadata prefix size overflow".to_string()))?;
    let mut output = String::new();
    reserve_string_growth(
        &mut output,
        length,
        memory,
        "ODT text:meta generated namespace prefix",
    )?;
    write!(&mut output, "meta{generated}").map_err(|_| {
        Error::InvalidFormat("ODT text:meta generated namespace prefix overflow".to_string())
    })?;
    Ok(output)
}

fn escaped_metadata_value(
    value: &str,
    memory: &mut MetadataMemory<'_>,
    resource: &'static str,
) -> Result<String> {
    let length = escaped_xml_len(value, true)?;
    let mut output = String::new();
    reserve_string_growth(&mut output, length, memory, resource)?;
    let mut cursor = 0usize;
    for (index, byte) in value.bytes().enumerate() {
        let replacement = match byte {
            b'&' => Some("&amp;"),
            b'<' => Some("&lt;"),
            b'>' => Some("&gt;"),
            b'"' => Some("&quot;"),
            b'\'' => Some("&apos;"),
            _ => None,
        };
        if let Some(replacement) = replacement {
            output.push_str(&value[cursor..index]);
            output.push_str(replacement);
            cursor = index + 1;
        }
    }
    output.push_str(&value[cursor..]);
    Ok(output)
}

fn canonical_metadata_prefix(namespace: &str) -> Option<&'static str> {
    match namespace {
        TEXTNS => Some("text"),
        XHTMLNS => Some("xhtml"),
        XMLNS => Some("xml"),
        _ => None,
    }
}

fn validate_xml_name(value: &str, name: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    };
    if !(first == '_' || first.is_alphabetic())
        || !chars.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
    {
        return Err(Error::InvalidFormat(format!("invalid XML name '{value}'")));
    }
    validate_xml_value(name, value, MAX_TEXT_META_BYTES)
}

fn is_valid_metadata_prefix(prefix: &str) -> bool {
    let Some(first) = prefix.as_bytes().first().copied() else {
        return false;
    };
    (first.is_ascii_alphabetic() || first == b'_')
        && prefix != "xml"
        && prefix != "xmlns"
        && prefix
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.' | b'-'))
}

fn resolved_namespace_with_memory(
    namespace: &ResolveResult<'_>,
    memory: &mut MetadataMemory<'_>,
) -> Result<Option<String>> {
    let Some(namespace) =
        crate::elements::xml::normalized_namespace_uri(namespace, "ODT metadata")?
    else {
        return Ok(None);
    };
    let namespace = std::str::from_utf8(namespace)
        .map_err(|_| Error::InvalidFormat("ODT metadata namespace URI is not UTF-8".into()))?;
    memory
        .clone_string(namespace, "ODT metadata namespace URI")
        .map(Some)
}

fn utf8_with_memory(
    value: &[u8],
    description: &str,
    memory: &mut MetadataMemory<'_>,
) -> Result<String> {
    let value = std::str::from_utf8(value)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 {description}")))?;
    memory.clone_string(value, "ODT metadata UTF-8 projection")
}

fn push_bounded_with_memory<T>(
    items: &mut Vec<T>,
    value: T,
    resource: &'static str,
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    memory.reserve_vec(items, 1, resource)?;
    items.push(value);
    Ok(())
}

fn reserve_vec_push<T>(
    items: &mut Vec<T>,
    value: T,
    resource: &'static str,
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    memory.reserve_vec(items, 1, resource)?;
    items.push(value);
    Ok(())
}

fn hash_capacity_to_buckets(capacity: usize, element_size: usize) -> Result<usize> {
    if capacity == 0 {
        return Ok(0);
    }
    // hashbrown's RawTable uses sixteen control-byte groups on the supported
    // targets.  This mirrors its capacity_to_buckets/load-factor calculation,
    // so the precharge covers the same backing layout as try_reserve.
    if capacity < 15 {
        let minimum = match element_size {
            0..=1 => 14,
            2..=3 => 7,
            _ => 3,
        };
        let capacity = capacity.max(minimum);
        return Ok(if capacity < 4 {
            4
        } else if capacity < 8 {
            8
        } else {
            16
        });
    }
    let adjusted = capacity
        .checked_mul(8)
        .ok_or_else(|| Error::InvalidFormat("metadata hash table size overflow".to_string()))?
        / 7;
    adjusted
        .checked_next_power_of_two()
        .ok_or_else(|| Error::InvalidFormat("metadata hash table size overflow".to_string()))
}

fn hash_table_allocation_size<T>(capacity: usize) -> Result<usize> {
    let buckets = hash_capacity_to_buckets(capacity, size_of::<T>())?;
    if buckets == 0 {
        return Ok(0);
    }
    let control_align = 16usize.max(align_of::<T>());
    let data_bytes = buckets
        .checked_mul(size_of::<T>())
        .ok_or_else(|| Error::InvalidFormat("metadata hash table size overflow".to_string()))?;
    let control_offset = data_bytes
        .checked_add(control_align - 1)
        .ok_or_else(|| Error::InvalidFormat("metadata hash table size overflow".to_string()))?
        & !(control_align - 1);
    control_offset
        .checked_add(buckets)
        .and_then(|size| size.checked_add(16))
        .ok_or_else(|| Error::InvalidFormat("metadata hash table size overflow".to_string()))
}

struct HashTableGrowth {
    old_bytes: usize,
    planned_bytes: usize,
    retained: Option<Reservation>,
    scratch: Option<Reservation>,
}

fn reserve_hash_table_growth<T>(
    length: usize,
    capacity: usize,
    memory: &mut MetadataMemory<'_>,
    resource: &'static str,
) -> Result<HashTableGrowth> {
    memory.check()?;
    let old_bytes = hash_table_allocation_size::<T>(capacity)?;
    let required = length
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
    if required <= capacity {
        return Ok(HashTableGrowth {
            old_bytes,
            planned_bytes: old_bytes,
            retained: None,
            scratch: None,
        });
    }
    let planned_capacity = hash_capacity_to_buckets(required, size_of::<T>())?;
    // hash_capacity_to_buckets returns a bucket count; derive the effective
    // capacity of that table before converting it back to bytes.
    let planned_capacity = if planned_capacity < 8 {
        planned_capacity - 1
    } else {
        (planned_capacity / 8) * 7
    };
    let planned_bytes = hash_table_allocation_size::<T>(planned_capacity)?;
    let requested_delta = planned_bytes
        .checked_sub(old_bytes)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
    let retained = memory.reserve_growth_scratch(requested_delta, resource)?;
    let scratch = memory.reserve_growth_scratch(old_bytes, resource)?;
    Ok(HashTableGrowth {
        old_bytes,
        planned_bytes,
        retained,
        scratch,
    })
}

fn reconcile_hash_table_growth<T>(
    old_capacity: usize,
    actual_capacity: usize,
    mut growth: HashTableGrowth,
    memory: &mut MetadataMemory<'_>,
    resource: &'static str,
) -> Result<()> {
    let old_bytes = hash_table_allocation_size::<T>(old_capacity)?;
    debug_assert_eq!(old_bytes, growth.old_bytes);
    let actual_bytes = hash_table_allocation_size::<T>(actual_capacity)?;
    if actual_bytes < old_bytes {
        return Err(Error::InvalidFormat(format!(
            "{resource} hash table capacity regressed"
        )));
    }
    if actual_bytes < growth.planned_bytes {
        drop(growth.retained.take());
        let actual_delta = actual_bytes - old_bytes;
        growth.retained = memory.reserve_growth_scratch(actual_delta, resource)?;
    } else if actual_bytes > growth.planned_bytes {
        if let (Some(retained), Some(budget)) = (growth.retained.as_mut(), memory.budget) {
            let extra = budget.reserve_bytes(actual_bytes - growth.planned_bytes, resource)?;
            retained.try_merge(extra).map_err(|_reservation| {
                Error::InvalidFormat(format!("{resource} reservation chain changed"))
            })?;
        }
    }
    if let Some(retained) = growth.retained.take() {
        memory.merge_reservation(retained, resource)?;
    }
    drop(growth.scratch);
    Ok(())
}

fn reserve_hash_set_insert(
    set: &mut HashSet<String>,
    value: &str,
    memory: &mut MetadataMemory<'_>,
    resource: &'static str,
) -> Result<()> {
    memory.check()?;
    if set.contains(value) {
        return Ok(());
    }
    let value = memory.clone_string(value, resource)?;
    let old_capacity = set.capacity();
    let growth = reserve_hash_table_growth::<String>(set.len(), old_capacity, memory, resource)?;
    if let Err(source) = set.try_reserve(1) {
        drop(growth.retained);
        drop(growth.scratch);
        return Err(Error::Allocation { resource, source });
    }
    reconcile_hash_table_growth::<String>(old_capacity, set.capacity(), growth, memory, resource)?;
    set.insert(value);
    Ok(())
}

fn reserve_hash_map_entry<K, V>(
    map: &mut HashMap<K, V>,
    key: &K,
    memory: &mut MetadataMemory<'_>,
    resource: &'static str,
) -> Result<()>
where
    K: std::hash::Hash + Eq,
{
    memory.check()?;
    if map.contains_key(key) {
        return Ok(());
    }
    let old_capacity = map.capacity();
    let growth = reserve_hash_table_growth::<(K, V)>(map.len(), old_capacity, memory, resource)?;
    if let Err(source) = map.try_reserve(1) {
        drop(growth.retained);
        drop(growth.scratch);
        return Err(Error::Allocation { resource, source });
    }
    reconcile_hash_table_growth::<(K, V)>(old_capacity, map.capacity(), growth, memory, resource)
}

fn metadata_attribute_decode_needs_owned(value: &[u8]) -> bool {
    value
        .iter()
        .any(|byte| matches!(byte, b'&' | b'\t' | b'\r' | b'\n'))
}

fn metadata_text_decode_needs_owned(value: &[u8]) -> bool {
    value.iter().any(|byte| matches!(byte, b'\r'))
}

#[derive(Clone, Copy)]
struct NamespaceFrame {
    previous_buffer_len: usize,
    previous_binding_count: usize,
}

fn namespace_tag_end(
    bytes: &[u8],
    mut index: usize,
    memory: &MetadataMemory<'_>,
) -> Result<Option<usize>> {
    let mut quote = None;
    while index < bytes.len() {
        if index & 0x0fff == 0 {
            memory.check()?;
        }
        let byte = bytes[index];
        if let Some(current) = quote {
            if byte == current {
                quote = None;
            }
        } else if matches!(byte, b'"' | b'\'') {
            quote = Some(byte);
        } else if byte == b'>' {
            return Ok(Some(index));
        }
        index += 1;
    }
    Ok(None)
}

fn namespace_find_delimiter(
    bytes: &[u8],
    mut index: usize,
    delimiter: &[u8],
    memory: &MetadataMemory<'_>,
) -> Result<Option<usize>> {
    while index
        .checked_add(delimiter.len())
        .is_some_and(|end| end <= bytes.len())
    {
        if index & 0x0fff == 0 {
            memory.check()?;
        }
        if bytes[index..].starts_with(delimiter) {
            return Ok(Some(index));
        }
        index += 1;
    }
    Ok(None)
}

fn namespace_buffer_extend(
    length: &mut usize,
    capacity: &mut usize,
    additional: usize,
) -> Result<()> {
    if additional == 0 {
        return Ok(());
    }
    let required = length
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("metadata namespace buffer overflow".to_string()))?;
    if required > *capacity {
        let doubled = capacity.checked_mul(2).unwrap_or(usize::MAX);
        let minimum = if *capacity == 0 { 8 } else { 0 };
        *capacity = required.max(doubled).max(minimum);
    }
    *length = required;
    Ok(())
}

fn namespace_tag_declarations(
    tag: &[u8],
    buffer_length: &mut usize,
    buffer_capacity: &mut usize,
    binding_length: &mut usize,
    binding_capacity: &mut usize,
    memory: &MetadataMemory<'_>,
) -> Result<usize> {
    let mut index = 1usize;
    while index < tag.len()
        && (tag[index].is_ascii_whitespace() || matches!(tag[index], b'/' | b'!'))
    {
        if index & 0x0fff == 0 {
            memory.check()?;
        }
        index += 1;
    }
    while index < tag.len()
        && !tag[index].is_ascii_whitespace()
        && !matches!(tag[index], b'/' | b'>')
    {
        if index & 0x0fff == 0 {
            memory.check()?;
        }
        index += 1;
    }
    let mut declaration_count = 0usize;
    while index < tag.len() {
        if index & 0x0fff == 0 {
            memory.check()?;
        }
        while index < tag.len()
            && (tag[index].is_ascii_whitespace() || matches!(tag[index], b'/' | b'>'))
        {
            index += 1;
        }
        if index >= tag.len() {
            break;
        }
        let name_start = index;
        while index < tag.len()
            && !tag[index].is_ascii_whitespace()
            && !matches!(tag[index], b'=' | b'/' | b'>')
        {
            index += 1;
        }
        let name = &tag[name_start..index];
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= tag.len() || tag[index] != b'=' {
            while index < tag.len() && !matches!(tag[index], b'/' | b'>') {
                index += 1;
            }
            continue;
        }
        index += 1;
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        let Some(quote) = tag.get(index).copied() else {
            break;
        };
        if !matches!(quote, b'"' | b'\'') {
            while index < tag.len() && !matches!(tag[index], b'/' | b'>') {
                index += 1;
            }
            continue;
        }
        index += 1;
        let value_start = index;
        while index < tag.len() && tag[index] != quote {
            index += 1;
        }
        let value_len = index.saturating_sub(value_start);
        if name == b"xmlns" {
            declaration_count = declaration_count.saturating_add(1);
            namespace_buffer_extend(buffer_length, buffer_capacity, value_len)?;
            namespace_buffer_extend_binding(binding_length, binding_capacity)?;
        } else if let Some(prefix) = name.strip_prefix(b"xmlns:") {
            if prefix != b"xml" && prefix != b"xmlns" {
                declaration_count = declaration_count.saturating_add(1);
                namespace_buffer_extend(buffer_length, buffer_capacity, prefix.len())?;
                namespace_buffer_extend(buffer_length, buffer_capacity, value_len)?;
                namespace_buffer_extend_binding(binding_length, binding_capacity)?;
            }
        }
        index = index.saturating_add(1);
    }
    Ok(declaration_count)
}

fn namespace_buffer_extend_binding(length: &mut usize, capacity: &mut usize) -> Result<()> {
    let required = length
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("metadata namespace bindings overflow".to_string()))?;
    if required > *capacity {
        let doubled = capacity.checked_mul(2).unwrap_or(usize::MAX);
        let minimum = if *capacity == 0 { 4 } else { 0 };
        *capacity = required.max(doubled).max(minimum);
    }
    *length = required;
    Ok(())
}

fn reserve_namespace_resolver_memory(xml: &str, memory: &mut MetadataMemory<'_>) -> Result<()> {
    let mut frames: Vec<NamespaceFrame> = Vec::new();
    let mut buffer_length = 0usize;
    let mut buffer_capacity = 0usize;
    namespace_buffer_extend(&mut buffer_length, &mut buffer_capacity, b"xml".len())?;
    namespace_buffer_extend(
        &mut buffer_length,
        &mut buffer_capacity,
        b"http://www.w3.org/XML/1998/namespace".len(),
    )?;
    namespace_buffer_extend(&mut buffer_length, &mut buffer_capacity, b"xmlns".len())?;
    namespace_buffer_extend(
        &mut buffer_length,
        &mut buffer_capacity,
        b"http://www.w3.org/2000/xmlns/".len(),
    )?;
    let mut current_binding_count = 0usize;
    let mut binding_capacity = 0usize;
    namespace_buffer_extend_binding(&mut current_binding_count, &mut binding_capacity)?;
    namespace_buffer_extend_binding(&mut current_binding_count, &mut binding_capacity)?;
    let reserved_namespace_bytes = buffer_length;
    let mut current_buffer_len = reserved_namespace_bytes;
    let mut index = 0usize;
    let bytes = xml.as_bytes();
    while index < bytes.len() {
        if index & 0x0fff == 0 {
            memory.check()?;
        }
        if bytes[index] != b'<' {
            index += 1;
            continue;
        }
        if bytes[index..].starts_with(b"<!--") {
            index = namespace_find_delimiter(bytes, index + 4, b"-->", memory)?
                .map_or(bytes.len(), |offset| offset + 3);
            continue;
        }
        if bytes[index..].starts_with(b"<![CDATA[") {
            index = namespace_find_delimiter(bytes, index + 9, b"]]>", memory)?
                .map_or(bytes.len(), |offset| offset + 3);
            continue;
        }
        if bytes[index..].starts_with(b"<?") {
            index = namespace_tag_end(bytes, index + 2, memory)?
                .map_or(bytes.len(), |offset| offset + 1);
            continue;
        }
        if bytes[index..].starts_with(b"</") {
            if let Some(frame) = frames.pop() {
                current_buffer_len = frame.previous_buffer_len;
                current_binding_count = frame.previous_binding_count;
            }
            index = namespace_tag_end(bytes, index + 2, memory)?
                .map_or(bytes.len(), |offset| offset + 1);
            continue;
        }
        let Some(end) = namespace_tag_end(bytes, index + 1, memory)? else {
            break;
        };
        let tag = &bytes[index..=end];
        let previous_buffer_len = current_buffer_len;
        let previous_binding_count = current_binding_count;
        let _declaration_count = namespace_tag_declarations(
            tag,
            &mut current_buffer_len,
            &mut buffer_capacity,
            &mut current_binding_count,
            &mut binding_capacity,
            memory,
        )?;
        let empty = tag[..tag.len().saturating_sub(1)]
            .iter()
            .rev()
            .find(|byte| !byte.is_ascii_whitespace())
            == Some(&b'/');
        if empty {
            current_buffer_len = previous_buffer_len;
            current_binding_count = previous_binding_count;
        } else {
            memory.reserve_vec(&mut frames, 1, "ODT metadata namespace planning stack")?;
            frames.push(NamespaceFrame {
                previous_buffer_len,
                previous_binding_count,
            });
        }
        index = end + 1;
    }
    let binding_bytes = binding_capacity
        .checked_mul(size_of::<(usize, usize, usize, usize)>())
        .ok_or_else(|| Error::InvalidFormat("metadata namespace state overflow".to_string()))?;
    memory.reserve_bytes(buffer_capacity, "ODT metadata namespace resolver buffer")?;
    memory.reserve_bytes(binding_bytes, "ODT metadata namespace resolver bindings")?;
    Ok(())
}

fn reserve_string_growth(
    value: &mut String,
    additional: usize,
    memory: &mut MetadataMemory<'_>,
    resource: &'static str,
) -> Result<()> {
    if additional == 0
        || value
            .len()
            .checked_add(additional)
            .is_some_and(|next| next <= value.capacity())
    {
        return Ok(());
    }
    let required = value
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
    let old_capacity = value.capacity();
    let requested_delta = required
        .checked_sub(old_capacity)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?;
    let mut retained = memory.reserve_growth_scratch(requested_delta, resource)?;
    let scratch = memory.reserve_growth_scratch(old_capacity, resource)?;
    value
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })?;
    let actual_capacity = value.capacity();
    if actual_capacity > required {
        if let (Some(retained), Some(budget)) = (retained.as_mut(), memory.budget) {
            let extra = budget.reserve_bytes(actual_capacity - required, resource)?;
            retained.try_merge(extra).map_err(|_reservation| {
                Error::InvalidFormat(format!("{resource} reservation chain changed"))
            })?;
        }
    } else if actual_capacity < required {
        // try_reserve_exact must satisfy the requested length, but keep the
        // branch checked so a future allocator implementation cannot silently
        // turn a precharge into an undercharge.
        drop(retained.take());
        retained = memory.reserve_growth_scratch(
            actual_capacity
                .checked_sub(old_capacity)
                .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflow")))?,
            resource,
        )?;
    }
    if let Some(retained) = retained {
        memory.merge_reservation(retained, resource)?;
    }
    drop(scratch);
    Ok(())
}

fn clone_rdfa_attributes(
    value: &RdfaAttributes,
    memory: &mut MetadataMemory<'_>,
) -> Result<RdfaAttributes> {
    Ok(RdfaAttributes {
        about: memory.clone_option_string(value.about.as_deref(), "ODF RDFa about projection")?,
        property: memory
            .clone_option_string(value.property.as_deref(), "ODF RDFa property projection")?,
        content: memory
            .clone_option_string(value.content.as_deref(), "ODF RDFa content projection")?,
        datatype: memory
            .clone_option_string(value.datatype.as_deref(), "ODF RDFa datatype projection")?,
    })
}

fn clone_rdfa_host(value: &RdfaHost, memory: &mut MetadataMemory<'_>) -> Result<RdfaHost> {
    Ok(match value {
        RdfaHost::Paragraph { index } => RdfaHost::Paragraph { index: *index },
        RdfaHost::Heading { index } => RdfaHost::Heading { index: *index },
        RdfaHost::BookmarkStart { name } => RdfaHost::BookmarkStart {
            name: memory.clone_option_string(name.as_deref(), "ODF RDFa bookmark name")?,
        },
        RdfaHost::TextMeta { index } => RdfaHost::TextMeta { index: *index },
        RdfaHost::Element {
            namespace_uri,
            local_name,
            occurrence,
        } => RdfaHost::Element {
            namespace_uri: memory.clone_string(namespace_uri, "ODF RDFa host namespace")?,
            local_name: memory.clone_string(local_name, "ODF RDFa host local name")?,
            occurrence: *occurrence,
        },
    })
}

fn clone_meta_field_attributes(
    value: &[MetaFieldAttribute],
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<MetaFieldAttribute>> {
    let mut output = Vec::new();
    if !value.is_empty() {
        memory.reserve_vec(
            &mut output,
            value.len(),
            "ODT text:meta attribute projection",
        )?;
    }
    for attribute in value {
        output.push(MetaFieldAttribute {
            namespace_uri: memory.clone_string(
                &attribute.namespace_uri,
                "ODT text:meta attribute namespace",
            )?,
            local_name: memory
                .clone_string(&attribute.local_name, "ODT text:meta attribute name")?,
            value: memory.clone_string(&attribute.value, "ODT text:meta attribute value")?,
        });
    }
    Ok(output)
}

// -------------------------------------------------------------------------
// Lossless content.xml editing helpers

#[derive(Debug, Clone)]
struct Span {
    start: usize,
    start_end: usize,
    end_start: usize,
    end: usize,
    namespace: String,
    local: String,
    rdfa_names: Vec<String>,
    rdfa: RdfaAttributes,
    name: Option<String>,
    namespace_declarations: Vec<String>,
}

struct ScannedSpans<'a> {
    spans: Vec<Span>,
    _memory: MetadataMemory<'a>,
}

/// Set RDFa on a paragraph selected in document order.
pub(crate) fn set_paragraph_rdfa(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
) -> Result<String> {
    set_paragraph_rdfa_with_limit(xml, position, value, MAX_XML_BYTES)
}

/// Set RDFa on a paragraph while charging the exact edited output.
pub(crate) fn set_paragraph_rdfa_with_limit(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
    maximum: usize,
) -> Result<String> {
    set_paragraph_rdfa_with_limit_and_budget(xml, position, value, maximum, None)
        .map(ChargedXml::into_string)
}

pub(crate) fn set_paragraph_rdfa_with_limit_and_budget(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    value.validate()?;
    let scanned = scan_spans_with_budget(
        xml,
        |namespace, local| namespace == TEXTNS && local == "p",
        budget,
    )?;
    let span = scanned.spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "paragraph position {} is out of range",
            position.get()
        ))
    })?;
    if span.rdfa == *value {
        bounded_output_len_with_limit(xml.len(), "RDFa exact no-op", maximum)?;
        let (mut output, memory) = allocate_xml(budget, xml.len(), "ODT RDFa exact no-op")?;
        output.push_str(xml);
        return Ok(ChargedXml {
            xml: output,
            memory,
        });
    }
    rewrite_rdfa_with_limit_and_budget(xml, span, value, maximum, budget)
}

/// Set RDFa on a uniquely named bookmark start.
pub(crate) fn set_bookmark_rdfa(xml: &str, name: &str, value: &RdfaAttributes) -> Result<String> {
    set_bookmark_rdfa_with_limit(xml, name, value, MAX_XML_BYTES)
}

/// Set RDFa on a uniquely named bookmark start while charging the exact output.
pub(crate) fn set_bookmark_rdfa_with_limit(
    xml: &str,
    name: &str,
    value: &RdfaAttributes,
    maximum: usize,
) -> Result<String> {
    set_bookmark_rdfa_with_limit_and_budget(xml, name, value, maximum, None)
        .map(ChargedXml::into_string)
}

pub(crate) fn set_bookmark_rdfa_with_limit_and_budget(
    xml: &str,
    name: &str,
    value: &RdfaAttributes,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    value.validate()?;
    let scanned = scan_spans_with_budget(
        xml,
        |namespace, local| namespace == TEXTNS && local == "bookmark-start",
        budget,
    )?;
    let mut matches = scanned
        .spans
        .iter()
        .filter(|span| span.name.as_deref() == Some(name));
    let span = matches
        .next()
        .ok_or_else(|| Error::InvalidFormat(format!("bookmark start '{name}' was not found")))?;
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "bookmark start '{name}' is not unique"
        )));
    }
    if span.rdfa == *value {
        bounded_output_len_with_limit(xml.len(), "RDFa exact no-op", maximum)?;
        let (mut output, memory) =
            allocate_xml(budget, xml.len(), "ODT bookmark RDFa exact no-op")?;
        output.push_str(xml);
        return Ok(ChargedXml {
            xml: output,
            memory,
        });
    }
    rewrite_rdfa_with_limit_and_budget(xml, span, value, maximum, budget)
}

/// Set RDFa on one inline `text:meta` occurrence.
pub(crate) fn set_text_meta_rdfa(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
) -> Result<String> {
    set_text_meta_rdfa_with_limit(xml, position, value, MAX_XML_BYTES)
}

/// Set RDFa on one inline `text:meta` while charging the exact output.
pub(crate) fn set_text_meta_rdfa_with_limit(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
    maximum: usize,
) -> Result<String> {
    set_text_meta_rdfa_with_limit_and_budget(xml, position, value, maximum, None)
        .map(ChargedXml::into_string)
}

pub(crate) fn set_text_meta_rdfa_with_limit_and_budget(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    value.validate()?;
    let scanned = scan_spans_with_budget(
        xml,
        |namespace, local| namespace == TEXTNS && local == "meta",
        budget,
    )?;
    let span = scanned.spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:meta position {} is out of range",
            position.get()
        ))
    })?;
    if span.rdfa == *value {
        bounded_output_len_with_limit(xml.len(), "RDFa exact no-op", maximum)?;
        let (mut output, memory) =
            allocate_xml(budget, xml.len(), "ODT text:meta RDFa exact no-op")?;
        output.push_str(xml);
        return Ok(ChargedXml {
            xml: output,
            memory,
        });
    }
    rewrite_rdfa_with_limit_and_budget(xml, span, value, maximum, budget)
}

/// Insert an inline `text:meta` before the selected paragraph's end tag.
pub(crate) fn insert_text_meta(xml: &str, paragraph: Position, value: &TextMeta) -> Result<String> {
    insert_text_meta_with_limit(xml, paragraph, value, MAX_XML_BYTES)
}

/// Insert an inline `text:meta` while charging every constructed fragment and output.
pub(crate) fn insert_text_meta_with_limit(
    xml: &str,
    paragraph: Position,
    value: &TextMeta,
    maximum: usize,
) -> Result<String> {
    insert_text_meta_with_limit_and_budget(xml, paragraph, value, maximum, None)
        .map(ChargedXml::into_string)
}

pub(crate) fn insert_text_meta_with_limit_and_budget(
    xml: &str,
    paragraph: Position,
    value: &TextMeta,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    let scanned = scan_spans_with_budget(
        xml,
        |namespace, local| namespace == TEXTNS && local == "p",
        budget,
    )?;
    let span = scanned.spans.get(paragraph.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "paragraph position {} is out of range",
            paragraph.get()
        ))
    })?;
    validate_text_meta_with_budget(value, budget)?;
    ensure_rdfa_namespace_context(&value.rdfa, &span.namespace_declarations)?;
    let serialized = value.to_xml_with_context_with_limit_and_budget(
        &span.namespace_declarations,
        maximum,
        budget,
    )?;
    let fragment = inject_namespace_declarations_with_limit_and_budget(
        &serialized.xml,
        &span.namespace_declarations,
        maximum,
        budget,
    )?;
    drop(serialized);
    if span.end_start == span.start_end {
        let source = xml
            .get(span.start..span.start_end)
            .ok_or_else(|| Error::InvalidFormat("invalid empty paragraph span".to_string()))?;
        if !source.trim_end().ends_with("/>") {
            return Err(Error::InvalidFormat(
                "invalid empty paragraph element".to_string(),
            ));
        }
        let close = source.len().checked_sub(2).ok_or_else(|| {
            Error::InvalidFormat("invalid empty paragraph closing delimiter".to_string())
        })?;
        let qname_end = source[1..]
            .find(|ch: char| ch.is_ascii_whitespace() || ch == '/')
            .map_or(close - 1, |index| index + 1);
        let qname_len = source[1..qname_end].len();
        let replacement_len = source
            .len()
            .checked_add(fragment.xml.len())
            .and_then(|length| length.checked_add(qname_len))
            .and_then(|length| length.checked_add(2))
            .ok_or_else(|| Error::InvalidFormat("text:meta insertion size overflow".to_string()))?;
        bounded_output_len_with_limit(replacement_len, "text:meta insertion", maximum)?;
        let (mut replacement, replacement_memory) =
            allocate_xml(budget, replacement_len, "ODT text:meta insertion")?;
        replacement.push_str(&source[..close]);
        replacement.push('>');
        replacement.push_str(&fragment.xml);
        replacement.push_str("</");
        replacement.push_str(&source[1..qname_end]);
        replacement.push('>');
        let candidate = splice_replace_with_limit_and_budget(
            xml,
            span.start,
            span.end,
            &replacement,
            maximum,
            budget,
        );
        drop(replacement_memory);
        return candidate;
    }
    splice_insert_with_limit_and_budget(xml, span.end_start, &fragment.xml, maximum, budget)
}

/// Replace one inline `text:meta` element.
pub(crate) fn replace_text_meta(xml: &str, position: Position, value: &TextMeta) -> Result<String> {
    replace_text_meta_with_limit(xml, position, value, MAX_XML_BYTES)
}

/// Replace one inline `text:meta` while charging every constructed fragment and output.
pub(crate) fn replace_text_meta_with_limit(
    xml: &str,
    position: Position,
    value: &TextMeta,
    maximum: usize,
) -> Result<String> {
    replace_text_meta_with_limit_and_budget(xml, position, value, maximum, None)
        .map(ChargedXml::into_string)
}

pub(crate) fn replace_text_meta_with_limit_and_budget(
    xml: &str,
    position: Position,
    value: &TextMeta,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    let scanned = scan_spans_with_budget(
        xml,
        |namespace, local| namespace == TEXTNS && local == "meta",
        budget,
    )?;
    let span = scanned.spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:meta position {} is out of range",
            position.get()
        ))
    })?;
    validate_text_meta_with_budget(value, budget)?;
    ensure_rdfa_namespace_context(&value.rdfa, &span.namespace_declarations)?;
    let existing = parse_text_meta_span(xml, span, maximum, budget)?;
    if text_meta_semantically_equal(&existing, value) {
        bounded_output_len_with_limit(xml.len(), "text:meta exact no-op", maximum)?;
        let (mut output, memory) = allocate_xml(budget, xml.len(), "ODT text:meta exact no-op")?;
        output.push_str(xml);
        return Ok(ChargedXml {
            xml: output,
            memory,
        });
    }
    let serialized = value.to_xml_with_context_with_limit_and_budget(
        &span.namespace_declarations,
        maximum,
        budget,
    )?;
    let fragment = inject_namespace_declarations_with_limit_and_budget(
        &serialized.xml,
        &span.namespace_declarations,
        maximum,
        budget,
    )?;
    drop(serialized);
    splice_replace_with_limit_and_budget(xml, span.start, span.end, &fragment.xml, maximum, budget)
}

fn parse_text_meta_span(
    xml: &str,
    span: &Span,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<TextMeta> {
    let raw = xml
        .get(span.start..span.end)
        .ok_or_else(|| Error::InvalidFormat("invalid text:meta span".to_string()))?;
    let owned = inject_namespace_declarations_with_limit_and_budget(
        raw,
        &span.namespace_declarations,
        maximum,
        budget,
    )?;
    let (parsed, _parsed_memory) =
        parse_part_with_budget(&owned.xml, MetadataPart::Content, budget)?;
    parsed
        .text_meta
        .into_iter()
        .next()
        .ok_or_else(|| Error::InvalidFormat("text:meta span did not parse as metadata".to_string()))
}

fn text_meta_semantically_equal(left: &TextMeta, right: &TextMeta) -> bool {
    left.xml_id == right.xml_id
        && left.rdfa == right.rdfa
        && left.attributes == right.attributes
        && left.content == right.content
}

/// Remove one inline `text:meta` element.
pub(crate) fn remove_text_meta(xml: &str, position: Position) -> Result<String> {
    remove_text_meta_with_limit(xml, position, MAX_XML_BYTES)
}

/// Remove one inline `text:meta` while charging the exact shortened output.
pub(crate) fn remove_text_meta_with_limit(
    xml: &str,
    position: Position,
    maximum: usize,
) -> Result<String> {
    remove_text_meta_with_limit_and_budget(xml, position, maximum, None)
        .map(ChargedXml::into_string)
}

pub(crate) fn remove_text_meta_with_limit_and_budget(
    xml: &str,
    position: Position,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    let scanned = scan_spans_with_budget(
        xml,
        |namespace, local| namespace == TEXTNS && local == "meta",
        budget,
    )?;
    let span = scanned.spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:meta position {} is out of range",
            position.get()
        ))
    })?;
    splice_replace_with_limit_and_budget(xml, span.start, span.end, "", maximum, budget)
}

fn rewrite_rdfa_with_limit_and_budget(
    xml: &str,
    span: &Span,
    value: &RdfaAttributes,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    value.validate()?;
    let source = xml
        .get(span.start..span.start_end)
        .ok_or_else(|| Error::InvalidFormat("invalid RDFa start-tag span".to_string()))?;
    let mut scratch = MetadataMemory::new(budget);
    let mut remove = HashSet::new();
    for name in &span.rdfa_names {
        reserve_hash_set_insert(
            &mut remove,
            name,
            &mut scratch,
            "ODT RDFa removed attribute names",
        )?;
    }
    let add = !value.is_empty();
    let rewritten = rewrite_start_tag_with_limit_and_budget(
        source,
        &remove,
        add,
        value,
        &span.namespace_declarations,
        maximum,
        budget,
        &mut scratch,
    )?;
    splice_replace_with_limit_and_budget(
        xml,
        span.start,
        span.start_end,
        &rewritten.xml,
        maximum,
        budget,
    )
}

fn rewrite_start_tag_with_limit_and_budget(
    source: &str,
    remove_names: &HashSet<String>,
    add_namespace: bool,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
    scratch: &mut MetadataMemory<'_>,
) -> Result<ChargedXml> {
    let bytes = source.as_bytes();
    let close = if source.ends_with("/>") {
        source.len() - 2
    } else {
        source.len() - 1
    };
    let output_len = rewritten_start_tag_len(
        source,
        remove_names,
        add_namespace,
        attrs,
        namespace_declarations,
        scratch,
    )?;
    bounded_output_len_with_limit(output_len, "RDFa start-tag rewrite", maximum)?;
    let (mut output, memory) = allocate_xml(budget, output_len, "ODT RDFa start-tag rewrite")?;
    let mut cursor;
    let mut name_end = 1usize;
    while name_end < close && !bytes[name_end].is_ascii_whitespace() {
        name_end += 1;
    }
    output.push_str(&source[..name_end]);
    cursor = name_end;
    while cursor < close {
        let token_start = cursor;
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= close {
            output.push_str(&source[token_start..close]);
            break;
        }
        let attr_start = cursor;
        while cursor < close
            && !bytes[cursor].is_ascii_whitespace()
            && bytes[cursor] != b'='
            && bytes[cursor] != b'/'
        {
            cursor += 1;
        }
        let attr_name = &source[attr_start..cursor];
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor < close && bytes[cursor] == b'=' {
            cursor += 1;
            while cursor < close && bytes[cursor].is_ascii_whitespace() {
                cursor += 1;
            }
            if cursor < close && (bytes[cursor] == b'"' || bytes[cursor] == b'\'') {
                let quote = bytes[cursor];
                cursor += 1;
                while cursor < close && bytes[cursor] != quote {
                    cursor += 1;
                }
                if cursor < close {
                    cursor += 1;
                }
            }
        }
        if !remove_names.contains(attr_name) {
            output.push_str(&source[token_start..cursor]);
        }
    }
    if add_namespace {
        match source_namespace_binding(source, "xhtml", scratch)? {
            Some(namespace) if namespace != XHTMLNS => {
                return Err(Error::InvalidFormat(
                    "RDFa edit cannot shadow the xhtml namespace prefix".to_string(),
                ));
            },
            Some(_) => {},
            None => {
                output.push_str(" xmlns:xhtml=\"");
                output.push_str(XHTMLNS);
                output.push('"');
            },
        }
        append_required_rdfa_namespaces(
            &mut output,
            source,
            attrs,
            namespace_declarations,
            scratch,
        )?;
    }
    push_rdfa_attributes(&mut output, attrs, scratch)?;
    output.push_str(&source[close..]);
    Ok(ChargedXml {
        xml: output,
        memory,
    })
}

fn rewritten_start_tag_len(
    source: &str,
    remove_names: &HashSet<String>,
    add_namespace: bool,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
    scratch: &mut MetadataMemory<'_>,
) -> Result<usize> {
    let bytes = source.as_bytes();
    let close = if source.ends_with("/>") {
        source.len() - 2
    } else {
        source.len() - 1
    };
    let mut length = 0usize;
    let mut cursor;
    let mut name_end = 1usize;
    while name_end < close && !bytes[name_end].is_ascii_whitespace() {
        name_end += 1;
    }
    add_size(&mut length, name_end)?;
    cursor = name_end;
    while cursor < close {
        let token_start = cursor;
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= close {
            add_size(&mut length, close - token_start)?;
            break;
        }
        let attr_start = cursor;
        while cursor < close
            && !bytes[cursor].is_ascii_whitespace()
            && bytes[cursor] != b'='
            && bytes[cursor] != b'/'
        {
            cursor += 1;
        }
        let attr_name = &source[attr_start..cursor];
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor < close && bytes[cursor] == b'=' {
            cursor += 1;
            while cursor < close && bytes[cursor].is_ascii_whitespace() {
                cursor += 1;
            }
            if cursor < close && (bytes[cursor] == b'"' || bytes[cursor] == b'\'') {
                let quote = bytes[cursor];
                cursor += 1;
                while cursor < close && bytes[cursor] != quote {
                    cursor += 1;
                }
                if cursor < close {
                    cursor += 1;
                }
            }
        }
        if !remove_names.contains(attr_name) {
            add_size(&mut length, cursor - token_start)?;
        }
    }
    if add_namespace {
        match source_namespace_binding(source, "xhtml", scratch)? {
            Some(namespace) if namespace != XHTMLNS => {
                return Err(Error::InvalidFormat(
                    "RDFa edit cannot shadow the xhtml namespace prefix".to_string(),
                ));
            },
            Some(_) => {},
            None => {
                add_size(&mut length, " xmlns:xhtml=\"\"".len() + XHTMLNS.len())?;
            },
        }
        for prefix in rdfa_prefixes_with_memory(attrs, scratch)? {
            if source_namespace_binding(source, &prefix, scratch)?.is_some()
                || namespace_binding(namespace_declarations, &prefix).is_some()
            {
                continue;
            }
            let Some(namespace) = canonical_rdfa_namespace(&prefix) else {
                return Err(Error::InvalidFormat(format!(
                    "RDFa CURIE prefix '{prefix}' has no in-scope namespace binding"
                )));
            };
            add_size(
                &mut length,
                " xmlns:=\"\"".len() + prefix.len() + namespace.len(),
            )?;
        }
        for (name, value) in [
            ("about", attrs.about.as_deref()),
            ("property", attrs.property.as_deref()),
            ("content", attrs.content.as_deref()),
            ("datatype", attrs.datatype.as_deref()),
        ] {
            if let Some(value) = value {
                let value_len = escaped_xml_len(value, true)?;
                let attribute_len = " xhtml:=\"\""
                    .len()
                    .checked_add(name.len())
                    .and_then(|size| size.checked_add(value_len))
                    .ok_or_else(|| {
                        Error::InvalidFormat("RDFa start-tag size overflow".to_string())
                    })?;
                add_size(&mut length, attribute_len)?;
            }
        }
    }
    add_size(&mut length, source.len() - close)?;
    Ok(length)
}

fn source_namespace_binding(
    source: &str,
    prefix: &str,
    memory: &mut MetadataMemory<'_>,
) -> Result<Option<String>> {
    reserve_namespace_resolver_memory(source, memory)?;
    let mut reader = ResolvedReader::from_xml(source);
    reader.config_mut().trim_text(false);
    loop {
        let (_resolved_event, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid RDFa target start tag: {error}"))
        })?;
        match event {
            Event::Start(source) | Event::Empty(source) => {
                for attribute in source.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid RDFa namespace declaration: {error}"))
                    })?;
                    if attribute.key.as_ref().strip_prefix(b"xmlns:") == Some(prefix.as_bytes()) {
                        let Some(namespace) =
                            reader
                                .resolver()
                                .bindings()
                                .find_map(|(declaration, namespace)| match declaration {
                                    PrefixDeclaration::Named(value)
                                        if value == prefix.as_bytes() =>
                                    {
                                        Some(namespace)
                                    },
                                    _ => None,
                                })
                        else {
                            return Ok(None);
                        };
                        let resolved = ResolveResult::Bound(namespace);
                        return resolved_namespace_with_memory(&resolved, memory);
                    }
                }
                return Ok(None);
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "missing RDFa target start tag".to_string(),
                ));
            },
            _ => {},
        }
    }
}

fn metadata_namespace_declarations(
    source: &BytesStart<'_>,
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<String>> {
    let mut output = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid metadata namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            let key = std::str::from_utf8(raw).map_err(|_| {
                Error::InvalidFormat("invalid metadata namespace declaration name".to_string())
            })?;
            let value = std::str::from_utf8(attribute.value.as_ref()).map_err(|_| {
                Error::InvalidFormat("invalid metadata namespace declaration value".to_string())
            })?;
            reserve_vec_push(
                &mut output,
                memory.namespace_declaration(key, value, "ODT metadata namespace declaration")?,
                "ODT metadata namespace declarations",
                memory,
            )?;
        }
    }
    Ok(output)
}

fn apply_metadata_namespace_declarations(
    scope: &mut Vec<(String, String)>,
    declarations: &[String],
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<(String, Option<String>)>> {
    let mut changes = Vec::new();
    for declaration in declarations {
        let (name, value) = declaration_parts(declaration).ok_or_else(|| {
            Error::InvalidFormat("invalid metadata namespace declaration span".to_string())
        })?;
        let previous = scope
            .iter()
            .find(|(current, _)| current == name)
            .map(|(_, value)| memory.clone_string(value, "ODT metadata namespace scope"))
            .transpose()?;
        reserve_vec_push(
            &mut changes,
            (
                memory.clone_string(name, "ODT metadata namespace scope name")?,
                previous,
            ),
            "ODT metadata namespace changes",
            memory,
        )?;
        replace_metadata_namespace(scope, name, value, memory)?;
    }
    Ok(changes)
}

fn restore_metadata_namespace_declarations(
    scope: &mut Vec<(String, String)>,
    changes: Vec<(String, Option<String>)>,
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    for (name, previous) in changes.into_iter().rev() {
        match previous {
            Some(value) => replace_metadata_namespace(scope, &name, &value, memory)?,
            None => scope.retain(|(current, _)| current != &name),
        }
    }
    Ok(())
}

fn replace_metadata_namespace(
    scope: &mut Vec<(String, String)>,
    name: &str,
    value: &str,
    memory: &mut MetadataMemory<'_>,
) -> Result<()> {
    if let Some((_, current)) = scope.iter_mut().find(|(current, _)| current == name) {
        *current = memory.clone_string(value, "ODT metadata namespace scope")?;
    } else {
        reserve_vec_push(
            scope,
            (
                memory.clone_string(name, "ODT metadata namespace scope name")?,
                memory.clone_string(value, "ODT metadata namespace scope")?,
            ),
            "ODT metadata namespace scope",
            memory,
        )?;
    }
    Ok(())
}

fn metadata_namespace_scope_to_raw(
    scope: &[(String, String)],
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<String>> {
    let mut output = Vec::new();
    for (name, value) in scope {
        reserve_vec_push(
            &mut output,
            memory.namespace_declaration(name, value, "ODT metadata namespace snapshot")?,
            "ODT metadata namespace snapshots",
            memory,
        )?;
    }
    Ok(output)
}

fn declaration_parts(declaration: &str) -> Option<(&str, &str)> {
    let declaration = declaration.trim();
    let (name, value) = declaration.split_once('=')?;
    let value = value.trim().strip_prefix('"')?.strip_suffix('"')?;
    Some((name.trim(), value))
}

fn inject_namespace_declarations_with_limit_and_budget(
    raw: &str,
    declarations: &[String],
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    let mut scan_memory = MetadataMemory::new(budget);
    reserve_namespace_resolver_memory(raw, &mut scan_memory)?;
    let (open_end, empty, present) = metadata_first_tag_span(raw, &mut scan_memory)?;
    let mut insert_len = 0usize;
    for declaration in declarations {
        let Some((name, _)) = declaration_parts(declaration) else {
            continue;
        };
        if !present.contains(name) {
            insert_len = insert_len.checked_add(declaration.len()).ok_or_else(|| {
                Error::InvalidFormat("metadata namespace insertion size overflow".to_string())
            })?;
        }
    }
    if insert_len == 0 {
        bounded_output_len_with_limit(raw.len(), "metadata namespace exact no-op", maximum)?;
        let (mut output, memory) =
            allocate_xml(budget, raw.len(), "ODT metadata namespace exact no-op")?;
        output.push_str(raw);
        return Ok(ChargedXml {
            xml: output,
            memory,
        });
    }
    let offset = if empty {
        open_end.checked_sub(2).ok_or_else(|| {
            Error::InvalidFormat("invalid metadata empty element span".to_string())
        })?
    } else {
        open_end
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("invalid metadata element span".to_string()))?
    };
    let output_len = raw.len().checked_add(insert_len).ok_or_else(|| {
        Error::InvalidFormat("metadata namespace insertion size overflow".to_string())
    })?;
    bounded_output_len_with_limit(output_len, "metadata namespace insertion", maximum)?;
    let (mut output, memory) =
        allocate_xml(budget, output_len, "ODT metadata namespace insertion")?;
    output.push_str(&raw[..offset]);
    for declaration in declarations {
        let Some((name, _)) = declaration_parts(declaration) else {
            continue;
        };
        if !present.contains(name) {
            output.push_str(declaration);
        }
    }
    output.push_str(&raw[offset..]);
    Ok(ChargedXml {
        xml: output,
        memory,
    })
}

fn metadata_first_tag_span(
    raw: &str,
    memory: &mut MetadataMemory<'_>,
) -> Result<(usize, bool, HashSet<String>)> {
    let mut reader = Reader::from_str(raw);
    reader.config_mut().trim_text(false);
    let mut seen = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid metadata XML: {error}")))?;
        seen = seen
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("metadata XML event count overflow".to_string()))?;
        if seen & 0x03ff == 0 {
            memory.check()?;
        }
        match event {
            Event::Start(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    false,
                    metadata_present_namespace_names(&source, memory)?,
                ));
            },
            Event::Empty(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    true,
                    metadata_present_namespace_names(&source, memory)?,
                ));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in metadata XML".to_string(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "missing metadata element root".to_string(),
                ));
            },
            _ => {},
        }
    }
}

fn metadata_present_namespace_names(
    source: &BytesStart<'_>,
    memory: &mut MetadataMemory<'_>,
) -> Result<HashSet<String>> {
    let mut output = HashSet::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid metadata namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            let name = std::str::from_utf8(raw).map_err(|_| {
                Error::InvalidFormat("invalid metadata namespace declaration name".to_string())
            })?;
            reserve_hash_set_insert(&mut output, name, memory, "ODT metadata namespace names")?;
        }
    }
    Ok(output)
}

fn scan_spans_with_budget<'a>(
    xml: &str,
    wanted: impl Fn(&str, &str) -> bool,
    budget: Option<&'a FlatMutationBudget>,
) -> Result<ScannedSpans<'a>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    let mut memory = MetadataMemory::new(budget);
    reserve_namespace_resolver_memory(xml, &mut memory)?;
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut stack: Vec<(Span, Vec<(String, Option<String>)>)> = Vec::new();
    let mut namespace_scope = Vec::new();
    let mut output = Vec::new();
    let mut seen = 0usize;
    loop {
        let event_position = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF XML: {error}")))?;
        let namespace = match &event {
            Event::Start(_) | Event::Empty(_) => {
                resolved_namespace_with_memory(&namespace, &mut memory)?.unwrap_or_default()
            },
            _ => String::new(),
        };
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref source) => {
                let local = utf8_with_memory(
                    source.local_name().as_ref(),
                    "edit target name",
                    &mut memory,
                )?;
                let start = event_position;
                let declarations = metadata_namespace_declarations(source, &mut memory)?;
                let changes = apply_metadata_namespace_declarations(
                    &mut namespace_scope,
                    &declarations,
                    &mut memory,
                )?;
                let attributes = parse_attributes(&reader, source, &mut memory)?;
                let rdfa_names = target_attribute_names(&reader, source, &mut memory)?;
                let name = target_attribute_value(&reader, source, TEXTNS, "name", &mut memory)?;
                let span = Span {
                    start,
                    start_end: event_end,
                    end_start: 0,
                    end: 0,
                    namespace: memory.clone_string(&namespace, "ODT metadata span namespace")?,
                    local: memory.clone_string(&local, "ODT metadata span name")?,
                    rdfa_names,
                    rdfa: rdfa_from_attributes(&attributes, &mut memory)?,
                    name,
                    namespace_declarations: metadata_namespace_scope_to_raw(
                        &namespace_scope,
                        &mut memory,
                    )?,
                };
                reserve_vec_push(
                    &mut stack,
                    (span, changes),
                    "ODT metadata span stack",
                    &mut memory,
                )?;
            },
            Event::Empty(ref source) => {
                let local = utf8_with_memory(
                    source.local_name().as_ref(),
                    "edit target name",
                    &mut memory,
                )?;
                let declarations = metadata_namespace_declarations(source, &mut memory)?;
                let changes = apply_metadata_namespace_declarations(
                    &mut namespace_scope,
                    &declarations,
                    &mut memory,
                )?;
                if wanted(&namespace, &local) {
                    let start = event_position;
                    let attributes = parse_attributes(&reader, source, &mut memory)?;
                    let rdfa_names = target_attribute_names(&reader, source, &mut memory)?;
                    let name =
                        target_attribute_value(&reader, source, TEXTNS, "name", &mut memory)?;
                    let span = Span {
                        start,
                        start_end: event_end,
                        end_start: event_end,
                        end: event_end,
                        namespace,
                        local,
                        rdfa_names,
                        rdfa: rdfa_from_attributes(&attributes, &mut memory)?,
                        name,
                        namespace_declarations: metadata_namespace_scope_to_raw(
                            &namespace_scope,
                            &mut memory,
                        )?,
                    };
                    push_bounded_with_memory(
                        &mut output,
                        span,
                        "ODT metadata edit spans",
                        &mut memory,
                    )?;
                }
                restore_metadata_namespace_declarations(
                    &mut namespace_scope,
                    changes,
                    &mut memory,
                )?;
            },
            Event::End(_) => {
                let end = event_end;
                let end_start = event_position;
                let Some((mut span, changes)) = stack.pop() else {
                    return Err(Error::InvalidFormat(
                        "metadata XML stack underflow while locating edit spans".to_string(),
                    ));
                };
                span.end_start = end_start;
                span.end = end;
                if wanted(&span.namespace, &span.local) {
                    push_bounded_with_memory(
                        &mut output,
                        span,
                        "ODT metadata edit spans",
                        &mut memory,
                    )?;
                }
                restore_metadata_namespace_declarations(
                    &mut namespace_scope,
                    changes,
                    &mut memory,
                )?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF XML edits".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        seen = seen
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("metadata XML event count overflow".to_string()))?;
        if seen > MAX_OCCURRENCES {
            return Err(Error::InvalidFormat(format!(
                "metadata XML exceeds {MAX_OCCURRENCES} events while locating edit spans"
            )));
        }
        if seen & 0x03ff == 0 {
            memory.check()?;
        }
    }
    if !stack.is_empty() || !namespace_scope.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete metadata XML while locating edit spans".to_string(),
        ));
    }
    output.sort_by_key(|span| span.start);
    Ok(ScannedSpans {
        spans: output,
        _memory: memory,
    })
}

fn target_attribute_names(
    reader: &ResolvedReader<'_>,
    source: &BytesStart<'_>,
    memory: &mut MetadataMemory<'_>,
) -> Result<Vec<String>> {
    let mut rdfa = Vec::new();
    for attr in source.attributes() {
        let attr =
            attr.map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let raw = utf8_with_memory(attr.key.as_ref(), "XML attribute name", memory)?;
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        let namespace = resolved_namespace_with_memory(&namespace, memory)?;
        if namespace.as_deref() == Some(XHTMLNS)
            && matches!(
                local.as_ref(),
                b"about" | b"property" | b"content" | b"datatype"
            )
        {
            reserve_vec_push(&mut rdfa, raw, "ODT RDFa attribute names", memory)?;
        }
    }
    Ok(rdfa)
}

fn target_attribute_value(
    reader: &ResolvedReader<'_>,
    source: &BytesStart<'_>,
    namespace: &str,
    local: &str,
    memory: &mut MetadataMemory<'_>,
) -> Result<Option<String>> {
    for attr in source.attributes() {
        let attr =
            attr.map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local_name) = reader.resolver().resolve_attribute(attr.key);
        if crate::elements::xml::normalized_namespace_matches(
            &resolved,
            namespace.as_bytes(),
            "ODT metadata target attribute",
        )? && local_name.as_ref() == local.as_bytes()
        {
            if metadata_attribute_decode_needs_owned(attr.value.as_ref()) {
                memory.reserve_bytes(
                    attr.value.len(),
                    "ODT metadata target attribute decode scratch",
                )?;
            }
            return attr
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                })
                .and_then(|value| {
                    memory
                        .clone_string(value.as_ref(), "ODT metadata target attribute value")
                        .map(Some)
                });
        }
    }
    Ok(None)
}

fn splice_insert_with_limit_and_budget(
    xml: &str,
    offset: usize,
    value: &str,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    if offset > xml.len() {
        return Err(Error::InvalidFormat(
            "invalid XML insertion offset".to_string(),
        ));
    }
    let output_len = xml
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("XML insertion size overflow".to_string()))?;
    bounded_output_len_with_limit(output_len, "XML insertion", maximum)?;
    let (mut output, memory) = allocate_xml(budget, output_len, "ODT XML insertion")?;
    output.push_str(&xml[..offset]);
    output.push_str(value);
    output.push_str(&xml[offset..]);
    Ok(ChargedXml {
        xml: output,
        memory,
    })
}

fn splice_replace_with_limit_and_budget(
    xml: &str,
    start: usize,
    end: usize,
    value: &str,
    maximum: usize,
    budget: Option<&FlatMutationBudget>,
) -> Result<ChargedXml> {
    if start > end || end > xml.len() {
        return Err(Error::InvalidFormat(
            "invalid XML replacement span".to_string(),
        ));
    }
    let removed = end
        .checked_sub(start)
        .ok_or_else(|| Error::InvalidFormat("XML replacement span underflow".to_string()))?;
    let output_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(value.len()))
        .ok_or_else(|| Error::InvalidFormat("XML replacement size overflow".to_string()))?;
    bounded_output_len_with_limit(output_len, "XML replacement", maximum)?;
    let (mut output, memory) = allocate_xml(budget, output_len, "ODT XML replacement")?;
    output.push_str(&xml[..start]);
    output.push_str(value);
    output.push_str(&xml[end..]);
    Ok(ChargedXml {
        xml: output,
        memory,
    })
}

fn bounded_output_len_with_limit(length: usize, operation: &str, maximum: usize) -> Result<()> {
    if length > maximum {
        if maximum < MAX_XML_BYTES {
            return Err(Error::ResourceLimit(ResourceLimit {
                resource: Resource::OutputBytes,
                observed: u64::try_from(length).unwrap_or(u64::MAX),
                limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                scope: operation.into(),
            }));
        }
        return Err(Error::InvalidFormat(format!(
            "{operation} exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    if length > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{operation} exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    Ok(())
}

fn metadata_output_len(
    value: &TextMeta,
    namespace_declarations: &[String],
    memory: &mut MetadataMemory<'_>,
) -> Result<usize> {
    let mut length =
        "<text:meta xmlns:text=\"\" xmlns:xhtml=\"\"".len() + TEXTNS.len() + XHTMLNS.len();
    if let Some(id) = &value.xml_id {
        add_size(&mut length, " xml:id=\"\"".len())?;
        add_size(&mut length, escaped_xml_len(id, true)?)?;
    }
    for prefix in rdfa_prefixes_with_memory(&value.rdfa, memory)? {
        if matches!(prefix.as_str(), "text" | "xhtml" | "xml")
            || namespace_binding(namespace_declarations, &prefix).is_some()
        {
            continue;
        }
        if let Some(namespace) = canonical_rdfa_namespace(&prefix) {
            add_size(
                &mut length,
                " xmlns:=\"\"".len() + prefix.len() + namespace.len(),
            )?;
        }
    }
    for (name, item) in [
        ("about", value.rdfa.about.as_deref()),
        ("property", value.rdfa.property.as_deref()),
        ("content", value.rdfa.content.as_deref()),
        ("datatype", value.rdfa.datatype.as_deref()),
    ] {
        if let Some(item) = item {
            add_size(&mut length, " xhtml:=\"\"".len() + name.len())?;
            add_size(&mut length, escaped_xml_len(item, true)?)?;
        }
    }
    for attribute in &value.attributes {
        // Metadata attributes can introduce a declaration for their retained
        // namespace.  The generated prefix itself is at most the `metaN`
        // form used by push_metadata_attributes; reserve a small fixed margin
        // while counting all escaped lexical values exactly.
        let namespace_size = escaped_xml_len(&attribute.namespace_uri, true)?;
        let local_size = attribute.local_name.len();
        let value_size = escaped_xml_len(&attribute.value, true)?;
        let attribute_size = 256usize
            .checked_add(namespace_size)
            .and_then(|size| size.checked_add(local_size))
            .and_then(|size| size.checked_add(value_size))
            .ok_or_else(|| {
                Error::InvalidFormat("text:meta serialization size overflow".to_string())
            })?;
        add_size(&mut length, attribute_size)?;
    }
    let mut node_count = 0usize;
    add_size(
        &mut length,
        metadata_nodes_output_len(value.content.nodes(), memory, &mut node_count)?,
    )?;
    add_size(&mut length, "\"/>".len())?;
    if !value.content.nodes().is_empty() {
        add_size(&mut length, "</text:meta>".len())?;
    }
    Ok(length)
}

fn metadata_nodes_output_len(
    nodes: &[MetaFieldNode],
    memory: &mut MetadataMemory<'_>,
    node_count: &mut usize,
) -> Result<usize> {
    let mut length = 0usize;
    for node in nodes {
        *node_count = node_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("text:meta output node count overflow".to_string())
        })?;
        if *node_count & 0x03ff == 0 {
            memory.check()?;
        }
        match node {
            MetaFieldNode::Text(value) => {
                if value.bytes().any(|byte| matches!(byte, b'&' | b'<' | b'>')) {
                    memory.reserve_bytes(
                        escaped_xml_len(value, false)?,
                        "ODT text:meta content escape scratch",
                    )?;
                }
                add_size(&mut length, escaped_xml_len(value, false)?)?
            },
            MetaFieldNode::Element(element) => {
                let prefix = canonical_meta_element_prefix(&element.namespace_uri)?;
                add_size(
                    &mut length,
                    1 + prefix.len()
                        + 1
                        + element.local_name.len()
                        + " xmlns:=\"\"".len()
                        + prefix.len()
                        + element.namespace_uri.len(),
                )?;
                let mut declared = Vec::new();
                memory.reserve_vec(
                    &mut declared,
                    element.attributes.len() + 1,
                    "ODT text:meta serialization namespace set",
                )?;
                declared.push(prefix);
                for attribute in &element.attributes {
                    let attribute_prefix = canonical_meta_element_prefix(&attribute.namespace_uri)?;
                    if attribute_prefix != "xml" && !declared.contains(&attribute_prefix) {
                        add_size(
                            &mut length,
                            " xmlns:=\"\"".len()
                                + attribute_prefix.len()
                                + attribute.namespace_uri.len(),
                        )?;
                        declared.push(attribute_prefix);
                    }
                    add_size(
                        &mut length,
                        1 + attribute_prefix.len()
                            + 1
                            + attribute.local_name.len()
                            + "=\"\"".len()
                            + escaped_xml_len(&attribute.value, true)?,
                    )?;
                    if attribute
                        .value
                        .bytes()
                        .any(|byte| matches!(byte, b'&' | b'<' | b'>' | b'"' | b'\''))
                    {
                        memory.reserve_bytes(
                            escaped_xml_len(&attribute.value, true)?,
                            "ODT text:meta content attribute escape scratch",
                        )?;
                    }
                }
                if element.children.is_empty() {
                    add_size(&mut length, 2)?;
                } else {
                    add_size(&mut length, 1)?;
                    add_size(
                        &mut length,
                        metadata_nodes_output_len(&element.children, memory, node_count)?,
                    )?;
                    add_size(
                        &mut length,
                        "</::>".len() + prefix.len() + element.local_name.len(),
                    )?;
                }
            },
        }
    }
    Ok(length)
}

fn escaped_xml_len(value: &str, attribute: bool) -> Result<usize> {
    let mut length = 0usize;
    for byte in value.bytes() {
        let extra = match byte {
            b'&' => 4,
            b'<' | b'>' => 3,
            b'"' if attribute => 5,
            b'\'' if attribute => 5,
            _ => 0,
        };
        add_size(&mut length, 1 + extra)?;
    }
    Ok(length)
}

fn canonical_meta_element_prefix(namespace: &str) -> Result<&'static str> {
    match namespace {
        TEXTNS => Ok("text"),
        OFFICENS => Ok("office"),
        STYLENS => Ok("style"),
        XLINKNS => Ok("xlink"),
        XMLNS => Ok("xml"),
        DRAWNS => Ok("draw"),
        TABLENS => Ok("table"),
        PRESENTATIONNS => Ok("presentation"),
        SVGNS => Ok("svg"),
        FONS => Ok("fo"),
        NUMBERNS => Ok("number"),
        METANS => Ok("meta"),
        DCNS => Ok("dc"),
        XHTMLNS => Ok("xhtml"),
        DR3DNS => Ok("dr3d"),
        FORMNS => Ok("form"),
        SCRIPTNS => Ok("script"),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported text:meta namespace '{namespace}'"
        ))),
    }
}

fn add_size(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::InvalidFormat("text:meta serialization size overflow".to_string()))?;
    Ok(())
}

// Keep the import visible to rustc when this module is compiled with a
// feature set that does not use the qualified-name helper directly.
#[allow(dead_code)]
fn _qualified_name(_: QName<'_>) {}

#[cfg(test)]
mod tests {
    use super::*;
    use std::num::{NonZeroU64, NonZeroUsize};

    use litchi_core::{
        Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
    };

    const XML: &str = r#"<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:xhtml="http://www.w3.org/1999/xhtml"
 xmlns:xml="http://www.w3.org/XML/1998/namespace">
 <office:body><office:text><text:p xhtml:about="urn:before">Hello<!--keep--><text:bookmark-start text:name="mark" xhtml:property="dc:title"/> <text:meta xml:id="m1" xhtml:property="dc:description">Meta <text:span text:style-name="Emph">text</text:span></text:meta></text:p></office:text></office:body>
</office:document-content>"#;

    fn metadata_budget(memory: u64) -> FlatMutationBudget {
        let root = Budget::root(
            "content metadata test",
            BudgetLimits::new(memory, 1 << 30, 1 << 30, 1_000_000, 4_096, 1_000_000_000),
        );
        let (_source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(memory.max(1)).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(root, cancellation, limits);
        FlatMutationBudget::new(&context, "content metadata test mutation").unwrap()
    }

    fn reserved_bytes(memory: &MetadataMemory<'_>) -> usize {
        memory.reservation.as_ref().map_or(0, |reservation| {
            usize::try_from(reservation.amount()).unwrap()
        })
    }

    #[test]
    fn span_scan_charges_namespace_and_span_state_before_allocation() {
        let xml = format!(
            r#"<root xmlns:t="{TEXTNS}" xmlns:a="urn:one" xmlns:b="urn:two"><t:p xmlns:a="urn:shadow" a:key="one" b:key="two"><t:span xmlns:c="urn:three" c:key="three"/></t:p><t:p xmlns:d="urn:four" d:key="four"/></root>"#
        );
        let probe_budget = metadata_budget(1 << 30);
        let probe = scan_spans_with_budget(
            &xml,
            |namespace, local| namespace == TEXTNS && local == "p",
            Some(&probe_budget),
        )
        .unwrap();
        let required = reserved_bytes(&probe._memory);
        assert!(required > xml.len());
        drop(probe);

        // The final lease retains the two span slots, while one old slot is
        // live during the second growth. Charge that exact peak scratch.
        let peak_required = required.checked_add(size_of::<Span>()).unwrap();
        let exact_budget = metadata_budget(u64::try_from(peak_required).unwrap());
        assert!(
            scan_spans_with_budget(
                &xml,
                |namespace, local| namespace == TEXTNS && local == "p",
                Some(&exact_budget),
            )
            .is_ok()
        );

        let under_budget = metadata_budget(u64::try_from(peak_required - 1).unwrap());
        assert!(matches!(
            scan_spans_with_budget(
                &xml,
                |namespace, local| namespace == TEXTNS && local == "p",
                Some(&under_budget),
            ),
            Err(Error::ResourceLimit(_))
        ));
    }

    #[test]
    fn metadata_projection_charges_node_and_string_ownership_before_parse() {
        let xml = format!(
            r#"<text:meta xmlns:text="{TEXTNS}" xmlns:xhtml="{XHTMLNS}" xhtml:property="dc:title"><text:span text:style-name="one">first <text:span text:style-name="two">second</text:span></text:span><text:soft-page-break/>tail</text:meta>"#
        );
        let probe_budget = metadata_budget(1 << 30);
        let mut probe_memory = MetadataMemory::new(Some(&probe_budget));
        let parsed = parse_part_inner(&xml, MetadataPart::Content, &mut probe_memory).unwrap();
        assert_eq!(
            parsed.text_meta[0].content.display_text(),
            "first secondtail"
        );
        let required = reserved_bytes(&probe_memory);
        assert!(required > xml.len());
        drop(parsed);
        drop(probe_memory);

        let exact_budget = metadata_budget(u64::try_from(required).unwrap());
        let (parsed, lease) =
            parse_part_with_budget(&xml, MetadataPart::Content, Some(&exact_budget)).unwrap();
        assert!(lease.is_some());
        assert_eq!(parsed.text_meta.len(), 1);
        drop(parsed);
        drop(lease);

        let under_budget = metadata_budget(u64::try_from(required - 1).unwrap());
        assert!(matches!(
            parse_part_with_budget(&xml, MetadataPart::Content, Some(&under_budget)),
            Err(Error::ResourceLimit(_))
        ));
    }

    #[test]
    fn parses_rdfa_and_text_meta_without_evaluating_values() {
        let parsed = parse_part(XML, MetadataPart::Content).expect("metadata parses");
        assert_eq!(parsed.text_meta.len(), 1);
        assert_eq!(parsed.text_meta[0].xml_id.as_deref(), Some("m1"));
        assert_eq!(
            parsed.text_meta[0].rdfa.property.as_deref(),
            Some("dc:description")
        );
        assert_eq!(parsed.text_meta[0].content.display_text(), "Meta text");
        assert!(parsed.rdfa.iter().any(|value| {
            matches!(value.host, RdfaHost::Paragraph { index: 0 })
                && value.attributes.about.as_deref() == Some("urn:before")
        }));
        assert!(parsed.rdfa.iter().any(|value| {
            matches!(value.host, RdfaHost::BookmarkStart { .. })
                && value.attributes.property.as_deref() == Some("dc:title")
        }));
    }

    #[test]
    fn escaped_namespace_uris_support_query_mutation_and_exact_noop() {
        let escaped = XML
            .replace(
                "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
                "urn:oasis:names:tc:opendocument:xmlns:office&#58;1.0",
            )
            .replace(
                "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
                "urn:oasis:names:tc:opendocument:xmlns:text&#58;1.0",
            )
            .replace(
                "http://www.w3.org/1999/xhtml",
                "http&#58;//www.w3.org/1999/xhtml",
            );
        let parsed = parse_part(&escaped, MetadataPart::Content).expect("metadata parses");
        assert_eq!(parsed.text_meta.len(), 1);
        assert!(parsed.rdfa.iter().any(|value| {
            matches!(value.host, RdfaHost::Paragraph { index: 0 })
                && value.attributes.about.as_deref() == Some("urn:before")
        }));
        assert!(parsed.rdfa.iter().any(|value| {
            matches!(
                value.host,
                RdfaHost::BookmarkStart {
                    name: Some(ref name)
                } if name == "mark"
            )
        }));

        let updated = set_paragraph_rdfa(
            &escaped,
            Position::new(0),
            &RdfaAttributes {
                about: Some("urn:after".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("escaped text namespace should locate the paragraph");
        assert!(
            updated.contains("xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text&#58;1.0\"")
        );
        assert!(updated.contains("xhtml:about=\"urn:after\""));
        assert_eq!(
            parse_part(&updated, MetadataPart::Content)
                .expect("mutated metadata reparses")
                .rdfa
                .iter()
                .find_map(|value| match value.host {
                    RdfaHost::Paragraph { index: 0 } => value.attributes.about.clone(),
                    _ => None,
                })
                .as_deref(),
            Some("urn:after")
        );

        let noop = replace_text_meta(&escaped, Position::new(0), &parsed.text_meta[0])
            .expect("semantically equal metadata should be an exact no-op");
        assert_eq!(noop, escaped);

        let malformed = escaped.replacen("&#58;", "&unknown;", 1);
        assert!(parse_part(&malformed, MetadataPart::Content).is_err());
    }

    #[test]
    fn rdfa_edit_preserves_unrelated_inline_markup() {
        let value = RdfaAttributes {
            about: Some("urn:after".to_string()),
            ..RdfaAttributes::default()
        };
        let updated = set_paragraph_rdfa(XML, Position::new(0), &value).expect("paragraph edit");
        assert!(updated.contains("xhtml:about=\"urn:after\""));
        assert!(updated.contains("<!--keep-->"));
        assert!(updated.contains("text:meta"));
        let parsed = parse_part(&updated, MetadataPart::Content).expect("reparse");
        assert_eq!(parsed.text_meta[0].content.display_text(), "Meta text");
    }

    #[test]
    fn rdfa_limit_accepts_exact_output_and_rejects_one_byte_under() {
        let value = RdfaAttributes {
            about: Some("urn:after".to_string()),
            ..RdfaAttributes::default()
        };
        let expected = set_paragraph_rdfa(XML, Position::new(0), &value).unwrap();
        assert_eq!(
            set_paragraph_rdfa_with_limit(XML, Position::new(0), &value, expected.len()).unwrap(),
            expected
        );
        assert!(matches!(
            set_paragraph_rdfa_with_limit(XML, Position::new(0), &value, expected.len() - 1),
            Err(Error::ResourceLimit(_))
        ));
    }

    #[test]
    fn empty_paragraph_insertion_uses_exact_replacement_size() {
        let xml = format!(r#"<text:p xmlns:text="{TEXTNS}"/>"#);
        let value = TextMeta::from_text("inserted").unwrap();
        let expected = insert_text_meta(&xml, Position::new(0), &value).unwrap();
        assert_eq!(
            insert_text_meta_with_limit(&xml, Position::new(0), &value, expected.len()).unwrap(),
            expected
        );
        assert!(matches!(
            insert_text_meta_with_limit(&xml, Position::new(0), &value, expected.len() - 1,),
            Err(Error::ResourceLimit(_))
        ));
    }

    #[test]
    fn text_meta_insert_and_remove_round_trip() {
        let value = TextMeta::from_text("added").expect("text metadata");
        let inserted = insert_text_meta(XML, Position::new(0), &value).expect("insert");
        assert_eq!(
            parse_part(&inserted, MetadataPart::Content)
                .unwrap()
                .text_meta
                .len(),
            2
        );
        let removed = remove_text_meta(&inserted, Position::new(1)).expect("remove");
        assert_eq!(
            parse_part(&removed, MetadataPart::Content)
                .unwrap()
                .text_meta
                .len(),
            1
        );
    }

    #[test]
    fn equivalent_text_meta_replacement_is_an_exact_noop() {
        let parsed = parse_part(XML, MetadataPart::Content).expect("metadata parses");
        assert_eq!(
            replace_text_meta(XML, Position::new(0), &parsed.text_meta[0]).unwrap(),
            XML
        );
    }

    #[test]
    fn text_meta_unknown_root_attributes_are_retained_and_rdfa_edits_are_lossless() {
        let source = XML.replace(
            "<text:meta xml:id=\"m1\"",
            "<text:meta xmlns:ext=\"urn:example:meta\" ext:keep=\"yes\" xml:id=\"m1\"",
        );
        let parsed = parse_part(&source, MetadataPart::Content).expect("metadata parses");
        assert_eq!(parsed.text_meta[0].attributes.len(), 1);
        assert_eq!(parsed.text_meta[0].attributes[0].value, "yes");
        let serialized = parsed.text_meta[0]
            .to_xml_with_context_with_limit(&[], MAX_XML_BYTES)
            .expect("metadata serializes");
        assert!(serialized.contains("ext:keep=\"yes\""));
        let updated = set_text_meta_rdfa(
            &source,
            Position::new(0),
            &RdfaAttributes {
                about: Some("urn:updated".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("RDFa edit");
        assert!(updated.contains("ext:keep=\"yes\""));
    }

    #[test]
    fn text_meta_accepts_soft_page_break_but_rejects_text_number() {
        let soft_break = XML.replace(
            "Meta <text:span text:style-name=\"Emph\">text</text:span>",
            "Meta<text:soft-page-break/><text:span text:style-name=\"Emph\">text</text:span>",
        );
        let parsed = parse_part(&soft_break, MetadataPart::Content).expect("soft break parses");
        assert_eq!(parsed.text_meta[0].content.display_text(), "Metatext");
        assert!(parsed.text_meta[0].content.nodes().iter().any(|node| {
            matches!(
                node,
                MetaFieldNode::Element(element)
                    if element.namespace_uri == TEXTNS && element.local_name == "soft-page-break"
            )
        }));

        let number = XML.replace(
            "Meta <text:span text:style-name=\"Emph\">text</text:span>",
            "Meta<text:number>1.</text:number>",
        );
        assert!(parse_part(&number, MetadataPart::Content).is_err());
    }

    #[test]
    fn text_meta_serialization_rejects_escaped_output_before_building_it() {
        let mut value = TextMeta::new(MetaFieldContent::new(Vec::new()).unwrap()).unwrap();
        value.attributes.push(MetadataAttribute {
            namespace_uri: String::new(),
            local_name: "flag".to_string(),
            value: "&".repeat(13 * 1024 * 1024),
            prefix: None,
        });
        assert!(
            value
                .to_xml_with_context_with_limit(&[], MAX_XML_BYTES)
                .is_err()
        );
    }

    #[test]
    fn rdfa_namespace_detection_ignores_values_and_rejects_shadowing() {
        let source = XML.replace(
            "<text:p xhtml:about=\"urn:before\">",
            "<text:p data-note=\"xmlns:xhtml\">Hello</text:p><text:p>",
        );
        let updated = set_paragraph_rdfa(
            &source,
            Position::new(1),
            &RdfaAttributes {
                about: Some("urn:after".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("RDFa namespace is added despite a value containing its name");
        assert!(updated.contains("xmlns:xhtml=\"http://www.w3.org/1999/xhtml\""));

        let shadowed = XML.replace(
            "<text:p xhtml:about=\"urn:before\">",
            "<text:p xmlns:xhtml=\"urn:wrong\">",
        );
        assert!(
            set_paragraph_rdfa(
                &shadowed,
                Position::new(0),
                &RdfaAttributes {
                    about: Some("urn:after".to_string()),
                    ..RdfaAttributes::default()
                },
            )
            .is_err()
        );
    }

    #[test]
    fn rdfa_curie_edits_require_or_add_namespace_context() {
        let missing = set_paragraph_rdfa(
            XML,
            Position::new(0),
            &RdfaAttributes {
                property: Some("custom:term".to_string()),
                ..RdfaAttributes::default()
            },
        );
        assert!(missing.is_err());

        let source = XML.replace(
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\"",
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\" xmlns:custom=\"urn:custom\"",
        );
        let updated = set_paragraph_rdfa(
            &source,
            Position::new(0),
            &RdfaAttributes {
                property: Some("custom:term".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("in-scope CURIE prefix is preserved");
        assert!(updated.contains("xmlns:custom=\"urn:custom\""));
        assert!(updated.contains("xhtml:property=\"custom:term\""));

        let added = set_paragraph_rdfa(
            XML,
            Position::new(0),
            &RdfaAttributes {
                property: Some("dc:term".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("canonical DC CURIE prefix is bound");
        assert!(added.contains("xmlns:dc=\"http://purl.org/dc/elements/1.1/\""));

        let custom_source = XML.replace(
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\"",
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\" xmlns:custom=\"urn:custom\"",
        );
        let inserted = insert_text_meta(
            &custom_source,
            Position::new(0),
            &TextMeta {
                rdfa: RdfaAttributes {
                    property: Some("custom:term".to_string()),
                    ..RdfaAttributes::default()
                },
                ..TextMeta::from_text("custom").unwrap()
            },
        )
        .expect("custom CURIE context is retained for inserted metadata");
        assert!(inserted.contains("xmlns:custom=\"urn:custom\""));
        assert!(parse_part(&inserted, MetadataPart::Content).is_ok());
    }
}
