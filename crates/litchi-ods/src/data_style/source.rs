//! Borrowed, source-qualified data-style scanning.
//!
//! This module deliberately stops at the XML owner boundary.  It builds a
//! bounded catalog of data-style projections while retaining byte ranges into
//! the caller's XML for source-preserving edits.  Package and transaction
//! publication belong to the document owner.

use std::{borrow::Cow, collections::BTreeSet, ops::Range, str, sync::Arc};

use litchi_core::{Error, Resource, Result};
use litchi_odf_common::ResolvedReader;
use quick_xml::{
    XmlVersion,
    events::{BytesDecl, BytesPI, BytesStart, Event},
    name::{Namespace, Prefix, PrefixDeclaration, ResolveResult},
};

use super::{
    Attributes, Body, Data, Double, EmbeddedText, Entry, Family, Format, Fraction,
    MAX_STYLE_BODY_BYTES, MAX_STYLE_BODY_ELEMENTS, MAX_STYLE_CATALOG_ENTRIES,
    MAX_STYLE_METADATA_BYTES, MAX_STYLE_OWNER_BYTES, MAX_STYLE_OWNER_DEPTH,
    MAX_STYLE_OWNER_ELEMENTS, MAX_STYLE_TEXT_BYTES, Number, Op, Opaque, OpaqueBody, Owner, Patch,
    Scientific, Selector, TextEntry, Transliteration, TransliterationStyle,
};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const NUMBER: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS: &[u8] = b"http://www.w3.org/2000/xmlns/";
// quick-xml bounds declarations on one element; this aggregate cap keeps
// inherited resolver state finite without copying the active scope per node.
const MAX_ACTIVE_NAMESPACE_BINDINGS: usize = MAX_STYLE_OWNER_DEPTH * 256;

/// Caller-adjustable owner scan limits.  Hard ceilings are always applied.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) struct Limits {
    pub owner_bytes: usize,
    pub owner_elements: usize,
    pub owner_depth: usize,
    pub catalog_entries: usize,
    pub body_elements: usize,
    pub body_bytes: usize,
    pub text_bytes: usize,
    pub metadata_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            owner_bytes: MAX_STYLE_OWNER_BYTES,
            owner_elements: MAX_STYLE_OWNER_ELEMENTS,
            owner_depth: MAX_STYLE_OWNER_DEPTH,
            catalog_entries: MAX_STYLE_CATALOG_ENTRIES,
            body_elements: MAX_STYLE_BODY_ELEMENTS,
            body_bytes: MAX_STYLE_BODY_BYTES,
            text_bytes: MAX_STYLE_TEXT_BYTES,
            metadata_bytes: MAX_STYLE_METADATA_BYTES,
        }
    }
}

impl Limits {
    fn bounded(self) -> Self {
        Self {
            owner_bytes: self.owner_bytes.min(MAX_STYLE_OWNER_BYTES),
            owner_elements: self.owner_elements.min(MAX_STYLE_OWNER_ELEMENTS),
            owner_depth: self.owner_depth.min(MAX_STYLE_OWNER_DEPTH),
            catalog_entries: self.catalog_entries.min(MAX_STYLE_CATALOG_ENTRIES),
            body_elements: self.body_elements.min(MAX_STYLE_BODY_ELEMENTS),
            body_bytes: self.body_bytes.min(MAX_STYLE_BODY_BYTES),
            text_bytes: self.text_bytes.min(MAX_STYLE_TEXT_BYTES),
            metadata_bytes: self.metadata_bytes.min(MAX_STYLE_METADATA_BYTES),
        }
    }
}

/// A borrowed catalog for one direct XML owner.
#[allow(dead_code)]
#[derive(Debug)]
pub(crate) struct Catalog<'a> {
    source: &'a [u8],
    owner: Owner,
    owner_range: Option<Range<usize>>,
    entries: Vec<SourceEntry<'a>>,
}

#[allow(dead_code)]
impl<'a> Catalog<'a> {
    /// Return the owner represented by this catalog.
    #[must_use]
    pub(crate) const fn owner(&self) -> Owner {
        self.owner
    }

    /// Return the original owner XML without copying it.
    #[must_use]
    pub(crate) fn source(&self) -> &'a [u8] {
        self.source
    }

    /// Return the selected direct owner range, if the owner existed.
    #[must_use]
    pub(crate) fn owner_range(&self) -> Option<Range<usize>> {
        self.owner_range.clone()
    }

    /// Borrow all recognized data-style entries in source order.
    #[must_use]
    pub(crate) fn entries(&self) -> &[SourceEntry<'a>] {
        &self.entries
    }

    /// Resolve a fully qualified selector to exactly one entry.
    pub(crate) fn lookup(&self, selector: Selector<'_>) -> Result<&SourceEntry<'a>> {
        selector.validate()?;
        let mut found = None;
        for entry in &self.entries {
            if entry.selector() == selector {
                if found.is_some() {
                    return invalid("duplicate data-style selector in source owner");
                }
                found = Some(entry);
            }
        }
        found.ok_or_else(|| invalid_error("data-style selector was not found"))
    }

    /// Lookup a name without inventing an owner or family.
    pub(crate) fn lookup_name(&self, name: &str) -> Result<&SourceEntry<'a>> {
        let mut found = None;
        for entry in &self.entries {
            if entry.name() == name {
                if found.is_some() {
                    return invalid("data-style name is ambiguous within source owner");
                }
                found = Some(entry);
            }
        }
        found.ok_or_else(|| invalid_error("data-style name was not found"))
    }
}

/// One source-qualified style projection.  The semantic entry owns only
/// bounded metadata; all source bytes remain borrowed through `Catalog`.
#[allow(dead_code)]
#[derive(Clone, Debug)]
pub(crate) struct SourceEntry<'a> {
    source: &'a [u8],
    entry: Entry,
    range: Range<usize>,
    opening_range: Range<usize>,
    body_range: Range<usize>,
    attributes: AttributeSlots,
}

#[allow(dead_code)]
impl<'a> SourceEntry<'a> {
    #[must_use]
    pub(crate) fn entry(&self) -> &Entry {
        &self.entry
    }

    #[must_use]
    pub(crate) fn selector(&self) -> Selector<'_> {
        self.entry.selector()
    }

    #[must_use]
    pub(crate) fn owner(&self) -> Owner {
        self.entry.owner()
    }

    #[must_use]
    pub(crate) fn family(&self) -> Family {
        self.entry.family()
    }

    #[must_use]
    pub(crate) fn name(&self) -> &str {
        self.entry.name()
    }

    #[must_use]
    pub(crate) fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    #[must_use]
    pub(crate) fn opening_range(&self) -> Range<usize> {
        self.opening_range.clone()
    }

    #[must_use]
    pub(crate) fn body_range(&self) -> Range<usize> {
        self.body_range.clone()
    }

    #[must_use = "the checked source slice is borrowed from the scanned source"]
    pub(crate) fn source(&self) -> Result<&'a [u8]> {
        self.source_slice(&self.range, "data-style source range")
    }

    #[must_use = "the checked body slice is borrowed from the scanned source"]
    pub(crate) fn raw_body(&self) -> Result<&'a [u8]> {
        self.source_slice(&self.body_range, "data-style body range")
    }

    fn source_slice(&self, range: &Range<usize>, label: &'static str) -> Result<&'a [u8]> {
        self.source
            .get(range.clone())
            .ok_or_else(|| invalid_error(format!("{label} is outside the source")))
    }

    /// Apply only common root attributes, preserving the selected body and
    /// every unselected byte.  The result is checked before allocation.
    pub(crate) fn patch_metadata(&self, patch: &Patch, max_output: usize) -> Result<Vec<u8>> {
        if !self.owner().is_mutable() {
            return Err(Error::Unsupported(
                "ODS common data styles are read-only".to_string(),
            ));
        }
        self.source_slice(&self.range, "data-style source range")?;
        patch_metadata(self, patch, max_output)
    }
}

/// Scan one XML owner using hard default limits.
#[cfg(test)]
pub(crate) fn scan(source: &[u8], owner: Owner) -> Result<Catalog<'_>> {
    scan_with_limits(source, owner, Limits::default())
}

/// Scan one XML owner with caller limits intersected with hard ceilings.
pub(crate) fn scan_with_limits(source: &[u8], owner: Owner, limits: Limits) -> Result<Catalog<'_>> {
    let limits = limits.bounded();
    if source.len() > limits.owner_bytes {
        return resource(
            Resource::InputBytes,
            source.len(),
            limits.owner_bytes,
            "ODS data-style owner exceeds its byte limit",
        );
    }
    let xml = str::from_utf8(source)
        .map_err(|_| invalid_error("ODS data-style owner is not UTF-8 XML"))?;
    let mut reader = ResolvedReader::from_xml(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_class = ElementKind::Unknown;
    let expected_container = match owner {
        Owner::ContentAutomatic => ElementKind::AutomaticStyles,
        Owner::CommonStyles => ElementKind::Styles,
        Owner::StylesAutomatic => ElementKind::AutomaticStyles,
    };
    let expected_document_root = match owner {
        Owner::ContentAutomatic => ElementKind::DocumentContent,
        Owner::CommonStyles | Owner::StylesAutomatic => ElementKind::DocumentStyles,
    };
    let mut owner_seen = false;
    let mut owner_depth = None;
    let mut owner_range = None;
    let mut active = None::<StyleCapture>;
    let mut entries = Vec::new();
    let mut element_count = 0usize;
    let mut declaration_seen = false;
    let mut prolog_misc_seen = false;

    loop {
        let event_start = checked_position(&reader)?;
        let (resolved_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let class = element_kind(&resolved_namespace, element.local_name().as_ref());
                reject_unbound_element_prefix(&resolved_namespace, element.name().prefix())?;
                let event_end = checked_position(&reader)?;
                let local_namespace_bindings =
                    validate_element_attributes(&reader, &element, &limits)?;
                validate_qname(element.name().as_ref(), "ODS data-style element name")?;
                element_count = element_count
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS style element counter overflow"))?;
                if element_count > limits.owner_elements {
                    return resource(
                        Resource::Objects,
                        element_count,
                        limits.owner_elements,
                        "ODS data-style owner exceeds its element limit",
                    );
                }
                let next_depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS data-style depth overflow"))?;
                let active_namespace_bindings = stack
                    .last()
                    .map_or(0, |frame| frame.namespace_bindings)
                    .checked_add(local_namespace_bindings)
                    .ok_or_else(|| invalid_error("ODS namespace binding counter overflow"))?;
                if active_namespace_bindings > MAX_ACTIVE_NAMESPACE_BINDINGS {
                    return resource(
                        Resource::Objects,
                        active_namespace_bindings,
                        MAX_ACTIVE_NAMESPACE_BINDINGS,
                        "ODS data-style active namespace bindings exceed their limit",
                    );
                }
                if next_depth > limits.owner_depth {
                    return resource(
                        Resource::Depth,
                        next_depth,
                        limits.owner_depth,
                        "ODS data-style owner exceeds its depth limit",
                    );
                }
                if !root_seen {
                    root_seen = true;
                    root_class = class;
                    if class != expected_document_root {
                        return invalid(format!(
                            "ODS data-style owner has an invalid document root for {owner:?}"
                        ));
                    }
                } else if root_closed {
                    return invalid("ODS data-style owner has multiple document roots");
                }
                let direct_root_child = stack.len() == 1;
                if direct_root_child && class == expected_container {
                    if owner_seen {
                        return invalid("duplicate direct data-style owner container");
                    }
                    owner_seen = true;
                    owner_depth = Some(next_depth);
                    owner_range = Some(event_start..usize::MAX);
                }
                let direct_owner_child = owner_depth == Some(stack.len());
                let style_root = direct_owner_child && is_style_root(class);
                if style_root {
                    if entries.len() >= limits.catalog_entries {
                        return resource(
                            Resource::Objects,
                            entries.len().saturating_add(1),
                            limits.catalog_entries,
                            "ODS data-style catalog exceeds its entry limit",
                        );
                    }
                    let capture = StyleCapture::new(
                        owner,
                        class,
                        event_start,
                        event_end,
                        next_depth,
                        parse_root_attributes(&reader, &element, &limits)?,
                    )?;
                    active = Some(capture);
                } else if let Some(capture) = active.as_mut() {
                    capture.event_bytes(event_start, event_end, &limits)?;
                    if stack.len() == capture.root_depth {
                        capture.start_child(
                            &reader,
                            &element,
                            class,
                            event_start,
                            event_end,
                            &limits,
                        )?;
                    } else {
                        capture.start_nested(
                            &reader,
                            &element,
                            class,
                            event_start,
                            event_end,
                            &limits,
                        )?;
                    }
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("ODS data-style element stack", source))?;
                stack.push(Frame {
                    class,
                    start: event_start,
                    namespace_bindings: active_namespace_bindings,
                });
            },
            Event::Empty(element) => {
                let class = element_kind(&resolved_namespace, element.local_name().as_ref());
                reject_unbound_element_prefix(&resolved_namespace, element.name().prefix())?;
                let event_end = checked_position(&reader)?;
                let local_namespace_bindings =
                    validate_element_attributes(&reader, &element, &limits)?;
                validate_qname(element.name().as_ref(), "ODS data-style element name")?;
                element_count = element_count
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS style element counter overflow"))?;
                if element_count > limits.owner_elements {
                    return resource(
                        Resource::Objects,
                        element_count,
                        limits.owner_elements,
                        "ODS data-style owner exceeds its element limit",
                    );
                }
                let next_depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODS data-style depth overflow"))?;
                let active_namespace_bindings = stack
                    .last()
                    .map_or(0, |frame| frame.namespace_bindings)
                    .checked_add(local_namespace_bindings)
                    .ok_or_else(|| invalid_error("ODS namespace binding counter overflow"))?;
                if active_namespace_bindings > MAX_ACTIVE_NAMESPACE_BINDINGS {
                    return resource(
                        Resource::Objects,
                        active_namespace_bindings,
                        MAX_ACTIVE_NAMESPACE_BINDINGS,
                        "ODS data-style active namespace bindings exceed their limit",
                    );
                }
                if next_depth > limits.owner_depth {
                    return resource(
                        Resource::Depth,
                        next_depth,
                        limits.owner_depth,
                        "ODS data-style owner exceeds its depth limit",
                    );
                }
                if !root_seen {
                    root_seen = true;
                    root_class = class;
                    if class != expected_document_root {
                        return invalid(format!(
                            "ODS data-style owner has an invalid document root for {owner:?}"
                        ));
                    }
                    root_closed = true;
                } else if root_closed {
                    return invalid("ODS data-style owner has multiple document roots");
                }
                let direct_root_child = stack.len() == 1;
                if direct_root_child && class == expected_container {
                    if owner_seen {
                        return invalid("duplicate direct data-style owner container");
                    }
                    owner_seen = true;
                    owner_range = Some(event_start..event_end);
                }
                let direct_owner_child = owner_depth == Some(stack.len());
                if direct_owner_child && is_style_root(class) {
                    if entries.len() >= limits.catalog_entries {
                        return resource(
                            Resource::Objects,
                            entries.len().saturating_add(1),
                            limits.catalog_entries,
                            "ODS data-style catalog exceeds its entry limit",
                        );
                    }
                    let capture = StyleCapture::new(
                        owner,
                        class,
                        event_start,
                        event_end,
                        next_depth,
                        parse_root_attributes(&reader, &element, &limits)?,
                    )?;
                    let entry = capture.finish(event_end, event_end, source)?;
                    entries
                        .try_reserve(1)
                        .map_err(|source| allocation("ODS data-style catalog", source))?;
                    entries.push(entry);
                } else if let Some(capture) = active.as_mut() {
                    capture.event_bytes(event_start, event_end, &limits)?;
                    if stack.len() == capture.root_depth {
                        let node = capture.node_from_element(
                            &reader,
                            &element,
                            class,
                            event_start,
                            event_end,
                            &limits,
                        )?;
                        capture.push_direct(node, limits.body_elements)?;
                    } else {
                        let node = capture.node_from_element(
                            &reader,
                            &element,
                            class,
                            event_start,
                            event_end,
                            &limits,
                        )?;
                        capture.push_nested(node, limits.body_elements)?;
                    }
                }
            },
            Event::End(element) => {
                let event_end = checked_position(&reader)?;
                validate_qname(element.name().as_ref(), "ODS data-style end name")?;
                let frame_class = stack
                    .last()
                    .ok_or_else(|| invalid_error("ODS data-style XML has an unmatched end"))?
                    .class;
                if let Some(capture) = active.as_mut() {
                    if stack.len() > capture.root_depth {
                        capture.event_bytes(event_start, event_end, &limits)?;
                        capture.end_nested()?;
                    } else if stack.len() == capture.root_depth {
                        let capture = active.take().ok_or_else(|| {
                            invalid_error("ODS data-style active capture disappeared")
                        })?;
                        let entry = capture.finish(event_end, event_start, source)?;
                        entries
                            .try_reserve(1)
                            .map_err(|source| allocation("ODS data-style catalog", source))?;
                        entries.push(entry);
                    }
                }
                if frame_class == root_class && stack.len() == 1 {
                    root_closed = true;
                }
                if let Some(depth) = owner_depth {
                    if stack.len() == depth {
                        if let Some(range) = owner_range.as_mut() {
                            range.end = event_end;
                        }
                        owner_depth = None;
                    }
                }
                stack.pop();
                let _ = element;
            },
            Event::Text(text) => {
                let event_end = checked_position(&reader)?;
                if let Some(capture) = active.as_mut() {
                    capture.event_bytes(event_start, event_end, &limits)?;
                }
                validate_raw_text(text.as_ref(), "ODS data-style text")?;
                let decoded = text.xml_content(XmlVersion::Implicit1_0).map_err(|error| {
                    invalid_error(format!("invalid ODS data-style text: {error}"))
                })?;
                validate_xml_characters(decoded.as_ref(), "ODS data-style text")?;
                if !root_seen {
                    prolog_misc_seen = true;
                }
                if let Some(capture) = active.as_mut() {
                    capture.text(&decoded, &limits)?;
                } else if owner_depth == Some(stack.len())
                    && !is_xml_whitespace_only(decoded.as_ref())
                {
                    return invalid("ODS data-style owner container contains non-whitespace text");
                } else if root_closed || !root_seen {
                    if !is_xml_whitespace_only(decoded.as_ref()) {
                        return invalid(
                            "ODS data-style owner has non-whitespace text outside its root",
                        );
                    }
                }
            },
            Event::CData(text) => {
                let event_end = checked_position(&reader)?;
                if let Some(capture) = active.as_mut() {
                    capture.event_bytes(event_start, event_end, &limits)?;
                }
                validate_raw_text(text.as_ref(), "ODS data-style CDATA")?;
                let decoded = text.xml_content(XmlVersion::Implicit1_0).map_err(|error| {
                    invalid_error(format!("invalid ODS data-style CDATA: {error}"))
                })?;
                validate_xml_characters(decoded.as_ref(), "ODS data-style CDATA")?;
                if let Some(capture) = active.as_mut() {
                    capture.cdata(&decoded, &limits)?;
                } else if owner_depth == Some(stack.len())
                    && !is_xml_whitespace_only(decoded.as_ref())
                {
                    return invalid("ODS data-style owner container contains non-whitespace CDATA");
                } else if root_closed || !root_seen {
                    return invalid("ODS data-style owner has CDATA outside its root");
                }
            },
            Event::GeneralRef(reference) => {
                let event_end = checked_position(&reader)?;
                if let Some(capture) = active.as_mut() {
                    capture.event_bytes(event_start, event_end, &limits)?;
                }
                validate_general_ref(&reference)?;
                if let Some(capture) = active.as_mut() {
                    capture.reference(&reference, &limits)?;
                } else if owner_depth == Some(stack.len()) {
                    let whitespace = reference
                        .resolve_char_ref()
                        .map_err(|error| {
                            invalid_error(format!("invalid ODS data-style reference: {error}"))
                        })?
                        .is_some_and(super::is_xml_whitespace);
                    if !whitespace {
                        return invalid(
                            "ODS data-style owner container contains a character reference",
                        );
                    }
                } else if root_closed || !root_seen {
                    return invalid("ODS data-style owner has a reference outside its root");
                }
            },
            Event::Comment(comment) => {
                let event_end = checked_position(&reader)?;
                validate_comment(comment.as_ref())?;
                if !root_seen {
                    prolog_misc_seen = true;
                }
                if let Some(capture) = active.as_mut() {
                    capture.event_bytes(event_start, event_end, &limits)?;
                    capture.opaque = true;
                }
            },
            Event::PI(processing_instruction) => {
                let event_end = checked_position(&reader)?;
                validate_processing_instruction(&processing_instruction)?;
                if !root_seen {
                    prolog_misc_seen = true;
                }
                if let Some(capture) = active.as_mut() {
                    capture.event_bytes(event_start, event_end, &limits)?;
                    capture.opaque = true;
                }
            },
            Event::Decl(declaration) => {
                if root_seen || declaration_seen || prolog_misc_seen {
                    return invalid("ODS data-style XML declaration is out of order");
                }
                validate_xml_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return invalid("ODS data-style document type declarations are unsupported");
            },
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return invalid("ODS data-style owner XML root is incomplete");
    }
    if let Some(range) = owner_range.as_mut()
        && range.end == usize::MAX
    {
        range.end = source.len();
    }
    let mut identities = BTreeSet::<(Family, String)>::new();
    for entry in &entries {
        let identity = (entry.family(), entry.name().to_owned());
        if !identities.insert(identity) {
            return invalid("duplicate data-style name within one source owner and family");
        }
    }
    Ok(Catalog {
        source,
        owner,
        owner_range,
        entries,
    })
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ElementKind {
    DocumentContent,
    DocumentStyles,
    AutomaticStyles,
    Styles,
    NumberStyle,
    DateStyle,
    TimeStyle,
    CurrencyStyle,
    PercentageStyle,
    BooleanStyle,
    TextStyle,
    Number,
    Scientific,
    Fraction,
    EmbeddedText,
    Text,
    TextWithFill,
    FillCharacter,
    CurrencySymbol,
    Year,
    Month,
    Day,
    Hours,
    Minutes,
    Seconds,
    Boolean,
    StyleMap,
    Unknown,
}

impl ElementKind {
    const fn family(self) -> Option<Family> {
        match self {
            Self::NumberStyle => Some(Family::Number),
            Self::DateStyle => Some(Family::Date),
            Self::TimeStyle => Some(Family::Time),
            Self::CurrencyStyle => Some(Family::Currency),
            Self::PercentageStyle => Some(Family::Percentage),
            Self::BooleanStyle => Some(Family::Boolean),
            Self::TextStyle => Some(Family::Text),
            _ => None,
        }
    }
}

fn element_kind(namespace: &ResolveResult<'_>, local: &[u8]) -> ElementKind {
    if namespace_matches(namespace, OFFICE) {
        return match local {
            b"document-content" => ElementKind::DocumentContent,
            b"document-styles" => ElementKind::DocumentStyles,
            b"automatic-styles" => ElementKind::AutomaticStyles,
            b"styles" => ElementKind::Styles,
            _ => ElementKind::Unknown,
        };
    }
    if namespace_matches(namespace, NUMBER) {
        return match local {
            b"number-style" => ElementKind::NumberStyle,
            b"date-style" => ElementKind::DateStyle,
            b"time-style" => ElementKind::TimeStyle,
            b"currency-style" => ElementKind::CurrencyStyle,
            b"percentage-style" => ElementKind::PercentageStyle,
            b"boolean-style" => ElementKind::BooleanStyle,
            b"text-style" => ElementKind::TextStyle,
            b"number" => ElementKind::Number,
            b"scientific-number" => ElementKind::Scientific,
            b"fraction" => ElementKind::Fraction,
            b"embedded-text" => ElementKind::EmbeddedText,
            b"text" => ElementKind::Text,
            b"text-with-fillchar" => ElementKind::TextWithFill,
            b"fill-character" => ElementKind::FillCharacter,
            b"currency-symbol" => ElementKind::CurrencySymbol,
            b"year" => ElementKind::Year,
            b"month" => ElementKind::Month,
            b"day" => ElementKind::Day,
            b"hours" => ElementKind::Hours,
            b"minutes" => ElementKind::Minutes,
            b"seconds" => ElementKind::Seconds,
            b"boolean" => ElementKind::Boolean,
            _ => ElementKind::Unknown,
        };
    }
    if namespace_matches(namespace, STYLE) && local == b"map" {
        return ElementKind::StyleMap;
    }
    ElementKind::Unknown
}

const fn is_style_root(kind: ElementKind) -> bool {
    matches!(
        kind,
        ElementKind::NumberStyle
            | ElementKind::DateStyle
            | ElementKind::TimeStyle
            | ElementKind::CurrencyStyle
            | ElementKind::PercentageStyle
            | ElementKind::BooleanStyle
            | ElementKind::TextStyle
    )
}

fn bound_namespace<'a>(namespace: &ResolveResult<'a>) -> Option<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(value),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

fn normalized_bound_namespace<'a>(namespace: &ResolveResult<'a>) -> Option<Cow<'a, [u8]>> {
    // `ResolvedReader` has already applied XML attribute-value normalization
    // before adding declaration values to its resolver.  Re-decoding here
    // would turn a literal entity-looking URI such as `urn:a&amp;amp;` into
    // the wrong namespace on the second pass.
    Some(Cow::Borrowed(bound_namespace(namespace)?))
}

fn namespace_matches(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    normalized_bound_namespace(namespace).is_some_and(|value| value.as_ref() == expected)
}

fn namespace_value_matches(value: &[u8], expected: &[u8]) -> bool {
    value == expected
}

fn in_scope_prefix(reader: &ResolvedReader<'_>, expected: &[u8]) -> Result<Option<Vec<u8>>> {
    let mut selected = None::<&[u8]>;
    for (declaration, Namespace(value)) in reader.resolver().bindings() {
        let PrefixDeclaration::Named(prefix) = declaration else {
            continue;
        };
        if namespace_value_matches(value, expected)
            && selected.is_none_or(|current| {
                prefix.len() < current.len() || (prefix.len() == current.len() && prefix < current)
            })
        {
            selected = Some(prefix);
        }
    }
    selected
        .map(|prefix| bounded_bytes(prefix, "ODS data-style namespace prefix"))
        .transpose()
}

fn reject_unbound_element_prefix(
    namespace: &ResolveResult<'_>,
    prefix: Option<Prefix<'_>>,
) -> Result<()> {
    if let Some(prefix) = prefix {
        if matches!(namespace, ResolveResult::Unknown(_)) {
            return invalid("ODS data-style element has an unbound namespace prefix");
        }
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.is_empty()) {
            return invalid(format!(
                "ODS data-style element prefix '{}' is bound to an empty namespace",
                String::from_utf8_lossy(prefix.as_ref())
            ));
        }
    }
    Ok(())
}

fn reject_unbound_attribute_prefix(
    namespace: &ResolveResult<'_>,
    prefix: Option<&[u8]>,
) -> Result<()> {
    if let Some(prefix) = prefix {
        if matches!(namespace, ResolveResult::Unknown(_)) {
            return invalid("ODS data-style attribute has an unbound namespace prefix");
        }
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.is_empty()) {
            return invalid(format!(
                "ODS data-style attribute prefix '{}' is bound to an empty namespace",
                String::from_utf8_lossy(prefix)
            ));
        }
    }
    Ok(())
}

fn validate_element_attributes(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
    limits: &Limits,
) -> Result<usize> {
    let mut expanded = BTreeSet::<Vec<u8>>::new();
    let mut namespace_declarations = 0usize;
    let mut attributes = element.attributes();
    attributes.with_checks(true);
    for attribute in attributes {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid ODS data-style attribute: {error}")))?;
        if attribute.value.as_ref().len() > limits.text_bytes {
            return resource(
                Resource::InputBytes,
                attribute.value.as_ref().len(),
                limits.text_bytes,
                "ODS data-style attribute exceeds its text limit",
            );
        }
        let key = attribute.key;
        let key_bytes = key.as_ref();
        validate_qname(key_bytes, "ODS data-style attribute name")?;
        if key_bytes == b"xmlns" || key_bytes.starts_with(b"xmlns:") {
            namespace_declarations = namespace_declarations
                .checked_add(1)
                .ok_or_else(|| invalid_error("ODS namespace declaration counter overflow"))?;
            validate_namespace_declaration(reader, &attribute, key_bytes)?;
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(key);
        reject_unbound_attribute_prefix(&namespace, qname_prefix(key_bytes))?;
        let namespace_bytes = normalized_bound_namespace(&namespace).unwrap_or(Cow::Borrowed(&[]));
        let mut identity = Vec::new();
        identity
            .try_reserve(
                namespace_bytes
                    .len()
                    .saturating_add(local.as_ref().len())
                    .saturating_add(1),
            )
            .map_err(|source| allocation("ODS style attribute identity", source))?;
        identity.extend_from_slice(namespace_bytes.as_ref());
        identity.push(0);
        identity.extend_from_slice(local.as_ref());
        if !expanded.insert(identity) {
            return invalid("duplicate ODF data-style attribute");
        }
        validate_raw_attribute_value(&attribute)?;
    }
    Ok(namespace_declarations)
}

#[allow(deprecated)]
fn validate_namespace_declaration(
    reader: &ResolvedReader<'_>,
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    key: &[u8],
) -> Result<()> {
    let prefix = key.strip_prefix(b"xmlns:");
    if let Some(prefix) = prefix {
        if prefix.is_empty() {
            return invalid("ODS data-style namespace declaration has an empty prefix");
        }
        if prefix.len() > MAX_STYLE_TEXT_BYTES {
            return resource(
                Resource::InputBytes,
                prefix.len(),
                MAX_STYLE_TEXT_BYTES,
                "ODS data-style namespace prefix exceeds its limit",
            );
        }
    }
    let value = attribute
        .decode_and_unescape_value_with(reader.decoder(), resolve_predefined_entity)
        .map_err(|error| invalid_error(format!("invalid ODS namespace declaration: {error}")))?;
    validate_xml_characters(value.as_ref(), "ODS data-style namespace declaration")?;
    if value.len() > MAX_STYLE_TEXT_BYTES {
        return resource(
            Resource::InputBytes,
            value.len(),
            MAX_STYLE_TEXT_BYTES,
            "ODS data-style namespace declaration exceeds its limit",
        );
    }
    let value = value.as_bytes();
    match prefix {
        None => {
            if value == XML || value == XMLNS {
                return invalid("ODS default namespace uses a reserved namespace URI");
            }
        },
        Some(b"xml") if value == XML => {},
        Some(b"xml") => {
            return invalid("ODS xml prefix is bound to the wrong namespace URI");
        },
        Some(b"xmlns") => {
            return invalid("ODS xmlns prefix declaration is reserved");
        },
        Some(prefix) => {
            if value.is_empty() {
                return invalid("ODS prefixed namespace declaration has an empty URI");
            }
            if value == XML || value == XMLNS {
                return invalid("ODS namespace prefix uses a reserved namespace URI");
            }
            if prefix == b"xml" {
                return invalid("ODS xml prefix is bound to the wrong namespace URI");
            }
        },
    }
    Ok(())
}

fn validate_raw_attribute_value(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<()> {
    let raw = attribute.value.as_ref();
    if raw.contains(&b'<') {
        return invalid("ODS data-style attribute contains a raw '<'");
    }
    let value =
        str::from_utf8(raw).map_err(|_| invalid_error("ODS data-style attribute is not UTF-8"))?;
    let bytes = value.as_bytes();
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        if bytes[cursor] != b'&' {
            let character = value[cursor..]
                .chars()
                .next()
                .ok_or_else(|| invalid_error("ODS data-style attribute has invalid UTF-8"))?;
            validate_xml_character(character, "ODS data-style attribute")?;
            cursor += character.len_utf8();
            continue;
        }
        let end = bytes[cursor + 1..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|offset| cursor + 1 + offset)
            .ok_or_else(|| {
                invalid_error("ODS data-style attribute has an unterminated reference")
            })?;
        if end == cursor + 1 {
            return invalid("ODS data-style attribute has an empty reference");
        }
        let name = &value[cursor + 1..end];
        if let Some(name) = name.strip_prefix("#x").or_else(|| name.strip_prefix("#X")) {
            let code = u32::from_str_radix(name, 16).map_err(|_| {
                invalid_error("ODS data-style attribute has an invalid character reference")
            })?;
            let character = char::from_u32(code).ok_or_else(|| {
                invalid_error("ODS data-style attribute has an invalid character reference")
            })?;
            validate_xml_character(character, "ODS data-style attribute")?;
        } else if let Some(name) = name.strip_prefix('#') {
            let code = name.parse::<u32>().map_err(|_| {
                invalid_error("ODS data-style attribute has an invalid character reference")
            })?;
            let character = char::from_u32(code).ok_or_else(|| {
                invalid_error("ODS data-style attribute has an invalid character reference")
            })?;
            validate_xml_character(character, "ODS data-style attribute")?;
        } else if resolve_predefined_entity(name).is_none() {
            return invalid("ODS data-style attribute has an unsupported named entity");
        }
        cursor = end + 1;
    }
    Ok(())
}

fn validate_general_ref(reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
    if resolve_general_ref(reference)?.is_some() {
        return Ok(());
    }
    invalid("ODS data-style named entity reference is unsupported")
}

fn validate_xml_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let raw = declaration.as_ref();
    let raw = str::from_utf8(raw)
        .map_err(|_| invalid_error("ODS data-style XML declaration is not UTF-8"))?;
    if !raw.starts_with("xml") {
        return invalid("ODS data-style XML declaration has an invalid target");
    }
    let start = BytesStart::from_content(raw, 3);
    let mut attributes = start.attributes();
    attributes.with_checks(true);
    let mut stage = 0u8;
    let mut count = 0usize;
    for attribute in attributes {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid ODS XML declaration: {error}")))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS XML declaration attribute counter overflow"))?;
        let key = attribute.key.as_ref();
        let value = attribute.value.as_ref();
        if value.contains(&b'&') {
            return invalid("ODS XML declaration contains an entity reference");
        }
        match (stage, key) {
            (0, b"version") => {
                if value != b"1.0" && value != b"1.1" {
                    return invalid("ODS XML declaration has an unsupported version");
                }
                stage = 1;
            },
            (0, _) => return invalid("ODS XML declaration version is not first"),
            (1, b"encoding") => {
                if !valid_encoding_name(value)
                    || !(value.eq_ignore_ascii_case(b"UTF-8")
                        || value.eq_ignore_ascii_case(b"US-ASCII"))
                {
                    return invalid("ODS XML declaration has an unsupported encoding");
                }
                stage = 2;
            },
            (1, b"standalone") => {
                if value != b"yes" && value != b"no" {
                    return invalid("ODS XML declaration has an invalid standalone value");
                }
                stage = 3;
            },
            (2, b"standalone") => {
                if value != b"yes" && value != b"no" {
                    return invalid("ODS XML declaration has an invalid standalone value");
                }
                stage = 3;
            },
            (_, b"version") | (_, b"encoding") | (_, b"standalone") => {
                return invalid("ODS XML declaration contains a duplicate or out-of-order field");
            },
            _ => return invalid("ODS XML declaration contains an unknown field"),
        }
    }
    if count == 0 || stage == 0 {
        return invalid("ODS XML declaration is missing version");
    }
    Ok(())
}

fn valid_encoding_name(value: &[u8]) -> bool {
    let Some(first) = value.first().copied() else {
        return false;
    };
    if !first.is_ascii_alphabetic() {
        return false;
    }
    value[1..]
        .iter()
        .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn validate_comment(value: &[u8]) -> Result<()> {
    if value.windows(2).any(|window| window == b"--") || value.ends_with(b"-") {
        return invalid("ODS data-style comment has an invalid '--' sequence");
    }
    let value =
        str::from_utf8(value).map_err(|_| invalid_error("ODS data-style comment is not UTF-8"))?;
    validate_xml_characters(value, "ODS data-style comment")
}

fn validate_processing_instruction(instruction: &BytesPI<'_>) -> Result<()> {
    let target = instruction.target();
    if target.is_empty() || !valid_xml_name(target) {
        return invalid("ODS data-style processing instruction has an invalid target");
    }
    if target.eq_ignore_ascii_case(b"xml") {
        return invalid("ODS data-style processing instruction uses the reserved xml target");
    }
    let content = instruction.content();
    if !content.is_empty() && !content[0].is_ascii_whitespace() {
        return invalid("ODS data-style processing instruction lacks target whitespace");
    }
    if content.windows(2).any(|window| window == b"?>") {
        return invalid("ODS data-style processing instruction contains a raw ?> sequence");
    }
    let target = str::from_utf8(target)
        .map_err(|_| invalid_error("ODS data-style processing-instruction target is not UTF-8"))?;
    validate_xml_characters(target, "ODS data-style processing-instruction target")?;
    let content = str::from_utf8(content)
        .map_err(|_| invalid_error("ODS data-style processing instruction is not UTF-8"))?;
    validate_xml_characters(content, "ODS data-style processing instruction")
}

fn resolve_general_ref(reference: &quick_xml::events::BytesRef<'_>) -> Result<Option<char>> {
    if let Some(value) = reference
        .resolve_char_ref()
        .map_err(|error| invalid_error(format!("invalid ODS data-style reference: {error}")))?
    {
        validate_xml_character(value, "ODS data-style character reference")?;
        return Ok(Some(value));
    }
    let name: &[u8] = reference.as_ref();
    let value = match name {
        b"amp" => Some('&'),
        b"lt" => Some('<'),
        b"gt" => Some('>'),
        b"apos" => Some('\''),
        b"quot" => Some('"'),
        _ => None,
    };
    if let Some(value) = value {
        validate_xml_character(value, "ODS data-style character reference")?;
    }
    Ok(value)
}

fn validate_qname(value: &[u8], label: &str) -> Result<()> {
    let mut parts = value.split(|byte| *byte == b':');
    let first = parts
        .next()
        .ok_or_else(|| invalid_error(format!("{label} is empty")))?;
    let second = parts.next();
    if parts.next().is_some() || first.is_empty() || second.is_some_and(|part| part.is_empty()) {
        return invalid(format!("{label} is not a valid QName"));
    }
    if !valid_ncname(first) || second.is_some_and(|part| !valid_ncname(part)) {
        return invalid(format!("{label} is not a valid QName"));
    }
    Ok(())
}

fn valid_ncname(value: &[u8]) -> bool {
    let Ok(value) = str::from_utf8(value) else {
        return false;
    };
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !valid_ncname_start(first) {
        return false;
    }
    chars.all(valid_ncname_char)
}

fn valid_xml_name(value: &[u8]) -> bool {
    let Ok(value) = str::from_utf8(value) else {
        return false;
    };
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if first != ':' && !valid_ncname_start(first) {
        return false;
    }
    chars.all(|character| character == ':' || valid_ncname_char(character))
}

fn valid_ncname_start(value: char) -> bool {
    matches!(value, 'A'..='Z' | 'a'..='z' | '_')
        || matches!(
            value as u32,
            0xC0..=0xD6
                | 0xD8..=0xF6
                | 0xF8..=0x2FF
                | 0x370..=0x37D
                | 0x37F..=0x1FFF
                | 0x200C..=0x200D
                | 0x2070..=0x218F
                | 0x2C00..=0x2FEF
                | 0x3001..=0xD7FF
                | 0xF900..=0xFDCF
                | 0xFDF0..=0xFFFD
                | 0x10000..=0xEFFFF
        )
}

fn valid_ncname_char(value: char) -> bool {
    valid_ncname_start(value)
        || matches!(value, '-' | '.' | '0'..='9')
        || matches!(value as u32, 0xB7 | 0x300..=0x36F | 0x203F..=0x2040)
}

fn is_xml_whitespace_only(value: &str) -> bool {
    value.chars().all(super::is_xml_whitespace)
}

#[derive(Clone, Copy, Debug)]
struct Frame {
    class: ElementKind,
    #[allow(dead_code)]
    start: usize,
    namespace_bindings: usize,
}

#[derive(Clone, Debug, Default)]
struct AttributeSlots {
    slots: Vec<AttributeSlot>,
    style_prefix: Option<Vec<u8>>,
    number_prefix: Option<Vec<u8>>,
    // The binding may be declared on the root or inherited from an owner
    // ancestor.  Keep this as availability rather than declaration locality:
    // metadata insertion must not add a redundant xmlns when the alias is
    // already in scope.
    style_binding_available: bool,
    number_binding_available: bool,
}

#[derive(Clone, Debug)]
struct AttributeSlot {
    field: Field,
    key: Vec<u8>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Field {
    DisplayName,
    Language,
    Country,
    Script,
    RfcLanguageTag,
    Title,
    Volatile,
    TransliterationFormat,
    TransliterationLanguage,
    TransliterationCountry,
    TransliterationStyle,
}

#[derive(Clone, Debug, Default)]
struct ParsedRoot {
    name: String,
    attributes: Attributes,
    slots: AttributeSlots,
    opaque: bool,
}

fn parse_root_attributes(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
    limits: &Limits,
) -> Result<ParsedRoot> {
    let mut parsed = parse_attributes(reader, element, AttributeContext::Root, limits)?;
    let Some(name) = parsed.name else {
        return invalid("ODS data-style root is missing style:name");
    };
    if parsed.slots.style_prefix.is_none() {
        if let Some(prefix) = in_scope_prefix(reader, STYLE)? {
            parsed.slots.style_prefix = Some(prefix);
        }
    }
    if parsed.slots.style_prefix.is_some() {
        // The prefix may be inherited from the owner/container.  This flag
        // means "already bound in scope", not "declared on this root".
        parsed.slots.style_binding_available = true;
    }
    if parsed.slots.number_prefix.is_none() {
        if let Some(prefix) = in_scope_prefix(reader, NUMBER)? {
            parsed.slots.number_prefix = Some(prefix);
        }
    }
    if parsed.slots.number_prefix.is_some() {
        parsed.slots.number_binding_available = true;
    }
    let attributes = Attributes {
        display_name: parsed.display_name,
        language: parsed.language,
        country: parsed.country,
        script: parsed.script,
        rfc_language_tag: parsed.rfc_language_tag,
        title: parsed.title,
        volatile: parsed.volatile,
        transliteration: Transliteration {
            format: parsed.transliteration_format,
            language: parsed.transliteration_language,
            country: parsed.transliteration_country,
            style: parsed.transliteration_style,
        },
    };
    attributes.validate()?;
    let metadata_bytes = metadata_size(&attributes)?;
    if metadata_bytes > limits.metadata_bytes {
        return resource(
            Resource::InputBytes,
            metadata_bytes,
            limits.metadata_bytes,
            "ODS data-style metadata exceeds its byte limit",
        );
    }
    Ok(ParsedRoot {
        name,
        attributes,
        slots: parsed.slots,
        opaque: parsed.opaque,
    })
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum AttributeContext {
    Root,
}

#[derive(Clone, Debug, Default)]
struct ParsedAttributes {
    name: Option<String>,
    display_name: Option<String>,
    language: Option<String>,
    country: Option<String>,
    script: Option<String>,
    rfc_language_tag: Option<String>,
    title: Option<String>,
    volatile: Option<bool>,
    transliteration_format: Option<String>,
    transliteration_language: Option<String>,
    transliteration_country: Option<String>,
    transliteration_style: Option<TransliterationStyle>,
    slots: AttributeSlots,
    opaque: bool,
}

#[derive(Clone, Debug)]
struct NodeCapture {
    class: ElementKind,
    fields: NumberFields,
    text: String,
    children: Vec<NodeCapture>,
    opaque: bool,
    empty: bool,
}

impl NodeCapture {
    fn new(
        reader: &ResolvedReader<'_>,
        element: &BytesStart<'_>,
        class: ElementKind,
        limits: &Limits,
        empty: bool,
    ) -> Result<Self> {
        let (fields, opaque) = parse_body_attributes(reader, element, class, limits)?;
        Ok(Self {
            class,
            fields,
            text: String::new(),
            children: Vec::new(),
            opaque,
            empty,
        })
    }

    fn append_text(&mut self, value: &str, limits: &Limits) -> Result<()> {
        if !matches!(
            self.class,
            ElementKind::Text
                | ElementKind::FillCharacter
                | ElementKind::CurrencySymbol
                | ElementKind::EmbeddedText
        ) {
            if !is_xml_whitespace_only(value) {
                self.opaque = true;
            }
            return Ok(());
        }
        let next = self
            .text
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid_error("ODS data-style text size overflow"))?;
        if next > limits.text_bytes {
            return resource(
                Resource::InputBytes,
                next,
                limits.text_bytes,
                "ODS data-style text exceeds its limit",
            );
        }
        self.text
            .try_reserve(value.len())
            .map_err(|source| allocation("ODS data-style text", source))?;
        self.text.push_str(value);
        Ok(())
    }
}

struct StyleCapture {
    owner: Owner,
    class: ElementKind,
    name: String,
    attributes: Attributes,
    root_open_end: usize,
    root_start: usize,
    root_depth: usize,
    body_bytes: usize,
    body_elements: usize,
    opaque: bool,
    root_slots: AttributeSlots,
    children: Vec<NodeCapture>,
    node_stack: Vec<NodeCapture>,
}

impl StyleCapture {
    fn new(
        owner: Owner,
        class: ElementKind,
        root_start: usize,
        root_open_end: usize,
        root_depth: usize,
        root: ParsedRoot,
    ) -> Result<Self> {
        super::validate_name(&root.name, "data-style name")?;
        let mut children = Vec::new();
        children
            .try_reserve(1)
            .map_err(|source| allocation("ODS data-style body children", source))?;
        Ok(Self {
            owner,
            class,
            name: root.name,
            attributes: root.attributes,
            root_open_end,
            root_start,
            root_depth,
            body_bytes: 0,
            body_elements: 0,
            opaque: root.opaque,
            root_slots: root.slots,
            children,
            node_stack: Vec::new(),
        })
    }

    fn event_bytes(&mut self, start: usize, end: usize, limits: &Limits) -> Result<()> {
        let bytes = end.saturating_sub(start);
        self.body_bytes = self
            .body_bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid_error("ODS data-style body size overflow"))?;
        if self.body_bytes > limits.body_bytes {
            return resource(
                Resource::InputBytes,
                self.body_bytes,
                limits.body_bytes,
                "ODS data-style body exceeds its byte limit",
            );
        }
        Ok(())
    }

    fn ensure_node_capacity(&self, limit: usize) -> Result<()> {
        if self.body_elements >= limit {
            return resource(
                Resource::Objects,
                self.body_elements.saturating_add(1),
                limit,
                "ODS data-style body exceeds its element limit",
            );
        }
        Ok(())
    }

    fn start_child(
        &mut self,
        reader: &ResolvedReader<'_>,
        element: &BytesStart<'_>,
        class: ElementKind,
        _start: usize,
        _end: usize,
        limits: &Limits,
    ) -> Result<()> {
        self.ensure_node_capacity(limits.body_elements)?;
        self.push_node(
            NodeCapture::new(reader, element, class, limits, false)?,
            limits.body_elements,
        )
    }

    fn start_nested(
        &mut self,
        reader: &ResolvedReader<'_>,
        element: &BytesStart<'_>,
        class: ElementKind,
        _start: usize,
        _end: usize,
        limits: &Limits,
    ) -> Result<()> {
        self.ensure_node_capacity(limits.body_elements)?;
        self.push_node(
            NodeCapture::new(reader, element, class, limits, false)?,
            limits.body_elements,
        )
    }

    fn node_from_element(
        &self,
        reader: &ResolvedReader<'_>,
        element: &BytesStart<'_>,
        class: ElementKind,
        _start: usize,
        _end: usize,
        limits: &Limits,
    ) -> Result<NodeCapture> {
        self.ensure_node_capacity(limits.body_elements)?;
        NodeCapture::new(reader, element, class, limits, true)
    }

    fn push_node(&mut self, node: NodeCapture, limit: usize) -> Result<()> {
        self.body_elements = self
            .body_elements
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS data-style body element counter overflow"))?;
        if self.body_elements > limit {
            return resource(
                Resource::Objects,
                self.body_elements,
                limit,
                "ODS data-style body exceeds its element limit",
            );
        }
        self.node_stack
            .try_reserve(1)
            .map_err(|source| allocation("ODS data-style body stack", source))?;
        self.node_stack.push(node);
        Ok(())
    }

    fn push_direct(&mut self, node: NodeCapture, limit: usize) -> Result<()> {
        self.body_elements = self
            .body_elements
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS data-style body element counter overflow"))?;
        if self.body_elements > limit {
            return resource(
                Resource::Objects,
                self.body_elements,
                limit,
                "ODS data-style body exceeds its element limit",
            );
        }
        self.children
            .try_reserve(1)
            .map_err(|source| allocation("ODS data-style body children", source))?;
        self.children.push(node);
        Ok(())
    }

    fn push_nested(&mut self, node: NodeCapture, limit: usize) -> Result<()> {
        self.body_elements = self
            .body_elements
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODS data-style body element counter overflow"))?;
        if self.body_elements > limit {
            return resource(
                Resource::Objects,
                self.body_elements,
                limit,
                "ODS data-style body exceeds its element limit",
            );
        }
        if let Some(parent) = self.node_stack.last_mut() {
            parent
                .children
                .try_reserve(1)
                .map_err(|source| allocation("ODS data-style body children", source))?;
            parent.children.push(node);
        } else {
            self.children
                .try_reserve(1)
                .map_err(|source| allocation("ODS data-style body children", source))?;
            self.children.push(node);
        }
        Ok(())
    }

    fn end_nested(&mut self) -> Result<()> {
        let node = self
            .node_stack
            .pop()
            .ok_or_else(|| invalid_error("ODS data-style body stack underflow"))?;
        if let Some(parent) = self.node_stack.last_mut() {
            parent
                .children
                .try_reserve(1)
                .map_err(|source| allocation("ODS data-style body children", source))?;
            parent.children.push(node);
        } else {
            self.children
                .try_reserve(1)
                .map_err(|source| allocation("ODS data-style body children", source))?;
            self.children.push(node);
        }
        Ok(())
    }

    fn text(&mut self, value: &str, limits: &Limits) -> Result<()> {
        if let Some(node) = self.node_stack.last_mut() {
            node.append_text(value, limits)
        } else if !is_xml_whitespace_only(value) {
            self.opaque = true;
            Ok(())
        } else {
            Ok(())
        }
    }

    fn cdata(&mut self, value: &str, limits: &Limits) -> Result<()> {
        self.opaque = true;
        self.text(value, limits)
    }

    fn reference(
        &mut self,
        reference: &quick_xml::events::BytesRef<'_>,
        limits: &Limits,
    ) -> Result<()> {
        let value = resolve_general_ref(reference)?;
        let Some(value) = value else {
            return invalid("ODS data-style named entity reference is unsupported");
        };
        validate_xml_character(value, "ODS data-style character reference")?;
        let mut text = [0u8; 4];
        let value = value.encode_utf8(&mut text);
        self.text(value, limits)
    }

    fn finish<'a>(
        self,
        end: usize,
        close_start: usize,
        source: &'a [u8],
    ) -> Result<SourceEntry<'a>> {
        let entry = if self.opaque {
            opaque_entry(self.owner, self.class, self.name, self.attributes)?
        } else {
            typed_entry(
                self.owner,
                self.class,
                self.name,
                self.attributes,
                &self.children,
            )?
        };
        Ok(SourceEntry {
            source,
            entry,
            range: self.root_start..end,
            opening_range: self.root_start..self.root_open_end,
            body_range: self.root_open_end..close_start,
            attributes: self.root_slots,
        })
    }
}

fn opaque_entry(
    owner: Owner,
    class: ElementKind,
    name: String,
    attributes: Attributes,
) -> Result<Entry> {
    let family = class
        .family()
        .ok_or_else(|| invalid_error("ODS data-style root has no family"))?;
    let value = Opaque {
        owner,
        name,
        family,
        attributes,
        body: OpaqueBody::Preserved,
    };
    value.validate()?;
    if family == Family::Text {
        return Ok(Entry::Text(TextEntry {
            owner,
            name: value.name,
            attributes: value.attributes,
            body: OpaqueBody::Preserved,
        }));
    }
    Ok(Entry::Opaque(value))
}

fn typed_entry(
    owner: Owner,
    class: ElementKind,
    name: String,
    attributes: Attributes,
    children: &[NodeCapture],
) -> Result<Entry> {
    match class {
        ElementKind::NumberStyle => {
            let Some((leading, format, trailing)) = parse_number_children(children)? else {
                return opaque_entry(owner, class, name, attributes);
            };
            let value = Number {
                name,
                attributes,
                leading,
                format,
                trailing,
            };
            value.validate()?;
            Ok(Entry::number_at(owner, value))
        },
        ElementKind::DateStyle => {
            if !matches_date_children(children) {
                return opaque_entry(owner, class, name, attributes);
            }
            Ok(Entry::existing_at(
                owner,
                Data {
                    name,
                    attributes,
                    body: Body::Date,
                },
            ))
        },
        ElementKind::TimeStyle => {
            let Some(decimal_places) = matches_time_children(children) else {
                return opaque_entry(owner, class, name, attributes);
            };
            Ok(Entry::existing_at(
                owner,
                Data {
                    name,
                    attributes,
                    body: Body::Time { decimal_places },
                },
            ))
        },
        ElementKind::CurrencyStyle => {
            let Some((symbol, number)) = matches_currency_children(children)? else {
                return opaque_entry(owner, class, name, attributes);
            };
            Ok(Entry::existing_at(
                owner,
                Data {
                    name,
                    attributes,
                    body: Body::Currency { symbol, number },
                },
            ))
        },
        ElementKind::PercentageStyle => {
            let Some(number) = matches_percentage_children(children)? else {
                return opaque_entry(owner, class, name, attributes);
            };
            Ok(Entry::existing_at(
                owner,
                Data {
                    name,
                    attributes,
                    body: Body::Percentage { number },
                },
            ))
        },
        ElementKind::BooleanStyle => {
            if children.len() != 1
                || children[0].class != ElementKind::Boolean
                || !children[0].empty
            {
                return opaque_entry(owner, class, name, attributes);
            }
            Ok(Entry::existing_at(
                owner,
                Data {
                    name,
                    attributes,
                    body: Body::Boolean,
                },
            ))
        },
        ElementKind::TextStyle => opaque_entry(owner, class, name, attributes),
        _ => invalid("ODS data-style typed root family is invalid"),
    }
}

fn parse_number_children(
    children: &[NodeCapture],
) -> Result<Option<(Option<super::Affix>, Option<Format>, Option<super::Affix>)>> {
    let mut index = 0usize;
    let leading = parse_affix(children, &mut index)?;
    let format = if index == children.len() {
        None
    } else {
        let child = &children[index];
        let value = match child.class {
            ElementKind::Number => decimal_from_node(child)?.map(Format::Decimal),
            ElementKind::Scientific => scientific_from_node(child)?.map(Format::Scientific),
            ElementKind::Fraction => fraction_from_node(child)?.map(Format::Fraction),
            ElementKind::Text | ElementKind::FillCharacter => {
                return invalid("ODS number-style has adjacent text particles");
            },
            _ => return Ok(None),
        };
        index += 1;
        let Some(value) = value else {
            return Ok(None);
        };
        Some(value)
    };
    let trailing = if format.is_some() {
        parse_affix(children, &mut index)?
    } else {
        None
    };
    if index != children.len() {
        let remaining = &children[index..];
        if remaining.first().is_some_and(|node| {
            matches!(
                node.class,
                ElementKind::Text
                    | ElementKind::FillCharacter
                    | ElementKind::Number
                    | ElementKind::Scientific
                    | ElementKind::Fraction
            )
        }) {
            return invalid("ODS number-style has adjacent text particles");
        }
        return Ok(None);
    }
    if format.is_none()
        && leading.is_some()
        && children.iter().any(|node| {
            matches!(
                node.class,
                ElementKind::Number | ElementKind::Scientific | ElementKind::Fraction
            )
        })
    {
        return invalid("ODS number-style number particle order is invalid");
    }
    Ok(Some((leading, format, trailing)))
}

fn parse_affix(children: &[NodeCapture], index: &mut usize) -> Result<Option<super::Affix>> {
    let Some(first) = children.get(*index) else {
        return Ok(None);
    };
    let mut text = None;
    let mut fill = None;
    let mut text_after_fill = None;
    match first.class {
        ElementKind::Text => {
            if first.opaque {
                return Ok(None);
            }
            text = Some(first.text.as_str());
            *index += 1;
        },
        ElementKind::FillCharacter => {
            if first.opaque {
                return Ok(None);
            }
            fill = Some(first.text.as_str());
            *index += 1;
        },
        _ => return Ok(None),
    }
    if fill.is_none()
        && children
            .get(*index)
            .is_some_and(|node| node.class == ElementKind::FillCharacter)
    {
        let fill_node = &children[*index];
        if fill_node.opaque {
            return Ok(None);
        }
        fill = Some(fill_node.text.as_str());
        *index += 1;
        if children
            .get(*index)
            .is_some_and(|node| node.class == ElementKind::Text)
        {
            let text_node = &children[*index];
            if text_node.opaque {
                return Ok(None);
            }
            text_after_fill = Some(text_node.text.as_str());
            *index += 1;
        }
    } else if fill.is_some()
        && children
            .get(*index)
            .is_some_and(|node| node.class == ElementKind::Text)
    {
        let text_node = &children[*index];
        if text_node.opaque {
            return Ok(None);
        }
        text_after_fill = Some(text_node.text.as_str());
        *index += 1;
    }
    Ok(Some(super::Affix::try_from_borrowed(
        text,
        fill,
        text_after_fill,
    )?))
}

fn decimal_from_node(node: &NodeCapture) -> Result<Option<super::Decimal>> {
    if node.opaque
        || node
            .children
            .iter()
            .any(|child| child.class != ElementKind::EmbeddedText || child.opaque)
    {
        return Ok(None);
    }
    let mut value = super::Decimal {
        decimal_places: node.fields.decimal_places,
        min_decimal_places: node.fields.min_decimal_places,
        min_integer_digits: node.fields.min_integer_digits,
        grouping: node.fields.grouping,
        decimal_replacement: node.fields.decimal_replacement.clone(),
        display_factor: node.fields.display_factor.clone(),
        embedded_text: Vec::new(),
    };
    if node.children.len() > MAX_STYLE_BODY_ELEMENTS {
        return Ok(None);
    }
    value
        .embedded_text
        .try_reserve(node.children.len())
        .map_err(|source| allocation("ODS data-style embedded text", source))?;
    for child in &node.children {
        let Some(position) = child.fields.position else {
            return Ok(None);
        };
        let Ok(embedded) = EmbeddedText::new(position, &child.text) else {
            return Ok(None);
        };
        value.embedded_text.push(embedded);
    }
    value.validate()?;
    Ok(Some(value))
}

fn scientific_from_node(node: &NodeCapture) -> Result<Option<Scientific>> {
    if node.opaque || !node.children.is_empty() {
        return Ok(None);
    }
    let value = Scientific {
        decimal_places: node.fields.decimal_places,
        min_decimal_places: node.fields.min_decimal_places,
        min_integer_digits: node.fields.min_integer_digits,
        grouping: node.fields.grouping,
        min_exponent_digits: node.fields.min_exponent_digits,
        exponent_interval: node.fields.exponent_interval,
        forced_exponent_sign: node.fields.forced_exponent_sign,
    };
    value.validate()?;
    Ok(Some(value))
}

fn fraction_from_node(node: &NodeCapture) -> Result<Option<Fraction>> {
    if node.opaque || !node.children.is_empty() {
        return Ok(None);
    }
    let value = Fraction {
        min_numerator_digits: node.fields.min_numerator_digits,
        min_denominator_digits: node.fields.min_denominator_digits,
        denominator_value: node.fields.denominator_value,
        max_denominator_value: node.fields.max_denominator_value,
        min_integer_digits: node.fields.min_integer_digits,
        grouping: node.fields.grouping,
    };
    value.validate()?;
    Ok(Some(value))
}

fn matches_date_children(children: &[NodeCapture]) -> bool {
    if children.len() != 5 {
        return false;
    }
    matches!(
        (&children[0], &children[1], &children[2], &children[3], &children[4]),
        (
            NodeCapture { class: ElementKind::Year, fields: NumberFields { style: year_style, .. }, children: year_children, opaque: false, .. },
            NodeCapture { class: ElementKind::Text, text, children: text_children, opaque: false, .. },
            NodeCapture { class: ElementKind::Month, fields: NumberFields { style: month_style, .. }, children: month_children, opaque: false, .. },
            NodeCapture { class: ElementKind::Text, text: second_text, children: second_children, opaque: false, .. },
            NodeCapture { class: ElementKind::Day, fields: NumberFields { style: day_style, .. }, children: day_children, opaque: false, .. },
        ) if year_style.as_deref().is_none_or(|value| value == "long")
            && month_style.as_deref().is_none_or(|value| value == "long")
            && day_style.as_deref().is_none_or(|value| value == "long")
            && year_children.is_empty()
            && month_children.is_empty()
            && day_children.is_empty()
            && text_children.is_empty()
            && second_children.is_empty()
            && text == "-"
            && second_text == "-"
    )
}

fn matches_time_children(children: &[NodeCapture]) -> Option<Option<i64>> {
    if children.len() != 5 {
        return None;
    }
    let [hours, first, minutes, second, seconds] = children else {
        return None;
    };
    if hours.class != ElementKind::Hours
        || minutes.class != ElementKind::Minutes
        || seconds.class != ElementKind::Seconds
        || first.class != ElementKind::Text
        || second.class != ElementKind::Text
        || first.text != ":"
        || second.text != ":"
        || hours.opaque
        || minutes.opaque
        || seconds.opaque
        || !hours.children.is_empty()
        || !minutes.children.is_empty()
        || !seconds.children.is_empty()
    {
        return None;
    }
    if !matches!(hours.fields.style.as_deref(), None | Some("long"))
        || !matches!(minutes.fields.style.as_deref(), None | Some("long"))
        || !matches!(seconds.fields.style.as_deref(), None | Some("long"))
    {
        return None;
    }
    Some(seconds.fields.decimal_places)
}

fn matches_currency_children(children: &[NodeCapture]) -> Result<Option<(String, super::Decimal)>> {
    if children.len() != 2
        || children[0].class != ElementKind::CurrencySymbol
        || children[1].class != ElementKind::Number
    {
        return Ok(None);
    }
    let symbol = &children[0];
    if symbol.opaque || !symbol.children.is_empty() {
        return Ok(None);
    }
    let Some(number) = decimal_from_node(&children[1])? else {
        return Ok(None);
    };
    Ok(Some((symbol.text.clone(), number)))
}

fn matches_percentage_children(children: &[NodeCapture]) -> Result<Option<super::Decimal>> {
    if children.len() != 2
        || children[0].class != ElementKind::Number
        || children[1].class != ElementKind::Text
    {
        return Ok(None);
    }
    if children[1].text != "%" || children[1].opaque || !children[1].children.is_empty() {
        return Ok(None);
    }
    decimal_from_node(&children[0])
}

#[derive(Clone, Debug, Default)]
struct NumberFields {
    decimal_places: Option<i64>,
    min_decimal_places: Option<i64>,
    min_integer_digits: Option<i64>,
    grouping: Option<bool>,
    decimal_replacement: Option<String>,
    display_factor: Option<Double>,
    min_exponent_digits: Option<i64>,
    exponent_interval: Option<u64>,
    forced_exponent_sign: Option<bool>,
    min_numerator_digits: Option<i64>,
    min_denominator_digits: Option<i64>,
    denominator_value: Option<i64>,
    max_denominator_value: Option<u64>,
    position: Option<i64>,
    style: Option<String>,
}

fn parse_attributes(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
    context: AttributeContext,
    limits: &Limits,
) -> Result<ParsedAttributes> {
    let mut parsed = ParsedAttributes::default();
    let mut expanded = BTreeSet::<Vec<u8>>::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid ODS style attribute: {error}")))?;
        let key = attribute.key;
        let key_bytes = key.as_ref();
        if key_bytes == b"xmlns" || key_bytes.starts_with(b"xmlns:") {
            if let Some(prefix) = key_bytes.strip_prefix(b"xmlns:") {
                let value = decode_attribute(reader, &attribute, limits, "namespace declaration")?;
                if value.len() > MAX_STYLE_TEXT_BYTES {
                    return resource(
                        Resource::InputBytes,
                        value.len(),
                        MAX_STYLE_TEXT_BYTES,
                        "ODS data-style namespace declaration exceeds its limit",
                    );
                }
                if value == String::from_utf8_lossy(STYLE) {
                    parsed.slots.style_binding_available = true;
                }
                if value == String::from_utf8_lossy(NUMBER) {
                    parsed.slots.number_binding_available = true;
                }
                let prefix = bounded_bytes(prefix, "ODS data-style namespace prefix")?;
                if value == String::from_utf8_lossy(STYLE) {
                    parsed.slots.style_prefix = Some(prefix.clone());
                }
                if value == String::from_utf8_lossy(NUMBER) {
                    parsed.slots.number_prefix = Some(prefix);
                }
            }
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(key);
        reject_unbound_attribute_prefix(&namespace, qname_prefix(key_bytes))?;
        let namespace_bytes = normalized_bound_namespace(&namespace).unwrap_or(Cow::Borrowed(&[]));
        let mut identity = Vec::new();
        identity
            .try_reserve(
                namespace_bytes
                    .len()
                    .saturating_add(local.as_ref().len())
                    .saturating_add(1),
            )
            .map_err(|source| allocation("ODS style attribute identity", source))?;
        identity.extend_from_slice(namespace_bytes.as_ref());
        identity.push(0);
        identity.extend_from_slice(local.as_ref());
        if !expanded.insert(identity) {
            return invalid("duplicate ODF data-style attribute");
        }
        let value = decode_attribute(reader, &attribute, limits, "data-style attribute")?;
        if namespace_bytes.as_ref() == STYLE && local.as_ref() == b"name" {
            if !matches!(context, AttributeContext::Root) {
                parsed.opaque = true;
                continue;
            }
            if parsed.name.is_some() {
                return invalid("duplicate ODF data-style style:name attribute");
            }
            parsed.name = Some(value);
            if parsed.slots.style_prefix.is_none() {
                parsed.slots.style_prefix = qname_prefix(key_bytes).map(|prefix| prefix.to_vec());
            }
            continue;
        }
        let field = field_for(namespace_bytes.as_ref(), local.as_ref(), context);
        let Some(field) = field else {
            parsed.opaque = true;
            continue;
        };
        let slot = attribute_slot(&attribute, field)?;
        if parsed
            .slots
            .slots
            .iter()
            .any(|existing| existing.field == field)
        {
            return invalid("duplicate ODF data-style attribute");
        }
        parsed.slots.slots.push(slot);
        let value = normalize_root_attribute_value(field, value)?;
        if namespace_bytes.as_ref() == STYLE && parsed.slots.style_prefix.is_none() {
            parsed.slots.style_prefix = qname_prefix(key_bytes).map(|prefix| prefix.to_vec());
        }
        if namespace_bytes.as_ref() == NUMBER && parsed.slots.number_prefix.is_none() {
            parsed.slots.number_prefix = qname_prefix(key_bytes).map(|prefix| prefix.to_vec());
        }
        match field {
            Field::DisplayName => parsed.display_name = Some(value),
            Field::Language => parsed.language = Some(value),
            Field::Country => parsed.country = Some(value),
            Field::Script => parsed.script = Some(value),
            Field::RfcLanguageTag => parsed.rfc_language_tag = Some(value),
            Field::Title => parsed.title = Some(value),
            Field::Volatile => {
                parsed.volatile = Some(super::parse_boolean(&value, "style:volatile")?)
            },
            Field::TransliterationFormat => parsed.transliteration_format = Some(value),
            Field::TransliterationLanguage => parsed.transliteration_language = Some(value),
            Field::TransliterationCountry => parsed.transliteration_country = Some(value),
            Field::TransliterationStyle => {
                parsed.transliteration_style = Some(value.parse::<TransliterationStyle>()?)
            },
        }
    }
    // The metadata fields are validated after all values have been collected;
    // body scalar parsing is performed by `parse_body_attributes`.
    Ok(parsed)
}

fn qname_prefix(value: &[u8]) -> Option<&[u8]> {
    value
        .splitn(2, |byte| *byte == b':')
        .next()
        .and_then(|prefix| {
            if prefix.len() == value.len() {
                None
            } else {
                Some(prefix)
            }
        })
}

fn normalize_root_attribute_value(field: Field, value: String) -> Result<String> {
    if matches!(
        field,
        Field::Language
            | Field::Country
            | Field::Script
            | Field::RfcLanguageTag
            | Field::TransliterationLanguage
            | Field::TransliterationCountry
    ) {
        return Ok(
            super::collapse_xml_whitespace(&value, "ODS data-style token attribute")?.into_owned(),
        );
    }
    Ok(value)
}

fn field_for(namespace: &[u8], local: &[u8], context: AttributeContext) -> Option<Field> {
    match (namespace, local) {
        (STYLE, b"name") if matches!(context, AttributeContext::Root) => None,
        (STYLE, b"display-name") => Some(Field::DisplayName),
        (NUMBER, b"language") => Some(Field::Language),
        (NUMBER, b"country") => Some(Field::Country),
        (NUMBER, b"script") => Some(Field::Script),
        (NUMBER, b"rfc-language-tag") => Some(Field::RfcLanguageTag),
        (NUMBER, b"title") => Some(Field::Title),
        (STYLE, b"volatile") => Some(Field::Volatile),
        (NUMBER, b"transliteration-format") => Some(Field::TransliterationFormat),
        (NUMBER, b"transliteration-language") => Some(Field::TransliterationLanguage),
        (NUMBER, b"transliteration-country") => Some(Field::TransliterationCountry),
        (NUMBER, b"transliteration-style") => Some(Field::TransliterationStyle),
        _ => None,
    }
}

fn parse_body_attributes(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
    class: ElementKind,
    limits: &Limits,
) -> Result<(NumberFields, bool)> {
    let mut fields = NumberFields::default();
    let mut opaque = false;
    let mut expanded = BTreeSet::<Vec<u8>>::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid ODS body attribute: {error}")))?;
        let key = attribute.key;
        let key_bytes = key.as_ref();
        if key_bytes == b"xmlns" || key_bytes.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(key);
        reject_unbound_attribute_prefix(&namespace, qname_prefix(key_bytes))?;
        let namespace_bytes = normalized_bound_namespace(&namespace).unwrap_or(Cow::Borrowed(&[]));
        let mut identity = Vec::new();
        identity
            .try_reserve(
                namespace_bytes
                    .len()
                    .saturating_add(local.as_ref().len())
                    .saturating_add(1),
            )
            .map_err(|source| allocation("ODS style body attribute identity", source))?;
        identity.extend_from_slice(namespace_bytes.as_ref());
        identity.push(0);
        identity.extend_from_slice(local.as_ref());
        if !expanded.insert(identity) {
            return invalid("duplicate ODF data-style body attribute");
        }
        let value = decode_attribute(reader, &attribute, limits, "data-style body attribute")?;
        let recognized = namespace_bytes.as_ref() == NUMBER;
        if !recognized {
            opaque = true;
            continue;
        }
        match local.as_ref() {
            b"decimal-places" => {
                let parsed = parse_i64_or_opaque(&value, "number:decimal-places", &mut opaque)?;
                if matches!(
                    class,
                    ElementKind::Number | ElementKind::Scientific | ElementKind::Seconds
                ) {
                    fields.decimal_places = parsed;
                } else {
                    opaque = true;
                }
            },
            b"min-decimal-places" => {
                let parsed = parse_i64_or_opaque(&value, "number:min-decimal-places", &mut opaque)?;
                if matches!(class, ElementKind::Number | ElementKind::Scientific) {
                    fields.min_decimal_places = parsed;
                } else {
                    opaque = true;
                }
            },
            b"min-integer-digits" => {
                let parsed = parse_i64_or_opaque(&value, "number:min-integer-digits", &mut opaque)?;
                if matches!(
                    class,
                    ElementKind::Number | ElementKind::Scientific | ElementKind::Fraction
                ) {
                    fields.min_integer_digits = parsed;
                } else {
                    opaque = true;
                }
            },
            b"grouping" => {
                let parsed = super::parse_boolean(&value, "number:grouping")?;
                if matches!(
                    class,
                    ElementKind::Number | ElementKind::Scientific | ElementKind::Fraction
                ) {
                    fields.grouping = Some(parsed);
                } else {
                    opaque = true;
                }
            },
            b"decimal-replacement" => {
                let parsed = bounded_string(&value, "number:decimal-replacement")?;
                if class == ElementKind::Number {
                    fields.decimal_replacement = Some(parsed);
                } else {
                    opaque = true;
                }
            },
            b"display-factor" => {
                let parsed = match Double::from_lexical(&value) {
                    Ok(value) => value,
                    Err(Error::Unsupported(_)) => {
                        opaque = true;
                        continue;
                    },
                    Err(error) => return Err(error),
                };
                if class == ElementKind::Number {
                    fields.display_factor = Some(parsed);
                } else {
                    opaque = true;
                }
            },
            b"min-exponent-digits" => {
                let parsed =
                    parse_i64_or_opaque(&value, "number:min-exponent-digits", &mut opaque)?;
                if class == ElementKind::Scientific {
                    fields.min_exponent_digits = parsed;
                } else {
                    opaque = true;
                }
            },
            b"exponent-interval" => {
                let parsed =
                    parse_positive_or_opaque(&value, "number:exponent-interval", &mut opaque)?;
                if class == ElementKind::Scientific {
                    fields.exponent_interval = parsed;
                } else {
                    opaque = true;
                }
            },
            b"forced-exponent-sign" => {
                let parsed = super::parse_boolean(&value, "number:forced-exponent-sign")?;
                if class == ElementKind::Scientific {
                    fields.forced_exponent_sign = Some(parsed);
                } else {
                    opaque = true;
                }
            },
            b"min-numerator-digits" => {
                let parsed =
                    parse_i64_or_opaque(&value, "number:min-numerator-digits", &mut opaque)?;
                if class == ElementKind::Fraction {
                    fields.min_numerator_digits = parsed;
                } else {
                    opaque = true;
                }
            },
            b"min-denominator-digits" => {
                let parsed =
                    parse_i64_or_opaque(&value, "number:min-denominator-digits", &mut opaque)?;
                if class == ElementKind::Fraction {
                    fields.min_denominator_digits = parsed;
                } else {
                    opaque = true;
                }
            },
            b"denominator-value" => {
                let parsed = parse_i64_or_opaque(&value, "number:denominator-value", &mut opaque)?;
                if class == ElementKind::Fraction {
                    fields.denominator_value = parsed;
                } else {
                    opaque = true;
                }
            },
            b"max-denominator-value" => {
                let parsed =
                    parse_positive_or_opaque(&value, "number:max-denominator-value", &mut opaque)?;
                if class == ElementKind::Fraction {
                    fields.max_denominator_value = parsed;
                } else {
                    opaque = true;
                }
            },
            b"position" => {
                let parsed = parse_i64_or_opaque(&value, "number:position", &mut opaque)?;
                if class == ElementKind::EmbeddedText {
                    fields.position = parsed;
                } else {
                    opaque = true;
                }
            },
            b"style" => {
                let parsed = bounded_string(&value, "number:style")?;
                if !matches!(parsed.as_str(), "short" | "long") {
                    return invalid("invalid ODF number:style value");
                }
                if matches!(
                    class,
                    ElementKind::Year
                        | ElementKind::Month
                        | ElementKind::Day
                        | ElementKind::Hours
                        | ElementKind::Minutes
                        | ElementKind::Seconds
                ) {
                    fields.style = Some(parsed);
                } else {
                    opaque = true;
                }
            },
            _ => opaque = true,
        }
    }
    Ok((fields, opaque))
}

fn parse_i64_or_opaque(value: &str, label: &str, opaque: &mut bool) -> Result<Option<i64>> {
    match super::parse_integer(value, label) {
        Ok(value) => Ok(Some(value)),
        Err(Error::Unsupported(_)) => {
            *opaque = true;
            Ok(None)
        },
        Err(error) => Err(error),
    }
}

fn parse_positive_or_opaque(value: &str, label: &str, opaque: &mut bool) -> Result<Option<u64>> {
    match super::parse_positive(value, label) {
        Ok(value) => Ok(Some(value)),
        Err(Error::Unsupported(_)) => {
            *opaque = true;
            Ok(None)
        },
        Err(error) => Err(error),
    }
}

fn attribute_slot(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    field: Field,
) -> Result<AttributeSlot> {
    Ok(AttributeSlot {
        field,
        key: attribute.key.as_ref().to_vec(),
    })
}

fn decode_attribute(
    reader: &ResolvedReader<'_>,
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    limits: &Limits,
    label: &str,
) -> Result<String> {
    let raw = attribute.value.as_ref();
    if raw.len() > limits.text_bytes {
        return resource(
            Resource::InputBytes,
            raw.len(),
            limits.text_bytes,
            format!("{label} exceeds its text limit"),
        );
    }
    if raw.contains(&b'<') {
        return invalid("ODS data-style attribute contains a raw '<'");
    }
    let value = attribute
        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
        .map_err(|error| invalid_error(format!("invalid {label}: {error}")))?;
    validate_xml_characters(value.as_ref(), label)?;
    if value.len() > limits.text_bytes {
        return resource(
            Resource::InputBytes,
            value.len(),
            limits.text_bytes,
            format!("{label} exceeds its text limit"),
        );
    }
    bounded_string(value.as_ref(), label)
}

fn bounded_string(value: &str, label: &str) -> Result<String> {
    if value.len() > MAX_STYLE_TEXT_BYTES {
        return resource(
            Resource::InputBytes,
            value.len(),
            MAX_STYLE_TEXT_BYTES,
            format!("{label} exceeds its text limit"),
        );
    }
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("ODS data-style text", source))?;
    output.push_str(value);
    Ok(output)
}

fn bounded_bytes(value: &[u8], label: &str) -> Result<Vec<u8>> {
    if value.len() > MAX_STYLE_TEXT_BYTES {
        return resource(
            Resource::InputBytes,
            value.len(),
            MAX_STYLE_TEXT_BYTES,
            label,
        );
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("ODS data-style namespace prefix", source))?;
    output.extend_from_slice(value);
    Ok(output)
}

fn validate_raw_text(value: &[u8], label: &str) -> Result<()> {
    if value.windows(3).any(|window| window == b"]]>") {
        return invalid(format!("{label} contains a raw ]]> delimiter"));
    }
    Ok(())
}

fn validate_xml_character(value: char, label: &str) -> Result<()> {
    let code = value as u32;
    if matches!(code, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF) {
        Ok(())
    } else {
        invalid(format!("{label} contains an invalid XML character"))
    }
}

fn validate_xml_characters(value: &str, label: &str) -> Result<()> {
    for character in value.chars() {
        validate_xml_character(character, label)?;
    }
    Ok(())
}

fn resolve_predefined_entity(value: &str) -> Option<&'static str> {
    match value {
        "amp" => Some("&"),
        "lt" => Some("<"),
        "gt" => Some(">"),
        "apos" => Some("'"),
        "quot" => Some("\""),
        _ => None,
    }
}

fn metadata_size(value: &Attributes) -> Result<usize> {
    let mut size = 0usize;
    for item in [
        value.display_name.as_deref(),
        value.language.as_deref(),
        value.country.as_deref(),
        value.script.as_deref(),
        value.rfc_language_tag.as_deref(),
        value.title.as_deref(),
        value.transliteration.format.as_deref(),
        value.transliteration.language.as_deref(),
        value.transliteration.country.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        size = size
            .checked_add(item.len())
            .ok_or_else(|| invalid_error("ODS data-style metadata size overflow"))?;
    }
    Ok(size)
}

#[derive(Clone, Debug)]
struct RawAttribute {
    key: Range<usize>,
    value: Range<usize>,
    field: Option<Field>,
}

fn patch_metadata(entry: &SourceEntry<'_>, patch: &Patch, max_output: usize) -> Result<Vec<u8>> {
    let source = entry.source;
    patch.validate()?;
    let mut candidate_attributes = entry_attributes(&entry.entry).clone();
    patch.apply(&mut candidate_attributes)?;
    let opening = source
        .get(entry.opening_range.clone())
        .ok_or_else(|| invalid_error("ODS data-style source opening range is invalid"))?;
    let opening_start = entry.opening_range.start;
    let mut raw = raw_attributes(opening)?;
    for attribute in &mut raw {
        let key = &source[entry.opening_range.clone()][attribute.key.clone()];
        attribute.field = entry
            .attributes
            .slots
            .iter()
            .find(|slot| slot.key.as_slice() == key)
            .map(|slot| slot.field);
        attribute.key = (attribute.key.start + opening_start)..(attribute.key.end + opening_start);
        attribute.value =
            (attribute.value.start + opening_start)..(attribute.value.end + opening_start);
    }
    let mut edits = Vec::<(Range<usize>, String)>::new();
    let mut insertion = String::new();
    let mut has_style_decl = entry.attributes.style_binding_available;
    let mut has_number_decl = entry.attributes.number_binding_available;

    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::DisplayName,
        &patch.display_name,
        namespace_for(Field::DisplayName),
        &entry.attributes.style_prefix,
        &mut has_style_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::Language,
        &patch.language,
        namespace_for(Field::Language),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::Country,
        &patch.country,
        namespace_for(Field::Country),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::Script,
        &patch.script,
        namespace_for(Field::Script),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::RfcLanguageTag,
        &patch.rfc_language_tag,
        namespace_for(Field::RfcLanguageTag),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::Title,
        &patch.title,
        namespace_for(Field::Title),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    apply_bool_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        Field::Volatile,
        &patch.volatile,
        namespace_for(Field::Volatile),
        &entry.attributes.style_prefix,
        &mut has_style_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::TransliterationFormat,
        &patch.transliteration_format,
        namespace_for(Field::TransliterationFormat),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::TransliterationLanguage,
        &patch.transliteration_language,
        namespace_for(Field::TransliterationLanguage),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    apply_string_patch(
        &mut edits,
        &mut insertion,
        &raw,
        source,
        &entry.attributes,
        Field::TransliterationCountry,
        &patch.transliteration_country,
        namespace_for(Field::TransliterationCountry),
        &entry.attributes.number_prefix,
        &mut has_number_decl,
        max_output,
    )?;
    match &patch.transliteration_style {
        Op::Keep => {},
        Op::Set(value) => apply_set_or_clear(
            &mut edits,
            &mut insertion,
            &raw,
            source,
            Field::TransliterationStyle,
            value.as_str(),
            false,
            namespace_for(Field::TransliterationStyle),
            &entry.attributes.number_prefix,
            &mut has_number_decl,
            max_output,
        )?,
        Op::Clear => apply_set_or_clear(
            &mut edits,
            &mut insertion,
            &raw,
            source,
            Field::TransliterationStyle,
            "",
            true,
            namespace_for(Field::TransliterationStyle),
            &entry.attributes.number_prefix,
            &mut has_number_decl,
            max_output,
        )?,
    }
    if !insertion.is_empty() {
        let position = opening
            .iter()
            .rposition(|byte| *byte == b'>')
            .ok_or_else(|| invalid_error("ODS data-style opening tag has no close"))?;
        let position = if position > 0 && opening[position - 1] == b'/' {
            position - 1
        } else {
            position
        };
        let position = opening_start
            .checked_add(position)
            .ok_or_else(|| invalid_error("ODS data-style insertion position overflow"))?;
        push_edit(&mut edits, position..position, insertion)?;
    }
    if edits.is_empty() {
        return bounded_copy(source, max_output);
    }
    edits.sort_by_key(|(range, _)| range.start);
    let mut delta = 0isize;
    for (range, replacement) in &edits {
        let removed = range.end.saturating_sub(range.start);
        delta = delta
            .checked_add(replacement.len() as isize)
            .and_then(|value| value.checked_sub(removed as isize))
            .ok_or_else(|| invalid_error("ODS data-style output size overflow"))?;
    }
    let output_len = if delta >= 0 {
        source
            .len()
            .checked_add(delta as usize)
            .ok_or_else(|| invalid_error("ODS data-style output size overflow"))?
    } else {
        source
            .len()
            .checked_sub(delta.unsigned_abs())
            .ok_or_else(|| invalid_error("ODS data-style output size underflow"))?
    };
    if output_len > max_output {
        return resource(
            Resource::OutputBytes,
            output_len,
            max_output,
            "ODS data-style output exceeds its byte limit",
        );
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("ODS data-style patched source", source))?;
    let mut cursor = 0usize;
    for (range, replacement) in edits {
        if range.start < cursor || range.end > source.len() {
            return invalid("overlapping ODS data-style source edits");
        }
        output.extend_from_slice(&source[cursor..range.start]);
        output.extend_from_slice(replacement.as_bytes());
        cursor = range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn entry_attributes(entry: &Entry) -> &Attributes {
    match entry {
        Entry::Number(value) => &value.value.attributes,
        Entry::Existing(value) => &value.value.attributes,
        Entry::Opaque(value) => &value.attributes,
        Entry::Text(value) => &value.attributes,
    }
}

fn apply_string_patch(
    edits: &mut Vec<(Range<usize>, String)>,
    insertion: &mut String,
    raw: &[RawAttribute],
    source: &[u8],
    _attributes: &AttributeSlots,
    field: Field,
    operation: &Op<String>,
    namespace: &'static str,
    prefix: &Option<Vec<u8>>,
    has_decl: &mut bool,
    max_output: usize,
) -> Result<()> {
    match operation {
        Op::Keep => Ok(()),
        Op::Set(value) => apply_set_or_clear(
            edits, insertion, raw, source, field, value, false, namespace, prefix, has_decl,
            max_output,
        ),
        Op::Clear => apply_set_or_clear(
            edits, insertion, raw, source, field, "", true, namespace, prefix, has_decl, max_output,
        ),
    }
}

fn apply_bool_patch(
    edits: &mut Vec<(Range<usize>, String)>,
    insertion: &mut String,
    raw: &[RawAttribute],
    source: &[u8],
    field: Field,
    operation: &Op<bool>,
    namespace: &'static str,
    prefix: &Option<Vec<u8>>,
    has_decl: &mut bool,
    max_output: usize,
) -> Result<()> {
    match operation {
        Op::Keep => Ok(()),
        Op::Set(value) => apply_set_or_clear(
            edits,
            insertion,
            raw,
            source,
            field,
            if *value { "true" } else { "false" },
            false,
            namespace,
            prefix,
            has_decl,
            max_output,
        ),
        Op::Clear => apply_set_or_clear(
            edits, insertion, raw, source, field, "", true, namespace, prefix, has_decl, max_output,
        ),
    }
}

fn attribute_semantically_equals(
    source: &[u8],
    attribute: &RawAttribute,
    field: Field,
    expected: &str,
) -> Result<bool> {
    let raw = source
        .get(attribute.value.clone())
        .ok_or_else(|| invalid_error("ODS data-style attribute range is outside the source"))?;
    if raw.contains(&b'<') {
        return invalid("ODS data-style attribute contains a raw '<'");
    }
    let actual = decode_attribute_for_semantics(raw)?;
    if field == Field::Volatile {
        return Ok(super::parse_boolean(&actual, "style:volatile")?
            == super::parse_boolean(expected, "style:volatile")?);
    }
    let actual = normalize_semantic_attribute(field, &actual)?;
    // Values supplied to a semantic patch are already decoded model values.
    // In particular, a literal tab in the caller's value is not the same as
    // the space produced by XML attribute-value normalization.  Only token
    // fields collapse whitespace; schema:string fields retain their exact
    // characters so a semantic no-op cannot rewrite a distinct lexical value.
    let expected = normalize_semantic_attribute(field, expected)?;
    Ok(actual.as_ref() == expected.as_ref())
}

fn normalize_semantic_attribute<'a>(field: Field, value: &'a str) -> Result<Cow<'a, str>> {
    if matches!(
        field,
        Field::Language
            | Field::Country
            | Field::Script
            | Field::RfcLanguageTag
            | Field::TransliterationLanguage
            | Field::TransliterationCountry
    ) {
        return super::collapse_xml_whitespace(value, "ODS data-style token attribute");
    }
    Ok(Cow::Borrowed(value))
}

fn decode_attribute_for_semantics(raw: &[u8]) -> Result<String> {
    let raw =
        str::from_utf8(raw).map_err(|_| invalid_error("ODS data-style attribute is not UTF-8"))?;
    let bytes = raw.as_bytes();
    let mut output = String::new();
    output
        .try_reserve_exact(raw.len())
        .map_err(|source| allocation("ODS data-style attribute value", source))?;
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        if bytes[cursor] != b'&' {
            let character = raw[cursor..]
                .chars()
                .next()
                .ok_or_else(|| invalid_error("ODS data-style attribute has invalid UTF-8"))?;
            validate_xml_character(character, "ODS data-style attribute")?;
            if character == '\r' {
                // XML end-of-line normalization folds a literal CRLF pair
                // before attribute whitespace normalization.  A numeric
                // character reference reaches this loop through the `&`
                // branch and therefore retains both referenced characters.
                output.push(' ');
                cursor += if bytes.get(cursor + 1) == Some(&b'\n') {
                    2
                } else {
                    1
                };
            } else {
                output.push(if matches!(character, '\t' | '\n') {
                    ' '
                } else {
                    character
                });
                cursor += character.len_utf8();
            }
            continue;
        }
        let end = bytes[cursor + 1..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|offset| cursor + 1 + offset)
            .ok_or_else(|| {
                invalid_error("ODS data-style attribute has an unterminated reference")
            })?;
        if end == cursor + 1 {
            return invalid("ODS data-style attribute has an empty reference");
        }
        let name = &raw[cursor + 1..end];
        let character =
            if let Some(value) = name.strip_prefix("#x").or_else(|| name.strip_prefix("#X")) {
                let code = u32::from_str_radix(value, 16).map_err(|_| {
                    invalid_error("ODS data-style attribute has an invalid character reference")
                })?;
                char::from_u32(code).ok_or_else(|| {
                    invalid_error("ODS data-style attribute has an invalid character reference")
                })?
            } else if let Some(value) = name.strip_prefix('#') {
                let code = value.parse::<u32>().map_err(|_| {
                    invalid_error("ODS data-style attribute has an invalid character reference")
                })?;
                char::from_u32(code).ok_or_else(|| {
                    invalid_error("ODS data-style attribute has an invalid character reference")
                })?
            } else {
                let Some(value) = resolve_predefined_entity(name) else {
                    return invalid("ODS data-style attribute has an unsupported named entity");
                };
                value
                    .chars()
                    .next()
                    .expect("predefined entity is one character")
            };
        validate_xml_character(character, "ODS data-style attribute")?;
        output.push(character);
        cursor = end + 1;
    }
    Ok(output)
}

fn apply_set_or_clear(
    edits: &mut Vec<(Range<usize>, String)>,
    insertion: &mut String,
    raw: &[RawAttribute],
    source: &[u8],
    field: Field,
    value: &str,
    clear: bool,
    namespace: &'static str,
    prefix: &Option<Vec<u8>>,
    has_decl: &mut bool,
    max_output: usize,
) -> Result<()> {
    if let Some(attribute) = raw.iter().find(|attribute| attribute.field == Some(field)) {
        if clear {
            push_edit(
                edits,
                attribute.key.start..attribute.value.end.saturating_add(1),
                String::new(),
            )?;
        } else if !attribute_semantically_equals(source, attribute, field, value)? {
            push_edit(
                edits,
                attribute.value.clone(),
                escape_xml_bounded(value, max_output)?,
            )?;
        }
        return Ok(());
    }
    if clear {
        return Ok(());
    }
    let prefix = prefix
        .as_deref()
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| {
            if namespace == "style" {
                b"style"
            } else {
                b"number"
            }
        });
    let prefix = str::from_utf8(prefix)
        .map_err(|_| invalid_error("ODS data-style namespace prefix is not UTF-8"))?;
    let namespace_uri = if namespace == "style" {
        str::from_utf8(STYLE).unwrap_or_default()
    } else {
        str::from_utf8(NUMBER).unwrap_or_default()
    };
    let escaped = escape_xml_bounded(value, max_output)?;
    if !*has_decl {
        append_bounded(insertion, " xmlns:", max_output)?;
        append_bounded(insertion, prefix, max_output)?;
        append_bounded(insertion, "=\"", max_output)?;
        append_bounded(insertion, namespace_uri, max_output)?;
        append_bounded(insertion, "\"", max_output)?;
        *has_decl = true;
    }
    append_bounded(insertion, " ", max_output)?;
    append_bounded(insertion, prefix, max_output)?;
    append_bounded(insertion, ":", max_output)?;
    append_bounded(insertion, field_local(field), max_output)?;
    append_bounded(insertion, "=\"", max_output)?;
    append_bounded(insertion, &escaped, max_output)?;
    append_bounded(insertion, "\"", max_output)?;
    Ok(())
}

fn push_edit(
    edits: &mut Vec<(Range<usize>, String)>,
    range: Range<usize>,
    value: String,
) -> Result<()> {
    edits
        .try_reserve(1)
        .map_err(|source| allocation("ODS data-style source edits", source))?;
    edits.push((range, value));
    Ok(())
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let extra = match character {
            '&' => 5,
            '<' | '>' => 4,
            '\'' | '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        length
            .checked_add(extra)
            .ok_or_else(|| invalid_error("ODS data-style escaped value size overflow"))
    })
}

fn escape_xml_bounded(value: &str, max_output: usize) -> Result<String> {
    let length = escaped_xml_len(value)?;
    if length > max_output {
        return resource(
            Resource::OutputBytes,
            length,
            max_output,
            "ODS data-style replacement exceeds its output limit",
        );
    }
    let mut escaped = String::new();
    escaped
        .try_reserve_exact(length)
        .map_err(|source| allocation("ODS data-style escaped value", source))?;
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '\'' => escaped.push_str("&apos;"),
            '"' => escaped.push_str("&quot;"),
            '\t' => escaped.push_str("&#x9;"),
            '\n' => escaped.push_str("&#xA;"),
            '\r' => escaped.push_str("&#xD;"),
            _ => escaped.push(character),
        }
    }
    Ok(escaped)
}

fn append_bounded(target: &mut String, value: &str, max_output: usize) -> Result<()> {
    let length = target
        .len()
        .checked_add(value.len())
        .ok_or_else(|| invalid_error("ODS data-style insertion size overflow"))?;
    if length > max_output {
        return resource(
            Resource::OutputBytes,
            length,
            max_output,
            "ODS data-style insertion exceeds its output limit",
        );
    }
    target
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("ODS data-style source insertion", source))?;
    target.push_str(value);
    Ok(())
}

fn namespace_for(field: Field) -> &'static str {
    match field {
        Field::DisplayName | Field::Volatile => "style",
        _ => "number",
    }
}

fn field_local(field: Field) -> &'static str {
    match field {
        Field::DisplayName => "display-name",
        Field::Language => "language",
        Field::Country => "country",
        Field::Script => "script",
        Field::RfcLanguageTag => "rfc-language-tag",
        Field::Title => "title",
        Field::Volatile => "volatile",
        Field::TransliterationFormat => "transliteration-format",
        Field::TransliterationLanguage => "transliteration-language",
        Field::TransliterationCountry => "transliteration-country",
        Field::TransliterationStyle => "transliteration-style",
    }
}

fn raw_attributes(opening: &[u8]) -> Result<Vec<RawAttribute>> {
    let mut output = Vec::new();
    let mut cursor = 1usize;
    while cursor < opening.len()
        && !opening[cursor].is_ascii_whitespace()
        && !matches!(opening[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    while cursor < opening.len() {
        while cursor < opening.len() && opening[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= opening.len() || opening[cursor] == b'>' || opening[cursor] == b'/' {
            break;
        }
        let key_start = cursor;
        while cursor < opening.len()
            && !opening[cursor].is_ascii_whitespace()
            && !matches!(opening[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let key_end = cursor;
        while cursor < opening.len() && opening[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= opening.len() || opening[cursor] != b'=' {
            return invalid("ODS data-style source attribute lacks equals");
        }
        cursor += 1;
        while cursor < opening.len() && opening[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *opening
            .get(cursor)
            .ok_or_else(|| invalid_error("ODS data-style source attribute lacks quote"))?;
        if quote != b'"' && quote != b'\'' {
            return invalid("ODS data-style source attribute value lacks quote");
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < opening.len() && opening[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= opening.len() {
            return invalid("ODS data-style source attribute value is unterminated");
        }
        output
            .try_reserve(1)
            .map_err(|source| allocation("ODS data-style raw attributes", source))?;
        output.push(RawAttribute {
            key: key_start..key_end,
            value: value_start..value_end,
            field: None,
        });
        cursor += 1;
    }
    Ok(output)
}

fn bounded_copy(source: &[u8], max_output: usize) -> Result<Vec<u8>> {
    if source.len() > max_output {
        return resource(
            Resource::OutputBytes,
            source.len(),
            max_output,
            "ODS data-style output exceeds its byte limit",
        );
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| allocation("ODS data-style source copy", source))?;
    output.extend_from_slice(source);
    Ok(output)
}

fn checked_position(reader: &ResolvedReader<'_>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid_error("ODS data-style XML position exceeds usize"))
}

fn xml_error(error: quick_xml::Error) -> Error {
    invalid_error(format!("ODS data-style XML parsing error: {error}"))
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn resource<T>(
    resource: Resource,
    observed: usize,
    limit: usize,
    message: impl Into<String>,
) -> Result<T> {
    Err(Error::ResourceLimit(litchi_core::ResourceLimit {
        resource,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        limit: u64::try_from(limit).unwrap_or(u64::MAX),
        scope: Arc::from(message.into()),
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const NUMBER_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
    const XML_STR: &str = "http://www.w3.org/XML/1998/namespace";

    fn aliased_content() -> Vec<u8> {
        format!(
            r#"<?xml version="1.0"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}">
  <office:automatic-styles>
    <n:number-style s:name="Aliased"><n:scientific-number n:decimal-places="2"/></n:number-style>
  </office:automatic-styles>
  <office:body><office:spreadsheet/></office:body>
</office:document-content>"#
        )
        .into_bytes()
    }

    #[test]
    fn source_scan_resolves_inherited_aliases_and_preserves_body() -> Result<()> {
        let source = aliased_content();
        let catalog = scan(&source, Owner::ContentAutomatic)?;
        let entry = catalog.lookup(Selector::automatic("Aliased", Family::Number))?;
        assert_eq!(
            entry.raw_body()?,
            b"<n:scientific-number n:decimal-places=\"2\"/>"
        );
        let patch = Patch::default().set_title("patched")?;
        let changed = entry.patch_metadata(&patch, source.len() + 256)?;
        assert!(
            changed
                .windows(b"s:name=\"Aliased\"".len())
                .any(|window| { window == b"s:name=\"Aliased\"" })
        );
        assert!(
            changed
                .windows(b"n:scientific-number".len())
                .any(|window| { window == b"n:scientific-number" })
        );
        let reopened = scan(&changed, Owner::ContentAutomatic)?;
        assert_eq!(
            reopened
                .lookup(Selector::automatic("Aliased", Family::Number))?
                .name(),
            "Aliased"
        );
        let mut stale = entry.clone();
        stale.range = 0..source.len().saturating_add(1);
        assert!(stale.source().is_err());
        stale.body_range = source.len().saturating_sub(1)..source.len().saturating_add(1);
        assert!(stale.raw_body().is_err());
        Ok(())
    }

    #[test]
    fn source_scan_rejects_duplicate_family_name() {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Same"/><n:number-style s:name="Same"/></office:automatic-styles><office:body/></office:document-content>"#
        );
        assert!(scan(source.as_bytes(), Owner::ContentAutomatic).is_err());
    }

    #[test]
    fn number_particles_and_embedded_text_are_projected() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Fmt"><n:text>$</n:text><n:fill-character>*</n:fill-character><n:text> </n:text><n:number n:decimal-places="2"><n:embedded-text n:position="1">#</n:embedded-text></n:number><n:text>%</n:text></n:number-style></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        let entry = catalog.lookup(Selector::automatic("Fmt", Family::Number))?;
        let Entry::Number(number) = entry.entry() else {
            panic!("expected typed number style");
        };
        let leading = number.value.leading.as_ref().expect("leading affix");
        assert_eq!(leading.text.as_deref(), Some("$"));
        assert_eq!(leading.fill_character.as_deref(), Some("*"));
        assert_eq!(leading.text_after_fill.as_deref(), Some(" "));
        assert_eq!(
            number
                .value
                .trailing
                .as_ref()
                .and_then(|value| value.text.as_deref()),
            Some("%")
        );
        let named_entity = source.replace("<n:text>$", "<n:text>&amp;");
        let entity_catalog = scan(named_entity.as_bytes(), Owner::ContentAutomatic)?;
        let entity_entry = entity_catalog.lookup(Selector::automatic("Fmt", Family::Number))?;
        let Entry::Number(entity_number) = entity_entry.entry() else {
            panic!("expected typed named entity number style");
        };
        assert_eq!(
            entity_number
                .value
                .leading
                .as_ref()
                .and_then(|value| value.text.as_deref()),
            Some("&")
        );
        Ok(())
    }

    #[test]
    fn valid_integer_overflow_is_opaque_but_malformed_integer_is_invalid() -> Result<()> {
        let overflow = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Wide"><n:number n:decimal-places="9223372036854775808"/></n:number-style></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(overflow.as_bytes(), Owner::ContentAutomatic)?;
        assert!(matches!(
            catalog
                .lookup(Selector::automatic("Wide", Family::Number))?
                .entry(),
            Entry::Opaque(_)
        ));
        let malformed = overflow.replace("9223372036854775808", "not-an-integer");
        assert!(scan(malformed.as_bytes(), Owner::ContentAutomatic).is_err());
        Ok(())
    }

    #[test]
    fn known_body_scalars_validate_even_when_used_on_the_wrong_particle() -> Result<()> {
        let valid = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Wrong"><n:year n:grouping="true"/></n:number-style></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(valid.as_bytes(), Owner::ContentAutomatic)?;
        assert!(matches!(
            catalog
                .lookup(Selector::automatic("Wrong", Family::Number))?
                .entry(),
            Entry::Opaque(_)
        ));

        let malformed = valid.replace("grouping=\"true\"", "grouping=\"maybe\"");
        assert!(scan(malformed.as_bytes(), Owner::ContentAutomatic).is_err());
        Ok(())
    }

    #[test]
    fn malformed_number_particle_order_is_rejected() {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Bad"><n:number/><n:number/></n:number-style></office:automatic-styles><office:body/></office:document-content>"#
        );
        assert!(scan(source.as_bytes(), Owner::ContentAutomatic).is_err());
    }

    #[test]
    fn owner_whitespace_uses_xml_character_domain() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles>&#x20;<n:number-style s:name="One"/>&#x9;</office:automatic-styles><office:body/></office:document-content>"#
        );
        assert_eq!(
            scan(source.as_bytes(), Owner::ContentAutomatic)?
                .entries()
                .len(),
            1
        );
        let invalid = source.replace("&#x9;", "\u{00a0}");
        assert!(scan(invalid.as_bytes(), Owner::ContentAutomatic).is_err());
        Ok(())
    }

    #[test]
    fn unbound_qualified_names_are_rejected() {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Bad" x:title="x"/></office:automatic-styles><office:body/></office:document-content>"#
        );
        assert!(scan(source.as_bytes(), Owner::ContentAutomatic).is_err());
    }

    #[test]
    fn caller_body_limits_are_checked_before_body_reservation() {
        let source = aliased_content();
        let limits = Limits {
            body_elements: 0,
            ..Limits::default()
        };
        assert!(scan_with_limits(source.as_slice(), Owner::ContentAutomatic, limits).is_err());
    }

    #[test]
    fn styles_after_owner_close_are_not_catalogued() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Selected"/></office:automatic-styles><office:body><office:automatic-styles><n:number-style s:name="Nested"/></office:automatic-styles></office:body></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        assert_eq!(catalog.entries().len(), 1);
        assert_eq!(catalog.entries()[0].name(), "Selected");
        let owner_range = catalog.owner_range().expect("direct owner range");
        assert_eq!(
            &source.as_bytes()[owner_range],
            b"<office:automatic-styles><n:number-style s:name=\"Selected\"/></office:automatic-styles>"
                .as_slice()
        );
        Ok(())
    }

    #[test]
    fn self_closing_owner_does_not_admit_later_lookalikes() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles/><office:body><n:number-style s:name="Lookalike"/></office:body></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        assert!(
            catalog.entries().is_empty(),
            "catalogued names: {:?}",
            catalog
                .entries()
                .iter()
                .map(SourceEntry::name)
                .collect::<Vec<_>>()
        );
        assert_eq!(
            catalog.owner_range().map(|range| &source.as_bytes()[range]),
            Some(b"<office:automatic-styles/>".as_slice())
        );
        Ok(())
    }

    #[test]
    fn namespace_and_xml_grammar_are_checked_before_projection() {
        let base = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="One"><n:text>value</n:text></n:number-style></office:automatic-styles><office:body/></office:document-content>"#
        );
        let reserved_default = base.replace(
            &format!("xmlns:office=\"{OFFICE_NS}\""),
            &format!("xmlns:office=\"{OFFICE_NS}\" xmlns=\"{XML_STR}\""),
        );
        assert!(scan(reserved_default.as_bytes(), Owner::ContentAutomatic).is_err());
        let raw_delimiter = base.replace("value", "]]>");
        assert!(scan(raw_delimiter.as_bytes(), Owner::ContentAutomatic).is_err());
        let unknown_entity = base.replace("value", "&unknown;");
        assert!(scan(unknown_entity.as_bytes(), Owner::ContentAutomatic).is_err());
        let invalid_qname = base.replace("<n:number-style", "<1number-style");
        assert!(scan(invalid_qname.as_bytes(), Owner::ContentAutomatic).is_err());

        let explicit_xml = base.replace(
            &format!("xmlns:office=\"{OFFICE_NS}\""),
            &format!("xmlns:office=\"{OFFICE_NS}\" xmlns:xml=\"{XML_STR}\""),
        );
        assert!(scan(explicit_xml.as_bytes(), Owner::ContentAutomatic).is_ok());
    }

    #[test]
    fn source_owner_requires_the_matching_document_root() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="One"/></office:automatic-styles><office:body/></office:document-content>"#
        );
        let styles = content.replace("document-content", "document-styles");
        assert!(scan(content.as_bytes(), Owner::StylesAutomatic).is_err());
        assert!(scan(styles.as_bytes(), Owner::ContentAutomatic).is_err());
    }

    #[test]
    fn semantic_metadata_whitespace_preserves_source_distinctions() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="One" s:display-name="a&#x9;b" n:title="a&#xA;b"/></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        let entry = catalog.lookup(Selector::automatic("One", Family::Number))?;
        let patch = Patch::default().set_display_name("a b")?.set_title("a b")?;
        let changed = entry.patch_metadata(&patch, source.len() + 64)?;
        assert_ne!(changed, source.as_bytes());
        assert!(
            String::from_utf8(changed.clone())
                .expect("patched source is UTF-8")
                .contains("s:display-name=\"a b\" n:title=\"a b\"")
        );
        scan(&changed, Owner::ContentAutomatic)?;

        // Literal XML attribute whitespace is normalized to a space by XML
        // parsing, so this is a semantic no-op against the same caller value.
        let literal = source.replace("a&#x9;b", "a\tb");
        let literal_catalog = scan(literal.as_bytes(), Owner::ContentAutomatic)?;
        let literal_entry = literal_catalog.lookup(Selector::automatic("One", Family::Number))?;
        assert_eq!(
            literal_entry
                .patch_metadata(&Patch::default().set_display_name("a b")?, literal.len(),)?,
            literal.as_bytes()
        );

        // XML line-end normalization folds a literal CRLF pair to one space
        // before attribute-value normalization.  The source bytes therefore
        // remain an exact semantic no-op for the single-space model value.
        let crlf = source.replace("a&#x9;b", "a\r\nb");
        let crlf_catalog = scan(crlf.as_bytes(), Owner::ContentAutomatic)?;
        let crlf_entry = crlf_catalog.lookup(Selector::automatic("One", Family::Number))?;
        assert_eq!(
            crlf_entry.patch_metadata(&Patch::default().set_display_name("a b")?, crlf.len(),)?,
            crlf.as_bytes()
        );
        Ok(())
    }

    #[test]
    fn metadata_attribute_controls_are_written_as_references_and_read_back() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="One"/></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        let entry = catalog.lookup(Selector::automatic("One", Family::Number))?;
        let patch = Patch::default().set_title("\t\n\r")?;
        let changed = entry.patch_metadata(&patch, source.len() + 128)?;
        let changed_text = String::from_utf8(changed.clone()).expect("patched source is UTF-8");
        assert!(changed_text.contains("n:title=\"&#x9;&#xA;&#xD;\""));
        let reopened = scan(&changed, Owner::ContentAutomatic)?;
        let reopened = reopened.lookup(Selector::automatic("One", Family::Number))?;
        match reopened.entry() {
            Entry::Number(value) => {
                assert_eq!(value.value.attributes.title.as_deref(), Some("\t\n\r"))
            },
            Entry::Opaque(value) => assert_eq!(value.attributes.title.as_deref(), Some("\t\n\r")),
            Entry::Text(value) => assert_eq!(value.attributes.title.as_deref(), Some("\t\n\r")),
            Entry::Existing(value) => {
                assert_eq!(value.value.attributes.title.as_deref(), Some("\t\n\r"))
            },
        }
        Ok(())
    }

    #[test]
    fn malformed_opaque_markup_is_rejected_outside_selected_owner() {
        let malformed_comment = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles/><office:body><!--bad--comment--></office:body></office:document-content>"#
        );
        assert!(scan(malformed_comment.as_bytes(), Owner::ContentAutomatic).is_err());

        let malformed_pi = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles/><office:body><?1bad data?></office:body></office:document-content>"#
        );
        assert!(scan(malformed_pi.as_bytes(), Owner::ContentAutomatic).is_err());

        let unknown_reference = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles/><office:body>&unknown;</office:body></office:document-content>"#
        );
        assert!(scan(unknown_reference.as_bytes(), Owner::ContentAutomatic).is_err());

        let invalid_character_reference = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles/><office:body>&#x1;</office:body></office:document-content>"#
        );
        assert!(
            scan(
                invalid_character_reference.as_bytes(),
                Owner::ContentAutomatic
            )
            .is_err()
        );
    }

    #[test]
    fn escaped_namespace_uri_aliases_are_resolved_and_expanded_duplicates_rejected() -> Result<()> {
        let base = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="One"><n:number n:decimal-places="2"/></n:number-style></office:automatic-styles><office:body/></office:document-content>"#
        );
        let escaped = base
            .replace(OFFICE_NS, &OFFICE_NS.replace(':', "&#58;"))
            .replace(STYLE_NS, &STYLE_NS.replace(':', "&#58;"))
            .replace(NUMBER_NS, &NUMBER_NS.replace(':', "&#58;"));
        let catalog = scan(escaped.as_bytes(), Owner::ContentAutomatic)?;
        assert_eq!(
            catalog
                .lookup(Selector::automatic("One", Family::Number))?
                .name(),
            "One"
        );

        let duplicate = escaped
            .replace(
                &format!("xmlns:n=\"{}\"", NUMBER_NS.replace(':', "&#58;")),
                &format!(
                    "xmlns:n=\"{}\" xmlns:n2=\"{}\"",
                    NUMBER_NS.replace(':', "&#58;"),
                    NUMBER_NS.replace(':', "&#58;")
                ),
            )
            .replace(
                "s:name=\"One\"",
                "s:name=\"One\" n:language=\"en\" n2:language=\"fr\"",
            );
        assert!(scan(duplicate.as_bytes(), Owner::ContentAutomatic).is_err());
        Ok(())
    }

    #[test]
    fn metadata_patch_reuses_inherited_shortest_namespace_aliases() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}"><office:automatic-styles xmlns:s="{STYLE_NS}" xmlns:t="{STYLE_NS}" xmlns:n="{NUMBER_NS}" xmlns:z="{NUMBER_NS}"><n:number-style s:name="One"/></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        let entry = catalog.lookup(Selector::automatic("One", Family::Number))?;
        let patch = Patch::default().set_title("inherited")?;
        let changed = entry.patch_metadata(&patch, source.len() + 128)?;
        let style_start = changed
            .windows(b"<n:number-style".len())
            .position(|window| window == b"<n:number-style")
            .expect("style root");
        let style_end = changed[style_start..]
            .iter()
            .position(|byte| *byte == b'>')
            .map(|offset| style_start + offset)
            .expect("style opening close");
        let opening = &changed[style_start..=style_end];
        assert!(
            opening
                .windows(b"n:title=\"inherited\"".len())
                .any(|window| { window == b"n:title=\"inherited\"" })
        );
        assert!(
            !opening
                .windows(b"xmlns:n".len())
                .any(|window| window == b"xmlns:n")
        );
        scan(&changed, Owner::ContentAutomatic)?
            .lookup(Selector::automatic("One", Family::Number))?;
        Ok(())
    }

    #[test]
    fn styles_xml_automatic_and_common_owners_have_distinct_containers() -> Result<()> {
        let source = format!(
            r#"<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Automatic"/></office:automatic-styles><office:styles><n:number-style s:name="Common"/></office:styles></office:document-styles>"#
        );
        assert_eq!(
            scan(source.as_bytes(), Owner::StylesAutomatic)?
                .lookup_name("Automatic")?
                .name(),
            "Automatic"
        );
        assert_eq!(
            scan(source.as_bytes(), Owner::CommonStyles)?
                .lookup_name("Common")?
                .name(),
            "Common"
        );
        assert!(
            scan(source.as_bytes(), Owner::StylesAutomatic)
                .unwrap()
                .lookup_name("Common")
                .is_err()
        );
        Ok(())
    }

    #[test]
    fn metadata_patch_distinguishes_empty_set_clear_and_semantic_noop() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="One" n:title="A&amp;B" s:volatile=" true "/></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        let entry = catalog.lookup(Selector::automatic("One", Family::Number))?;
        let semantic_noop = Patch::default().set_title("A&B")?.set_volatile(true);
        assert_eq!(
            entry.patch_metadata(&semantic_noop, source.len())?,
            source.as_bytes()
        );

        let set_empty = Patch::default().set_title("")?;
        let set_empty_source = entry.patch_metadata(&set_empty, source.len())?;
        assert!(
            String::from_utf8(set_empty_source)
                .expect("patched source is UTF-8")
                .contains("n:title=\"\"")
        );

        let clear = Patch::default().clear_title();
        let clear_source = entry.patch_metadata(&clear, source.len())?;
        assert!(
            !String::from_utf8(clear_source)
                .expect("patched source is UTF-8")
                .contains("n:title=")
        );
        Ok(())
    }

    #[test]
    fn token_whitespace_is_semantic_for_metadata_patch() -> Result<()> {
        let source = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="One" n:language=" en "/></office:automatic-styles><office:body/></office:document-content>"#
        );
        let catalog = scan(source.as_bytes(), Owner::ContentAutomatic)?;
        let entry = catalog.lookup(Selector::automatic("One", Family::Number))?;
        let patch = Patch::default().set_language("en")?;
        assert_eq!(
            entry.patch_metadata(&patch, source.len())?,
            source.as_bytes()
        );
        Ok(())
    }

    #[test]
    fn number_style_lexical_domain_is_checked_before_wrong_particle_opaque() -> Result<()> {
        let valid = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}"><office:automatic-styles><n:number-style s:name="Wrong"><n:year n:style="short"/></n:number-style></office:automatic-styles><office:body/></office:document-content>"#
        );
        assert!(matches!(
            scan(valid.as_bytes(), Owner::ContentAutomatic)?
                .lookup(Selector::automatic("Wrong", Family::Number))?
                .entry(),
            Entry::Opaque(_)
        ));
        let malformed = valid.replace("style=\"short\"", "style=\"bogus\"");
        assert!(scan(malformed.as_bytes(), Owner::ContentAutomatic).is_err());
        Ok(())
    }
}
