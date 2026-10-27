//! XML event reader with normalized namespace declaration values.
//!
//! `quick_xml::reader::NsReader` resolves namespace declarations before it
//! returns the event.  Its resolver therefore sees the lexical attribute
//! bytes.  XML namespace values are attribute values, however, so character
//! references must be expanded before reserved-prefix checks.  This adapter
//! keeps the original event bytes and source positions while applying the
//! same resolver operations to normalized declaration values.

use crate::Error;
use crate::limits::{LimitKind, RawXmlLimitKind};
use quick_xml::{
    events::{BytesDecl, BytesStart, Event},
    name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult},
    reader::Reader,
};
use std::{collections::HashSet, io, str};

use super::stream_xml::{
    XmlStreamLimits, reject_unknown_namespace, validate_event_bytes, validate_name, validate_qname,
    validate_references,
};

type Result<T> = std::result::Result<T, Error>;

/// A source-backed XML reader that resolves namespace declarations after XML
/// attribute-value normalization.
///
/// Events borrow the caller's event buffer exactly as they do with
/// [`quick_xml::reader::NsReader`].  Namespace declarations are copied into
/// the resolver's bounded internal table, while the original start-tag bytes
/// remain untouched for source-range consumers.
#[derive(Debug, Clone)]
pub struct ResolvedReader<'a> {
    reader: Reader<&'a [u8]>,
    resolver: NamespaceResolver,
    pending_pop: bool,
    active_bindings: usize,
    active_bytes: usize,
    scope_bindings: Vec<usize>,
    scope_bytes: Vec<usize>,
    attributes: usize,
    events: usize,
    limits: XmlStreamLimits,
    check_attribute_errors: bool,
    allow_empty_prefixed_namespace: bool,
    document: DocumentValidation,
}

#[derive(Debug, Clone, Copy, Default)]
struct DocumentValidation {
    enabled: bool,
    allow_unknown_namespaces: bool,
    depth: usize,
    root_seen: bool,
    root_closed: bool,
    declaration_seen: bool,
    saw_any_event: bool,
}

/// Preserve typed reader failures while adding a stable parser context to
/// malformed-wire errors.  In particular, resource-limit and allocation
/// failures must not be flattened into a generic invalid-format message by a
/// family codec that consumes this reader.
#[must_use]
pub fn map_resolved_reader_error(error: Error, context: &'static str) -> litchi_core::Error {
    match litchi_core::Error::from(error) {
        litchi_core::Error::InvalidFormat(message) => {
            litchi_core::Error::InvalidFormat(format!("{context}: {message}"))
        },
        litchi_core::Error::ParseError(message) => {
            litchi_core::Error::ParseError(format!("{context}: {message}"))
        },
        other => other,
    }
}

impl<'a> ResolvedReader<'a> {
    /// Construct a reader over one UTF-8 XML source under an explicit policy.
    pub fn from_xml(source: &'a str, limits: XmlStreamLimits) -> Result<Self> {
        if source.len() as u64 > limits.max_bytes() {
            return Err(limit_error(
                RawXmlDimension::InputBytes,
                source.len() as u64,
                limits.max_bytes(),
            ));
        }
        Ok(Self {
            reader: Reader::from_str(source),
            resolver: NamespaceResolver::default(),
            pending_pop: false,
            active_bindings: 0,
            active_bytes: 0,
            scope_bindings: Vec::new(),
            scope_bytes: Vec::new(),
            attributes: 0,
            events: 0,
            limits,
            check_attribute_errors: true,
            allow_empty_prefixed_namespace: false,
            document: DocumentValidation::default(),
        })
    }

    /// Construct a reader for one complete XML document.
    ///
    /// Unlike [`Self::from_xml`], this mode restores the bounded scanner's
    /// document-level validation while retaining the source-backed, zero-copy
    /// event stream. Callers that intentionally parse an XML fragment should
    /// use [`Self::from_xml`] instead.
    pub fn from_xml_document(source: &'a str, limits: XmlStreamLimits) -> Result<Self> {
        let mut reader = Self::from_xml(source, limits)?;
        reader.enable_document_validation();
        Ok(reader)
    }

    /// Enable complete-document validation before reading the first event.
    pub fn enable_document_validation(&mut self) {
        self.document.enabled = true;
        self.document.allow_unknown_namespaces = false;
        self.reader.config_mut().check_end_names = true;
        self.reader.config_mut().check_comments = true;
    }

    /// Enable complete XML structure validation while retaining opaque
    /// prefixes inherited from the containing package owner.
    ///
    /// A typed reader for an extracted fragment must reject malformed XML,
    /// duplicate roots, and malformed document-level declarations, but it
    /// cannot reject a vendor prefix whose declaration lives on the package
    /// owner's root.
    /// This mode applies the same structural and resource checks as
    /// [`Self::enable_document_validation`] while treating such prefixes as
    /// opaque names.
    pub fn enable_fragment_validation(&mut self) {
        self.document.enabled = true;
        self.document.allow_unknown_namespaces = true;
        self.reader.config_mut().check_end_names = true;
        self.reader.config_mut().check_comments = true;
    }

    /// Return the parser configuration.
    #[must_use]
    pub const fn config(&self) -> &quick_xml::reader::Config {
        self.reader.config()
    }

    /// Mutably access the parser configuration.
    pub fn config_mut(&mut self) -> &mut quick_xml::reader::Config {
        self.reader.config_mut()
    }

    /// Return the namespace resolver for the current event scope.
    #[must_use]
    pub const fn resolver(&self) -> &NamespaceResolver {
        &self.resolver
    }

    /// Return the current source position.
    #[must_use]
    pub const fn buffer_position(&self) -> u64 {
        self.reader.buffer_position()
    }

    /// Return the number of events admitted by this reader, including EOF.
    #[must_use]
    pub const fn event_count(&self) -> usize {
        self.events
    }

    /// Return the aggregate number of attributes admitted by this reader,
    /// including XML declaration attributes when document validation is on.
    #[must_use]
    pub const fn attribute_count(&self) -> usize {
        self.attributes
    }

    /// Return the decoder selected by the XML reader.
    #[must_use]
    pub const fn decoder(&self) -> quick_xml::encoding::Decoder {
        self.reader.decoder()
    }

    /// Read and resolve the next event while preserving its original bytes.
    ///
    /// Namespace declaration values are normalized before they are passed to
    /// `NamespaceResolver::add`, which admits legal escaped aliases such as
    /// `xmlns:xml="http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace"`.
    pub fn read_resolved_event_into<'b>(
        &mut self,
        buffer: &'b mut Vec<u8>,
    ) -> Result<(ResolveResult<'_>, Event<'b>)> {
        self.pop_pending_scope();
        self.admit_event()?;
        let start = self.reader.buffer_position();
        let event = self
            .reader
            .read_event_into(buffer)
            .map_err(map_reader_error)?;
        self.check_token_bytes(start)?;
        self.resolve_event(event, start)
    }

    /// Read and resolve the next event directly from the borrowed XML input.
    ///
    /// This is the zero-copy counterpart to
    /// [`Self::read_resolved_event_into`].  It is intended for source-backed
    /// callers whose event buffer is already the original input slice.
    pub fn read_resolved_event(&mut self) -> Result<(ResolveResult<'_>, Event<'a>)> {
        self.pop_pending_scope();
        self.admit_event()?;
        let start = self.reader.buffer_position();
        let event = self.reader.read_event().map_err(map_reader_error)?;
        self.check_token_bytes(start)?;
        self.resolve_event(event, start)
    }

    /// Return the policy carried by this reader.
    #[must_use]
    pub const fn limits(&self) -> XmlStreamLimits {
        self.limits
    }

    /// Configure whether malformed attributes abort the read.
    ///
    /// Namespace-aware callers normally leave this enabled. A legacy source
    /// scanner may deliberately stop namespace discovery at the first
    /// malformed attribute and let its enclosing validator preserve that
    /// scanner's historical error precedence.
    pub const fn set_attribute_error_checks(&mut self, enabled: bool) {
        self.check_attribute_errors = enabled;
    }

    /// Configure whether an empty URI may shadow a named prefix.
    ///
    /// XML namespace readers normally reject this form.  A legacy ODF
    /// scanner can retain it as an unbound prefix while preserving the
    /// historical deep-event behavior.
    pub const fn set_allow_empty_prefixed_namespace(&mut self, enabled: bool) {
        self.allow_empty_prefixed_namespace = enabled;
    }

    fn admit_event(&mut self) -> Result<()> {
        let observed = self.events.saturating_add(1);
        if observed > self.limits.max_events() {
            return Err(limit_error(
                RawXmlDimension::Events,
                observed as u64,
                self.limits.max_events() as u64,
            ));
        }
        self.events = observed;
        Ok(())
    }

    fn check_token_bytes(&self, start: u64) -> Result<()> {
        let observed = self.reader.buffer_position().saturating_sub(start);
        if observed > self.limits.max_token_bytes() as u64 {
            return Err(limit_error(
                RawXmlDimension::TokenBytes,
                observed,
                self.limits.max_token_bytes() as u64,
            ));
        }
        Ok(())
    }

    fn pop_pending_scope(&mut self) {
        if self.pending_pop {
            self.pop_scope();
            self.pending_pop = false;
        }
    }

    fn resolve_event<'event>(
        &mut self,
        event: Event<'event>,
        offset: u64,
    ) -> Result<(ResolveResult<'_>, Event<'event>)> {
        match &event {
            Event::Start(element) => self.push(element)?,
            Event::Empty(element) => {
                self.push(element)?;
                self.pending_pop = true;
            },
            Event::End(_) => {
                self.pending_pop = true;
            },
            _ => {},
        }
        let resolved = self.resolver.resolve_event(event);
        if self.document.enabled {
            let allow_unknown_namespaces = self.document.allow_unknown_namespaces;
            validate_document_event(
                &mut self.document,
                &self.resolver,
                &mut self.attributes,
                self.limits,
                allow_unknown_namespaces,
                &resolved.0,
                &resolved.1,
                offset,
            )?;
        }
        Ok(resolved)
    }

    fn push(&mut self, element: &BytesStart<'_>) -> Result<()> {
        let attempted_depth = self.scope_bindings.len().saturating_add(1);
        if attempted_depth > self.limits.max_depth() {
            return Err(limit_error(
                RawXmlDimension::Depth,
                attempted_depth as u64,
                self.limits.max_depth() as u64,
            ));
        }
        if attempted_depth > self.limits.max_namespace_scope_depth() {
            return Err(limit_error(
                RawXmlDimension::NamespaceScopeDepth,
                attempted_depth as u64,
                self.limits.max_namespace_scope_depth() as u64,
            ));
        }
        self.scope_bindings.try_reserve(1).map_err(|source| {
            Error::from(litchi_core::Error::Allocation {
                resource: "ODF namespace scope",
                source,
            })
        })?;
        self.scope_bytes.try_reserve(1).map_err(|source| {
            Error::from(litchi_core::Error::Allocation {
                resource: "ODF namespace scope bytes",
                source,
            })
        })?;
        let next_level = self.resolver.level().checked_add(1).ok_or_else(|| {
            limit_error(
                RawXmlDimension::NamespaceScopeDepth,
                attempted_depth as u64,
                self.limits.max_namespace_scope_depth() as u64,
            )
        })?;
        self.resolver.set_level(next_level);
        let mut declarations = 0usize;
        let mut declaration_bytes = 0usize;
        let mut attributes = element.attributes();
        // With document validation on, its expanded-name check refuses every
        // duplicate attribute (two equal qualified names have one expanded
        // name), so the qualified-name check is not repeated here
        // (record 0377a).
        attributes.with_checks(self.check_attribute_errors && !self.document.enabled);
        for attribute in attributes {
            let attribute = match attribute {
                Ok(attribute) => attribute,
                Err(_error) if !self.check_attribute_errors => break,
                Err(error) => {
                    return Err(Error::from(litchi_core::Error::InvalidFormat(format!(
                        "XML attribute parsing error: {error}"
                    ))));
                },
            };
            let observed_attributes = self.attributes.saturating_add(1);
            if observed_attributes > self.limits.max_attributes() {
                return Err(limit_error(
                    RawXmlDimension::Attributes,
                    observed_attributes as u64,
                    self.limits.max_attributes() as u64,
                ));
            }
            self.attributes = observed_attributes;
            let Some(prefix) = attribute.key.as_namespace_binding() else {
                continue;
            };
            let observed_declarations = declarations.saturating_add(1);
            if observed_declarations > self.limits.max_namespace_declarations() {
                return Err(limit_error(
                    RawXmlDimension::NamespaceDeclarations,
                    observed_declarations as u64,
                    self.limits.max_namespace_declarations() as u64,
                ));
            }
            declarations = observed_declarations;
            let prefix_bytes = match prefix {
                PrefixDeclaration::Default => None,
                PrefixDeclaration::Named(value) => {
                    if value.is_empty() {
                        return Err(Error::from(litchi_core::Error::InvalidFormat(
                            "ODF namespace declaration has an empty prefix".to_string(),
                        )));
                    }
                    if value.len() > self.limits.max_namespace_prefix_bytes() {
                        return Err(limit_error(
                            RawXmlDimension::NamespacePrefixBytes,
                            value.len() as u64,
                            self.limits.max_namespace_prefix_bytes() as u64,
                        ));
                    }
                    Some(value)
                },
            };
            let active_bindings =
                self.active_bindings
                    .checked_add(declarations)
                    .ok_or_else(|| {
                        limit_error(
                            RawXmlDimension::NamespaceBindings,
                            u64::MAX,
                            self.limits.max_namespace_bindings() as u64,
                        )
                    })?;
            if active_bindings > self.limits.max_namespace_bindings() {
                return Err(limit_error(
                    RawXmlDimension::NamespaceBindings,
                    active_bindings as u64,
                    self.limits.max_namespace_bindings() as u64,
                ));
            }
            // Normalize from the lexical bytes before constructing the
            // resolver namespace.  In particular, this admits escaped URI
            // aliases while preserving the original event bytes and avoids
            // an unbounded `Attribute::normalized_value` allocation.
            let normalized = crate::namespace::normalize_namespace_uri(
                attribute.value.as_ref(),
                self.reader.decoder(),
                "ODF namespace declaration",
                self.limits.max_namespace_uri_bytes(),
            )?;
            let binding_bytes = prefix_bytes
                .map_or(0, |prefix| prefix.len())
                .checked_add(normalized.len())
                .ok_or_else(|| {
                    limit_error(
                        RawXmlDimension::NamespaceBytes,
                        u64::MAX,
                        self.limits.max_namespace_bytes() as u64,
                    )
                })?;
            declaration_bytes = declaration_bytes
                .checked_add(binding_bytes)
                .ok_or_else(|| {
                    limit_error(
                        RawXmlDimension::NamespaceBytes,
                        u64::MAX,
                        self.limits.max_namespace_bytes() as u64,
                    )
                })?;
            let active_bytes = self
                .active_bytes
                .checked_add(declaration_bytes)
                .ok_or_else(|| {
                    limit_error(
                        RawXmlDimension::NamespaceBytes,
                        u64::MAX,
                        self.limits.max_namespace_bytes() as u64,
                    )
                })?;
            if active_bytes > self.limits.max_namespace_bytes() {
                return Err(limit_error(
                    RawXmlDimension::NamespaceBytes,
                    active_bytes as u64,
                    self.limits.max_namespace_bytes() as u64,
                ));
            }
            if !(self.allow_empty_prefixed_namespace
                && prefix_bytes.is_some()
                && normalized.is_empty())
            {
                crate::namespace::validate_normalized_namespace_binding(
                    prefix_bytes,
                    normalized.as_ref(),
                    "ODF namespace declaration",
                )?;
            }
            self.resolver
                .add(prefix, Namespace(normalized.as_bytes()))
                .map_err(|error| {
                    Error::from(litchi_core::Error::InvalidFormat(format!(
                        "ODF namespace declaration: {error}"
                    )))
                })?;
        }
        self.active_bindings += declarations;
        self.active_bytes += declaration_bytes;
        self.scope_bindings.push(declarations);
        self.scope_bytes.push(declaration_bytes);
        Ok(())
    }

    fn pop_scope(&mut self) {
        self.resolver.pop();
        if let Some(declarations) = self.scope_bindings.pop() {
            self.active_bindings = self.active_bindings.saturating_sub(declarations);
        }
        if let Some(bytes) = self.scope_bytes.pop() {
            self.active_bytes = self.active_bytes.saturating_sub(bytes);
        }
    }
}

fn validate_document_event(
    state: &mut DocumentValidation,
    resolver: &NamespaceResolver,
    attributes: &mut usize,
    limits: XmlStreamLimits,
    allow_unknown_namespaces: bool,
    namespace: &ResolveResult<'_>,
    event: &Event<'_>,
    offset: u64,
) -> Result<()> {
    validate_event_bytes(event, offset)?;
    match event {
        Event::Start(element) => {
            if state.root_closed {
                return Err(invalid_xml(offset, "content appears after the XML root"));
            }
            if state.depth >= limits.max_depth() {
                return Err(limit_error(
                    RawXmlDimension::Depth,
                    state.depth.saturating_add(1) as u64,
                    limits.max_depth() as u64,
                ));
            }
            if state.depth == 0 {
                if state.root_seen {
                    return Err(invalid_xml(offset, "XML document has more than one root"));
                }
                state.root_seen = true;
            }
            validate_start_event(
                resolver,
                namespace,
                element,
                offset,
                allow_unknown_namespaces,
            )?;
            state.depth += 1;
        },
        Event::Empty(element) => {
            if state.root_closed {
                return Err(invalid_xml(offset, "content appears after the XML root"));
            }
            if state.depth >= limits.max_depth() {
                return Err(limit_error(
                    RawXmlDimension::Depth,
                    state.depth.saturating_add(1) as u64,
                    limits.max_depth() as u64,
                ));
            }
            if state.depth == 0 {
                if state.root_seen {
                    return Err(invalid_xml(offset, "XML document has more than one root"));
                }
                state.root_seen = true;
                state.root_closed = true;
            }
            validate_start_event(
                resolver,
                namespace,
                element,
                offset,
                allow_unknown_namespaces,
            )?;
        },
        Event::End(element) => {
            if state.depth == 0 {
                return Err(invalid_xml(offset, "XML document has an unmatched end tag"));
            }
            validate_qname(element.name().as_ref(), offset, "end element")?;
            if !allow_unknown_namespaces {
                reject_unknown_namespace(namespace, offset)?;
            }
            state.depth -= 1;
            if state.depth == 0 {
                state.root_closed = true;
            }
        },
        Event::Decl(declaration) => {
            if state.saw_any_event || state.declaration_seen || state.root_seen || state.depth != 0
            {
                return Err(invalid_xml(
                    offset,
                    "XML declaration is not the first document event",
                ));
            }
            state.declaration_seen = true;
            validate_declaration(resolver, declaration, attributes, limits, offset)?;
        },
        Event::DocType(_) => {
            return Err(invalid_xml(
                offset,
                "DTD and DOCTYPE declarations are prohibited",
            ));
        },
        Event::GeneralRef(reference) => {
            if state.depth == 0 || state.root_closed {
                return Err(invalid_xml(
                    offset,
                    "entity reference appears outside the XML root",
                ));
            }
            if !crate::validation::valid_xml_reference(reference) {
                return Err(invalid_xml(
                    offset,
                    "undeclared or invalid XML entity reference",
                ));
            }
        },
        Event::CData(_) => {
            if state.depth == 0 || state.root_closed {
                return Err(invalid_xml(offset, "CDATA appears outside the XML root"));
            }
        },
        Event::Text(text) => {
            let bytes: &[u8] = text;
            if state.depth == 0
                && !bytes
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
            {
                return Err(invalid_xml(
                    offset,
                    "non-whitespace text appears outside the XML root",
                ));
            }
        },
        Event::Comment(_) | Event::PI(_) => {
            if let Event::PI(instruction) = event {
                let target = instruction.target();
                validate_name(target, offset, "processing-instruction target")?;
                if target.eq_ignore_ascii_case(b"xml") {
                    return Err(invalid_xml(
                        offset,
                        "processing-instruction target 'xml' is reserved",
                    ));
                }
            }
        },
        Event::Eof => {
            if !state.root_seen || !state.root_closed || state.depth != 0 {
                return Err(invalid_xml(
                    offset,
                    "XML document has no single complete root element",
                ));
            }
            return Ok(());
        },
    }
    state.saw_any_event = true;
    Ok(())
}

fn validate_start_event(
    resolver: &NamespaceResolver,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    offset: u64,
    allow_unknown_namespaces: bool,
) -> Result<()> {
    validate_qname(element.name().as_ref(), offset, "element")?;
    if element
        .name()
        .prefix()
        .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
    {
        return Err(invalid_xml(
            offset,
            "the xmlns prefix is reserved for namespace declarations",
        ));
    }
    if !allow_unknown_namespaces {
        reject_unknown_namespace(namespace, offset)?;
    }

    // Expanded names seen on this element, compared linearly: elements have
    // few attributes, and the borrowed keys avoid the per-attribute
    // allocations and hashing of a set (record 0377a). The duplicate rule is
    // unchanged.
    let mut expanded_names: Vec<(ExpandedNamespace<'_>, &[u8])> = Vec::new();
    let mut expanded_set: Option<HashSet<(ExpandedNamespace<'_>, &[u8])>> = None;
    // quick-xml's duplicate check compares qualified names; the
    // expanded-name check below refuses every such duplicate as well, so
    // the attributes are read without it (record 0377a).
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute
            .map_err(|error| invalid_xml(offset, format!("malformed XML attribute: {error}")))?;
        validate_qname(attribute.key.as_ref(), offset, "attribute")?;
        let value = attribute.value.as_ref();
        // One vectorizable pass classifies the value (record 0377a); the
        // checks below then run only where it found something, in the same
        // order and with the same errors as the four scans they replace.
        let mut non_ascii = false;
        let mut special = false;
        for &byte in value {
            non_ascii |= byte >= 0x80;
            special |= (byte < 0x20) | (byte == b'<') | (byte == b'&');
        }
        if non_ascii {
            str::from_utf8(value).map_err(|error| {
                invalid_xml(offset, format!("XML attribute is not UTF-8: {error}"))
            })?;
        }
        if special || non_ascii {
            if value.contains(&b'<') {
                return Err(invalid_xml(
                    offset,
                    "XML attribute values cannot contain an unescaped '<'",
                ));
            }
            if value
                .iter()
                .any(|byte| !matches!(*byte, 0x09 | 0x0A | 0x0D | 0x20..=0x7F) && *byte < 0x80)
            {
                return Err(invalid_xml(
                    offset,
                    "XML attribute contains a forbidden control",
                ));
            }
            validate_references(value, offset)?;
        }
        let (resolved, local) = resolver.resolve_attribute(attribute.key);
        let namespace = match resolved {
            ResolveResult::Bound(Namespace(uri)) => ExpandedNamespace::Bound(uri),
            ResolveResult::Unbound => ExpandedNamespace::Unbound,
            ResolveResult::Unknown(prefix) if allow_unknown_namespaces => {
                // Keep unknown inherited prefixes in a separate key space so
                // duplicate checking remains meaningful without pretending
                // that an opaque prefix has a resolved URI.
                let mut namespace = Vec::new();
                namespace.try_reserve(prefix.len() + 1).map_err(|source| {
                    Error::from(litchi_core::Error::Allocation {
                        resource: "ODF expanded attribute namespace",
                        source,
                    })
                })?;
                namespace.push(0);
                namespace.extend_from_slice(prefix.as_ref());
                ExpandedNamespace::Unknown(namespace)
            },
            ResolveResult::Unknown(_) => {
                reject_unknown_namespace(&resolved, offset)?;
                unreachable!("unknown namespace was rejected")
            },
        };
        let local = local.into_inner();
        let duplicate = match &mut expanded_set {
            Some(set) => !set.insert((namespace, local)),
            None => {
                let duplicate = expanded_names
                    .iter()
                    .any(|(seen, seen_local)| *seen_local == local && *seen == namespace);
                if !duplicate {
                    expanded_names.try_reserve(1).map_err(|source| {
                        Error::from(litchi_core::Error::Allocation {
                            resource: "ODF expanded attribute names",
                            source,
                        })
                    })?;
                    expanded_names.push((namespace, local));
                    // Many attributes: keep the comparison count linear.
                    if expanded_names.len() > LINEAR_EXPANDED_NAMES {
                        expanded_set = Some(expanded_names.drain(..).collect());
                    }
                }
                duplicate
            },
        };
        if duplicate {
            return Err(invalid_xml(offset, "duplicate XML attribute expanded name"));
        }
    }
    Ok(())
}

/// Attributes compared pairwise before the duplicate check switches to a set.
const LINEAR_EXPANDED_NAMES: usize = 16;

/// The namespace part of an attribute's expanded name, as the duplicate
/// check compares it: a bound URI, no namespace, or (when unknown prefixes
/// are admitted) the opaque prefix in its own key space.
#[derive(PartialEq, Eq, Hash)]
enum ExpandedNamespace<'r> {
    Bound(&'r [u8]),
    Unbound,
    Unknown(Vec<u8>),
}

fn validate_declaration(
    resolver: &NamespaceResolver,
    declaration: &BytesDecl<'_>,
    attributes: &mut usize,
    limits: XmlStreamLimits,
    offset: u64,
) -> Result<()> {
    declaration
        .xml_version()
        .map_err(|error| invalid_xml(offset, format!("invalid XML declaration: {error}")))?;
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(|error| {
            invalid_xml(offset, format!("invalid XML declaration encoding: {error}"))
        })?;
        if !encoding.eq_ignore_ascii_case(b"utf-8") {
            return Err(invalid_xml(
                offset,
                "XML declaration names an encoding other than UTF-8",
            ));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(|error| {
            invalid_xml(
                offset,
                format!("invalid XML declaration standalone: {error}"),
            )
        })?;
        if !matches!(standalone.as_ref(), b"yes" | b"no") {
            return Err(invalid_xml(
                offset,
                "XML declaration standalone must be 'yes' or 'no'",
            ));
        }
    }

    let content = str::from_utf8(declaration.as_ref())
        .map_err(|error| invalid_xml(offset, format!("XML declaration is not UTF-8: {error}")))?;
    let declaration_start = BytesStart::from_content(content, 3);
    let mut order = 0u8;
    for attribute in declaration_start.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid_xml(offset, format!("malformed XML declaration: {error}")))?;
        let observed = (*attributes).saturating_add(1);
        *attributes = observed;
        if observed > limits.max_attributes() {
            return Err(limit_error(
                RawXmlDimension::Attributes,
                observed as u64,
                limits.max_attributes() as u64,
            ));
        }
        validate_qname(attribute.key.as_ref(), offset, "declaration attribute")?;
        if attribute.key.prefix().is_some() {
            return Err(invalid_xml(
                offset,
                "XML declaration attributes must be unprefixed",
            ));
        }
        order = match (order, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return Err(invalid_xml(
                    offset,
                    "XML declaration attributes are missing, duplicated, or out of order",
                ));
            },
        };
        let (resolved, _) = resolver.resolve_attribute(attribute.key);
        reject_unknown_namespace(&resolved, offset)?;
        validate_references(attribute.value.as_ref(), offset)?;
    }
    if order == 0 {
        return Err(invalid_xml(
            offset,
            "XML declaration is missing its version",
        ));
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum RawXmlDimension {
    InputBytes,
    Depth,
    Events,
    Attributes,
    TokenBytes,
    NamespaceScopeDepth,
    NamespaceDeclarations,
    NamespaceBindings,
    NamespaceBytes,
    NamespacePrefixBytes,
}

fn limit_error(dimension: RawXmlDimension, observed: u64, maximum: u64) -> Error {
    Error::Limit {
        kind: LimitKind::RawXml(dimension.kind()),
        observed,
        maximum,
    }
}

impl RawXmlDimension {
    const fn kind(self) -> RawXmlLimitKind {
        match self {
            Self::InputBytes => RawXmlLimitKind::InputBytes,
            Self::Depth => RawXmlLimitKind::Depth,
            Self::Events => RawXmlLimitKind::Events,
            Self::Attributes => RawXmlLimitKind::Attributes,
            Self::TokenBytes => RawXmlLimitKind::TokenBytes,
            Self::NamespaceScopeDepth => RawXmlLimitKind::NamespaceScopeDepth,
            Self::NamespaceDeclarations => RawXmlLimitKind::NamespaceDeclarations,
            Self::NamespaceBindings => RawXmlLimitKind::NamespaceBindings,
            Self::NamespaceBytes => RawXmlLimitKind::NamespaceBytes,
            Self::NamespacePrefixBytes => RawXmlLimitKind::NamespacePrefixBytes,
        }
    }
}

fn map_reader_error(error: quick_xml::Error) -> Error {
    match error {
        quick_xml::Error::Io(source) => Error::from(litchi_core::Error::Io(io::Error::new(
            source.kind(),
            source.to_string(),
        ))),
        quick_xml::Error::Encoding(source) => Error::from(litchi_core::Error::InvalidFormat(
            format!("XML encoding error: {source}"),
        )),
        other => Error::from(litchi_core::Error::InvalidFormat(format!(
            "XML parsing error: {other}"
        ))),
    }
}

fn invalid_xml(offset: u64, message: impl Into<String>) -> Error {
    Error::from(super::stream_xml::invalid_xml(offset, message))
}
