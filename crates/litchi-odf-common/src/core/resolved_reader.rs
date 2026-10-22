//! XML event reader with normalized namespace declaration values.
//!
//! `quick_xml::reader::NsReader` resolves namespace declarations before it
//! returns the event.  Its resolver therefore sees the lexical attribute
//! bytes.  XML namespace values are attribute values, however, so character
//! references must be expanded before reserved-prefix checks.  This adapter
//! keeps the original event bytes and source positions while applying the
//! same resolver operations to normalized declaration values.

use litchi_core::{Error, Resource, ResourceLimit, Result};
use quick_xml::{
    events::{BytesStart, Event},
    name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult},
    reader::Reader,
};
use std::{io, sync::Arc};

use super::stream_xml::XmlStreamLimits;

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
}

/// Preserve typed reader failures while adding a stable parser context to
/// malformed-wire errors.  In particular, resource-limit and allocation
/// failures must not be flattened into a generic invalid-format message by a
/// family codec that consumes this reader.
#[must_use]
pub fn map_resolved_reader_error(error: Error, context: &'static str) -> Error {
    match error {
        Error::InvalidFormat(message) => Error::InvalidFormat(format!("{context}: {message}")),
        Error::ParseError(message) => Error::ParseError(format!("{context}: {message}")),
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
        })
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
        self.resolve_event(event)
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
        self.resolve_event(event)
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
        Ok(self.resolver.resolve_event(event))
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
        self.scope_bindings
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODF namespace scope",
                source,
            })?;
        self.scope_bytes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODF namespace scope bytes",
                source,
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
        attributes.with_checks(self.check_attribute_errors);
        for attribute in attributes {
            let attribute = match attribute {
                Ok(attribute) => attribute,
                Err(_error) if !self.check_attribute_errors => break,
                Err(error) => {
                    return Err(Error::InvalidFormat(format!(
                        "XML attribute parsing error: {error}"
                    )));
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
                        return Err(Error::InvalidFormat(
                            "ODF namespace declaration has an empty prefix".to_string(),
                        ));
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
                    Error::InvalidFormat(format!("ODF namespace declaration: {error}"))
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
    let scope = match dimension {
        RawXmlDimension::InputBytes => "ODF raw XML input bytes",
        RawXmlDimension::Depth => "ODF raw XML depth",
        RawXmlDimension::Events => "ODF raw XML events",
        RawXmlDimension::Attributes => "ODF raw XML attributes",
        RawXmlDimension::TokenBytes => "ODF raw XML token bytes",
        RawXmlDimension::NamespaceScopeDepth => "ODF raw XML namespace scope depth",
        RawXmlDimension::NamespaceDeclarations => "ODF raw XML namespace declarations",
        RawXmlDimension::NamespaceBindings => "ODF raw XML namespace bindings",
        RawXmlDimension::NamespaceBytes => "ODF raw XML namespace bytes",
        RawXmlDimension::NamespacePrefixBytes => "ODF raw XML namespace prefix bytes",
    };
    Error::ResourceLimit(ResourceLimit {
        resource: match dimension {
            RawXmlDimension::InputBytes => Resource::InputBytes,
            RawXmlDimension::Depth | RawXmlDimension::NamespaceScopeDepth => Resource::Depth,
            RawXmlDimension::Events
            | RawXmlDimension::Attributes
            | RawXmlDimension::NamespaceDeclarations
            | RawXmlDimension::NamespaceBindings => Resource::Objects,
            RawXmlDimension::TokenBytes
            | RawXmlDimension::NamespaceBytes
            | RawXmlDimension::NamespacePrefixBytes => Resource::Memory,
        },
        observed,
        limit: maximum,
        scope: Arc::from(scope),
    })
}

fn map_reader_error(error: quick_xml::Error) -> Error {
    match error {
        quick_xml::Error::Io(source) => {
            Error::Io(io::Error::new(source.kind(), source.to_string()))
        },
        quick_xml::Error::Encoding(source) => {
            Error::InvalidFormat(format!("XML encoding error: {source}"))
        },
        other => Error::InvalidFormat(format!("XML parsing error: {other}")),
    }
}
