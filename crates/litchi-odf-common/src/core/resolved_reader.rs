//! XML event reader with normalized namespace declaration values.
//!
//! `quick_xml::reader::NsReader` resolves namespace declarations before it
//! returns the event.  Its resolver therefore sees the lexical attribute
//! bytes.  XML namespace values are attribute values, however, so character
//! references must be expanded before reserved-prefix checks.  This adapter
//! keeps the original event bytes and source positions while applying the
//! same resolver operations to normalized declaration values.

use quick_xml::{
    errors::Result,
    events::{BytesStart, Event},
    name::{
        DEFAULT_MAX_DECLARATIONS_PER_ELEMENT, Namespace, NamespaceError, NamespaceResolver,
        PrefixDeclaration, ResolveResult,
    },
    reader::Reader,
};

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
}

// A document can legitimately inherit a declaration at each nested element,
// but an untrusted source must not be able to make the resolver's active table
// grow without bound.  This keeps the adapter's retained namespace state
// bounded while leaving the per-element quick-xml limit intact.
const MAX_ACTIVE_NAMESPACE_BINDINGS: usize =
    DEFAULT_MAX_DECLARATIONS_PER_ELEMENT.saturating_mul(1024);
const MAX_ACTIVE_NAMESPACE_BYTES: usize = 16 * 1024 * 1024;
const MAX_NAMESPACE_SCOPE_DEPTH: usize = 1024;

impl<'a> ResolvedReader<'a> {
    /// Construct a reader over one UTF-8 XML source.
    #[must_use]
    pub fn from_xml(source: &'a str) -> Self {
        Self {
            reader: Reader::from_str(source),
            resolver: NamespaceResolver::default(),
            pending_pop: false,
            active_bindings: 0,
            active_bytes: 0,
            scope_bindings: Vec::new(),
            scope_bytes: Vec::new(),
        }
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
        let event = self.reader.read_event_into(buffer)?;
        self.resolve_event(event)
    }

    /// Read and resolve the next event directly from the borrowed XML input.
    ///
    /// This is the zero-copy counterpart to
    /// [`Self::read_resolved_event_into`].  It is intended for source-backed
    /// callers whose event buffer is already the original input slice.
    pub fn read_resolved_event(&mut self) -> Result<(ResolveResult<'_>, Event<'a>)> {
        self.pop_pending_scope();
        let event = self.reader.read_event()?;
        self.resolve_event(event)
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
        if self.scope_bindings.len() >= MAX_NAMESPACE_SCOPE_DEPTH {
            return Err(NamespaceError::TooManyDeclarations(MAX_NAMESPACE_SCOPE_DEPTH).into());
        }
        self.scope_bindings
            .try_reserve(1)
            .map_err(|error| allocation_error("ODF namespace scope", error))?;
        self.scope_bytes
            .try_reserve(1)
            .map_err(|error| allocation_error("ODF namespace scope bytes", error))?;
        let next_level =
            self.resolver
                .level()
                .checked_add(1)
                .ok_or(NamespaceError::TooManyDeclarations(
                    MAX_ACTIVE_NAMESPACE_BINDINGS,
                ))?;
        self.resolver.set_level(next_level);
        let mut declarations = 0usize;
        let mut declaration_bytes = 0usize;
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute?;
            let Some(prefix) = attribute.key.as_namespace_binding() else {
                continue;
            };
            if declarations >= DEFAULT_MAX_DECLARATIONS_PER_ELEMENT {
                return Err(NamespaceError::TooManyDeclarations(
                    DEFAULT_MAX_DECLARATIONS_PER_ELEMENT,
                )
                .into());
            }
            declarations += 1;
            let prefix_bytes = match prefix {
                PrefixDeclaration::Default => None,
                PrefixDeclaration::Named(value) => {
                    if value.is_empty() {
                        return Err(NamespaceError::InvalidPrefixForXml(Vec::new()).into());
                    }
                    if value.len() > crate::namespace::MAX_NAMESPACE_URI_BYTES {
                        return Err(namespace_limit_error(
                            "ODF namespace prefix exceeds the namespace byte limit",
                        ));
                    }
                    Some(value)
                },
            };
            if self
                .active_bindings
                .checked_add(declarations)
                .is_none_or(|value| value > MAX_ACTIVE_NAMESPACE_BINDINGS)
            {
                return Err(
                    NamespaceError::TooManyDeclarations(MAX_ACTIVE_NAMESPACE_BINDINGS).into(),
                );
            }
            // Normalize from the lexical bytes before constructing the
            // resolver namespace.  In particular, this admits escaped URI
            // aliases while preserving the original event bytes and avoids
            // an unbounded `Attribute::normalized_value` allocation.
            let normalized = crate::namespace::normalize_namespace_uri(
                attribute.value.as_ref(),
                self.reader.decoder(),
                "ODF namespace declaration",
            )
            .map_err(namespace_error)?;
            let binding_bytes = prefix_bytes
                .map_or(0, |prefix| prefix.len())
                .checked_add(normalized.len())
                .ok_or_else(|| namespace_limit_error("ODF namespace byte count overflow"))?;
            declaration_bytes = declaration_bytes
                .checked_add(binding_bytes)
                .ok_or_else(|| namespace_limit_error("ODF namespace byte count overflow"))?;
            if self
                .active_bytes
                .checked_add(declaration_bytes)
                .is_none_or(|value| value > MAX_ACTIVE_NAMESPACE_BYTES)
            {
                return Err(namespace_limit_error(
                    "ODF active namespace bindings exceed their byte limit",
                ));
            }
            crate::namespace::validate_normalized_namespace_binding(
                prefix_bytes,
                normalized.as_ref(),
                "ODF namespace declaration",
            )
            .map_err(namespace_error)?;
            self.resolver
                .add(prefix, Namespace(normalized.as_bytes()))?;
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

fn namespace_error(error: litchi_core::Error) -> quick_xml::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidData, error.to_string()).into()
}

fn namespace_limit_error(message: &str) -> quick_xml::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidData, message.to_owned()).into()
}

fn allocation_error(
    resource: &'static str,
    error: std::collections::TryReserveError,
) -> quick_xml::Error {
    std::io::Error::new(
        std::io::ErrorKind::OutOfMemory,
        format!("allocation failed for {resource}: {error}"),
    )
    .into()
}
