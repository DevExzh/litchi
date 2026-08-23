//! Private namespace-binding tracking for the borrowing OOXML scanners.
//!
//! `quick_xml::reader::NsReader` performs two independent operations for each
//! event: the underlying `Reader` tokenizes the bytes, and its resolver keeps
//! the in-scope namespace bindings up to date.  The small scanners which only
//! need a few resolved element names can avoid the resolver's extra event
//! plumbing by driving a plain `Reader` and using this tracker instead.
//!
//! The implementation intentionally mirrors quick-xml 0.41's
//! `NamespaceResolver` rather than implementing a simplified XML namespace
//! model.  In particular, declaration ordering, reserved-prefix errors,
//! malformed-attribute handling, and deferred scope pops are observable by
//! callers and are therefore part of this module's contract.
//!
//! quick-xml 0.41 stores its nesting level as a `u16`; this tracker uses a
//! checked `u32` counter instead. The owner scanners impose a much smaller
//! structural depth bound, while a direct hidden-plumbing caller receives a
//! checked error instead of inheriting quick-xml's overflow panic/wrap path.

use memchr::memmem;
use quick_xml::events::BytesStart;
use quick_xml::name::{
    LocalName, Namespace, NamespaceError, Prefix, PrefixDeclaration, QName, ResolveResult,
};
use std::fmt;

const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";
const MAX_NS_DECLARATIONS_PER_ELEMENT: usize = 256;

/// Failure from the hidden namespace-tracker plumbing.
///
/// Namespace failures retain quick-xml's exact display text without exposing
/// quick-xml's error enum through the common crate's private hook. The depth
/// failure is unreachable for the bounded format scanners, but keeps a direct
/// caller from turning the checked level increment into a panic.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BindingTrackerError {
    /// A namespace declaration failed quick-xml's reserved-prefix or count rules.
    Namespace(String),
    /// The internal nesting counter cannot represent another open element.
    DepthOverflow,
}

impl fmt::Display for BindingTrackerError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Namespace(error) => formatter.write_str(error),
            Self::DepthOverflow => formatter.write_str("namespace binding depth exceeds u32"),
        }
    }
}

impl std::error::Error for BindingTrackerError {}

fn namespace_error(error: NamespaceError) -> BindingTrackerError {
    BindingTrackerError::Namespace(error.to_string())
}

#[derive(Debug)]
struct Binding {
    start: usize,
    prefix_len: usize,
    value_len: usize,
    level: u32,
}

impl Binding {
    fn prefix<'buffer>(&self, buffer: &'buffer [u8]) -> Option<&'buffer [u8]> {
        (self.prefix_len != 0).then(|| &buffer[self.start..self.start + self.prefix_len])
    }

    fn value<'buffer>(&self, buffer: &'buffer [u8]) -> &'buffer [u8] {
        &buffer[self.start + self.prefix_len..self.start + self.prefix_len + self.value_len]
    }
}

/// Namespace state used by the private borrowing OOXML scanners.
///
/// This type is unstable implementation plumbing, reachable only through the
/// hidden `private` common-crate namespace. It is intentionally not
/// re-exported by any format facade and does not add a public OOXML model or
/// package type.
#[derive(Debug)]
pub struct BindingTracker {
    buffer: Vec<u8>,
    bindings: Vec<Binding>,
    level: u32,
}

impl BindingTracker {
    /// Create a resolver with quick-xml's two predefined bindings.
    #[must_use]
    pub fn new() -> Self {
        let mut buffer = Vec::new();
        let mut bindings = Vec::new();
        for (prefix, uri) in [
            (b"xml".as_slice(), XML_NAMESPACE),
            (b"xmlns".as_slice(), XMLNS_NAMESPACE),
        ] {
            bindings.push(Binding {
                start: buffer.len(),
                prefix_len: prefix.len(),
                value_len: uri.len(),
                level: 0,
            });
            buffer.extend_from_slice(prefix);
            buffer.extend_from_slice(uri);
        }
        Self {
            buffer,
            bindings,
            level: 0,
        }
    }

    /// Apply the declarations on one `Start` or `Empty` event.
    ///
    /// The scan is deliberately the same `with_checks(false)` scan used by
    /// `NamespaceResolver::push`: malformed attributes stop the declaration
    /// scan silently, while a namespace error is returned immediately.  The
    /// cheap raw-attribute prefilter keeps declaration-free elements on the
    /// fast path.
    ///
    /// A declaration-free raw attribute slice shorter than `xmlns` cannot
    /// contain a namespace declaration.  The out-of-line scan also avoids
    /// putting its iterator and substring search on the common bare-tag path.
    #[inline]
    pub fn push(&mut self, element: &BytesStart<'_>) -> Result<(), BindingTrackerError> {
        self.level = self
            .level
            .checked_add(1)
            .ok_or(BindingTrackerError::DepthOverflow)?;
        if element.attributes_raw().len() < b"xmlns".len() {
            return Ok(());
        }
        self.push_scanned(element)
    }

    fn push_scanned(&mut self, element: &BytesStart<'_>) -> Result<(), BindingTrackerError> {
        if memmem::find(element.attributes_raw(), b"xmlns").is_none() {
            return Ok(());
        }
        let mut count = 0usize;
        for attribute in element.attributes().with_checks(false) {
            let Ok(attribute) = attribute else {
                break;
            };
            if let Some(prefix) = attribute.key.as_namespace_binding() {
                if count >= MAX_NS_DECLARATIONS_PER_ELEMENT {
                    return Err(namespace_error(NamespaceError::TooManyDeclarations(
                        MAX_NS_DECLARATIONS_PER_ELEMENT,
                    )));
                }
                count += 1;
                self.add(prefix, &attribute.value)?;
            }
        }
        Ok(())
    }

    fn add(
        &mut self,
        prefix: PrefixDeclaration<'_>,
        uri: &[u8],
    ) -> Result<(), BindingTrackerError> {
        let level = self.level;
        match prefix {
            PrefixDeclaration::Default => {
                let start = self.buffer.len();
                self.buffer.extend_from_slice(uri);
                self.bindings.push(Binding {
                    start,
                    prefix_len: 0,
                    value_len: uri.len(),
                    level,
                });
            },
            PrefixDeclaration::Named(b"xml") => {
                if uri != XML_NAMESPACE {
                    return Err(namespace_error(NamespaceError::InvalidXmlPrefixBind(
                        uri.to_vec(),
                    )));
                }
            },
            PrefixDeclaration::Named(b"xmlns") => {
                return Err(namespace_error(NamespaceError::InvalidXmlnsPrefixBind(
                    uri.to_vec(),
                )));
            },
            PrefixDeclaration::Named(prefix) => {
                if uri == XML_NAMESPACE {
                    return Err(namespace_error(NamespaceError::InvalidPrefixForXml(
                        prefix.to_vec(),
                    )));
                }
                if uri == XMLNS_NAMESPACE {
                    return Err(namespace_error(NamespaceError::InvalidPrefixForXmlns(
                        prefix.to_vec(),
                    )));
                }
                let start = self.buffer.len();
                self.buffer.extend_from_slice(prefix);
                self.buffer.extend_from_slice(uri);
                self.bindings.push(Binding {
                    start,
                    prefix_len: prefix.len(),
                    value_len: uri.len(),
                    level,
                });
            },
        }
        Ok(())
    }

    /// End the most recently opened namespace scope.
    pub fn pop(&mut self) {
        self.level = self.level.saturating_sub(1);
        match self
            .bindings
            .iter()
            .rposition(|binding| binding.level <= self.level)
        {
            None => {
                self.buffer.clear();
                self.bindings.clear();
            },
            Some(last_kept) => {
                if let Some(len) = self
                    .bindings
                    .get(last_kept + 1)
                    .map(|binding| binding.start)
                {
                    self.buffer.truncate(len);
                    self.bindings.truncate(last_kept + 1);
                }
            },
        }
    }

    /// Resolve an element name with the default namespace enabled.
    #[must_use]
    pub fn resolve_element<'name>(
        &self,
        name: QName<'name>,
    ) -> (ResolveResult<'_>, LocalName<'name>) {
        let (local_name, prefix) = name.decompose();
        (self.resolve_prefix(prefix), local_name)
    }

    /// Resolve an attribute name with the default namespace disabled.
    #[must_use]
    pub fn resolve_attribute<'name>(
        &self,
        name: QName<'name>,
    ) -> (ResolveResult<'_>, LocalName<'name>) {
        let (local_name, prefix) = name.decompose();
        match prefix {
            None => (ResolveResult::Unbound, local_name),
            Some(prefix) => (self.resolve_prefix(Some(prefix)), local_name),
        }
    }

    /// Resolve a prefix using the newest in-scope declaration.
    #[must_use]
    pub fn resolve_prefix(&self, prefix: Option<Prefix<'_>>) -> ResolveResult<'_> {
        let mut bindings = self.bindings.iter().rev();
        match prefix {
            None => match bindings.find(|binding| binding.prefix_len == 0) {
                Some(binding) if binding.value_len != 0 => {
                    ResolveResult::Bound(Namespace(binding.value(&self.buffer)))
                },
                _ => ResolveResult::Unbound,
            },
            Some(prefix) => {
                let prefix = prefix.into_inner();
                match bindings.find(|binding| binding.prefix(&self.buffer) == Some(prefix)) {
                    Some(binding) if binding.value_len != 0 => {
                        ResolveResult::Bound(Namespace(binding.value(&self.buffer)))
                    },
                    _ => ResolveResult::Unknown(prefix.to_vec()),
                }
            },
        }
    }
}

impl Default for BindingTracker {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    fn push_root(xml: &str) -> Result<(), BindingTrackerError> {
        let mut reader = Reader::from_str(xml);
        match reader
            .read_event()
            .expect("namespace-limit fixture is readable")
        {
            Event::Start(element) => BindingTracker::new().push(&element),
            event => panic!("expected a root start event, got {event:?}"),
        }
    }

    #[test]
    fn declaration_limit_accepts_exactly_256_and_rejects_257() {
        let attributes = |count: usize| {
            (0..count).fold(String::new(), |mut attributes, index| {
                attributes.push_str(&format!(r##" xmlns:p{index}="urn:{index}""##));
                attributes
            })
        };

        assert!(push_root(&format!("<root{}>", attributes(256))).is_ok());
        assert_eq!(
            push_root(&format!("<root{}>", attributes(257))),
            Err(BindingTrackerError::Namespace(
                NamespaceError::TooManyDeclarations(MAX_NS_DECLARATIONS_PER_ELEMENT).to_string(),
            ))
        );
    }
}
