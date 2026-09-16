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
use quick_xml::events::attributes::Attribute;
use quick_xml::name::{
    LocalName, Namespace, NamespaceError, NamespaceResolver, Prefix, PrefixDeclaration, QName,
    ResolveResult,
};
use std::borrow::Cow;
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

    /// How many declarations the tracker currently holds, including the two
    /// predefined reserved bindings.
    ///
    /// Callers that bound namespace growth compare this against their own
    /// limit after each [`push`](Self::push).
    #[must_use]
    pub fn declaration_count(&self) -> usize {
        self.bindings.len()
    }

    /// Visit every binding in scope, innermost first and each prefix at most
    /// once, as `(prefix, namespace)` with an empty prefix for the default
    /// namespace.
    ///
    /// The reserved `xml` and `xmlns` prefixes are never visited: re-declaring
    /// the first is redundant and re-declaring the second is forbidden. A
    /// prefix whose innermost declaration undeclares it (`xmlns=""`) is not
    /// visited either, and neither is the outer binding it hides, so the
    /// visited set is exactly what a fragment has to carry to resolve every
    /// name the way it resolves here.
    pub fn for_each_in_scope(&self, mut visit: impl FnMut(&[u8], &[u8])) {
        let mut seen: Vec<&[u8]> = Vec::new();
        for binding in self.bindings.iter().rev() {
            let prefix = binding.prefix(&self.buffer).unwrap_or_default();
            if matches!(prefix, b"xml" | b"xmlns") {
                continue;
            }
            if seen.contains(&prefix) {
                continue;
            }
            seen.push(prefix);
            if binding.value_len == 0 {
                continue;
            }
            visit(prefix, binding.value(&self.buffer));
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

/// Every namespace binding in scope in one quick-xml resolver, innermost first
/// and each prefix at most once, as `(prefix, namespace)` with an empty prefix
/// for the default namespace.
///
/// `NamespaceResolver::bindings` yields declarations outermost first and keeps
/// the ones an inner element shadows, so a caller that re-declares them
/// verbatim would emit the same prefix twice and pick the wrong binding. This
/// resolves the shadowing once, and drops a prefix whose innermost declaration
/// undeclares it (`xmlns=""`), so what it returns is exactly what a fragment
/// has to carry for every name in it to resolve the way it resolves here.
#[must_use]
pub fn in_scope_declarations(resolver: &NamespaceResolver) -> Vec<(Box<[u8]>, Box<[u8]>)> {
    let mut bindings: Vec<(Box<[u8]>, Box<[u8]>)> = Vec::new();
    for (prefix, namespace) in resolver.bindings() {
        let prefix: Box<[u8]> = match prefix {
            PrefixDeclaration::Default => Box::default(),
            PrefixDeclaration::Named(named) => named.into(),
        };
        let namespace: Box<[u8]> = namespace.as_ref().into();
        match bindings
            .iter()
            .position(|(candidate, _)| *candidate == prefix)
        {
            Some(index) => bindings[index] = (prefix, namespace),
            None => bindings.push((prefix, namespace)),
        }
    }
    bindings.retain(|(_, namespace)| !namespace.is_empty());
    bindings
}

/// `element` with every declaration of `bindings` it does not make itself
/// added to its own start tag.
///
/// This is the re-serializing counterpart of
/// [`mce::InScopeNamespaces::make_self_contained`](crate::mce::InScopeNamespaces::make_self_contained):
/// a capture that rebuilds a fragment through a `Writer` adds the inherited
/// declarations here, at the one element that becomes the fragment's root, so
/// the retained markup resolves standalone.
#[must_use]
pub fn with_in_scope_namespaces<'element>(
    element: &BytesStart<'element>,
    bindings: &[(Box<[u8]>, Box<[u8]>)],
) -> BytesStart<'element> {
    if bindings.is_empty() {
        return element.clone();
    }
    let mut declared: Vec<&[u8]> = Vec::new();
    for attribute in element.attributes().with_checks(false) {
        let Ok(attribute) = attribute else {
            break;
        };
        match attribute.key.as_namespace_binding() {
            Some(PrefixDeclaration::Default) => declared.push(b""),
            Some(PrefixDeclaration::Named(prefix)) => declared.push(prefix),
            None => {},
        }
    }
    let mut rewritten = element.clone();
    for (prefix, namespace) in bindings {
        if declared.contains(&prefix.as_ref()) {
            continue;
        }
        let mut key = Vec::from(b"xmlns".as_slice());
        if !prefix.is_empty() {
            key.push(b':');
            key.extend_from_slice(prefix);
        }
        rewritten.push_attribute(Attribute {
            key: QName(&key),
            value: Cow::Owned(escape_attribute_value(namespace)),
        });
    }
    rewritten.into_owned()
}

/// Quote a namespace URI as a double-quoted attribute value.
///
/// A resolver keeps a declaration's bytes exactly as the source wrote them
/// between its quotes, so they are already in attribute-value form and a `&`
/// must not be escaped a second time. Only a literal `"`, legal inside a
/// single-quoted source declaration, has to change.
fn escape_attribute_value(value: &[u8]) -> Vec<u8> {
    let mut escaped = Vec::with_capacity(value.len());
    for byte in value {
        if *byte == b'"' {
            escaped.extend_from_slice(b"&quot;");
        } else {
            escaped.push(*byte);
        }
    }
    escaped
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
