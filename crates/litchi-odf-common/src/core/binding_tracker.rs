//! Namespace-binding tracking over a plain quick-xml
//! [`quick_xml::reader::Reader`].
//!
//! `NsReader` combines tokenization with namespace maintenance.  The bounded
//! ODF scanners use a borrowing plain `Reader` instead, and drive this small
//! tracker when they need resolved element names.  The implementation mirrors
//! quick-xml 0.41's `NamespaceResolver`, including declaration ordering,
//! reserved-prefix errors, malformed-attribute handling, and deferred scope
//! pops.
//!
//! quick-xml stores its nesting level as a `u16`; this tracker uses a checked
//! `u32` counter so a deeply nested input cannot trigger that implementation's
//! debug overflow panic or release-mode wraparound.

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
/// quick-xml's error enum through the common crate's private hook.  The depth
/// failure keeps a direct caller from turning the checked level increment
/// into a panic.
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

/// Namespace state used by the hidden borrowing ODF scanners.
///
/// This type is reachable only through `core::private`.  It intentionally
/// mirrors quick-xml's resolver rather than implementing a simplified XML
/// namespace model:
///
/// - declarations are scanned in attribute order, with malformed attributes
///   stopping the scan silently;
/// - more than 256 declarations on one element is rejected;
/// - reserved-prefix errors retain quick-xml's precedence and display text;
/// - emptied bindings shadow outer bindings until the scope is popped; and
/// - `pop` is caller-deferred until the next read, so an end event remains in
///   its own namespace scope.
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

    /// Apply declarations on one `Start` or `Empty` event.
    ///
    /// The scan is deliberately the same `with_checks(false)` scan used by
    /// quick-xml's `NamespaceResolver::push`: malformed attributes stop the
    /// declaration scan silently, while namespace errors are returned
    /// immediately.  The cheap raw-attribute prefilter keeps declaration-free
    /// elements on the fast path.
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
    ///
    /// Callers intentionally defer this operation until the next reader
    /// iteration, matching quick-xml's `NsReader` behavior for end and empty
    /// events.
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

    const XML_URI: &str = "http://www.w3.org/XML/1998/namespace";
    const XMLNS_URI: &str = "http://www.w3.org/2000/xmlns/";

    fn push_root(xml: &str) -> Result<(), BindingTrackerError> {
        let mut reader = Reader::from_str(xml);
        match reader.read_event().expect("namespace fixture is readable") {
            Event::Start(element) => BindingTracker::new().push(&element),
            event => panic!("expected a root start event, got {event:?}"),
        }
    }

    fn expected_namespace(error: NamespaceError) -> BindingTrackerError {
        BindingTrackerError::Namespace(error.to_string())
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
            Err(expected_namespace(NamespaceError::TooManyDeclarations(
                MAX_NS_DECLARATIONS_PER_ELEMENT,
            )))
        );
    }

    #[test]
    fn reserved_namespace_bindings_keep_quick_xml_errors() {
        assert_eq!(
            push_root(r#"<root xmlns:xml="urn:wrong">"#),
            Err(expected_namespace(NamespaceError::InvalidXmlPrefixBind(
                b"urn:wrong".to_vec(),
            )))
        );
        assert_eq!(
            push_root(r#"<root xmlns:xmlns="urn:wrong">"#),
            Err(expected_namespace(NamespaceError::InvalidXmlnsPrefixBind(
                b"urn:wrong".to_vec(),
            )))
        );
        assert_eq!(
            push_root(&format!(r#"<root xmlns:foo="{XML_URI}">"#)),
            Err(expected_namespace(NamespaceError::InvalidPrefixForXml(
                b"foo".to_vec(),
            )))
        );
        assert_eq!(
            push_root(&format!(r#"<root xmlns:foo="{XMLNS_URI}">"#)),
            Err(expected_namespace(NamespaceError::InvalidPrefixForXmlns(
                b"foo".to_vec(),
            )))
        );
    }

    fn resolved_uri(tracker: &BindingTracker, name: QName<'_>) -> Option<Vec<u8>> {
        match tracker.resolve_element(name).0 {
            ResolveResult::Bound(Namespace(uri)) => Some(uri.to_vec()),
            ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
        }
    }

    #[test]
    fn scope_and_deferred_pops_restore_outer_bindings() {
        let xml = r#"<root xmlns:p="urn:outer"><p:child xmlns:p="urn:inner"><p:leaf/></p:child><p:after/></root>"#;
        let mut reader = Reader::from_str(xml);
        let mut tracker = BindingTracker::new();
        let mut pending_pop = false;
        let mut observed = Vec::new();

        loop {
            if pending_pop {
                tracker.pop();
                pending_pop = false;
            }
            let event = reader.read_event().expect("scope fixture is readable");
            match &event {
                Event::Start(element) => {
                    tracker.push(element).expect("scope declaration is valid");
                    observed.push(resolved_uri(&tracker, element.name()));
                },
                Event::Empty(element) => {
                    tracker.push(element).expect("scope declaration is valid");
                    observed.push(resolved_uri(&tracker, element.name()));
                    pending_pop = true;
                },
                Event::End(_) => pending_pop = true,
                Event::Eof => break,
                _ => {},
            }
        }

        assert_eq!(
            observed,
            vec![
                None,
                Some(b"urn:inner".to_vec()),
                Some(b"urn:inner".to_vec()),
                Some(b"urn:outer".to_vec()),
            ]
        );
    }

    #[test]
    fn unbinding_shadows_outer_prefix_until_scope_pop() {
        #[derive(Debug, Eq, PartialEq)]
        enum OwnedResolution {
            Bound(Vec<u8>),
            Unbound,
            Unknown,
        }

        fn owned_resolution(result: ResolveResult<'_>) -> OwnedResolution {
            match result {
                ResolveResult::Bound(Namespace(uri)) => OwnedResolution::Bound(uri.to_vec()),
                ResolveResult::Unbound => OwnedResolution::Unbound,
                ResolveResult::Unknown(_) => OwnedResolution::Unknown,
            }
        }

        let xml = r#"<root xmlns:p="urn:outer"><p:child xmlns:p=""><p:inside/></p:child><p:after/></root>"#;
        let mut reader = Reader::from_str(xml);
        let mut tracker = BindingTracker::new();
        let mut pending_pop = false;
        let mut observed = Vec::new();

        loop {
            if pending_pop {
                tracker.pop();
                pending_pop = false;
            }
            let event = reader.read_event().expect("unbinding fixture is readable");
            match &event {
                Event::Start(element) => {
                    tracker
                        .push(element)
                        .expect("unbinding declaration is valid");
                    observed.push(owned_resolution(
                        tracker.resolve_prefix(element.name().prefix()),
                    ));
                },
                Event::Empty(element) => {
                    tracker.push(element).expect("empty scope is valid");
                    observed.push(owned_resolution(
                        tracker.resolve_prefix(element.name().prefix()),
                    ));
                    pending_pop = true;
                },
                Event::End(_) => pending_pop = true,
                Event::Eof => break,
                _ => {},
            }
        }

        assert_eq!(
            observed,
            vec![
                OwnedResolution::Unbound,
                OwnedResolution::Unknown,
                OwnedResolution::Unknown,
                OwnedResolution::Bound(b"urn:outer".to_vec()),
            ]
        );
    }

    #[test]
    fn depths_above_quick_xml_u16_limit_drain_without_panic() {
        let element = BytesStart::new("x");
        let mut tracker = BindingTracker::new();
        for _ in 0..=u16::MAX {
            tracker
                .push(&element)
                .expect("u32 depth accepts 65,536 scopes");
        }
        for _ in 0..=u16::MAX {
            tracker.pop();
        }
        tracker
            .push(&element)
            .expect("tracker remains usable after deep scope drain");
    }
}
