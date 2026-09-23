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
//!
//! Resolution answers exactly what quick-xml's newest-first linear search
//! answers. Change 0754 put a small cache in front of that search: the
//! positions of the few prefixes resolved most recently, and of the innermost
//! default-namespace declaration. Every change to the binding list clears the
//! cache, and a cached position is used only after the binding there is
//! checked to carry the requested prefix, so a hit returns the binding the
//! search would have found. The cache is a fixed-size array compared byte for
//! byte, with no hashing: a lookup costs at most those few comparisons more
//! than the search it fronts, and a document cannot steer it into a worse
//! case than the search's own.

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::Attribute;
use quick_xml::name::{
    LocalName, Namespace, NamespaceError, NamespaceResolver, Prefix, PrefixDeclaration, QName,
    ResolveResult,
};
use std::borrow::Cow;
use std::cell::Cell;
use std::fmt;

const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";
const MAX_NS_DECLARATIONS_PER_ELEMENT: usize = 256;
/// How many recently resolved prefixes the resolution cache remembers.
const RECENT_PREFIXES: usize = 4;

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

    /// Whether this binding declares the non-default `prefix`: exactly
    /// `self.prefix(buffer) == Some(prefix)`.
    ///
    /// The lengths are compared first; a prefix of at most eight bytes, which
    /// covers the prefixes OOXML producers write, is then compared byte by
    /// byte.
    #[inline]
    fn declares(&self, buffer: &[u8], prefix: &[u8]) -> bool {
        if self.prefix_len == 0 || self.prefix_len != prefix.len() {
            return false;
        }
        let Some(stored) = buffer.get(self.start..self.start + self.prefix_len) else {
            return false;
        };
        if stored.len() <= 8 {
            stored.iter().zip(prefix).all(|(left, right)| left == right)
        } else {
            stored == prefix
        }
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
    /// Positions in `bindings` of the innermost declarations of recently
    /// resolved prefixes. Cleared by every change to `bindings`.
    recent: [Cell<Option<usize>>; RECENT_PREFIXES],
    /// The `recent` slot the next cache miss replaces.
    recent_next: Cell<usize>,
    /// Where the innermost default-namespace declaration is, once looked up.
    /// Cleared by every change to `bindings`.
    default: Cell<DefaultLookup>,
}

/// The cached answer to "which binding is the innermost default-namespace
/// declaration".
#[derive(Clone, Copy, Debug)]
enum DefaultLookup {
    /// Not looked up since the binding list last changed.
    Unresolved,
    /// No binding declares the default namespace.
    Absent,
    /// The innermost default-namespace declaration is at this position.
    At(usize),
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
            recent: std::array::from_fn(|_| Cell::new(None)),
            recent_next: Cell::new(0),
            default: Cell::new(DefaultLookup::Unresolved),
        }
    }

    /// Forget every cached resolution; called whenever `bindings` changes.
    fn invalidate_resolutions(&mut self) {
        for slot in &mut self.recent {
            *slot.get_mut() = None;
        }
        *self.default.get_mut() = DefaultLookup::Unresolved;
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
        if !contains_xmlns(element.attributes_raw()) {
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
                self.invalidate_resolutions();
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
                self.invalidate_resolutions();
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
                self.invalidate_resolutions();
            },
            Some(last_kept) => {
                if let Some(len) = self
                    .bindings
                    .get(last_kept + 1)
                    .map(|binding| binding.start)
                {
                    self.buffer.truncate(len);
                    self.bindings.truncate(last_kept + 1);
                    self.invalidate_resolutions();
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
        self.resolve_prefix_bytes(prefix.map(Prefix::into_inner))
    }

    /// Resolve an element name, given as its raw prefix bytes, with the
    /// default namespace enabled.
    ///
    /// This is [`Self::resolve_prefix`] for a caller that has already split
    /// the qualified name with [`split_qualified_name`]: `None` is an
    /// unprefixed name, and `Some` carries the bytes before the first colon,
    /// exactly what `QName::prefix` returns.
    #[must_use]
    pub fn resolve_prefix_bytes(&self, prefix: Option<&[u8]>) -> ResolveResult<'_> {
        match prefix {
            None => match self.innermost_default() {
                Some(binding) if binding.value_len != 0 => {
                    ResolveResult::Bound(Namespace(binding.value(&self.buffer)))
                },
                _ => ResolveResult::Unbound,
            },
            Some(prefix) => match self.innermost_prefixed(prefix) {
                Some(binding) if binding.value_len != 0 => {
                    ResolveResult::Bound(Namespace(binding.value(&self.buffer)))
                },
                _ => ResolveResult::Unknown(prefix.to_vec()),
            },
        }
    }

    /// The newest binding that declares the default namespace, which is the
    /// first one quick-xml's newest-first search meets.
    fn innermost_default(&self) -> Option<&Binding> {
        let position = match self.default.get() {
            DefaultLookup::At(position) => Some(position),
            DefaultLookup::Absent => None,
            DefaultLookup::Unresolved => {
                let found = self
                    .bindings
                    .iter()
                    .rposition(|binding| binding.prefix_len == 0);
                self.default
                    .set(found.map_or(DefaultLookup::Absent, DefaultLookup::At));
                found
            },
        };
        position.and_then(|position| self.bindings.get(position))
    }

    /// The newest binding that declares `prefix`, which is the first one
    /// quick-xml's newest-first search meets.
    ///
    /// A cached position is trusted only after the binding there is checked
    /// to declare `prefix`; the binding list has not changed since it was
    /// cached (every change clears the cache), so no newer declaration of
    /// `prefix` can exist, and the answer is the search's. A miss runs the
    /// search and remembers a found position in the least recently filled
    /// slot. An absent prefix is not cached: it has no binding to point at.
    fn innermost_prefixed(&self, prefix: &[u8]) -> Option<&Binding> {
        for slot in &self.recent {
            if let Some(position) = slot.get()
                && let Some(binding) = self.bindings.get(position)
                && binding.declares(&self.buffer, prefix)
            {
                return Some(binding);
            }
        }
        let position = self
            .bindings
            .iter()
            .rposition(|binding| binding.declares(&self.buffer, prefix))?;
        let slot = self.recent_next.get() % RECENT_PREFIXES;
        if let Some(entry) = self.recent.get(slot) {
            entry.set(Some(position));
        }
        self.recent_next.set((slot + 1) % RECENT_PREFIXES);
        self.bindings.get(position)
    }
}

/// Split a raw qualified name at its first colon, as `QName::decompose` does.
///
/// Returns the prefix bytes (`None` when the name has no colon) and the local
/// name: `QName::prefix` and `QName::local_name` for the same bytes. It is a
/// plain byte scan, so a caller can split a name once and use both halves
/// without a second search.
#[must_use]
pub fn split_qualified_name(name: &[u8]) -> (Option<&[u8]>, &[u8]) {
    match name.iter().position(|byte| *byte == b':') {
        Some(colon) => (name.get(..colon), name.get(colon + 1..).unwrap_or_default()),
        None => (None, name),
    }
}

/// Whether raw attribute bytes contain `xmlns`, which every namespace
/// declaration's key starts with.
///
/// Every key is a slice of the raw attribute bytes, so when this is false no
/// attribute can declare a namespace. A plain window scan with no searcher to
/// prepare, linear in the tag's length for any input.
fn contains_xmlns(raw: &[u8]) -> bool {
    raw.windows(b"xmlns".len()).any(|window| window == b"xmlns")
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

    /// A small deterministic generator for reproducible random documents.
    struct Random(u64);

    impl Random {
        fn below(&mut self, bound: usize) -> usize {
            // xorshift64*
            self.0 ^= self.0 >> 12;
            self.0 ^= self.0 << 25;
            self.0 ^= self.0 >> 27;
            let value = self.0.wrapping_mul(0x2545_F491_4F6C_DD1D);
            usize::try_from(value % u64::try_from(bound.max(1)).unwrap()).unwrap()
        }
    }

    const PREFIXES: &[&str] = &["a", "b", "w", "p1", "p2", "p3", "p4", "p5", "p6", "xml"];
    const NAMESPACES: &[&str] = &[
        "",
        "urn:one",
        "urn:two",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ];

    fn random_name(random: &mut Random) -> String {
        match random.below(12) {
            0 => "plain".to_owned(),
            1 => ":empty-prefix".to_owned(),
            2 => "two:colons:here".to_owned(),
            3 => "trailing:".to_owned(),
            _ => format!("{}:e", PREFIXES[random.below(PREFIXES.len())]),
        }
    }

    fn random_declarations(random: &mut Random, allow_errors: bool) -> String {
        let mut attributes = String::new();
        for _ in 0..random.below(4) {
            let namespace = NAMESPACES[random.below(NAMESPACES.len())];
            match random.below(if allow_errors { 12 } else { 10 }) {
                0 | 1 => attributes.push_str(&format!(" xmlns=\"{namespace}\"")),
                10 => attributes.push_str(" xmlns:xmlns=\"urn:bad\""),
                11 => attributes.push_str(" xmlns:xml=\"urn:bad\""),
                _ => {
                    let prefix = PREFIXES[random.below(PREFIXES.len() - 1)];
                    attributes.push_str(&format!(" xmlns:{prefix}=\"{namespace}\""));
                },
            }
        }
        // A prefixed and an unprefixed ordinary attribute.
        let attribute = PREFIXES[random.below(PREFIXES.len())];
        attributes.push_str(&format!(" {attribute}:at=\"v\" at=\"v\""));
        attributes
    }

    fn random_document(random: &mut Random, allow_errors: bool) -> String {
        fn element(random: &mut Random, depth: usize, allow_errors: bool, out: &mut String) {
            let name = random_name(random);
            let declarations = random_declarations(random, allow_errors);
            if depth >= 6 || random.below(4) == 0 {
                out.push_str(&format!("<{name}{declarations}/>"));
                return;
            }
            out.push_str(&format!("<{name}{declarations}>"));
            for _ in 0..random.below(4) {
                element(random, depth + 1, allow_errors, out);
            }
            out.push_str(&format!("</{name}>"));
        }
        let mut out = String::new();
        element(random, 0, allow_errors, &mut out);
        out
    }

    /// Every resolution, and the first error, of one document walked with
    /// quick-xml's own resolver.
    fn resolver_trace(xml: &str) -> Vec<String> {
        let mut reader = quick_xml::reader::NsReader::from_str(xml);
        let mut trace = Vec::new();
        loop {
            match reader.read_resolved_event() {
                Ok((namespace, event)) => {
                    let namespace = format!("{namespace:?}");
                    let attributes = match &event {
                        Event::Start(element) | Event::Empty(element) => element
                            .attributes()
                            .with_checks(false)
                            .flatten()
                            .map(|attribute| {
                                format!("{:?}", reader.resolver().resolve_attribute(attribute.key))
                            })
                            .collect::<Vec<_>>(),
                        _ => Vec::new(),
                    };
                    trace.push(format!("{namespace} {attributes:?}"));
                    if matches!(event, Event::Eof) {
                        return trace;
                    }
                },
                Err(error) => {
                    trace.push(format!("error {error}"));
                    return trace;
                },
            }
        }
    }

    /// The same walk with a plain reader and the tracker, resolving element
    /// names both through `resolve_element` and through the byte-level
    /// `resolve_prefix_bytes` after `split_qualified_name`.
    fn tracker_trace(xml: &str) -> Vec<String> {
        let mut reader = Reader::from_str(xml);
        let mut tracker = BindingTracker::new();
        let mut pending_pop = false;
        let mut trace = Vec::new();
        loop {
            if pending_pop {
                tracker.pop();
                pending_pop = false;
            }
            let event = match reader.read_event() {
                Ok(event) => event,
                Err(error) => {
                    trace.push(format!("error {error}"));
                    return trace;
                },
            };
            let name = match &event {
                Event::Start(element) | Event::Empty(element) => {
                    if let Err(error) = tracker.push(element) {
                        trace.push(format!("error {error}"));
                        return trace;
                    }
                    pending_pop = matches!(event, Event::Empty(_));
                    Some(element.name())
                },
                Event::End(element) => {
                    pending_pop = true;
                    Some(element.name())
                },
                _ => None,
            };
            let namespace = match name {
                Some(name) => {
                    let resolved = tracker.resolve_element(name).0;
                    let (prefix, local) = split_qualified_name(name.into_inner());
                    assert_eq!(local, name.local_name().into_inner());
                    assert_eq!(
                        format!("{:?}", tracker.resolve_prefix_bytes(prefix)),
                        format!("{resolved:?}"),
                        "byte-level resolution diverged for {:?} in {xml}",
                        String::from_utf8_lossy(name.into_inner())
                    );
                    resolved
                },
                None => ResolveResult::Unbound,
            };
            let attributes = match &event {
                Event::Start(element) | Event::Empty(element) => element
                    .attributes()
                    .with_checks(false)
                    .flatten()
                    .map(|attribute| format!("{:?}", tracker.resolve_attribute(attribute.key)))
                    .collect::<Vec<_>>(),
                _ => Vec::new(),
            };
            trace.push(format!("{namespace:?} {attributes:?}"));
            if matches!(event, Event::Eof) {
                return trace;
            }
        }
    }

    #[test]
    fn cached_resolution_matches_quick_xml_on_random_documents() {
        let mut random = Random(0x0754_B1D5_7AC4_0001);
        let mut resolutions = 0usize;
        let mut errors = 0usize;
        for round in 0..3_000 {
            let xml = random_document(&mut random, round % 5 == 0);
            let expected = resolver_trace(&xml);
            let actual = tracker_trace(&xml);
            assert_eq!(actual, expected, "resolution diverged on {xml}");
            resolutions += expected.len();
            errors += usize::from(
                expected
                    .last()
                    .is_some_and(|last| last.starts_with("error")),
            );
        }
        // The documents shadow, undeclare and fail often enough to matter.
        assert!(
            resolutions > 30_000 && errors > 50,
            "{resolutions} {errors}"
        );
    }

    #[test]
    fn cache_eviction_and_invalidation_follow_the_innermost_declaration() {
        let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let documents = [
            // More prefixes than cache slots, resolved round robin.
            r#"<a:r xmlns:a="urn:a" xmlns:b="urn:b" xmlns:c="urn:c" xmlns:d="urn:d" xmlns:e="urn:e" xmlns:f="urn:f"><a:x/><b:x/><c:x/><d:x/><e:x/><f:x/><a:x/><f:x/><c:x/><e:x/></a:r>"#
                .to_owned(),
            // A cached prefix redeclared deeper, then popped back out.
            format!(
                r#"<w:r xmlns:w="{w}"><w:x/><w:s xmlns:w="urn:inner"><w:x/><w:t xmlns:w=""><w:x/></w:t><w:x/></w:s><w:x/></w:r>"#
            ),
            // The default namespace cached, shadowed, undeclared and restored.
            format!(
                r#"<r xmlns="{w}"><x/><s xmlns="urn:inner"><x/><t xmlns=""><x/></t><x/></s><x/></r>"#
            ),
            // A declaration on an empty element lives only for that element.
            format!(r#"<w:r xmlns:w="{w}"><w:x/><w:y xmlns:w="urn:empty"/><w:x/></w:r>"#),
            // An end tag resolves in the scope its start tag opened.
            format!(r#"<w:r xmlns:w="{w}"><w:t xmlns:w="urn:end">x</w:t><w:x/></w:r>"#),
        ];
        for xml in &documents {
            assert_eq!(tracker_trace(xml), resolver_trace(xml), "diverged on {xml}");
        }
    }

    #[test]
    fn the_xmlns_prefilter_finds_every_declaration() {
        for raw in [
            &b" xmlns=\"u\""[..],
            b" xmlns:w=\"u\"",
            b" x=\"1\" xmlns:w=\"u\"",
            b"xmlns",
            b" xxmlns:w=\"u\"",
            b" a=\"xmlnx\" xmlns=\"u\"",
        ] {
            assert!(contains_xmlns(raw), "{}", String::from_utf8_lossy(raw));
        }
        for raw in [
            &b""[..],
            b" xml:space=\"preserve\"",
            b" x=\"xmln\"",
            b"xmln",
            b" w:val=\"x\" w:x=\"xmlnx\"",
        ] {
            assert!(!contains_xmlns(raw), "{}", String::from_utf8_lossy(raw));
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
