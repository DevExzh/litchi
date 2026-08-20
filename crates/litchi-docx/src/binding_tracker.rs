//! Namespace-binding tracking over a plain quick-xml [`Reader`].
//!
//! [`BindingTracker`] replicates the binding maintenance `NsReader` performs
//! in `process_event`, so a parse loop can drive a plain [`Reader`] (with
//! borrowing reads) and resolve names with byte-identical semantics.  Ported
//! from `litchi-odt`'s change-0224/0227 tracker for change 0229 (the
//! `extract_word_text` port in [`crate::paragraph`]); kept crate-private and
//! litchi-docx-local, mirroring the ODT precedent. When litchi-pptx (or
//! another OOXML owner) needs the same machinery, the two copies should be
//! reconciled and lifted to `litchi-ooxml-common` in one change.
//!
//! [`Reader`]: quick_xml::reader::Reader

use memchr::memmem;
use quick_xml::events::BytesStart;
use quick_xml::name::{
    LocalName, Namespace, NamespaceError, Prefix, PrefixDeclaration, QName, ResolveResult,
};

/// The reserved namespace URIs whose binding rules quick-xml's
/// `NamespaceResolver::add` enforces. Crate-private copies of quick-xml's
/// private `RESERVED_NAMESPACE_XML` / `RESERVED_NAMESPACE_XMLNS` constants.
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

/// Parity with quick-xml's `DEFAULT_MAX_DECLARATIONS_PER_ELEMENT`.
const MAX_NS_DECLARATIONS_PER_ELEMENT: usize = 256;

/// A single in-scope prefix→URI binding, indexing the tracker's flat byte
/// buffer exactly like quick-xml's `NamespaceBinding`: the prefix occupies
/// `buffer[start..start + prefix_len]` (`prefix_len == 0` for a default
/// `xmlns="..."` declaration) and the URI follows it for `value_len` bytes.
/// Values are copied into the one shared buffer (as quick-xml copies them);
/// pushes only happen on elements that declare bindings, so the copies are
/// rare and tiny.
#[derive(Debug)]
struct Binding {
    /// Offset in the tracker buffer where the prefix starts.
    start: usize,
    /// Prefix length; zero marks a default-namespace declaration.
    prefix_len: usize,
    /// URI length; the URI starts at `start + prefix_len`.
    value_len: usize,
    /// Nesting level at which the binding was declared; the declaring element
    /// is included, i.e. a declaration on the document root has `level = 1`.
    level: u32,
}

impl Binding {
    /// The bound prefix bytes, or `None` for a default-namespace declaration.
    fn prefix<'b>(&self, buffer: &'b [u8]) -> Option<&'b [u8]> {
        (self.prefix_len != 0).then(|| &buffer[self.start..self.start + self.prefix_len])
    }

    /// The bound namespace URI bytes.
    fn value<'b>(&self, buffer: &'b [u8]) -> &'b [u8] {
        &buffer[self.start + self.prefix_len..self.start + self.prefix_len + self.value_len]
    }
}

/// Hand-rolled namespace binding tracker replacing `NsReader`'s binding
/// maintenance over a plain [`Reader`] (change 0229, ported from the
/// litchi-odt change-0224 tracker).
///
/// Byte-exactness contract with quick-xml 0.41 `NsReader` (`ns_reader.rs`
/// `process_event`/`read_event_impl` and `name.rs` `NamespaceResolver`):
///
/// - **push** runs for every `Start`/`Empty` event at any depth, before the
///   event is handed to the handlers, so a namespace error preempts the
///   event exactly as `read_event` returning `Err` did. The attribute scan
///   uses the same `attributes().with_checks(false)` iterator and stops
///   SILENTLY at the first malformed attribute, keeping bindings declared
///   before it.
/// - **declaration limit**: more than [`MAX_NS_DECLARATIONS_PER_ELEMENT`]
///   `xmlns` attributes on one tag fail with
///   `NamespaceError::TooManyDeclarations`, checked per declaration in
///   attribute order.
/// - **reserved prefixes**: `add` rejects, in attribute order and with the
///   same precedence, binding `xml` to a foreign URI
///   (`InvalidXmlPrefixBind`), declaring `xmlns` (`InvalidXmlnsPrefixBind`),
///   binding another prefix to the xml URI (`InvalidPrefixForXml`), and
///   binding any prefix to the xmlns URI (`InvalidPrefixForXmlns`). Binding
///   `xml` to its reserved URI is a no-op. Errors are real
///   [`NamespaceError`] values, so their `Display` (which is also
///   `quick_xml::Error::Namespace`'s `Display`) is byte-identical by
///   construction.
/// - **unbinding**: `xmlns=""` / `xmlns:p=""` push a value-less binding;
///   resolution maps an emptied default to `Unbound` and an emptied or
///   missing prefix to `Unknown(prefix.to_vec())`, matching
///   `resolve_prefix`.
/// - **pending pop**: the caller applies [`BindingTracker::pop`] at the top
///   of the next read iteration, never while delivering the `End`/`Empty`
///   event itself, so an end tag still resolves in its own scope — the
///   deferred `pending_pop` semantics of `NsReader::read_event_impl`.
/// - **pre-bound entries**: `xml` and `xmlns` are bound at level 0 before any
///   push, like `NamespaceResolver::default`.
/// - **resolution**: `resolve_element` decomposes at the first `:` and scans
///   bindings in reverse declaration order (last binding wins), the
///   `resolve(name, use_default = true)` algorithm.
///
/// The one deliberate divergence: quick-xml counts nesting in a `u16`
/// (wrapping in release, panicking in debug past depth 65 535); the tracker
/// uses a `u32`. This is unobservable here — the text scan's own element
/// nesting is capped at 128 — and strictly removes a panic path.
///
/// The `xmlns` prefilter: when a tag's raw attribute bytes contain no
/// `xmlns` substring, no attribute key can start with `xmlns`, so the push
/// can neither add a binding nor fail (every push error requires an
/// `xmlns`-series key, and the malformed-attribute break is unobservable
/// when no such key exists). The level increment is then the entire push.
/// A raw attribute slice shorter than the needle cannot contain the
/// substring, so it skips even the substring search.
///
/// [`Reader`]: quick_xml::reader::Reader
#[derive(Debug)]
pub(crate) struct BindingTracker {
    /// Flat byte buffer holding every in-scope prefix and URI, mirroring
    /// `NamespaceResolver`'s single-buffer layout: one amortized allocation
    /// for all names instead of two allocations per binding.
    buffer: Vec<u8>,
    /// Bindings in declaration order; levels are non-decreasing because a
    /// push always appends at the current level and a pop truncates.
    bindings: Vec<Binding>,
    level: u32,
}

impl BindingTracker {
    pub(crate) fn new() -> Self {
        // `NamespaceResolver::default` pre-binds the reserved prefixes; mirror
        // its allocation pattern (one flat buffer, one binding stack, no
        // per-entry allocations).
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

    /// Replicate `NamespaceResolver::push` for one `Start`/`Empty` element.
    ///
    /// Inline fast path: a raw attribute slice shorter than `xmlns` can
    /// never contain the substring, so the prefilter and the scan it gates
    /// are both skipped and the level increment is the entire push.
    #[inline]
    pub(crate) fn push(&mut self, element: &BytesStart<'_>) -> Result<(), NamespaceError> {
        self.level += 1;
        if element.attributes_raw().len() < b"xmlns".len() {
            return Ok(());
        }
        self.push_scanned(element)
    }

    /// The prefilter-gated attribute scan of [`BindingTracker::push`], kept
    /// out of line so the bare-tag fast path carries none of the scan
    /// machinery (its stack frame, the `memmem` search, the attribute
    /// iterator).
    fn push_scanned(&mut self, element: &BytesStart<'_>) -> Result<(), NamespaceError> {
        if memmem::find(element.attributes_raw(), b"xmlns").is_none() {
            return Ok(());
        }
        let mut count = 0usize;
        for attribute in element.attributes().with_checks(false) {
            // `with_checks(false)` plus a silent break on the first malformed
            // attribute: the exact scan `NamespaceResolver::push` performs.
            let Ok(attribute) = attribute else {
                break;
            };
            if let Some(prefix) = attribute.key.as_namespace_binding() {
                if count >= MAX_NS_DECLARATIONS_PER_ELEMENT {
                    return Err(NamespaceError::TooManyDeclarations(
                        MAX_NS_DECLARATIONS_PER_ELEMENT,
                    ));
                }
                count += 1;
                self.add(prefix, &attribute.value)?;
            }
        }
        Ok(())
    }

    /// Replicate `NamespaceResolver::add` for one `xmlns` declaration.
    fn add(&mut self, prefix: PrefixDeclaration<'_>, uri: &[u8]) -> Result<(), NamespaceError> {
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
                    return Err(NamespaceError::InvalidXmlPrefixBind(uri.to_vec()));
                }
                // Binding `xml` to its reserved URI adds no entry.
            },
            PrefixDeclaration::Named(b"xmlns") => {
                return Err(NamespaceError::InvalidXmlnsPrefixBind(uri.to_vec()));
            },
            PrefixDeclaration::Named(prefix) => {
                if uri == XML_NAMESPACE {
                    return Err(NamespaceError::InvalidPrefixForXml(prefix.to_vec()));
                }
                if uri == XMLNS_NAMESPACE {
                    return Err(NamespaceError::InvalidPrefixForXmlns(prefix.to_vec()));
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

    /// Replicate `NamespaceResolver::pop` (`set_level(level - 1)`): drop every
    /// binding declared deeper than the new level, truncating the flat buffer
    /// at the first dropped binding's start exactly as `set_level` does.
    pub(crate) fn pop(&mut self) {
        self.level = self.level.saturating_sub(1);
        // From the back (most deeply nested scope), look for the first scope
        // that is still valid — the `set_level` scan.
        match self
            .bindings
            .iter()
            .rposition(|binding| binding.level <= self.level)
        {
            // None of the bindings are valid: remove all of them.
            None => {
                self.buffer.clear();
                self.bindings.clear();
            },
            // Drop all bindings past the last valid one, only when there is
            // something to drop (`set_level`'s `get(last + 1)` guard).
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

    /// Replicate `NamespaceResolver::resolve(name, use_default = true)`.
    pub(crate) fn resolve_element<'n>(
        &self,
        name: QName<'n>,
    ) -> (ResolveResult<'_>, LocalName<'n>) {
        let (local_name, prefix) = name.decompose();
        (self.resolve_prefix(prefix), local_name)
    }

    /// Replicate `NamespaceResolver::resolve_prefix`: scan bindings in reverse
    /// declaration order so the last binding wins; an emptied or missing
    /// prefixed binding resolves to `Unknown`, an emptied default to
    /// `Unbound`, without scanning further back.
    fn resolve_prefix(&self, prefix: Option<Prefix<'_>>) -> ResolveResult<'_> {
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
