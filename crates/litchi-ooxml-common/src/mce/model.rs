//! Typed vocabulary and bounded policies for markup-compatibility preprocessing.

use std::{
    borrow::{Borrow, Cow},
    cmp::Ordering,
    collections::HashSet,
    collections::TryReserveError,
    fmt,
    hash::{Hash, Hasher},
    ops::Deref,
    sync::Arc,
};
use thiserror::Error as ThisError;

/// Markup Compatibility namespace from ISO/IEC 29500-3.
pub const NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(crate) const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";

/// XML namespace used by namespace declaration attributes.
pub const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// An expanded XML name used by MCE preservation and extension policies.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Name {
    pub namespace: String,
    pub local_name: String,
}

/// The namespace URI of an [`ExpandedName`] delivered by the MCE stream.
///
/// A document chooses its namespace URIs, and one declaration can put any
/// number of names in its namespace. The stream copies a URI once, when a
/// declaration binds it, and every name it expands in that namespace refers
/// to that copy, so a name costs the same whatever the length of its URI:
/// cloning a `NamespaceUri` never copies the URI.
///
/// It reads as its text: it dereferences to `str`, and compares, orders,
/// hashes, displays and debug-prints exactly as that text does. The empty
/// URI is no namespace.
#[derive(Clone, Default)]
pub struct NamespaceUri(Storage);

#[derive(Clone, Default)]
enum Storage {
    /// No namespace: the empty URI.
    #[default]
    None,
    /// A URI with static text, such as the `xml` or `xmlns` namespace.
    Static(&'static str),
    /// A URI copied from a document, shared by the names in its namespace.
    Shared(Arc<String>),
}

impl NamespaceUri {
    /// No namespace: the empty URI.
    pub const NONE: Self = Self(Storage::None);

    /// A namespace URI with static text; `""` is no namespace.
    #[must_use]
    pub const fn from_static(uri: &'static str) -> Self {
        if uri.is_empty() {
            Self::NONE
        } else {
            Self(Storage::Static(uri))
        }
    }

    /// Copy `uri` into storage that its clones share.
    ///
    /// # Errors
    ///
    /// The allocation failure when the copy cannot be reserved.
    pub(crate) fn try_copy(uri: &str) -> Result<Self, TryReserveError> {
        if uri.is_empty() {
            return Ok(Self::NONE);
        }
        let mut text = String::new();
        text.try_reserve_exact(uri.len())?;
        text.push_str(uri);
        Ok(Self(Storage::Shared(Arc::new(text))))
    }

    /// The URI; empty for no namespace.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match &self.0 {
            Storage::None => "",
            Storage::Static(text) => text,
            Storage::Shared(text) => text.as_str(),
        }
    }

    /// Whether this is no namespace, the empty URI.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        matches!(self.0, Storage::None)
    }

    /// Whether `self` and `other` refer to one copy of their URI, which
    /// implies that they are equal. Names the stream expands in one
    /// namespace share a copy while a declaration keeps it in scope.
    #[must_use]
    pub fn shares_storage_with(&self, other: &Self) -> bool {
        match (&self.0, &other.0) {
            (Storage::None, Storage::None) => true,
            (Storage::Static(left), Storage::Static(right)) => core::ptr::eq(*left, *right),
            (Storage::Shared(left), Storage::Shared(right)) => Arc::ptr_eq(left, right),
            _ => false,
        }
    }
}

impl Deref for NamespaceUri {
    type Target = str;

    fn deref(&self) -> &str {
        self.as_str()
    }
}

impl AsRef<str> for NamespaceUri {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl Borrow<str> for NamespaceUri {
    fn borrow(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Debug for NamespaceUri {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Debug::fmt(self.as_str(), f)
    }
}

impl fmt::Display for NamespaceUri {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(self.as_str(), f)
    }
}

impl PartialEq for NamespaceUri {
    fn eq(&self, other: &Self) -> bool {
        self.shares_storage_with(other) || self.as_str() == other.as_str()
    }
}

impl Eq for NamespaceUri {}

impl PartialOrd for NamespaceUri {
    fn partial_cmp(&self, other: &Self) -> Option<Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for NamespaceUri {
    fn cmp(&self, other: &Self) -> Ordering {
        if self.shares_storage_with(other) {
            return Ordering::Equal;
        }
        self.as_str().cmp(other.as_str())
    }
}

impl Hash for NamespaceUri {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.as_str().hash(state);
    }
}

impl PartialEq<str> for NamespaceUri {
    fn eq(&self, other: &str) -> bool {
        self.as_str() == other
    }
}

impl PartialEq<&str> for NamespaceUri {
    fn eq(&self, other: &&str) -> bool {
        self.as_str() == *other
    }
}

impl PartialEq<String> for NamespaceUri {
    fn eq(&self, other: &String) -> bool {
        self.as_str() == other.as_str()
    }
}

impl PartialEq<NamespaceUri> for str {
    fn eq(&self, other: &NamespaceUri) -> bool {
        self == other.as_str()
    }
}

impl PartialEq<NamespaceUri> for &str {
    fn eq(&self, other: &NamespaceUri) -> bool {
        *self == other.as_str()
    }
}

impl PartialEq<NamespaceUri> for String {
    fn eq(&self, other: &NamespaceUri) -> bool {
        self.as_str() == other.as_str()
    }
}

impl From<&str> for NamespaceUri {
    fn from(uri: &str) -> Self {
        Self::from(uri.to_owned())
    }
}

impl From<String> for NamespaceUri {
    fn from(uri: String) -> Self {
        if uri.is_empty() {
            Self::NONE
        } else {
            Self(Storage::Shared(Arc::new(uri)))
        }
    }
}

impl From<NamespaceUri> for String {
    fn from(uri: NamespaceUri) -> Self {
        match uri.0 {
            Storage::None => Self::new(),
            Storage::Static(text) => text.to_owned(),
            Storage::Shared(text) => {
                Arc::try_unwrap(text).unwrap_or_else(|text| text.as_str().to_owned())
            },
        }
    }
}

/// A namespace-expanded element or attribute name delivered by the MCE
/// stream.
///
/// Its namespace is a [`NamespaceUri`], which shares the stream's one copy of
/// the URI, so the stream's events cost the same whatever the length of the
/// URIs a document chooses. [`Name`] is the owned form that the processing
/// policy ([`Capabilities`]) takes.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct ExpandedName {
    /// The namespace URI; empty when the name is in no namespace.
    pub namespace: NamespaceUri,
    /// The local part of the name.
    pub local_name: String,
}

impl ExpandedName {
    /// The name `local_name` in `namespace`.
    #[must_use]
    pub fn new(namespace: impl Into<NamespaceUri>, local_name: impl Into<String>) -> Self {
        Self {
            namespace: namespace.into(),
            local_name: local_name.into(),
        }
    }
}

/// Namespaces understood by a caller and extension elements retained as opaque
/// branches during preprocessing.
#[derive(Debug, Clone)]
pub struct Capabilities {
    pub(crate) understood: HashSet<String>,
    pub(crate) extensions: HashSet<Name>,
}

impl Capabilities {
    /// Create an empty capability set.
    #[must_use]
    pub fn new() -> Self {
        Self {
            understood: HashSet::new(),
            extensions: HashSet::new(),
        }
    }

    /// Create the baseline namespaces required by the OOXML profile.
    #[must_use]
    pub fn ooxml_baseline() -> Self {
        let mut capabilities = Self::new();
        for namespace in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "http://purl.oclc.org/ooxml/drawingml/chart",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
            "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "http://purl.oclc.org/ooxml/officeDocument/math",
            "urn:schemas-microsoft-com:vml",
            "urn:schemas-microsoft-com:office:office",
            XML_NS,
        ] {
            capabilities.understood.insert(namespace.into());
        }
        capabilities
    }

    /// Mark one namespace as understood by the processing profile.
    pub fn understand_namespace(&mut self, namespace: impl Into<String>) -> &mut Self {
        self.understood.insert(namespace.into());
        self
    }

    /// Retain one extension element as an opaque branch.
    pub fn preserve_extension_element(&mut self, name: Name) -> &mut Self {
        self.extensions.insert(name);
        self
    }

    /// Test whether a namespace is understood by this profile.
    #[must_use]
    pub fn understands(&self, namespace: &str) -> bool {
        self.understood.contains(namespace)
    }
}

impl Default for Capabilities {
    fn default() -> Self {
        Self::ooxml_baseline()
    }
}

/// Default for [`Limits::max_attributes_per_element`].
///
/// The largest element in the repository's real OOXML and ODF packages
/// carries 43 attributes, and Word's roots about 40 namespace declarations;
/// the widest ECMA-376 element types declare about 70 attributes.
pub const DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT: usize = 1024;

/// Immutable ceiling for [`Limits::max_attributes_per_element`] and for the
/// stream's per-event attribute limit, matching the publication audit's.
///
/// quick-xml's duplicate-name check costs up to quadratic time in a tag's
/// attributes when their names are chosen to collide, so no configuration may
/// admit more. The in-memory processor applies at most this many whatever the
/// field holds; the stream's `validate` refuses a larger value.
pub const ATTRIBUTES_PER_ELEMENT_CEILING: usize = 4096;

/// Bounds for one markup-compatibility preprocessing operation.
#[derive(Debug, Clone)]
pub struct Limits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_depth: usize,
    pub max_namespace_bindings: usize,
    pub max_directive_tokens: usize,
    pub max_choices_per_alternate: usize,
    /// Maximum attributes on one start or empty-element tag, namespace
    /// declarations included.
    ///
    /// Checked as each attribute is read, so a larger tag is refused at its
    /// first surplus attribute. The parser's duplicate-name check costs up to
    /// quadratic time in a tag's attribute count when the names are chosen
    /// to collide, and this bound caps that cost per tag. A value above
    /// [`ATTRIBUTES_PER_ELEMENT_CEILING`] acts as the ceiling.
    pub max_attributes_per_element: usize,
}

/// Resource policy for retaining source offsets through MCE preprocessing.
///
/// Source and returned coordinates are byte offsets into the caller's original
/// XML. The processing field bounds the marked intermediate document as it
/// passes through the MCE processor.
#[derive(Debug, Clone)]
pub struct OffsetLimits {
    /// Maximum raw source XML accepted by active offset selection.
    pub max_source_bytes: usize,
    /// Maximum number of source offsets accepted in one call.
    pub max_offsets: usize,
    /// Maximum marked intermediate XML retained during branch selection.
    pub max_marked_bytes: usize,
    /// Bounds applied by the semantic MCE processor.
    pub processing: Limits,
}

impl Default for OffsetLimits {
    fn default() -> Self {
        let processing = Limits::default();
        Self {
            max_source_bytes: processing.max_input_bytes,
            max_offsets: 1_000_000,
            max_marked_bytes: processing.max_input_bytes,
            processing,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: 256 * 1024 * 1024,
            max_output_bytes: 512 * 1024 * 1024,
            max_depth: 256,
            max_namespace_bindings: 4096,
            max_directive_tokens: 4096,
            max_choices_per_alternate: 1024,
            max_attributes_per_element: DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
        }
    }
}

/// Processing counters for one MCE output.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Report {
    pub alternate_content_count: usize,
    pub selected_choices: usize,
    pub selected_fallbacks: usize,
    pub ignored_elements: usize,
    pub ignored_attributes: usize,
    pub preserved_elements: usize,
    pub preserved_attributes: usize,
    pub unwrapped_elements: usize,
}

/// Processed XML and counters describing the selected/retained branches.
#[derive(Debug)]
pub struct Output<'a> {
    pub xml: Cow<'a, [u8]>,
    pub report: Report,
}

/// A malformed, unsupported, or resource-limited MCE document.
#[derive(Debug, ThisError, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    #[error("non-conformant markup compatibility XML: {0}")]
    NonConformant(String),
    #[error("unsupported namespace required by MustUnderstand: {0}")]
    MustUnderstand(String),
    #[error("markup compatibility resource limit exceeded: {0}")]
    LimitExceeded(String),
    #[error("markup compatibility XML error: {0}")]
    Xml(String),

    /// A bounded intermediate buffer could not be allocated.
    #[error("markup compatibility allocation failed for {resource}")]
    Allocation {
        /// Intermediate representation that could not reserve storage.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },
}
