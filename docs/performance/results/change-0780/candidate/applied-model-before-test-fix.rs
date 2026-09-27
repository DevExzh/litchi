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
/// to that copy. Cloning a `NamespaceUri` never copies the URI; public text
/// hashing and comparisons between independent allocations can still read it.
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
/// the URI, avoiding a URI copy for every event name. Registered-extension
/// checks and caller observers can still read or copy that text. [`Name`] is
/// the owned form that the processing
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

/// The owned or static set of namespaces understood by a profile.
#[derive(Debug, Clone)]
enum NamespaceSet {
    /// The fixed Transitional/Strict OOXML vocabulary.
    Baseline,
    /// Namespaces explicitly registered by a caller.
    Explicit(HashSet<String>),
}

/// Namespaces understood by a caller and extension elements retained as opaque
/// branches during preprocessing.
#[derive(Debug, Clone)]
pub struct Capabilities {
    understood: NamespaceSet,
    pub(crate) extensions: HashSet<Name>,
}

impl Capabilities {
    /// Create an empty capability set.
    #[must_use]
    pub fn new() -> Self {
        Self {
            understood: NamespaceSet::Explicit(HashSet::new()),
            extensions: HashSet::new(),
        }
    }

    /// Create the baseline namespaces required by the OOXML profile.
    #[must_use]
    pub fn ooxml_baseline() -> Self {
        Self {
            understood: NamespaceSet::Baseline,
            extensions: HashSet::new(),
        }
    }

    /// Mark one namespace as understood by the processing profile.
    pub fn understand_namespace(&mut self, namespace: impl Into<String>) -> &mut Self {
        let namespace = namespace.into();
        if matches!(&self.understood, NamespaceSet::Baseline)
            && is_ooxml_baseline_namespace(&namespace)
        {
            return self;
        }

        if matches!(&self.understood, NamespaceSet::Baseline) {
            let mut understood = HashSet::with_capacity(OOXML_BASELINE_NAMESPACES.len() + 1);
            for &known in OOXML_BASELINE_NAMESPACES {
                understood.insert(known.to_owned());
            }
            understood.insert(namespace);
            self.understood = NamespaceSet::Explicit(understood);
        } else if let NamespaceSet::Explicit(understood) = &mut self.understood {
            understood.insert(namespace);
        }
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
        match &self.understood {
            NamespaceSet::Baseline => is_ooxml_baseline_namespace(namespace),
            NamespaceSet::Explicit(understood) => understood.contains(namespace),
        }
    }
}

/// The fixed OOXML profile vocabulary. Keeping these as static string
/// references means `Capabilities::ooxml_baseline` owns no URI strings or
/// hash table; custom registrations are materialized only when needed.
static OOXML_BASELINE_NAMESPACES: &[&str; 17] = &[
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
];

/// Whether `namespace` belongs to the fixed Transitional/Strict OOXML
/// vocabulary understood by [`Capabilities::ooxml_baseline`].
fn is_ooxml_baseline_namespace(namespace: &str) -> bool {
    OOXML_BASELINE_NAMESPACES.contains(&namespace)
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

#[cfg(test)]
mod capabilities_tests {
    use super::{Capabilities, Error, Limits, NAMESPACE, Name, NamespaceSet, Report};
    use crate::mce::StreamLimits;
    use crate::mce::{process_markup_compatibility, process_markup_compatibility_stream};
    use std::io::Cursor;

    const OLD_BASELINE_NAMESPACES: [&str; 17] = [
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
        "http://www.w3.org/XML/1998/namespace",
    ];

    fn old_explicit_baseline() -> Capabilities {
        let mut capabilities = Capabilities::new();
        for &namespace in OLD_BASELINE_NAMESPACES {
            capabilities.understand_namespace(namespace);
        }
        capabilities
    }

    fn legacy_result(
        source: &[u8],
        capabilities: &Capabilities,
        limits: &Limits,
    ) -> Result<(Vec<u8>, Report), Error> {
        process_markup_compatibility(source, capabilities, limits)
            .map(|output| (output.xml.into_owned(), output.report))
    }

    fn stream_signature(source: &[u8], capabilities: &Capabilities) -> (String, Vec<String>) {
        let mut input = Cursor::new(source);
        let mut events = Vec::new();
        let outcome = process_markup_compatibility_stream(
            &mut input,
            capabilities,
            &StreamLimits::default(),
            |event| {
                events.push(format!("{event:?}"));
                Ok::<(), std::convert::Infallible>(())
            },
        );
        let outcome = match outcome {
            Ok(report) => format!("ok {report:?}"),
            Err(error) => format!("err {error:?}"),
        };
        (outcome, events)
    }

    #[test]
    fn baseline_matches_the_old_explicit_seventeen_namespace_set() {
        let baseline = Capabilities::ooxml_baseline();
        let explicit = old_explicit_baseline();
        let default_capabilities = Capabilities::default();
        let probes = [
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
            "http://www.w3.org/XML/1998/namespace",
            "",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main/",
            "http://purl.oclc.org/ooxml/wordprocessingml/main/",
            "urn:schemas-microsoft-com:vmlx",
            "urn:custom",
        ];

        for namespace in probes {
            assert_eq!(
                baseline.understands(namespace),
                explicit.understands(namespace),
                "membership differs for {namespace:?}"
            );
            assert_eq!(
                default_capabilities.understands(namespace),
                baseline.understands(namespace),
                "default membership differs for {namespace:?}"
            );
        }
        assert!(matches!(&baseline.understood, NamespaceSet::Baseline));
        assert!(matches!(&explicit.understood, NamespaceSet::Explicit(_)));
        assert!(!Capabilities::new().understands("urn:custom"));
        assert!(default_capabilities.understands(OLD_BASELINE_NAMESPACES[0]));
    }

    #[test]
    fn static_namespace_tag_preserves_the_current_capability_owner_size() {
        use std::{collections::HashSet, mem::size_of};

        assert_eq!(
            size_of::<NamespaceSet>(),
            size_of::<HashSet<String>>(),
            "the owner budget must be revisited if the tag loses HashSet's niche"
        );
        assert_eq!(
            size_of::<Capabilities>(),
            size_of::<HashSet<String>>() + size_of::<HashSet<Name>>(),
            "the Capabilities owner must remain two hash-set owners"
        );
    }

    #[test]
    fn baseline_and_explicit_profiles_have_identical_legacy_and_stream_behavior() {
        let source = format!(
            r#"<r xmlns:mc="{NAMESPACE}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" mc:Ignorable="w s"><mc:AlternateContent><mc:Choice Requires="w"><w:body/></mc:Choice><mc:Choice Requires="s"><s:body/></mc:Choice><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent></r>"#
        );
        let baseline = Capabilities::ooxml_baseline();
        let explicit = old_explicit_baseline();
        let limits = Limits::default();

        assert_eq!(
            legacy_result(source.as_bytes(), &baseline, &limits),
            legacy_result(source.as_bytes(), &explicit, &limits)
        );
        assert!(
            String::from_utf8(
                legacy_result(source.as_bytes(), &baseline, &limits)
                    .unwrap()
                    .0
            )
            .unwrap()
            .contains("<w:body>")
        );
        assert_eq!(
            stream_signature(source.as_bytes(), &baseline),
            stream_signature(source.as_bytes(), &explicit)
        );
    }

    #[test]
    fn custom_registration_clone_isolation_and_extension_preservation_are_unchanged() {
        let mut custom = Capabilities::new();
        custom.understand_namespace("urn:custom");
        assert!(custom.understands("urn:custom"));
        assert!(!custom.understands(OLD_BASELINE_NAMESPACES[0]));

        let mut clone = custom.clone();
        custom.understand_namespace("urn:original-only");
        clone.understand_namespace("urn:clone-only");
        assert!(custom.understands("urn:original-only"));
        assert!(!clone.understands("urn:original-only"));
        assert!(clone.understands("urn:clone-only"));
        assert!(!custom.understands("urn:clone-only"));

        let must_understand = format!(
            r#"<r xmlns:mc="{NAMESPACE}" xmlns:c="urn:custom" mc:MustUnderstand="c"><c:item/></r>"#
        );
        let mut baseline_custom = Capabilities::ooxml_baseline();
        baseline_custom.understand_namespace("urn:custom");
        let mut explicit_custom = old_explicit_baseline();
        explicit_custom.understand_namespace("urn:custom");
        let limits = Limits::default();
        assert_eq!(
            legacy_result(must_understand.as_bytes(), &baseline_custom, &limits),
            legacy_result(must_understand.as_bytes(), &explicit_custom, &limits)
        );
        assert!(legacy_result(must_understand.as_bytes(), &baseline_custom, &limits).is_ok());
        assert!(matches!(
            legacy_result(
                must_understand.as_bytes(),
                &Capabilities::ooxml_baseline(),
                &limits
            ),
            Err(Error::MustUnderstand(_))
        ));

        let source = format!(
            r#"<r xmlns:mc="{NAMESPACE}" xmlns:e="urn:extension"><e:opaque><child/></e:opaque></r>"#
        );
        let extension = Name {
            namespace: "urn:extension".to_owned(),
            local_name: "opaque".to_owned(),
        };
        let mut baseline = Capabilities::ooxml_baseline();
        baseline.preserve_extension_element(extension.clone());
        let mut explicit = old_explicit_baseline();
        explicit.preserve_extension_element(extension);
        assert_eq!(
            legacy_result(source.as_bytes(), &baseline, &limits),
            legacy_result(source.as_bytes(), &explicit, &limits)
        );
        assert_eq!(
            stream_signature(source.as_bytes(), &baseline),
            stream_signature(source.as_bytes(), &explicit)
        );
    }

    #[test]
    fn malformed_unbound_and_limited_inputs_match_the_old_profile() {
        let baseline = Capabilities::ooxml_baseline();
        let explicit = old_explicit_baseline();
        let cases = [
            br#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#.as_slice(),
            br#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="missing"/>"#.as_slice(),
        ];
        for source in cases {
            assert_eq!(
                legacy_result(source, &baseline, &Limits::default()),
                legacy_result(source, &explicit, &Limits::default())
            );
            assert_eq!(
                stream_signature(source, &baseline),
                stream_signature(source, &explicit)
            );
        }

        let source = br#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><child/></r>"#;
        let limits = Limits {
            max_input_bytes: source.len() - 1,
            ..Limits::default()
        };
        assert_eq!(
            legacy_result(source, &baseline, &limits),
            legacy_result(source, &explicit, &limits)
        );
    }
}
