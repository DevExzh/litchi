//! Typed vocabulary and bounded policies for markup-compatibility preprocessing.

use std::{borrow::Cow, collections::HashSet, collections::TryReserveError};
use thiserror::Error as ThisError;

/// Markup Compatibility namespace from ISO/IEC 29500-3.
pub const NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(crate) const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";

/// An expanded XML name used by MCE preservation and extension policies.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Name {
    pub namespace: String,
    pub local_name: String,
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

/// Bounds for one markup-compatibility preprocessing operation.
#[derive(Debug, Clone)]
pub struct Limits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_depth: usize,
    pub max_namespace_bindings: usize,
    pub max_directive_tokens: usize,
    pub max_choices_per_alternate: usize,
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
