//! `ODF` namespace vocabulary and qualified-name resolution.
//!
//! This module provides support for `XML` namespaces, including qualified names,
//! namespace context, and namespace-aware comparisons shared by every `ODF`
//! document family.
//!
//! # Implementation Status
//!
//! ✅ COMPLETED: All `ODF` `1.2` namespaces (40+ namespaces)
//! ✅ COMPLETED: Extension namespaces (`LibreOffice`, `OpenOffice`, `KOffice`)
//! ✅ COMPLETED: Standard web namespaces (`XML`, `XLink`, `SVG`, `MathML`, etc.)
//!
//! # References
//!
//! - `odfpy`: `3rdparty/odfpy/odf/namespaces.py` (lines 24–111)

use litchi_core::{Error, Result};
use phf::{Map, phf_map};
use quick_xml::XmlVersion;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;
use std::collections::HashMap;

// ============================================================================
// NAMESPACE CONSTANTS
// ============================================================================
// Reference: odfpy/odf/namespaces.py lines 24-66

// Some namespace vocabulary is retained for completeness even when a specific
// document family does not currently consume it.

/// Animation namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const ANIMNS: &str = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0";

/// Chart namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const CHARTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";

/// `OpenOffice` chart extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const CHARTOOONS: &str = "http://openoffice.org/2010/chart";

/// Configuration namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const CONFIGNS: &str = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

/// CSS3 text extensions
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const CSS3TNS: &str = "http://www.w3.org/TR/css3-text/";

/// Database namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const DBNS: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";

/// Dublin Core namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const DCNS: &str = "http://purl.org/dc/elements/1.1/";

/// DOM events namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const DOMNS: &str = "http://www.w3.org/2001/xml-events";

/// 3D drawing namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const DR3DNS: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";

/// Drawing namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const DRAWNS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

/// `OpenOffice` field extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const FIELDNS: &str = "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0";

/// XSL-FO compatible namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const FONS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

/// Form namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const FORMNS: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";

/// OOXML-ODF form interoperability
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const FORMXNS: &str = "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0";

/// GRDDL namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const GRDDLNS: &str = "http://www.w3.org/2003/g/data-view#";

/// `KOffice` extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const KOFFICENS: &str = "http://www.koffice.org/2005/";

/// `LibreOffice` extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const LOEXTNS: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

/// Manifest namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const MANIFESTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

/// `MathML` namespace.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const MATHNS: &str = "http://www.w3.org/1998/Math/MathML";

/// Metadata namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const METANS: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

/// Number/data style namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const NUMBERNS: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

/// Office namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const OFFICENS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// `OpenFormula` namespace (`ODF` 1.2).
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const OFNS: &str = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";

/// `OpenOffice` Calc extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const OOOCNS: &str = "http://openoffice.org/2004/calc";

/// `OpenOffice` general extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const OOONS: &str = "http://openoffice.org/2004/office";

/// `OpenOffice` Writer extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const OOOWNS: &str = "http://openoffice.org/2004/writer";

/// Presentation namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const PRESENTATIONNS: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

/// `RDFa` namespace.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const RDFANS: &str = "http://docs.oasis-open.org/opendocument/meta/rdfa#";

/// Report namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const RPTNS: &str = "http://openoffice.org/2005/report";

/// Script namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const SCRIPTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";

/// SMIL compatible namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const SMILNS: &str = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";

/// Style namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const STYLENS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/// SVG compatible namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const SVGNS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

/// Table namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const TABLENS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

/// `OpenOffice` table extensions.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const TABLEOOONS: &str = "http://openoffice.org/2009/table";

/// Text namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const TEXTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/// `XForms` namespace.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const XFORMSNS: &str = "http://www.w3.org/2002/xforms";

/// XHTML namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const XHTMLNS: &str = "http://www.w3.org/1999/xhtml";

/// `XLink` namespace.
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const XLINKNS: &str = "http://www.w3.org/1999/xlink";

/// XML namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const XMLNS: &str = "http://www.w3.org/XML/1998/namespace";

/// The namespace URI reserved for `xmlns` declarations.
///
/// This URI is never a valid default namespace or an ordinary application
/// prefix binding.  `quick-xml` validates most prefixed forms while parsing,
/// but it intentionally leaves the default declaration available to callers,
/// so the shared ODF helper validates it as well.
pub const XMLNS_DECLARATION_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// Maximum decoded namespace URI size accepted by the shared resolver.
///
/// Namespace declarations in ODF are short tokens.  The ceiling is kept
/// separate from the caller's XML-part limit so entity expansion is bounded
/// before a destination string is allocated.
pub const MAX_NAMESPACE_URI_BYTES: usize = 64 * 1024;

/// XML Schema namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const XSDNS: &str = "http://www.w3.org/2001/XMLSchema";

/// XML Schema instance namespace
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const XSINS: &str = "http://www.w3.org/2001/XMLSchema-instance";

/// Calc extensions (`LibreOffice`).
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const CALCEXTNS: &str = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

/// Drawing extensions (`OpenOffice`).
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const DRAWOOONS: &str = "http://openoffice.org/2010/draw";

/// Office extensions (`OpenOffice`).
#[allow(
    dead_code,
    reason = "Retained as part of the complete ODF namespace vocabulary."
)]
pub const OFFICEOOONS: &str = "http://openoffice.org/2009/office";

// ============================================================================
// NAMESPACE MAPPING (compile-time perfect hash map)
// ============================================================================
// Reference: odfpy/odf/namespaces.py lines 68-111

/// `URI` to prefix mapping (compile-time perfect hash map for zero-cost lookups).
static URI_TO_PREFIX: Map<&'static str, &'static str> = phf_map! {
    "urn:oasis:names:tc:opendocument:xmlns:animation:1.0" => "anim",
    "urn:oasis:names:tc:opendocument:xmlns:chart:1.0" => "chart",
    "http://openoffice.org/2010/chart" => "chartooo",
    "urn:oasis:names:tc:opendocument:xmlns:config:1.0" => "config",
    "http://www.w3.org/TR/css3-text/" => "css3t",
    "urn:oasis:names:tc:opendocument:xmlns:database:1.0" => "db",
    "http://purl.org/dc/elements/1.1/" => "dc",
    "http://www.w3.org/2001/xml-events" => "dom",
    "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" => "dr3d",
    "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" => "draw",
    "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" => "field",
    "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" => "fo",
    "urn:oasis:names:tc:opendocument:xmlns:form:1.0" => "form",
    "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" => "formx",
    "http://www.w3.org/2003/g/data-view#" => "grddl",
    "http://www.koffice.org/2005/" => "koffice",
    "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" => "loext",
    "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" => "manifest",
    "http://www.w3.org/1998/Math/MathML" => "math",
    "urn:oasis:names:tc:opendocument:xmlns:meta:1.0" => "meta",
    "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" => "number",
    "urn:oasis:names:tc:opendocument:xmlns:office:1.0" => "office",
    "urn:oasis:names:tc:opendocument:xmlns:of:1.2" => "of",
    "http://openoffice.org/2004/office" => "ooo",
    "http://openoffice.org/2004/writer" => "ooow",
    "http://openoffice.org/2004/calc" => "oooc",
    "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" => "presentation",
    "http://docs.oasis-open.org/opendocument/meta/rdfa#" => "rdfa",
    "http://openoffice.org/2005/report" => "rpt",
    "urn:oasis:names:tc:opendocument:xmlns:script:1.0" => "script",
    "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" => "smil",
    "urn:oasis:names:tc:opendocument:xmlns:style:1.0" => "style",
    "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" => "svg",
    "urn:oasis:names:tc:opendocument:xmlns:table:1.0" => "table",
    "http://openoffice.org/2009/table" => "tableooo",
    "urn:oasis:names:tc:opendocument:xmlns:text:1.0" => "text",
    "http://www.w3.org/2002/xforms" => "xforms",
    "http://www.w3.org/1999/xlink" => "xlink",
    "http://www.w3.org/1999/xhtml" => "xhtml",
    "http://www.w3.org/XML/1998/namespace" => "xml",
    "http://www.w3.org/2001/XMLSchema" => "xsd",
    "http://www.w3.org/2001/XMLSchema-instance" => "xsi",
    "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" => "calcext",
    "http://openoffice.org/2010/draw" => "drawooo",
    "http://openoffice.org/2009/office" => "officeooo",
};

/// Prefix to `URI` mapping (compile-time perfect hash map for zero-cost lookups).
static PREFIX_TO_URI: Map<&'static str, &'static str> = phf_map! {
    "anim" => "urn:oasis:names:tc:opendocument:xmlns:animation:1.0",
    "chart" => "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    "chartooo" => "http://openoffice.org/2010/chart",
    "config" => "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    "css3t" => "http://www.w3.org/TR/css3-text/",
    "db" => "urn:oasis:names:tc:opendocument:xmlns:database:1.0",
    "dc" => "http://purl.org/dc/elements/1.1/",
    "dom" => "http://www.w3.org/2001/xml-events",
    "dr3d" => "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
    "draw" => "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "field" => "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
    "fo" => "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "form" => "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    "formx" => "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
    "grddl" => "http://www.w3.org/2003/g/data-view#",
    "koffice" => "http://www.koffice.org/2005/",
    "loext" => "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    "manifest" => "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    "math" => "http://www.w3.org/1998/Math/MathML",
    "meta" => "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "number" => "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "office" => "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "of" => "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
    "ooo" => "http://openoffice.org/2004/office",
    "ooow" => "http://openoffice.org/2004/writer",
    "oooc" => "http://openoffice.org/2004/calc",
    "presentation" => "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    "rdfa" => "http://docs.oasis-open.org/opendocument/meta/rdfa#",
    "rpt" => "http://openoffice.org/2005/report",
    "script" => "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    "smil" => "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0",
    "style" => "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "svg" => "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "table" => "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "tableooo" => "http://openoffice.org/2009/table",
    "text" => "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "xforms" => "http://www.w3.org/2002/xforms",
    "xlink" => "http://www.w3.org/1999/xlink",
    "xhtml" => "http://www.w3.org/1999/xhtml",
    "xml" => "http://www.w3.org/XML/1998/namespace",
    "xsd" => "http://www.w3.org/2001/XMLSchema",
    "xsi" => "http://www.w3.org/2001/XMLSchema-instance",
    "calcext" => "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    "drawooo" => "http://openoffice.org/2010/draw",
    "officeooo" => "http://openoffice.org/2009/office",
};

// ============================================================================
// QUALIFIED NAME
// ============================================================================

/// Qualified name with namespace support
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QualifiedName {
    /// Namespace `URI`.
    pub namespace_uri: Option<String>,
    /// Local name (without prefix)
    pub local_name: String,
    /// Full qualified name (with prefix if present)
    pub qualified_name: String,
}

impl QualifiedName {
    fn try_copy(value: &str, resource: &'static str) -> Result<String> {
        let mut output = String::new();
        output
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation { resource, source })?;
        output.push_str(value);
        Ok(output)
    }

    /// Parse a qualified name while reporting allocation failure.
    pub fn try_from_string(name: &str) -> Result<Self> {
        let qualified_name = Self::try_copy(name, "ODF qualified name")?;
        if let Some(colon_pos) = name.find(':') {
            let prefix = &name[..colon_pos];
            let local_name = Self::try_copy(&name[colon_pos + 1..], "ODF local name")?;
            let namespace_uri = PREFIX_TO_URI
                .get(prefix)
                .map(|uri| Self::try_copy(uri, "ODF namespace URI"))
                .transpose()?;
            Ok(Self {
                namespace_uri,
                local_name,
                qualified_name,
            })
        } else {
            Ok(Self {
                namespace_uri: None,
                local_name: Self::try_copy(name, "ODF local name")?,
                qualified_name,
            })
        }
    }

    /// Create a new qualified name
    ///
    /// Note: A clone of `local_name` is necessary when no prefix is needed,
    /// as both fields must be owned strings in the struct.
    #[must_use]
    pub fn new(namespace_uri: Option<String>, local_name: String) -> Self {
        let qualified_name = match namespace_uri {
            Some(ref uri) => {
                // For common ODF namespaces, use standard prefixes
                let prefix = Self::uri_to_prefix(uri);
                if prefix.is_empty() {
                    // Clone needed: local_name used in both fields
                    local_name.clone()
                } else {
                    format!("{prefix}:{local_name}")
                }
            },
            // Clone needed: local_name used in both fields
            None => local_name.clone(),
        };

        Self {
            namespace_uri,
            local_name,
            qualified_name,
        }
    }

    /// Parse qualified name from string
    #[must_use]
    pub fn from_string(name: &str) -> Self {
        if let Some(colon_pos) = name.find(':') {
            let prefix = &name[..colon_pos];
            let local_name = &name[colon_pos + 1..];

            // Try to resolve common prefixes to URIs
            let namespace_uri = Self::prefix_to_uri(prefix);

            Self {
                namespace_uri,
                local_name: local_name.to_string(),
                qualified_name: name.to_string(),
            }
        } else {
            Self {
                namespace_uri: None,
                local_name: name.to_string(),
                qualified_name: name.to_string(),
            }
        }
    }

    /// Convert a namespace `URI` to a standard prefix using a compile-time
    /// perfect hash map.
    #[inline]
    fn uri_to_prefix(uri: &str) -> &'static str {
        URI_TO_PREFIX.get(uri).copied().unwrap_or("")
    }

    /// Convert a prefix to a namespace `URI` using a compile-time perfect hash
    /// map.
    #[inline]
    fn prefix_to_uri(prefix: &str) -> Option<String> {
        PREFIX_TO_URI.get(prefix).map(ToString::to_string)
    }

    /// Check if this name matches another qualified name
    #[must_use]
    pub fn matches(&self, other: &QualifiedName) -> bool {
        self.namespace_uri == other.namespace_uri && self.local_name == other.local_name
    }

    /// Check if this name matches a string (with optional namespace resolution)
    #[must_use]
    pub fn matches_str(&self, name: &str, namespace_context: Option<&NamespaceContext>) -> bool {
        let other = QualifiedName::from_string_with_context(name, namespace_context);
        self.matches(&other)
    }
}

impl From<&str> for QualifiedName {
    fn from(name: &str) -> Self {
        Self::from_string(name)
    }
}

impl std::fmt::Display for QualifiedName {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.qualified_name)
    }
}

/// Namespace context for resolving prefixes to `URI`s.
#[derive(Debug, Clone, Default)]
#[allow(
    clippy::module_name_repetitions,
    reason = "The public type name makes its role explicit at call sites."
)]
pub struct NamespaceContext {
    /// Mapping from a prefix to a namespace `URI`.
    pub prefixes: HashMap<String, String>,
    /// Default namespace `URI`.
    pub default_namespace: Option<String>,
}

impl NamespaceContext {
    /// Add a namespace declaration
    pub fn add_namespace(&mut self, namespace_attribute: &str, uri: &str) {
        if namespace_attribute == "xmlns" {
            self.default_namespace = Some(uri.to_string());
        } else if let Some(prefix) = namespace_attribute.strip_prefix("xmlns:") {
            self.prefixes.insert(prefix.to_string(), uri.to_string());
        }
    }

    /// Resolve prefix to namespace URI
    #[must_use]
    pub fn resolve_prefix(&self, prefix: &str) -> Option<&str> {
        self.prefixes.get(prefix).map(String::as_str)
    }

    /// Get default namespace
    #[must_use]
    pub fn default_namespace(&self) -> Option<&str> {
        self.default_namespace.as_deref()
    }

    /// Parse qualified name with this context
    #[must_use]
    pub fn parse_qualified_name(&self, name: &str) -> QualifiedName {
        QualifiedName::from_string_with_context(name, Some(self))
    }
}

/// Helper implementation for `QualifiedName`.
impl QualifiedName {
    fn from_string_with_context(name: &str, context: Option<&NamespaceContext>) -> Self {
        if let Some(colon_pos) = name.find(':') {
            let prefix = &name[..colon_pos];
            let local_name = &name[colon_pos + 1..];

            let namespace_uri = if let Some(ctx) = context {
                ctx.resolve_prefix(prefix).map(str::to_string)
            } else {
                Self::prefix_to_uri(prefix)
            };

            Self {
                namespace_uri,
                local_name: local_name.to_string(),
                qualified_name: name.to_string(),
            }
        } else {
            // No prefix - check for default namespace
            let namespace_uri = if let Some(ctx) = context {
                ctx.default_namespace().map(str::to_string)
            } else {
                None
            };

            Self {
                namespace_uri,
                local_name: name.to_string(),
                qualified_name: name.to_string(),
            }
        }
    }
}

pub(crate) fn is_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

/// Decode and XML 1.0-normalize a resolved namespace URI.
///
/// `Namespace` deliberately exposes quick-xml's lexical attribute bytes.  In
/// particular, a declaration such as `urn:example&#58;forms` is semantically
/// the same URI as `urn:example:forms`, but direct byte comparison misses it.
/// This helper performs the XML 1.0 attribute-value steps for predefined and
/// numeric character references and rejects undeclared or malformed entities.
/// It preserves the borrowed fast path when the declaration has no decoding or
/// normalization work.  The input and normalized output are bounded before a
/// destination `String` is allocated.
pub fn resolved_namespace_uri<'uri>(
    namespace: &ResolveResult<'uri>,
    decoder: quick_xml::Decoder,
    context: &str,
) -> Result<Option<Cow<'uri, str>>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(_) => Err(invalid_namespace(
            context,
            "the namespace prefix is not bound",
        )),
        ResolveResult::Bound(Namespace(raw)) => {
            let uri = decode_namespace_uri(raw, decoder, context)?;
            if uri.as_ref() == XMLNS_DECLARATION_NAMESPACE {
                return Err(invalid_namespace(
                    context,
                    "the XMLNS namespace URI is reserved for namespace declarations",
                ));
            }
            Ok(Some(uri))
        },
    }
}

/// Compare a resolved namespace against a canonical URI.
///
/// The expected value may be supplied as either a string or a byte slice.  A
/// malformed resolved namespace is an error rather than a non-match, so a
/// caller cannot silently treat an invalid or unknown prefix as opaque typed
/// content.
pub fn namespace_matches<E>(
    namespace: &ResolveResult<'_>,
    expected: E,
    decoder: quick_xml::Decoder,
    context: &str,
) -> Result<bool>
where
    E: AsRef<[u8]>,
{
    let Some(actual) = resolved_namespace_uri(namespace, decoder, context)? else {
        return Ok(false);
    };
    Ok(actual.as_bytes() == expected.as_ref())
}

/// Validate a namespace declaration while retaining its prefix context.
///
/// `ResolveResult` contains the resolved URI but not the prefix that produced
/// it.  Callers scanning `xmlns` declarations should use this function when
/// they have that lexical prefix available.  `None` means the default
/// namespace.  The returned URI is the same borrowed/owned value produced by
/// [`resolved_namespace_uri`].
pub fn validate_namespace_binding<'uri>(
    prefix: Option<&[u8]>,
    namespace: &ResolveResult<'uri>,
    decoder: quick_xml::Decoder,
    context: &str,
) -> Result<Option<Cow<'uri, str>>> {
    if let Some(prefix) = prefix {
        if prefix.is_empty() {
            return Err(invalid_namespace(
                context,
                "a named namespace binding must have a non-empty prefix",
            ));
        }
        if prefix == b"xmlns" {
            return Err(invalid_namespace(
                context,
                "the xmlns prefix is reserved and cannot be declared",
            ));
        }
    }
    let resolved = resolved_namespace_uri(namespace, decoder, context)?;
    let Some(uri) = resolved.as_ref() else {
        if prefix.is_some() {
            return Err(invalid_namespace(
                context,
                "a named namespace binding cannot be unbound",
            ));
        }
        return Ok(None);
    };

    validate_normalized_namespace_binding(prefix, uri, context)?;
    Ok(resolved)
}

/// Validate a namespace binding whose URI has already undergone XML
/// attribute-value normalization.
///
/// This is kept separate from [`validate_namespace_binding`] so readers that
/// normalize a declaration before adding it to a resolver do not normalize it
/// a second time when applying the reserved-prefix rules.
pub fn validate_normalized_namespace_binding(
    prefix: Option<&[u8]>,
    uri: &str,
    context: &str,
) -> Result<()> {
    if let Some(prefix) = prefix {
        if prefix.is_empty() {
            return Err(invalid_namespace(
                context,
                "a named namespace binding must have a non-empty prefix",
            ));
        }
        if prefix == b"xmlns" {
            return Err(invalid_namespace(
                context,
                "the xmlns prefix is reserved and cannot be declared",
            ));
        }
    }
    let prefix = prefix.unwrap_or_default();
    if uri.is_empty() && !prefix.is_empty() {
        return Err(invalid_namespace(
            context,
            "a non-empty prefix cannot be bound to the empty namespace URI",
        ));
    }
    if uri == XMLNS_DECLARATION_NAMESPACE {
        return Err(invalid_namespace(
            context,
            "the XMLNS namespace URI is reserved for namespace declarations",
        ));
    }
    if uri == XMLNS {
        if prefix != b"xml" {
            return Err(invalid_namespace(
                context,
                "the XML namespace URI may only be bound to the xml prefix",
            ));
        }
    } else if prefix == b"xml" {
        return Err(invalid_namespace(
            context,
            "the xml prefix must be bound to the XML namespace URI",
        ));
    }
    Ok(())
}

/// Normalize one lexical namespace declaration value with the shared byte,
/// entity-expansion, and XML-character bounds.
pub fn normalize_namespace_uri<'uri>(
    raw: &'uri [u8],
    decoder: quick_xml::Decoder,
    context: &str,
) -> Result<Cow<'uri, str>> {
    decode_namespace_uri(raw, decoder, context)
}

fn invalid_namespace(context: &str, reason: &str) -> Error {
    Error::InvalidFormat(format!("invalid {context} namespace URI: {reason}"))
}

fn decode_namespace_uri<'uri>(
    raw: &'uri [u8],
    decoder: quick_xml::Decoder,
    context: &str,
) -> Result<Cow<'uri, str>> {
    ensure_utf8_decoder(decoder, context)?;
    if raw.len() > MAX_NAMESPACE_URI_BYTES {
        return Err(invalid_namespace(
            context,
            "the lexical value exceeds the namespace URI limit",
        ));
    }
    if raw.contains(&b'<') {
        return Err(invalid_namespace(
            context,
            "the lexical value contains an unescaped '<'",
        ));
    }

    let decoded = decoder.decode(raw).map_err(|error| {
        invalid_namespace(
            context,
            &format!("the value is not valid in the XML encoding: {error}"),
        )
    })?;
    if decoded.len() > MAX_NAMESPACE_URI_BYTES {
        return Err(invalid_namespace(
            context,
            "the decoded value exceeds the namespace URI limit",
        ));
    }

    let (normalized_len, changed) = namespace_normalized_len(&decoded, context)?;
    if normalized_len > MAX_NAMESPACE_URI_BYTES {
        return Err(invalid_namespace(
            context,
            "entity expansion exceeds the namespace URI limit",
        ));
    }
    if !changed {
        return Ok(decoded);
    }

    let mut output = String::new();
    output
        .try_reserve_exact(normalized_len)
        .map_err(|source| Error::Allocation {
            resource: "ODF namespace URI normalization",
            source,
        })?;
    append_normalized_namespace(&mut output, &decoded, context)?;
    debug_assert_eq!(output.len(), normalized_len);
    Ok(Cow::Owned(output))
}

fn ensure_utf8_decoder(decoder: quick_xml::Decoder, context: &str) -> Result<()> {
    let valid_utf8 = decoder
        .decode(b"\xC3\xA9")
        .map(|value| value.as_ref() == "é")
        .unwrap_or(false);
    let rejects_invalid_utf8 = decoder.decode(b"\xFF").is_err();
    if valid_utf8 && rejects_invalid_utf8 {
        Ok(())
    } else {
        Err(invalid_namespace(
            context,
            "namespace URI resolution requires an XML UTF-8 decoder",
        ))
    }
}

fn namespace_normalized_len(value: &str, context: &str) -> Result<(usize, bool)> {
    let bytes = value.as_bytes();
    let mut index = 0;
    let mut length = 0usize;
    let mut changed = false;
    while index < bytes.len() {
        if bytes[index] == b'&' {
            let relative_end = memchr::memchr(b';', &bytes[index + 1..]).ok_or_else(|| {
                invalid_namespace(context, "an entity reference is not terminated")
            })?;
            let end = index + 1 + relative_end;
            let name = value
                .get(index + 1..end)
                .ok_or_else(|| invalid_namespace(context, "an entity reference is not UTF-8"))?;
            let replacement = namespace_entity(name, context)?;
            length = length
                .checked_add(namespace_entity_len(replacement))
                .ok_or_else(|| invalid_namespace(context, "entity expansion length overflow"))?;
            changed = true;
            index = end + 1;
            continue;
        }

        let character = value[index..]
            .chars()
            .next()
            .ok_or_else(|| invalid_namespace(context, "invalid UTF-8 boundary"))?;
        validate_xml10_character(character, context)?;
        if character == '\r' && value[index + character.len_utf8()..].starts_with('\n') {
            length = length
                .checked_add(1)
                .ok_or_else(|| invalid_namespace(context, "normalized length overflow"))?;
            changed = true;
            index += 2;
            continue;
        }
        let output = normalized_literal_character(character);
        length = length
            .checked_add(output.len_utf8())
            .ok_or_else(|| invalid_namespace(context, "normalized length overflow"))?;
        changed |= output != character;
        index += character.len_utf8();
    }
    Ok((length, changed))
}

fn append_normalized_namespace(output: &mut String, value: &str, context: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'&' {
            let relative_end = memchr::memchr(b';', &bytes[index + 1..]).ok_or_else(|| {
                invalid_namespace(context, "an entity reference is not terminated")
            })?;
            let end = index + 1 + relative_end;
            let name = value
                .get(index + 1..end)
                .ok_or_else(|| invalid_namespace(context, "an entity reference is not UTF-8"))?;
            match namespace_entity(name, context)? {
                NamespaceEntity::Named(replacement) => output.push_str(replacement),
                NamespaceEntity::Character(character) => output.push(character),
            }
            index = end + 1;
            continue;
        }

        let character = value[index..]
            .chars()
            .next()
            .ok_or_else(|| invalid_namespace(context, "invalid UTF-8 boundary"))?;
        validate_xml10_character(character, context)?;
        if character == '\r' && value[index + character.len_utf8()..].starts_with('\n') {
            output.push(' ');
            index += 2;
            continue;
        }
        output.push(normalized_literal_character(character));
        index += character.len_utf8();
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum NamespaceEntity {
    Named(&'static str),
    Character(char),
}

fn namespace_entity(name: &str, context: &str) -> Result<NamespaceEntity> {
    if let Some(value) = quick_xml::escape::resolve_xml_entity(name) {
        return Ok(NamespaceEntity::Named(value));
    }

    let (radix, digits) = if let Some(value) = name.strip_prefix("#x") {
        (16, value)
    } else if let Some(value) = name.strip_prefix('#') {
        (10, value)
    } else {
        return Err(invalid_namespace(
            context,
            "the entity reference is not a predefined XML entity",
        ));
    };
    if digits.is_empty() {
        return Err(invalid_namespace(
            context,
            "the character reference has no digits",
        ));
    }
    let valid_digits = if radix == 16 {
        digits.bytes().all(|value| value.is_ascii_hexdigit())
    } else {
        digits.bytes().all(|value| value.is_ascii_digit())
    };
    if !valid_digits {
        return Err(invalid_namespace(
            context,
            "the character reference contains invalid digits",
        ));
    }
    let codepoint = u32::from_str_radix(digits, radix)
        .map_err(|_| invalid_namespace(context, "the character reference is not numeric"))?;
    let character = char::from_u32(codepoint).ok_or_else(|| {
        invalid_namespace(context, "the character reference is not a Unicode scalar")
    })?;
    validate_xml10_character(character, context)?;
    Ok(NamespaceEntity::Character(character))
}

fn namespace_entity_len(entity: NamespaceEntity) -> usize {
    match entity {
        NamespaceEntity::Named(value) => value.len(),
        NamespaceEntity::Character(value) => value.len_utf8(),
    }
}

fn normalized_literal_character(character: char) -> char {
    match character {
        '\t' | '\n' | '\r' => ' ',
        value => value,
    }
}

fn validate_xml10_character(character: char, context: &str) -> Result<()> {
    let value = character as u32;
    if matches!(
        value,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    ) {
        Ok(())
    } else {
        Err(invalid_namespace(
            context,
            "the value contains a character forbidden by XML 1.0",
        ))
    }
}

pub(crate) fn namespaced_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
    context: &str,
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid {context} attribute: {error}"))
        })?;
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if is_bound(&namespace, expected_namespace) && local_name.as_ref() == expected_local_name {
            if value.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate expanded {context} attribute '{}'",
                    String::from_utf8_lossy(expected_local_name)
                )));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid {context} attribute value: {error}"))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_namespace_constants() {
        // Verify some key namespace constants exist
        assert_eq!(OFFICENS, "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
        assert_eq!(TEXTNS, "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
        assert_eq!(TABLENS, "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
        assert_eq!(STYLENS, "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
    }

    #[test]
    fn test_uri_to_prefix_mapping() {
        assert_eq!(URI_TO_PREFIX.get(OFFICENS), Some(&"office"));
        assert_eq!(URI_TO_PREFIX.get(TEXTNS), Some(&"text"));
        assert_eq!(URI_TO_PREFIX.get(TABLENS), Some(&"table"));
        assert_eq!(URI_TO_PREFIX.get(STYLENS), Some(&"style"));
    }

    #[test]
    fn test_prefix_to_uri_mapping() {
        assert_eq!(PREFIX_TO_URI.get("office"), Some(&OFFICENS));
        assert_eq!(PREFIX_TO_URI.get("text"), Some(&TEXTNS));
        assert_eq!(PREFIX_TO_URI.get("table"), Some(&TABLENS));
        assert_eq!(PREFIX_TO_URI.get("style"), Some(&STYLENS));
    }

    #[test]
    fn test_qualified_name_new() {
        let qn = QualifiedName::new(Some(OFFICENS.to_string()), "body".to_string());
        assert_eq!(qn.local_name, "body");
        assert_eq!(qn.qualified_name, "office:body");
        assert_eq!(qn.namespace_uri, Some(OFFICENS.to_string()));
    }

    #[test]
    fn test_qualified_name_new_no_namespace() {
        let qn = QualifiedName::new(None, "body".to_string());
        assert_eq!(qn.local_name, "body");
        assert_eq!(qn.qualified_name, "body");
        assert_eq!(qn.namespace_uri, None);
    }

    #[test]
    fn test_qualified_name_from_string() {
        let qn = QualifiedName::from_string("office:body");
        assert_eq!(qn.local_name, "body");
        assert_eq!(qn.qualified_name, "office:body");
        assert_eq!(qn.namespace_uri, Some(OFFICENS.to_string()));
    }

    #[test]
    fn test_qualified_name_from_string_no_prefix() {
        let qn = QualifiedName::from_string("body");
        assert_eq!(qn.local_name, "body");
        assert_eq!(qn.qualified_name, "body");
        assert_eq!(qn.namespace_uri, None);
    }

    #[test]
    fn test_qualified_name_from_str() {
        let qn: QualifiedName = "text:p".into();
        assert_eq!(qn.local_name, "p");
        assert_eq!(qn.qualified_name, "text:p");
        assert_eq!(qn.namespace_uri, Some(TEXTNS.to_string()));
    }

    #[test]
    fn test_qualified_name_display() {
        let qn = QualifiedName::from_string("table:table");
        assert_eq!(format!("{qn}"), "table:table");
    }

    #[test]
    fn test_qualified_name_matches() {
        let qn1 = QualifiedName::from_string("office:body");
        let qn2 = QualifiedName::new(Some(OFFICENS.to_string()), "body".to_string());
        let qn3 = QualifiedName::from_string("text:p");

        assert!(qn1.matches(&qn2));
        assert!(!qn1.matches(&qn3));
    }

    #[test]
    fn test_namespace_context_default() {
        let ctx = NamespaceContext::default();
        assert!(ctx.default_namespace.is_none());
        assert!(ctx.prefixes.is_empty());
    }

    #[test]
    fn test_namespace_context_add_namespace() {
        let mut ctx = NamespaceContext::default();
        ctx.add_namespace("xmlns:text", TEXTNS);

        assert_eq!(ctx.resolve_prefix("text"), Some(TEXTNS));
    }

    #[test]
    fn test_namespace_context_add_default_namespace() {
        let mut ctx = NamespaceContext::default();
        ctx.add_namespace("xmlns", OFFICENS);

        assert_eq!(ctx.default_namespace(), Some(OFFICENS));
    }

    #[test]
    fn test_namespace_context_resolve_prefix() {
        let mut ctx = NamespaceContext::default();
        ctx.add_namespace("xmlns:table", TABLENS);
        ctx.add_namespace("xmlns:text", TEXTNS);

        assert_eq!(ctx.resolve_prefix("table"), Some(TABLENS));
        assert_eq!(ctx.resolve_prefix("text"), Some(TEXTNS));
        assert_eq!(ctx.resolve_prefix("office"), None);
    }

    #[test]
    fn test_qualified_name_with_context() {
        let mut ctx = NamespaceContext::default();
        ctx.add_namespace("xmlns:custom", "http://example.com/custom");

        let qn = ctx.parse_qualified_name("custom:element");
        assert_eq!(qn.local_name, "element");
        assert_eq!(
            qn.namespace_uri,
            Some("http://example.com/custom".to_string())
        );
    }

    #[test]
    fn test_qualified_name_with_default_namespace() {
        let mut ctx = NamespaceContext::default();
        ctx.add_namespace("xmlns", OFFICENS);

        let qn = ctx.parse_qualified_name("body");
        assert_eq!(qn.local_name, "body");
        assert_eq!(qn.namespace_uri, Some(OFFICENS.to_string()));
    }
}
