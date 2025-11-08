//! Namespace handling utilities for ODF XML elements.
//!
//! This module provides support for XML namespaces, including qualified names,
//! namespace context, and namespace-aware operations.
//!
//! # Implementation Status
//!
//! ✅ COMPLETED: All ODF 1.2 namespaces (40+ namespaces)
//! ✅ COMPLETED: Extension namespaces (LibreOffice, OpenOffice, KOffice)
//! ✅ COMPLETED: Standard web namespaces (XML, XLink, SVG, MathML, etc.)
//!
//! # References
//!
//! - odfpy: `3rdparty/odfpy/odf/namespaces.py` (lines 24-111)

use phf::{Map, phf_map};
use std::collections::HashMap;

// ============================================================================
// NAMESPACE CONSTANTS
// ============================================================================
// Reference: odfpy/odf/namespaces.py lines 24-66

// Allow dead_code for namespace constants as they are provided as public API for users
// These constants are part of the public API and may not all be used internally

/// Animation namespace
#[allow(dead_code)]
pub const ANIMNS: &str = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0";

/// Chart namespace
#[allow(dead_code)]
pub const CHARTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";

/// OpenOffice chart extensions
#[allow(dead_code)]
pub const CHARTOOONS: &str = "http://openoffice.org/2010/chart";

/// Configuration namespace
#[allow(dead_code)]
pub const CONFIGNS: &str = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

/// CSS3 text extensions
#[allow(dead_code)]
pub const CSS3TNS: &str = "http://www.w3.org/TR/css3-text/";

/// Database namespace
#[allow(dead_code)]
pub const DBNS: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";

/// Dublin Core namespace
#[allow(dead_code)]
pub const DCNS: &str = "http://purl.org/dc/elements/1.1/";

/// DOM events namespace
#[allow(dead_code)]
pub const DOMNS: &str = "http://www.w3.org/2001/xml-events";

/// 3D drawing namespace
#[allow(dead_code)]
pub const DR3DNS: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";

/// Drawing namespace
#[allow(dead_code)]
pub const DRAWNS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

/// OpenOffice field extensions
#[allow(dead_code)]
pub const FIELDNS: &str = "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0";

/// XSL-FO compatible namespace
#[allow(dead_code)]
pub const FONS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

/// Form namespace
#[allow(dead_code)]
pub const FORMNS: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";

/// OOXML-ODF form interoperability
#[allow(dead_code)]
pub const FORMXNS: &str = "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0";

/// GRDDL namespace
#[allow(dead_code)]
pub const GRDDLNS: &str = "http://www.w3.org/2003/g/data-view#";

/// KOffice extensions
#[allow(dead_code)]
pub const KOFFICENS: &str = "http://www.koffice.org/2005/";

/// LibreOffice extensions
#[allow(dead_code)]
pub const LOEXTNS: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

/// Manifest namespace
#[allow(dead_code)]
pub const MANIFESTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

/// MathML namespace
#[allow(dead_code)]
pub const MATHNS: &str = "http://www.w3.org/1998/Math/MathML";

/// Metadata namespace
#[allow(dead_code)]
pub const METANS: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

/// Number/data style namespace
#[allow(dead_code)]
pub const NUMBERNS: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

/// Office namespace
#[allow(dead_code)]
pub const OFFICENS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// OpenFormula namespace (ODF 1.2)
#[allow(dead_code)]
pub const OFNS: &str = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";

/// OpenOffice Calc extensions
#[allow(dead_code)]
pub const OOOCNS: &str = "http://openoffice.org/2004/calc";

/// OpenOffice general extensions
#[allow(dead_code)]
pub const OOONS: &str = "http://openoffice.org/2004/office";

/// OpenOffice Writer extensions
#[allow(dead_code)]
pub const OOOWNS: &str = "http://openoffice.org/2004/writer";

/// Presentation namespace
#[allow(dead_code)]
pub const PRESENTATIONNS: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

/// RDFa namespace
#[allow(dead_code)]
pub const RDFANS: &str = "http://docs.oasis-open.org/opendocument/meta/rdfa#";

/// Report namespace
#[allow(dead_code)]
pub const RPTNS: &str = "http://openoffice.org/2005/report";

/// Script namespace
#[allow(dead_code)]
pub const SCRIPTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";

/// SMIL compatible namespace
#[allow(dead_code)]
pub const SMILNS: &str = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";

/// Style namespace
#[allow(dead_code)]
pub const STYLENS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/// SVG compatible namespace
#[allow(dead_code)]
pub const SVGNS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

/// Table namespace
#[allow(dead_code)]
pub const TABLENS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

/// OpenOffice table extensions
#[allow(dead_code)]
pub const TABLEOOONS: &str = "http://openoffice.org/2009/table";

/// Text namespace
#[allow(dead_code)]
pub const TEXTNS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/// XForms namespace
#[allow(dead_code)]
pub const XFORMSNS: &str = "http://www.w3.org/2002/xforms";

/// XHTML namespace
#[allow(dead_code)]
pub const XHTMLNS: &str = "http://www.w3.org/1999/xhtml";

/// XLink namespace
#[allow(dead_code)]
pub const XLINKNS: &str = "http://www.w3.org/1999/xlink";

/// XML namespace
#[allow(dead_code)]
pub const XMLNS: &str = "http://www.w3.org/XML/1998/namespace";

/// XML Schema namespace
#[allow(dead_code)]
pub const XSDNS: &str = "http://www.w3.org/2001/XMLSchema";

/// XML Schema instance namespace
#[allow(dead_code)]
pub const XSINS: &str = "http://www.w3.org/2001/XMLSchema-instance";

/// Calc extensions (LibreOffice)
#[allow(dead_code)]
pub const CALCEXTNS: &str = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

/// Drawing extensions (OpenOffice)
#[allow(dead_code)]
pub const DRAWOOONS: &str = "http://openoffice.org/2010/draw";

/// Office extensions (OpenOffice)
#[allow(dead_code)]
pub const OFFICEOOONS: &str = "http://openoffice.org/2009/office";

// ============================================================================
// NAMESPACE MAPPING (compile-time perfect hash map)
// ============================================================================
// Reference: odfpy/odf/namespaces.py lines 68-111

/// URI to prefix mapping (compile-time perfect hash map for zero-cost lookups)
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

/// Prefix to URI mapping (compile-time perfect hash map for zero-cost lookups)
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
    /// Namespace URI
    pub namespace_uri: Option<String>,
    /// Local name (without prefix)
    pub local_name: String,
    /// Full qualified name (with prefix if present)
    pub qualified_name: String,
}

impl QualifiedName {
    /// Create a new qualified name
    ///
    /// Note: A clone of `local_name` is necessary when no prefix is needed,
    /// as both fields must be owned strings in the struct.
    pub fn new(namespace_uri: Option<String>, local_name: String) -> Self {
        let qualified_name = match namespace_uri {
            Some(ref uri) => {
                // For common ODF namespaces, use standard prefixes
                let prefix = Self::uri_to_prefix(uri);
                if prefix.is_empty() {
                    // Clone needed: local_name used in both fields
                    local_name.clone()
                } else {
                    format!("{}:{}", prefix, local_name)
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

    /// Convert namespace URI to standard prefix using compile-time perfect hash map
    #[inline]
    fn uri_to_prefix(uri: &str) -> &'static str {
        URI_TO_PREFIX.get(uri).copied().unwrap_or("")
    }

    /// Convert prefix to namespace URI using compile-time perfect hash map
    #[inline]
    fn prefix_to_uri(prefix: &str) -> Option<String> {
        PREFIX_TO_URI.get(prefix).map(|s| (*s).to_string())
    }

    /// Check if this name matches another qualified name
    pub fn matches(&self, other: &QualifiedName) -> bool {
        self.namespace_uri == other.namespace_uri && self.local_name == other.local_name
    }

    /// Check if this name matches a string (with optional namespace resolution)
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
        write!(f, "{}", self.qualified_name)
    }
}

/// Namespace context for resolving prefixes to URIs
#[derive(Debug, Clone, Default)]
pub struct NamespaceContext {
    /// Mapping from prefix to namespace URI
    pub prefixes: HashMap<String, String>,
    /// Default namespace URI
    pub default_namespace: Option<String>,
}

impl NamespaceContext {
    /// Add a namespace declaration
    pub fn add_namespace(&mut self, prefix: &str, uri: &str) {
        if prefix == "xmlns" {
            self.default_namespace = Some(uri.to_string());
        } else if let Some(prefix) = prefix.strip_prefix("xmlns:") {
            self.prefixes.insert(prefix.to_string(), uri.to_string());
        }
    }

    /// Resolve prefix to namespace URI
    pub fn resolve_prefix(&self, prefix: &str) -> Option<&str> {
        self.prefixes.get(prefix).map(|s| s.as_str())
    }

    /// Get default namespace
    pub fn default_namespace(&self) -> Option<&str> {
        self.default_namespace.as_deref()
    }

    /// Parse qualified name with this context
    pub fn parse_qualified_name(&self, name: &str) -> QualifiedName {
        QualifiedName::from_string_with_context(name, Some(self))
    }
}

/// Helper implementation for QualifiedName
impl QualifiedName {
    fn from_string_with_context(name: &str, context: Option<&NamespaceContext>) -> Self {
        if let Some(colon_pos) = name.find(':') {
            let prefix = &name[..colon_pos];
            let local_name = &name[colon_pos + 1..];

            let namespace_uri = if let Some(ctx) = context {
                ctx.resolve_prefix(prefix).map(|s| s.to_string())
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
                ctx.default_namespace().map(|s| s.to_string())
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
