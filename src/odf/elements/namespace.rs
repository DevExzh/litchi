//! Namespace handling utilities for ODF XML elements.
//!
//! This module provides support for XML namespaces, including qualified names,
//! namespace context, and namespace-aware operations.

use std::collections::HashMap;

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

    /// Convert namespace URI to standard prefix
    fn uri_to_prefix(uri: &str) -> &'static str {
        match uri {
            "urn:oasis:names:tc:opendocument:xmlns:text:1.0" => "text",
            "urn:oasis:names:tc:opendocument:xmlns:style:1.0" => "style",
            "urn:oasis:names:tc:opendocument:xmlns:table:1.0" => "table",
            "urn:oasis:names:tc:opendocument:xmlns:draw:1.0" => "draw",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0" => "office",
            "urn:oasis:names:tc:opendocument:xmlns:meta:1.0" => "meta",
            "urn:oasis:names:tc:opendocument:xmlns:fo:1.0" => "fo",
            "http://www.w3.org/1999/xlink" => "xlink",
            "http://www.w3.org/XML/1998/namespace" => "xml",
            _ => "",
        }
    }

    /// Convert prefix to namespace URI
    fn prefix_to_uri(prefix: &str) -> Option<String> {
        match prefix {
            "text" => Some("urn:oasis:names:tc:opendocument:xmlns:text:1.0".to_string()),
            "style" => Some("urn:oasis:names:tc:opendocument:xmlns:style:1.0".to_string()),
            "table" => Some("urn:oasis:names:tc:opendocument:xmlns:table:1.0".to_string()),
            "draw" => Some("urn:oasis:names:tc:opendocument:xmlns:draw:1.0".to_string()),
            "office" => Some("urn:oasis:names:tc:opendocument:xmlns:office:1.0".to_string()),
            "meta" => Some("urn:oasis:names:tc:opendocument:xmlns:meta:1.0".to_string()),
            "fo" => Some("urn:oasis:names:tc:opendocument:xmlns:fo:1.0".to_string()),
            "xlink" => Some("http://www.w3.org/1999/xlink".to_string()),
            "xml" => Some("http://www.w3.org/XML/1998/namespace".to_string()),
            _ => None,
        }
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
