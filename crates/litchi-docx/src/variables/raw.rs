//! Mutable, package-neutral settings XML codec model.
//!
//! [`Tree`] owns one complete settings XML payload and its typed document
//! variables.  It is deliberately below `variables::raw`: it is a codec
//! model, not a DOCX package owner.  Mutations are staged in the typed model
//! and are serialized only when [`Tree::to_bytes`] is requested.

use super::codec;
use super::model::{
    MAX_DOCUMENT_VARIABLE_DEPTH, MAX_DOCUMENT_VARIABLE_NAME_CHARS,
    MAX_DOCUMENT_VARIABLE_VALUE_CHARS, MAX_DOCUMENT_VARIABLE_XML_BYTES, MAX_DOCUMENT_VARIABLES,
    Variables,
};
use crate::{Error, Result};

/// Caller-selected admission and output policy for the raw settings codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_xml_bytes: usize,
    max_depth: usize,
    max_variables: usize,
    max_name_chars: usize,
    max_value_chars: usize,
}

impl Limits {
    /// Construct a bounded raw settings policy.
    ///
    /// XML bytes and nesting must be nonzero; collection and character limits
    /// may be zero when an empty model is required. No value may exceed the
    /// format's hard bound. The limits are used for parsing, mutation, and
    /// serialization.
    pub fn new(
        max_xml_bytes: usize,
        max_depth: usize,
        max_variables: usize,
        max_name_chars: usize,
        max_value_chars: usize,
    ) -> Result<Self> {
        let limits = Self {
            max_xml_bytes,
            max_depth,
            max_variables,
            max_name_chars,
            max_value_chars,
        };
        limits.validate()?;
        Ok(limits)
    }

    /// Maximum accepted XML input or output bytes.
    #[must_use]
    pub const fn max_xml_bytes(self) -> usize {
        self.max_xml_bytes
    }

    /// Maximum accepted XML nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum document-variable count.
    #[must_use]
    pub const fn max_variables(self) -> usize {
        self.max_variables
    }

    /// Maximum variable-name length in Unicode scalar values.
    #[must_use]
    pub const fn max_name_chars(self) -> usize {
        self.max_name_chars
    }

    /// Maximum variable-value length in Unicode scalar values.
    #[must_use]
    pub const fn max_value_chars(self) -> usize {
        self.max_value_chars
    }

    fn validate(self) -> Result<()> {
        if self.max_xml_bytes == 0 || self.max_xml_bytes > MAX_DOCUMENT_VARIABLE_XML_BYTES {
            return Err(invalid("raw document-variable XML limit is out of range"));
        }
        if self.max_depth == 0 || self.max_depth > MAX_DOCUMENT_VARIABLE_DEPTH {
            return Err(invalid("raw document-variable depth limit is out of range"));
        }
        if self.max_variables > MAX_DOCUMENT_VARIABLES {
            return Err(invalid("raw document-variable count limit is out of range"));
        }
        if self.max_name_chars > MAX_DOCUMENT_VARIABLE_NAME_CHARS {
            return Err(invalid("raw document-variable name limit is out of range"));
        }
        if self.max_value_chars > MAX_DOCUMENT_VARIABLE_VALUE_CHARS {
            return Err(invalid("raw document-variable value limit is out of range"));
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_xml_bytes: MAX_DOCUMENT_VARIABLE_XML_BYTES,
            max_depth: MAX_DOCUMENT_VARIABLE_DEPTH,
            max_variables: MAX_DOCUMENT_VARIABLES,
            max_name_chars: MAX_DOCUMENT_VARIABLE_NAME_CHARS,
            max_value_chars: MAX_DOCUMENT_VARIABLE_VALUE_CHARS,
        }
    }
}

/// A mutable, package-neutral settings XML codec model.
#[derive(Debug, Clone)]
pub struct Tree {
    source_xml: Vec<u8>,
    source_depth: usize,
    base: Variables,
    variables: Variables,
}

impl Tree {
    /// Parse a complete settings XML payload under the required raw policy.
    ///
    /// The owned `Vec` input is intentional: the caller performs any input
    /// copy before this boundary, so an oversized slice is never copied before
    /// the required byte admission check.
    pub fn from_xml(xml: Vec<u8>, limits: Limits) -> Result<Self> {
        let (variables, source_depth) = codec::parse_variables_with_depth(&xml, &limits)?;
        let base = variables.try_clone()?;
        Ok(Self {
            source_xml: xml,
            source_depth,
            base,
            variables,
        })
    }

    /// Borrow the current typed variables.
    #[must_use]
    pub const fn variables(&self) -> &Variables {
        &self.variables
    }

    /// Return whether the model has staged changes.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.variables != self.base
    }

    /// Set or replace one variable without serializing the whole settings XML.
    pub fn set(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
        limits: Limits,
    ) -> Result<Option<String>> {
        self.variables.validate_with_limits(&limits)?;
        self.variables.insert_with_limits(name, value, &limits)
    }

    /// Remove one variable without serializing the whole settings XML.
    pub fn remove(&mut self, name: &str) -> Option<String> {
        self.variables.remove(name)
    }

    /// Remove every variable without serializing the whole settings XML.
    pub fn clear(&mut self) {
        self.variables.clear();
    }

    /// Replace the typed collection under the required raw policy.
    pub fn replace(&mut self, variables: Variables, limits: Limits) -> Result<()> {
        variables.validate_with_limits(&limits)?;
        self.variables = variables;
        Ok(())
    }

    /// Apply a clone-staged mutation atomically without serializing per edit.
    pub fn edit_variables(
        &mut self,
        limits: Limits,
        edit: impl FnOnce(&mut Variables) -> Result<()>,
    ) -> Result<()> {
        self.variables.validate_with_limits(&limits)?;
        let mut candidate = self.variables.try_clone()?;
        edit(&mut candidate)?;
        candidate.validate_with_limits(&limits)?;
        self.variables = candidate;
        Ok(())
    }

    /// Serialize the current tree under the required raw policy.
    pub fn to_bytes(&self, limits: Limits) -> Result<Vec<u8>> {
        self.variables.validate_with_limits(&limits)?;
        if self.source_depth > limits.max_depth() {
            return Err(limit_error(
                crate::settings::LimitKind::DocumentVariableDepth,
                self.source_depth,
                limits.max_depth(),
            ));
        }
        if !self.changed() {
            if self.source_xml.len() > limits.max_xml_bytes() {
                return Err(limit_error(
                    crate::settings::LimitKind::DocumentVariableXmlBytes,
                    self.source_xml.len(),
                    limits.max_xml_bytes(),
                ));
            }
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(self.source_xml.len())
                .map_err(|source| Error::Allocation {
                    resource: "raw document-variable XML",
                    source,
                })?;
            bytes.extend_from_slice(&self.source_xml);
            return Ok(bytes);
        }
        codec::rewrite(&self.source_xml, &self.base, &self.variables, &limits)
    }

    /// Discard staged values and restore the admitted source model.
    pub fn reset(&mut self) {
        self.variables = self.base.clone();
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit_error(kind: crate::settings::LimitKind, actual: usize, maximum: usize) -> Error {
    Error::SettingsLimit {
        kind,
        actual,
        maximum,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn limits(max_xml_bytes: usize) -> Limits {
        Limits::new(max_xml_bytes, 64, 16, 32, 128).expect("valid raw limits")
    }

    #[test]
    fn tree_batches_mutations_until_output_and_preserves_noop_bytes() {
        let source = format!(
            r#"<w:settings xmlns:w="{W}"><w:docVars><w:docVar w:name="old" w:val="one"/></w:docVars></w:settings>"#
        );
        let policy = limits(1024);
        let mut tree = Tree::from_xml(source.as_bytes().to_vec(), policy).unwrap();
        assert!(!tree.changed());
        assert_eq!(tree.to_bytes(policy).unwrap(), source.as_bytes());

        tree.set("new", "two", policy).unwrap();
        tree.remove("old");
        assert!(tree.changed());
        let output = tree.to_bytes(policy).unwrap();
        assert!(!output.is_empty());
        assert_eq!(tree.variables().get("new"), Some("two"));
        assert_eq!(tree.variables().get("old"), None);
    }

    #[test]
    fn tree_failed_mutations_are_atomic_and_output_limits_apply() {
        let source = format!(r#"<w:settings xmlns:w="{W}"/>"#);
        let policy = limits(256);
        let mut tree = Tree::from_xml(source.as_bytes().to_vec(), policy).unwrap();
        assert!(tree.set("", "invalid", policy).is_err());
        assert!(tree.variables().is_empty());
        tree.set("name", "x".repeat(128), policy).unwrap();
        assert!(tree.to_bytes(policy).is_err());
        tree.reset();
        assert_eq!(tree.to_bytes(policy).unwrap(), source.as_bytes());
    }

    #[test]
    fn unchanged_output_rechecks_a_stricter_caller_limit() {
        let source = format!(r#"<w:settings xmlns:w="{W}"/>"#);
        let admitted = limits(256);
        let tree = Tree::from_xml(source.as_bytes().to_vec(), admitted).unwrap();
        let stricter = limits(source.len() - 1);
        assert!(tree.to_bytes(stricter).is_err());
        assert!(!tree.changed());
    }

    #[test]
    fn unchanged_output_rechecks_depth_and_collection_limits() {
        let source = format!(
            r#"<w:settings xmlns:w="{W}"><w:docVars><w:docVar w:name="name" w:val="value"/></w:docVars></w:settings>"#
        );
        let admitted = limits(1024);
        let tree = Tree::from_xml(source.as_bytes().to_vec(), admitted).unwrap();

        let depth_tight = Limits::new(1024, 1, 16, 32, 128).unwrap();
        assert!(matches!(
            tree.to_bytes(depth_tight),
            Err(Error::SettingsLimit {
                kind: crate::settings::LimitKind::DocumentVariableDepth,
                ..
            })
        ));
        let count_tight = Limits::new(1024, 64, 0, 32, 128).unwrap();
        assert!(matches!(
            tree.to_bytes(count_tight),
            Err(Error::SettingsLimit {
                kind: crate::settings::LimitKind::DocumentVariableCount,
                ..
            })
        ));
    }
}
