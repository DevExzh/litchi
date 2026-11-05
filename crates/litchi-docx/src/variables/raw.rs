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
use std::borrow::Cow;

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
    admitted_variables: VariableLimits,
    base_variables: VariableLimits,
}

/// The part of a raw policy which can invalidate an already admitted typed
/// collection. XML size and source depth are checked at output; collection
/// limits can be retained independently so repeated staging does not rescan
/// the whole collection for every set.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct VariableLimits {
    max_variables: usize,
    max_name_chars: usize,
    max_value_chars: usize,
}

impl VariableLimits {
    const fn from_limits(limits: Limits) -> Self {
        Self {
            max_variables: limits.max_variables,
            max_name_chars: limits.max_name_chars,
            max_value_chars: limits.max_value_chars,
        }
    }

    const fn allows(self, limits: Limits) -> bool {
        limits.max_variables >= self.max_variables
            && limits.max_name_chars >= self.max_name_chars
            && limits.max_value_chars >= self.max_value_chars
    }
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
            admitted_variables: VariableLimits::from_limits(limits),
            base_variables: VariableLimits::from_limits(limits),
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
    pub fn set<'name, 'value, N, V>(
        &mut self,
        name: N,
        value: V,
        limits: Limits,
    ) -> Result<Option<String>>
    where
        N: Into<Cow<'name, str>>,
        V: Into<Cow<'value, str>>,
    {
        self.validate_variables_if_tighter(limits)?;
        let name = name.into();
        let value = value.into();
        super::model::validate_document_variable_with_limits(
            name.as_ref(),
            value.as_ref(),
            &limits,
        )?;
        let previous =
            self.variables
                .insert_prevalidated(name.into_owned(), value.into_owned(), &limits)?;
        self.admitted_variables = VariableLimits::from_limits(limits);
        Ok(previous)
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
        self.admitted_variables = VariableLimits::from_limits(limits);
        Ok(())
    }

    /// Apply a clone-staged mutation atomically without serializing per edit.
    pub fn edit_variables(
        &mut self,
        limits: Limits,
        edit: impl FnOnce(&mut Variables) -> Result<()>,
    ) -> Result<()> {
        self.validate_variables_if_tighter(limits)?;
        let mut candidate = self.variables.try_clone()?;
        edit(&mut candidate)?;
        candidate.validate_with_limits(&limits)?;
        self.variables = candidate;
        self.admitted_variables = VariableLimits::from_limits(limits);
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
        self.admitted_variables = self.base_variables;
    }

    fn validate_variables_if_tighter(&self, limits: Limits) -> Result<()> {
        if !self.admitted_variables.allows(limits) {
            self.variables.validate_with_limits(&limits)?;
        }
        Ok(())
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

    #[test]
    fn repeated_sets_reuse_an_admitted_collection_policy() {
        let policy = Limits::new(8 * 1024 * 1024, 64, 1_200, 32, 128).unwrap();
        let mut variables = Variables::new();
        for index in 0..1_000 {
            variables.insert(format!("name-{index}"), "value").unwrap();
        }
        let source = format!(
            r#"<w:settings xmlns:w="{W}">{}</w:settings>"#,
            variables.to_xml(&policy).unwrap()
        );
        let mut tree = Tree::from_xml(source.into_bytes(), policy).unwrap();

        crate::variables::model::reset_validation_count();
        for _ in 0..1_000 {
            tree.set("name-0", "updated", policy).unwrap();
        }
        assert_eq!(crate::variables::model::validation_count(), 0);

        let looser = Limits::new(8 * 1024 * 1024, 64, 2_000, 64, 256).unwrap();
        tree.set("name-0", "looser", looser).unwrap();
        assert_eq!(crate::variables::model::validation_count(), 0);

        let tighter = Limits::new(8 * 1024 * 1024, 64, 1_200, 32, 3).unwrap();
        assert!(tree.set("name-0", "tight", tighter).is_err());
        assert_eq!(crate::variables::model::validation_count(), 1);
    }

    #[test]
    fn cow_values_are_checked_before_ownership_conversion() {
        struct MismatchedInput {
            borrowed: &'static str,
            owned: &'static str,
        }

        impl AsRef<str> for MismatchedInput {
            fn as_ref(&self) -> &str {
                self.borrowed
            }
        }

        impl From<MismatchedInput> for String {
            fn from(_: MismatchedInput) -> String {
                panic!("the legacy owned view must never be selected")
            }
        }

        impl<'a> From<MismatchedInput> for Cow<'a, str> {
            fn from(input: MismatchedInput) -> Cow<'a, str> {
                Cow::Owned(input.owned.to_owned())
            }
        }

        let policy = Limits::new(1024, 64, 4, 8, 3).unwrap();
        let source = format!(r#"<w:settings xmlns:w="{W}"/>"#);
        let mut tree = Tree::from_xml(source.into_bytes(), policy).unwrap();
        let error = tree.set(
            MismatchedInput {
                borrowed: "name",
                owned: "name",
            },
            MismatchedInput {
                borrowed: "value",
                owned: "too-long",
            },
            policy,
        );
        assert!(matches!(
            error,
            Err(Error::SettingsLimit {
                kind: crate::settings::LimitKind::DocumentVariableValueChars,
                ..
            })
        ));
        assert!(!tree.changed());
    }

    #[test]
    fn owned_cow_values_move_into_the_collection_without_copying() {
        let policy = limits(1024);
        let source = format!(r#"<w:settings xmlns:w="{W}"/>"#);
        let mut tree = Tree::from_xml(source.into_bytes(), policy).unwrap();
        let name = String::from("owned-name");
        let value = String::from("owned-value");
        let name_pointer = name.as_ptr();
        let value_pointer = value.as_ptr();

        tree.set(name, value, policy).unwrap();

        let (stored_name, stored_value) = tree.variables().iter().next().unwrap();
        assert_eq!(stored_name.as_ptr(), name_pointer);
        assert_eq!(stored_value.as_ptr(), value_pointer);
    }
}
