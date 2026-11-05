#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
use crate::{Error, Result};

use super::raw::Limits;

#[cfg(test)]
use std::cell::Cell;

#[cfg(test)]
thread_local! {
    static VALIDATION_COUNT: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
pub(crate) fn reset_validation_count() {
    VALIDATION_COUNT.with(|count| count.set(0));
}

#[cfg(test)]
pub(crate) fn validation_count() -> usize {
    VALIDATION_COUNT.with(Cell::get)
}

/// Maximum number of document variables accepted by the bounded codec.
pub const MAX_DOCUMENT_VARIABLES: usize = 4096;
/// Maximum XML input accepted by the bounded document-variable codec.
pub const MAX_DOCUMENT_VARIABLE_XML_BYTES: usize = 8 * 1024 * 1024;
/// Maximum XML nesting depth accepted by the bounded document-variable codec.
pub const MAX_DOCUMENT_VARIABLE_DEPTH: usize = 64;
/// Maximum number of Unicode scalar values in a Word document-variable name.
pub const MAX_DOCUMENT_VARIABLE_NAME_CHARS: usize = 255;
/// Maximum number of Unicode scalar values in a Word document-variable value.
pub const MAX_DOCUMENT_VARIABLE_VALUE_CHARS: usize = 65_280;

/// Deterministic insertion-order collection of document variables.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Variables {
    variables: Vec<(String, String)>,
}

impl Variables {
    /// Create an empty collection.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            variables: Vec::new(),
        }
    }

    /// Get a variable value by its case-sensitive OOXML name.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&str> {
        self.variables
            .iter()
            .find(|(candidate, _)| candidate == name)
            .map(|(_, value)| value.as_str())
    }

    /// Check whether a variable exists.
    #[must_use]
    pub fn contains(&self, name: &str) -> bool {
        self.get(name).is_some()
    }

    /// Return variable names in deterministic insertion order.
    #[must_use]
    pub fn names(&self) -> Vec<&str> {
        self.variables
            .iter()
            .map(|(name, _)| name.as_str())
            .collect()
    }

    /// Iterate in deterministic insertion order.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.variables
            .iter()
            .map(|(name, value)| (name.as_str(), value.as_str()))
    }

    /// Number of variables.
    #[must_use]
    pub fn count(&self) -> usize {
        self.variables.len()
    }

    /// Whether the collection is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.variables.is_empty()
    }

    /// Insert or replace a variable without changing an existing entry's order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn insert(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        self.insert_with_limits(name, value, &Limits::default())
    }

    pub(crate) fn insert_with_limits(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
        limits: &Limits,
    ) -> Result<Option<String>> {
        let name = name.into();
        let value = value.into();
        validate_document_variable_with_limits(&name, &value, limits)?;
        self.insert_prevalidated(name, value, limits)
    }

    pub(super) fn insert_prevalidated(
        &mut self,
        name: String,
        value: String,
        limits: &Limits,
    ) -> Result<Option<String>> {
        if let Some((_, existing)) = self
            .variables
            .iter_mut()
            .find(|(candidate, _)| candidate == &name)
        {
            return Ok(Some(std::mem::replace(existing, value)));
        }
        if self.variables.len() >= limits.max_variables() {
            return Err(limit_error(
                crate::settings::LimitKind::DocumentVariableCount,
                self.variables.len().saturating_add(1),
                limits.max_variables(),
            ));
        }
        self.variables
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "document-variable collection",
                source,
            })?;
        self.variables.push((name, value));
        Ok(None)
    }

    /// Remove a variable while preserving the order of all remaining entries.
    pub fn remove(&mut self, name: &str) -> Option<String> {
        let index = self
            .variables
            .iter()
            .position(|(candidate, _)| candidate == name)?;
        Some(self.variables.remove(index).1)
    }

    /// Remove all variables.
    pub fn clear(&mut self) {
        self.variables.clear();
    }

    /// Validate all collection and attribute limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        self.validate_with_limits(&Limits::default())
    }

    pub(crate) fn validate_with_limits(&self, limits: &Limits) -> Result<()> {
        #[cfg(test)]
        VALIDATION_COUNT.with(|count| count.set(count.get().saturating_add(1)));
        if self.variables.len() > limits.max_variables() {
            return Err(limit_error(
                crate::settings::LimitKind::DocumentVariableCount,
                self.variables.len(),
                limits.max_variables(),
            ));
        }
        for (name, value) in &self.variables {
            validate_document_variable_with_limits(name, value, limits)?;
        }
        Ok(())
    }

    pub(super) fn push_parsed_with_limits(
        &mut self,
        name: String,
        value: String,
        limits: &Limits,
    ) -> Result<()> {
        validate_document_variable_with_limits(&name, &value, limits)?;
        if self.variables.len() >= limits.max_variables() {
            return Err(limit_error(
                crate::settings::LimitKind::DocumentVariableCount,
                self.variables.len().saturating_add(1),
                limits.max_variables(),
            ));
        }
        self.variables
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "parsed document-variable collection",
                source,
            })?;
        self.variables.push((name, value));
        Ok(())
    }

    pub(super) fn try_clone(&self) -> Result<Self> {
        let mut variables = Vec::new();
        variables
            .try_reserve_exact(self.variables.len())
            .map_err(|source| Error::Allocation {
                resource: "document-variable edit collection",
                source,
            })?;
        for (name, value) in &self.variables {
            variables.push((try_clone_string(name)?, try_clone_string(value)?));
        }
        Ok(Self { variables })
    }
}

impl Default for Variables {
    fn default() -> Self {
        Self::new()
    }
}

pub(super) fn validate_document_variable_with_limits(
    name: &str,
    value: &str,
    limits: &Limits,
) -> Result<()> {
    let name_chars = name.chars().count();
    if !(1..=limits.max_name_chars()).contains(&name_chars) {
        return Err(limit_error(
            crate::settings::LimitKind::DocumentVariableNameChars,
            name_chars,
            limits.max_name_chars(),
        ));
    }
    if value.chars().count() > limits.max_value_chars() {
        return Err(limit_error(
            crate::settings::LimitKind::DocumentVariableValueChars,
            value.chars().count(),
            limits.max_value_chars(),
        ));
    }
    Ok(())
}

fn try_clone_string(value: &str) -> Result<String> {
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation {
            resource: "document-variable string clone",
            source,
        })?;
    copy.push_str(value);
    Ok(copy)
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

    #[test]
    fn insertion_order_replace_remove_and_clear_are_deterministic() {
        let mut variables = Variables::new();
        assert_eq!(variables.insert("first", "one").unwrap(), None);
        assert_eq!(variables.insert("second", "two").unwrap(), None);
        assert_eq!(
            variables.insert("first", "updated").unwrap(),
            Some("one".into())
        );
        assert_eq!(variables.names(), vec!["first", "second"]);
        assert_eq!(variables.remove("second"), Some("two".into()));
        variables.clear();
        assert!(variables.is_empty());
    }

    #[test]
    fn enforces_word_name_value_and_count_boundaries() {
        let mut variables = Variables::new();
        assert!(variables.insert("", "value").is_err());
        assert!(variables.insert("名".repeat(255), "").is_ok());
        assert!(variables.insert("名".repeat(256), "value").is_err());
        assert!(variables.insert("maximum", "x".repeat(65_280)).is_ok());
        assert!(variables.insert("too-long", "x".repeat(65_281)).is_err());

        let mut count = Variables::new();
        for index in 0..MAX_DOCUMENT_VARIABLES {
            count.insert(format!("v{index}"), "x").unwrap();
        }
        assert!(count.insert("overflow", "x").is_err());
    }
}
