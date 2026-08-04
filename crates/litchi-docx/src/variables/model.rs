use crate::{Error, Result};

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
    pub const fn new() -> Self {
        Self {
            variables: Vec::new(),
        }
    }

    /// Get a variable value by its case-sensitive OOXML name.
    pub fn get(&self, name: &str) -> Option<&str> {
        self.variables
            .iter()
            .find(|(candidate, _)| candidate == name)
            .map(|(_, value)| value.as_str())
    }

    /// Check whether a variable exists.
    pub fn contains(&self, name: &str) -> bool {
        self.get(name).is_some()
    }

    /// Return variable names in deterministic insertion order.
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
    pub fn count(&self) -> usize {
        self.variables.len()
    }

    /// Whether the collection is empty.
    pub fn is_empty(&self) -> bool {
        self.variables.is_empty()
    }

    /// Insert or replace a variable without changing an existing entry's order.
    pub fn insert(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        let name = name.into();
        let value = value.into();
        validate_document_variable(&name, &value)?;
        if let Some((_, existing)) = self
            .variables
            .iter_mut()
            .find(|(candidate, _)| candidate == &name)
        {
            return Ok(Some(std::mem::replace(existing, value)));
        }
        if self.variables.len() >= MAX_DOCUMENT_VARIABLES {
            return Err(invalid(format!(
                "document variables exceed the {MAX_DOCUMENT_VARIABLES} entry limit"
            )));
        }
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
    pub fn validate(&self) -> Result<()> {
        if self.variables.len() > MAX_DOCUMENT_VARIABLES {
            return Err(invalid(format!(
                "document variables exceed the {MAX_DOCUMENT_VARIABLES} entry limit"
            )));
        }
        for (name, value) in &self.variables {
            validate_document_variable(name, value)?;
        }
        Ok(())
    }

    pub(super) fn push_parsed(&mut self, name: String, value: String) -> Result<()> {
        validate_document_variable(&name, &value)?;
        if self.contains(&name) {
            return Err(invalid(format!(
                "duplicate document variable name {name:?}"
            )));
        }
        if self.variables.len() >= MAX_DOCUMENT_VARIABLES {
            return Err(invalid(format!(
                "document variables exceed the {MAX_DOCUMENT_VARIABLES} entry limit"
            )));
        }
        self.variables.push((name, value));
        Ok(())
    }
}

impl Default for Variables {
    fn default() -> Self {
        Self::new()
    }
}

fn validate_document_variable(name: &str, value: &str) -> Result<()> {
    let name_chars = name.chars().count();
    if !(1..=MAX_DOCUMENT_VARIABLE_NAME_CHARS).contains(&name_chars) {
        return Err(invalid(format!(
            "document variable name must contain 1 to {MAX_DOCUMENT_VARIABLE_NAME_CHARS} characters"
        )));
    }
    if value.chars().count() > MAX_DOCUMENT_VARIABLE_VALUE_CHARS {
        return Err(invalid(format!(
            "document variable value exceeds {MAX_DOCUMENT_VARIABLE_VALUE_CHARS} characters"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
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
