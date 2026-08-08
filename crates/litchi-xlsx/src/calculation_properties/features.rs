//! Inert calculation-feature names and their ordered occurrences.

use std::fmt;
use std::str::FromStr;

use crate::error::{Result, invalid};

/// A validated, inert XML 1.0 string used as a calculation-feature name.
///
/// The value is data, not raw markup. Codecs must XML-escape it when writing.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Feature(String);

impl Feature {
    /// Creates a feature name containing only XML 1.0 characters.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_xml_string(&value)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }

    pub fn into_string(self) -> String {
        self.0
    }
}

impl AsRef<str> for Feature {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for Feature {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for Feature {
    type Err = crate::error::Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for Feature {
    type Error = crate::error::Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for Feature {
    type Error = crate::error::Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// A nonempty, ordered, duplicate-preserving collection of feature occurrences.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Features {
    values: Vec<Feature>,
}

impl Features {
    pub fn new(first: Feature) -> Self {
        Self {
            values: vec![first],
        }
    }

    pub fn try_from_vec(values: Vec<Feature>) -> Result<Self> {
        if values.is_empty() {
            return Err(invalid("calculation features must be nonempty"));
        }
        Ok(Self { values })
    }

    pub fn len(&self) -> usize {
        self.values.len()
    }

    pub fn as_slice(&self) -> &[Feature] {
        &self.values
    }

    pub fn iter(&self) -> std::slice::Iter<'_, Feature> {
        self.values.iter()
    }

    pub fn get(&self, index: usize) -> Option<&Feature> {
        self.values.get(index)
    }

    /// Returns the zero-based occurrence of `name` without normalizing it.
    pub fn occurrence(&self, name: &str, occurrence: usize) -> Option<&Feature> {
        self.values
            .iter()
            .filter(|feature| feature.as_str() == name)
            .nth(occurrence)
    }

    pub fn occurrence_count(&self, name: &str) -> usize {
        self.values
            .iter()
            .filter(|feature| feature.as_str() == name)
            .count()
    }

    pub fn push(&mut self, feature: Feature) {
        self.values.push(feature);
    }

    pub fn insert(&mut self, index: usize, feature: Feature) -> Result<()> {
        if index > self.values.len() {
            return Err(invalid("calculation feature index is out of bounds"));
        }
        self.values.insert(index, feature);
        Ok(())
    }

    pub fn replace(&mut self, index: usize, feature: Feature) -> Result<Feature> {
        let slot = self
            .values
            .get_mut(index)
            .ok_or_else(|| invalid("calculation feature index is out of bounds"))?;
        Ok(std::mem::replace(slot, feature))
    }

    pub fn replace_occurrence(
        &mut self,
        name: &str,
        occurrence: usize,
        feature: Feature,
    ) -> Result<Feature> {
        let index = self
            .occurrence_index(name, occurrence)
            .ok_or_else(|| invalid("calculation feature occurrence does not exist"))?;
        self.replace(index, feature)
    }

    /// Removes an occurrence, rejecting removal of the collection's last value.
    pub fn remove(&mut self, index: usize) -> Result<Feature> {
        if self.values.len() == 1 {
            return Err(invalid("calculation features must remain nonempty"));
        }
        if index >= self.values.len() {
            return Err(invalid("calculation feature index is out of bounds"));
        }
        Ok(self.values.remove(index))
    }

    pub fn remove_occurrence(&mut self, name: &str, occurrence: usize) -> Result<Feature> {
        let index = self
            .occurrence_index(name, occurrence)
            .ok_or_else(|| invalid("calculation feature occurrence does not exist"))?;
        self.remove(index)
    }

    pub fn into_vec(self) -> Vec<Feature> {
        self.values
    }

    fn occurrence_index(&self, name: &str, occurrence: usize) -> Option<usize> {
        self.values
            .iter()
            .enumerate()
            .filter(|(_, feature)| feature.as_str() == name)
            .nth(occurrence)
            .map(|(index, _)| index)
    }
}

impl TryFrom<Vec<Feature>> for Features {
    type Error = crate::error::Error;

    fn try_from(values: Vec<Feature>) -> Result<Self> {
        Self::try_from_vec(values)
    }
}

impl IntoIterator for Features {
    type Item = Feature;
    type IntoIter = std::vec::IntoIter<Feature>;

    fn into_iter(self) -> Self::IntoIter {
        self.values.into_iter()
    }
}

impl<'a> IntoIterator for &'a Features {
    type Item = &'a Feature;
    type IntoIter = std::slice::Iter<'a, Feature>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

fn validate_xml_string(value: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(invalid(
            "calculation feature name contains an invalid XML 1.0 character",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preserves_order_duplicates_and_occurrences() {
        let alpha = Feature::new("alpha").unwrap();
        let beta = Feature::new("beta").unwrap();
        let mut features = Features::new(alpha.clone());
        features.push(beta);
        features.push(alpha.clone());
        assert_eq!(features.occurrence_count("alpha"), 2);
        assert_eq!(features.occurrence("alpha", 1), Some(&alpha));
        assert_eq!(features.remove_occurrence("alpha", 0).unwrap(), alpha);
        assert_eq!(features.as_slice()[0].as_str(), "beta");
    }

    #[test]
    fn validates_names_and_nonempty_collection() {
        assert_eq!(Feature::new("").unwrap().as_str(), "");
        assert!(Feature::new("bad\u{0}name").is_err());
        assert!(Features::try_from_vec(Vec::new()).is_err());
    }
}
