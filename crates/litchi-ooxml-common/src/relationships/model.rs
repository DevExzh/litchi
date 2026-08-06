//! Semantic values for OOXML relationship attributes.

use std::fmt;

use crate::XmlError;

/// A checked OPC relationship identifier (`xsd:ID`/`ST_RelationshipId`).
///
/// Relationship identifiers are NCNames, not arbitrary strings. Keeping the
/// lexical check at the shared boundary lets package owners use a compact
/// value without repeating validation or carrying spare `String` capacity.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Id(Box<str>);

impl Id {
    /// Construct a checked relationship identifier.
    pub fn new(value: impl Into<String>) -> Result<Self, XmlError> {
        let value = value.into();
        if !crate::xml_name::is_ncname(&value) {
            return Err(XmlError::Invalid(format!(
                "invalid OOXML relationship identifier '{value}'"
            )));
        }
        Ok(Self(value.into_boxed_str()))
    }

    /// Borrow the identifier's lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Return the owned lexical value.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for Id {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for Id {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for Id {
    type Error = XmlError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl TryFrom<String> for Id {
    type Error = XmlError;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Id> for String {
    fn from(value: Id) -> Self {
        value.into_string()
    }
}
