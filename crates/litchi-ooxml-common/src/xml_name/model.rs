use std::fmt;
use thiserror::Error;

/// A validated XML Schema `NCName` lexical value.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct NcName(Box<str>);

impl NcName {
    /// Construct a validated name without retaining spare capacity.
    ///
    /// # Errors
    ///
    /// Returns `NameError::InvalidNcName` when the lexical value is not an
    /// XML `NCName`.
    pub fn new(value: impl Into<String>) -> Result<Self, NameError> {
        let lexical = value.into();
        crate::xml_name::validation::validate_ncname(&lexical)?;
        Ok(Self(lexical.into_boxed_str()))
    }

    /// Borrow the lexical name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for NcName {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for NcName {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for NcName {
    type Error = NameError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

/// A validated XML Schema `QName` lexical value.
///
/// The complete lexical value is stored once. Prefix and local-name accessors
/// borrow slices, avoiding a second allocation while keeping the API semantic.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct QualifiedName(Box<str>);

impl QualifiedName {
    pub(crate) fn from_parts(prefix: Option<&NcName>, local: &NcName) -> Self {
        let value = match prefix {
            Some(prefix_name) => format!("{prefix_name}:{local}"),
            None => local.as_str().to_owned(),
        };
        Self(value.into_boxed_str())
    }

    /// Borrow the complete lexical `QName`.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Borrow the optional namespace prefix.
    #[must_use]
    pub fn prefix(&self) -> Option<&str> {
        self.as_str().split_once(':').map(|(prefix, _)| prefix)
    }

    /// Borrow the local part.
    #[must_use]
    pub fn local(&self) -> &str {
        self.as_str()
            .split_once(':')
            .map_or_else(|| self.as_str(), |(_, local)| local)
    }
}

impl AsRef<str> for QualifiedName {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for QualifiedName {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for QualifiedName {
    type Error = NameError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        crate::xml_name::codec::parse(value)
    }
}

/// A lexical XML name failure.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum NameError {
    /// The value is not a valid XML `NCName`.
    #[error("invalid XML NCName '{0}'")]
    InvalidNcName(String),
    /// The value is not a valid XML `QName`.
    #[error("invalid XML QName '{0}'")]
    InvalidQualifiedName(String),
}
