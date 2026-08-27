//! Validated, archive-free names for Keynote slide tables.
//!
//! A table name is a semantic value, not a native object identifier or a
//! protobuf/string-field view. The package adapter validates names while
//! crossing the archive boundary and owns any transaction-specific wire
//! state. Keeping this type here means public reads and edits can exchange an
//! owned value without exposing package, archive, or native identifiers.

use std::fmt;

/// Why a slide-table name was rejected before it was stored.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Error {
    /// A table name must contain at least one UTF-8 byte.
    Empty,
    /// Native string fields use NUL as a terminator and cannot contain it.
    ContainsNul,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Empty => "Keynote slide-table names must not be empty",
            Self::ContainsNul => "Keynote slide-table names cannot contain NUL",
        })
    }
}

impl std::error::Error for Error {}

/// Result type for constructing a validated slide-table name.
pub type Result<T> = std::result::Result<T, Error>;

/// A validated, owned Keynote slide-table name.
//
// The value intentionally contains only UTF-8 semantic text. In particular,
// it does not retain a package reference, archive member, protobuf message,
// native object identifier, or selector position.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Name(Box<str>);

impl Name {
    /// Validate and own a slide-table name.
    //
    // Validation precedes allocation for borrowed input. A package owner
    // should use this constructor for names decoded from native payloads and
    // use [`TryFrom<String>`] when it already owns the decoded string.
    pub fn new(value: &str) -> Result<Self> {
        validate(value)?;
        Ok(Self(value.into()))
    }

    /// Borrow the validated table name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the validated name as an owned string.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for Name {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl std::borrow::Borrow<str> for Name {
    fn borrow(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<&str> for Name {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for Name {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        validate(&value)?;
        Ok(Self(value.into_boxed_str()))
    }
}

impl From<Name> for String {
    fn from(value: Name) -> Self {
        value.into_string()
    }
}

impl fmt::Display for Name {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

fn validate(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::Empty);
    }
    if value.contains('\0') {
        return Err(Error::ContainsNul);
    }
    Ok(())
}

/// Exact-source transaction types for a rooted Keynote slide-table name.
pub mod transaction {
    pub use crate::package::slide_table_name::{
        SlideTableNameCommit as Commit, SlideTableNameDiagnostics as Diagnostics,
        SlideTableNameEdit as Edit, SlideTableNameError as Error,
        SlideTableNameLimitKind as LimitKind, SlideTableNamePatch as Patch,
        SlideTableNamePath as Path,
    };
}
