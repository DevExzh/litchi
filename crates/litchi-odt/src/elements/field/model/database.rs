//! Database field identities and bounded source values.

#![allow(
    clippy::wildcard_imports,
    reason = "semantic field owners share the stable model facade namespace"
)]
use super::*;
/// One of the five OpenDocument database field elements.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DatabaseFieldKind {
    Display,
    Next,
    RowSelect,
    RowNumber,
    Name,
}

/// Kind of database object selected by a field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DatabaseTableType {
    Table,
    Query,
    Command,
}

impl DatabaseTableType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "table" => Ok(Self::Table),
            "query" => Ok(Self::Query),
            "command" => Ok(Self::Command),
            _ => Err(Error::InvalidFormat(format!(
                "invalid database table type '{value}'"
            ))),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Table => "table",
            Self::Query => "query",
            Self::Command => "command",
        }
    }
}

/// An inert `form:connection-resource`. The URI is never resolved or opened.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseConnectionResource {
    pub href: String,
    pub simple_link: bool,
}

/// Common source identity shared by all database fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseSource {
    pub database_name: Option<String>,
    pub table_name: String,
    pub table_type: Option<DatabaseTableType>,
    pub connection_resource: Option<DatabaseConnectionResource>,
}

/// Canonical, bounded XML Schema `nonNegativeInteger` without arithmetic semantics.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NonNegativeInteger(String);

impl NonNegativeInteger {
    pub fn new(lexical: &str) -> Result<Self> {
        let lexical = lexical.trim_matches(|ch| matches!(ch, ' ' | '\t' | '\n' | '\r'));
        let (negative, digits) = match lexical.as_bytes().first() {
            Some(b'+') => (false, &lexical[1..]),
            Some(b'-') => (true, &lexical[1..]),
            _ => (false, lexical),
        };
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(Error::InvalidFormat(format!(
                "invalid XML Schema nonNegativeInteger '{lexical}'"
            )));
        }
        if digits.len() > MAX_DATABASE_INTEGER_DIGITS {
            return Err(Error::InvalidFormat(format!(
                "nonNegativeInteger exceeds {MAX_DATABASE_INTEGER_DIGITS} digits"
            )));
        }
        if negative && digits.bytes().any(|byte| byte != b'0') {
            return Err(Error::InvalidFormat(format!(
                "negative value is not a nonNegativeInteger '{lexical}'"
            )));
        }
        let canonical = digits.trim_start_matches('0');
        Ok(Self(
            if canonical.is_empty() { "0" } else { canonical }.to_string(),
        ))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl DatabaseSource {
    /// ODF defaults `text:table-type` to `table`.
    pub fn effective_table_type(&self) -> DatabaseTableType {
        self.table_type.unwrap_or(DatabaseTableType::Table)
    }
}

/// Typed, non-executing database field metadata in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseField {
    pub kind: DatabaseFieldKind,
    pub source: DatabaseSource,
    pub column_name: Option<String>,
    pub condition: Option<String>,
    pub row_number: Option<NonNegativeInteger>,
    pub value: Option<NonNegativeInteger>,
    pub data_style_name: Option<String>,
    pub number_format: Option<String>,
    pub number_letter_sync: Option<bool>,
    pub display_text: String,
}
