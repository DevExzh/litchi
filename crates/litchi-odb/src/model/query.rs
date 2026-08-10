//! Stored query semantics.

use super::table::Column;
use litchi_core::{Error, Result};

/// A stored database query.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Query {
    name: String,
    command: String,
    escape_processing: Option<bool>,
    columns: Vec<Column>,
}

impl Query {
    /// Creates an inert stored-query declaration.
    #[must_use]
    pub fn new(name: impl Into<String>, command: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            command: command.into(),
            escape_processing: None,
            columns: Vec::new(),
        }
    }

    /// Sets the optional ODF escape-processing declaration.
    #[must_use]
    pub const fn with_escape_processing(mut self, value: Option<bool>) -> Self {
        self.escape_processing = value;
        self
    }

    /// Appends one inert query result-column presentation declaration.
    #[must_use]
    pub fn with_column(mut self, value: Column) -> Self {
        self.columns.push(value);
        self
    }

    pub(crate) fn parsed(name: String, command: String, escape_processing: Option<bool>) -> Self {
        Self {
            name,
            command,
            escape_processing,
            columns: Vec::new(),
        }
    }

    pub(crate) fn try_push_column(&mut self, value: Column) -> Result<()> {
        self.columns
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB query columns",
                source,
            })?;
        self.columns.push(value);
        Ok(())
    }

    /// Returns the query name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the query command text.
    #[must_use]
    pub fn command(&self) -> &str {
        &self.command
    }

    /// Returns the ODF escape-processing declaration, if the producer stored one.
    ///
    /// This metadata is descriptive only. Litchi never parses, connects to, or
    /// executes the command.
    #[must_use]
    pub const fn escape_processing(&self) -> Option<bool> {
        self.escape_processing
    }

    /// Returns inert query result-column declarations in source order.
    #[must_use]
    pub fn columns(&self) -> &[Column] {
        &self.columns
    }
}
