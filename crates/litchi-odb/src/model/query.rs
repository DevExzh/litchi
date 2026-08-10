//! Stored query semantics.

use super::table::Column;
use litchi_core::{Error, Result};

/// An inert table target updated by a stored query.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryUpdateTarget {
    name: String,
    schema: Option<String>,
    catalog: Option<String>,
}

impl QueryUpdateTarget {
    /// Creates a named update target.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            schema: None,
            catalog: None,
        }
    }

    /// Sets the optional schema qualifier.
    #[must_use]
    pub fn with_schema_name(mut self, value: impl Into<String>) -> Self {
        self.schema = Some(value.into());
        self
    }

    /// Sets the optional catalog qualifier.
    #[must_use]
    pub fn with_catalog_name(mut self, value: impl Into<String>) -> Self {
        self.catalog = Some(value.into());
        self
    }

    pub(crate) fn parsed(
        name: String,
        schema_name: Option<String>,
        catalog_name: Option<String>,
    ) -> Self {
        Self {
            name,
            schema: schema_name,
            catalog: catalog_name,
        }
    }

    /// Returns the target table name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the schema qualifier, if declared.
    #[must_use]
    pub fn schema_name(&self) -> Option<&str> {
        self.schema.as_deref()
    }

    /// Returns the catalog qualifier, if declared.
    #[must_use]
    pub fn catalog_name(&self) -> Option<&str> {
        self.catalog.as_deref()
    }
}

/// A stored database query.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Query {
    name: String,
    command: String,
    escape_processing: Option<bool>,
    columns: Vec<Column>,
    filter_statement: Option<String>,
    order_statement: Option<String>,
    update_target: Option<QueryUpdateTarget>,
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
            filter_statement: None,
            order_statement: None,
            update_target: None,
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

    /// Sets the inert filter command metadata.
    #[must_use]
    pub fn with_filter_statement(mut self, value: impl Into<String>) -> Self {
        self.filter_statement = Some(value.into());
        self
    }

    /// Sets the inert ordering command metadata.
    #[must_use]
    pub fn with_order_statement(mut self, value: impl Into<String>) -> Self {
        self.order_statement = Some(value.into());
        self
    }

    /// Sets the inert update-table target.
    #[must_use]
    pub fn with_update_target(mut self, value: QueryUpdateTarget) -> Self {
        self.update_target = Some(value);
        self
    }

    pub(crate) fn parsed(name: String, command: String, escape_processing: Option<bool>) -> Self {
        Self {
            name,
            command,
            escape_processing,
            columns: Vec::new(),
            filter_statement: None,
            order_statement: None,
            update_target: None,
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

    pub(crate) fn set_filter_statement(&mut self, value: String) -> Result<()> {
        set_once(&mut self.filter_statement, value, "filter statement")
    }

    pub(crate) fn set_order_statement(&mut self, value: String) -> Result<()> {
        set_once(&mut self.order_statement, value, "order statement")
    }

    pub(crate) fn set_update_target(&mut self, value: QueryUpdateTarget) -> Result<()> {
        if self.update_target.replace(value).is_some() {
            return Err(Error::InvalidFormat(
                "ODB query contains duplicate update targets".to_string(),
            ));
        }
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

    /// Returns the inert filter command, if declared.
    #[must_use]
    pub fn filter_statement(&self) -> Option<&str> {
        self.filter_statement.as_deref()
    }

    /// Returns the inert ordering command, if declared.
    #[must_use]
    pub fn order_statement(&self) -> Option<&str> {
        self.order_statement.as_deref()
    }

    /// Returns the inert update-table target, if declared.
    #[must_use]
    pub const fn update_target(&self) -> Option<&QueryUpdateTarget> {
        self.update_target.as_ref()
    }
}

fn set_once(target: &mut Option<String>, value: String, kind: &str) -> Result<()> {
    if target.replace(value).is_some() {
        return Err(Error::InvalidFormat(format!(
            "ODB query contains duplicate {kind}s"
        )));
    }
    Ok(())
}
