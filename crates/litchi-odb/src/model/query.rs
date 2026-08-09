//! Stored query semantics.

/// A stored database query.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Query {
    name: String,
    command: String,
    escape_processing: Option<bool>,
}

impl Query {
    pub fn new(name: impl Into<String>, command: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            command: command.into(),
            escape_processing: None,
        }
    }

    pub(crate) fn parsed(name: String, command: String, escape_processing: Option<bool>) -> Self {
        Self {
            name,
            command,
            escape_processing,
        }
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
}
