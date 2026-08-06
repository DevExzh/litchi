//! Stored query semantics.

/// A stored database query.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Query {
    name: String,
    command: String,
}

impl Query {
    pub fn new(name: impl Into<String>, command: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            command: command.into(),
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
}
