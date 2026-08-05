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

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn command(&self) -> &str {
        &self.command
    }
}
