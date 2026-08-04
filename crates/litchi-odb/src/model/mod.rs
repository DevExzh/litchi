//! Immutable semantic values for this document family.

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
/// A database connection target. Credentials are intentionally not modeled here.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Connection {
    File(String),
    Resource(String),
    Server { host: String, database: String },
}
