//! Database connection targets.

/// A database connection target.
///
/// Credentials are intentionally not modeled here; callers provide them to
/// the concrete database driver rather than persisting them in an ODB model.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Connection {
    File(String),
    Resource(String),
    Server { host: String, database: String },
}
