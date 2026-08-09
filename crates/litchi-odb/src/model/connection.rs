//! Database connection targets.

/// A database connection target declared by `db:connection-data`.
///
/// Credentials, driver configuration, and connection attempts are intentionally
/// not modeled here. Callers provide credentials to their concrete database
/// driver; Litchi only reads the inert ODF declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Connection {
    /// A `db:file-based-database` target.
    File(String),
    /// A `db:connection-resource` IRI.
    Resource(String),
    /// A `db:server-database` host/database pair.
    Server { host: String, database: String },
}

impl Connection {
    pub(crate) fn file(href: String) -> Self {
        Self::File(href)
    }

    pub(crate) fn resource(href: String) -> Self {
        Self::Resource(href)
    }

    pub(crate) fn server(host: String, database: String) -> Self {
        Self::Server { host, database }
    }
}
