//! Database connection targets.

/// The required metadata on a `db:file-based-database` target.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FileDatabaseTarget {
    href: String,
    media_type: String,
    extension: Option<String>,
}

impl FileDatabaseTarget {
    /// Creates a file target with its normative ODF media type.
    #[must_use]
    pub fn new(href: impl Into<String>, media_type: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            media_type: media_type.into(),
            extension: None,
        }
    }

    /// Retains the optional producer file extension.
    #[must_use]
    pub fn with_extension(mut self, value: Option<String>) -> Self {
        self.extension = value;
        self
    }

    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    #[must_use]
    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    #[must_use]
    pub fn extension(&self) -> Option<&str> {
        self.extension.as_deref()
    }
}

/// The address alternatives permitted by `db:server-database` when an address
/// is present. The ODF grammar also permits a target with neither alternative.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ServerDatabaseAddress {
    /// A host name with an optional positive port.
    Host { hostname: String, port: Option<u64> },
    /// A local socket name.
    LocalSocket(String),
}

/// A normative `db:server-database` target.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ServerDatabaseTarget {
    database_type: String,
    database_type_namespace: Option<String>,
    address: Option<ServerDatabaseAddress>,
    database_name: Option<String>,
}

impl ServerDatabaseTarget {
    /// Creates a server target with its required QName-like database type.
    #[must_use]
    pub fn new(database_type: impl Into<String>) -> Self {
        Self {
            database_type: database_type.into(),
            database_type_namespace: None,
            address: None,
            database_name: None,
        }
    }

    /// Binds the prefix used by the `db:type` QName to its source namespace.
    ///
    /// OpenDocument defines `db:type` as an XML Schema QName and does not
    /// assign a universal URI to application prefixes such as `sdbc`.  Fresh
    /// authors therefore provide the binding when they use a non-standard
    /// prefix.  The binding is retained when a parsed target is rewritten.
    #[must_use]
    pub fn with_database_type_namespace(mut self, value: impl Into<String>) -> Self {
        self.database_type_namespace = Some(value.into());
        self
    }

    /// Sets a host and optional port address.
    #[must_use]
    pub fn with_host(mut self, hostname: impl Into<String>, port: Option<u64>) -> Self {
        self.address = Some(ServerDatabaseAddress::Host {
            hostname: hostname.into(),
            port,
        });
        self
    }

    /// Sets a local socket address.
    #[must_use]
    pub fn with_local_socket(mut self, value: impl Into<String>) -> Self {
        self.address = Some(ServerDatabaseAddress::LocalSocket(value.into()));
        self
    }

    /// Sets the optional database name.
    #[must_use]
    pub fn with_database_name(mut self, value: Option<String>) -> Self {
        self.database_name = value;
        self
    }

    #[must_use]
    pub fn database_type(&self) -> &str {
        &self.database_type
    }

    /// Returns the namespace URI bound to the `db:type` prefix, when known.
    #[must_use]
    pub fn database_type_namespace(&self) -> Option<&str> {
        self.database_type_namespace.as_deref()
    }

    #[must_use]
    pub fn address(&self) -> Option<&ServerDatabaseAddress> {
        self.address.as_ref()
    }

    #[must_use]
    pub fn database_name(&self) -> Option<&str> {
        self.database_name.as_deref()
    }
}

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
    /// A fully described normative `db:file-based-database` target.
    FileTarget(FileDatabaseTarget),
    /// A `db:connection-resource` IRI.
    Resource(String),
    /// A legacy `db:server-database` host/database pair.
    ///
    /// This is retained for reading older sources that omit the normative
    /// `db:type` QName. Transactions preserve it on a semantic no-op but
    /// refuse to author it; use [`Connection::ServerTarget`] with an explicit
    /// QName namespace for fresh or changed output.
    Server { host: String, database: String },
    /// A fully described normative `db:server-database` target.
    ServerTarget(ServerDatabaseTarget),
}

impl Connection {
    pub(crate) fn file(href: String) -> Self {
        Self::File(href)
    }

    /// Creates a file target with required normative metadata.
    #[must_use]
    pub fn file_with_media_type(href: impl Into<String>, media_type: impl Into<String>) -> Self {
        Self::FileTarget(FileDatabaseTarget::new(href, media_type))
    }

    pub(crate) fn resource(href: String) -> Self {
        Self::Resource(href)
    }

    pub(crate) fn server(host: String, database: String) -> Self {
        Self::Server { host, database }
    }

    /// Creates a server target with an unbound QName-like database type.
    ///
    /// Fresh publication requires a namespace binding. Use
    /// [`Self::server_with_type_namespace`] when authoring a target.
    #[must_use]
    pub fn server_with_type(
        database_type: impl Into<String>,
        hostname: impl Into<String>,
        port: Option<u64>,
        database_name: Option<String>,
    ) -> Self {
        Self::ServerTarget(
            ServerDatabaseTarget::new(database_type)
                .with_host(hostname, port)
                .with_database_name(database_name),
        )
    }

    /// Creates a typed server target and binds the `db:type` QName prefix.
    #[must_use]
    pub fn server_with_type_namespace(
        database_type: impl Into<String>,
        database_type_namespace: impl Into<String>,
        hostname: impl Into<String>,
        port: Option<u64>,
        database_name: Option<String>,
    ) -> Self {
        Self::ServerTarget(
            ServerDatabaseTarget::new(database_type)
                .with_database_type_namespace(database_type_namespace)
                .with_host(hostname, port)
                .with_database_name(database_name),
        )
    }

    /// Creates a local-socket server target with an unbound QName-like type.
    ///
    /// Fresh publication requires a namespace binding. Use
    /// [`Self::server_with_local_socket_namespace`] when authoring a target.
    #[must_use]
    pub fn server_with_local_socket(
        database_type: impl Into<String>,
        local_socket: impl Into<String>,
        database_name: Option<String>,
    ) -> Self {
        Self::ServerTarget(
            ServerDatabaseTarget::new(database_type)
                .with_local_socket(local_socket)
                .with_database_name(database_name),
        )
    }

    /// Creates a local-socket server target and binds the `db:type` QName.
    #[must_use]
    pub fn server_with_local_socket_namespace(
        database_type: impl Into<String>,
        database_type_namespace: impl Into<String>,
        local_socket: impl Into<String>,
        database_name: Option<String>,
    ) -> Self {
        Self::ServerTarget(
            ServerDatabaseTarget::new(database_type)
                .with_database_type_namespace(database_type_namespace)
                .with_local_socket(local_socket)
                .with_database_name(database_name),
        )
    }
}
