//! Inert producer-extension subtrees.

/// One opaque direct child of `office:database` outside the standard office
/// and database namespaces.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ProducerExtension {
    pub(crate) namespace: String,
    pub(crate) local_name: String,
    pub(crate) xml: String,
}

impl ProducerExtension {
    /// Returns the resolved extension namespace URI.
    #[must_use]
    pub fn namespace(&self) -> &str {
        &self.namespace
    }

    /// Returns the namespace-local element name.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Returns the exact inert XML subtree.
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }
}
