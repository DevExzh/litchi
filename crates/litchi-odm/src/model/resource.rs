//! Inert package-resource graph for master documents.

use litchi_core::Position;

/// One safe package member and the linked sections which target it exactly.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    pub(crate) path: String,
    pub(crate) media_type: Option<String>,
    pub(crate) references: Vec<Position>,
}

impl Resource {
    /// Returns the safe package path without resolving or opening it.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Returns the manifest media type, when declared.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Returns linked-subdocument positions targeting this exact member.
    #[must_use]
    pub fn references(&self) -> &[Position] {
        &self.references
    }
}

/// Immutable package-resource graph.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Graph {
    pub(crate) resources: Vec<Resource>,
    pub(crate) missing: Vec<Position>,
}

impl Graph {
    /// Returns safe package members in package enumeration order.
    #[must_use]
    pub fn resources(&self) -> &[Resource] {
        &self.resources
    }

    /// Returns package-target link positions whose exact member is absent.
    #[must_use]
    pub fn missing(&self) -> &[Position] {
        &self.missing
    }
}
