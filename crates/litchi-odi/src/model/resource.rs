//! Package-local image resource inventory.

/// A package-local image reference discovered in `content.xml`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    frame: usize,
    href: String,
    path: String,
    media_type: Option<String>,
    present: bool,
}

/// A bounded package resource graph, including unreferenced inert members.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Graph {
    nodes: Vec<Node>,
    edges: Vec<Edge>,
}

impl Graph {
    pub(crate) fn new(nodes: Vec<Node>, edges: Vec<Edge>) -> Self {
        Self { nodes, edges }
    }

    /// Returns package and missing-target nodes in deterministic path order.
    #[must_use]
    pub fn nodes(&self) -> &[Node] {
        &self.nodes
    }

    /// Returns frame-to-resource references in semantic frame order.
    #[must_use]
    pub fn edges(&self) -> &[Edge] {
        &self.edges
    }
}

/// One resource-graph node.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Node {
    path: String,
    media_type: Option<String>,
    present: bool,
    referenced: bool,
}

impl Node {
    pub(crate) fn new(
        path: String,
        media_type: Option<String>,
        present: bool,
        referenced: bool,
    ) -> Self {
        Self {
            path,
            media_type,
            present,
            referenced,
        }
    }

    /// Returns the safe package path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Returns the manifest media type, if declared.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Returns whether this node exists in the archive.
    #[must_use]
    pub const fn is_present(&self) -> bool {
        self.present
    }

    /// Returns whether at least one document frame references this node.
    #[must_use]
    pub const fn is_referenced(&self) -> bool {
        self.referenced
    }
}

/// One frame-to-package-resource graph edge.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Edge {
    frame: usize,
    href: String,
    node: usize,
}

impl Edge {
    pub(crate) fn new(frame: usize, href: String, node: usize) -> Self {
        Self { frame, href, node }
    }

    /// Returns the source semantic frame position.
    #[must_use]
    pub const fn frame(&self) -> usize {
        self.frame
    }

    /// Returns the original inert URI reference.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Returns the target node position in [`Graph::nodes`].
    #[must_use]
    pub const fn node(&self) -> usize {
        self.node
    }
}

impl Resource {
    pub(crate) fn new(
        frame: usize,
        href: String,
        path: String,
        media_type: Option<String>,
        present: bool,
    ) -> Self {
        Self {
            frame,
            href,
            path,
            media_type,
            present,
        }
    }

    /// Returns the zero-based semantic frame position that references this resource.
    #[must_use]
    pub fn frame(&self) -> usize {
        self.frame
    }

    /// Returns the original inert `xlink:href` value.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Returns the safely resolved package path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Returns the manifest media type, if declared.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Returns whether the referenced package member exists.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.present
    }
}
