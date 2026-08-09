//! Image payload sources.

/// Image payload location.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Source {
    /// An inert URI reference. No I/O is performed implicitly.
    Linked(String),
    /// Decoded bytes from an inline `office:binary-data` child.
    Embedded(Vec<u8>),
}
