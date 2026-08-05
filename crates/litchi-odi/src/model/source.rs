//! Image payload sources.

/// Image payload location.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Source {
    Linked(String),
    Embedded(Vec<u8>),
}
