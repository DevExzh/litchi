//! Typed `ShtProps` requests and reversible change metadata.

use crate::chart::Props;

/// One source-checked semantic `ShtProps` change.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Change {
    before: Props,
    after: Props,
    offset: usize,
}

impl Change {
    pub(crate) const fn new(before: Props, after: Props, offset: usize) -> Self {
        Self {
            before,
            after,
            offset,
        }
    }

    /// Properties required before this change can be applied.
    pub const fn before(self) -> Props {
        self.before
    }

    /// Properties produced by this change.
    pub const fn after(self) -> Props {
        self.after
    }

    /// Source-relative byte offset of the fixed `ShtProps` record header.
    pub const fn offset(self) -> usize {
        self.offset
    }

    pub(crate) const fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
            offset: self.offset,
        }
    }
}

/// One staged `ShtProps` replacement, kept private to the transaction facade.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Request {
    pub(crate) value: Props,
    pub(crate) expected_offset: Option<usize>,
}
