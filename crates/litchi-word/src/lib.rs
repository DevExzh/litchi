//! Format-neutral Word-processing vocabulary.
//!
//! Concrete stories and nodes remain owned by DOC, DOCX, RTF, or ODT crates.

#![forbid(unsafe_code)]

use std::marker::PhantomData;

/// Origin of a text-bearing story.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StoryKind {
    Main,
    Header,
    Footer,
    Footnote,
    Endnote,
    Comment,
    TextBox,
    Unknown,
}

/// Which semantic projection is traversed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Visibility {
    #[default]
    Visible,
    Review,
    All,
}

/// Where an anchor remains when content is inserted exactly at its position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Affinity {
    Before,
    After,
}

/// Grapheme-cluster position marker.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Grapheme {}

/// Unicode-scalar position marker.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Scalar {}

/// UTF-16 code-unit position marker used by applicable Office encodings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Utf16 {}

/// A typed transient text position with deterministic boundary affinity.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Pos<Unit> {
    offset: u64,
    affinity: Affinity,
    unit: PhantomData<fn() -> Unit>,
}

impl<Unit> Pos<Unit> {
    pub const fn new(offset: u64, affinity: Affinity) -> Self {
        Self {
            offset,
            affinity,
            unit: PhantomData,
        }
    }

    pub const fn offset(self) -> u64 {
        self.offset
    }

    pub const fn affinity(self) -> Affinity {
        self.affinity
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn positions_keep_units_out_of_runtime_storage() {
        assert_eq!(
            std::mem::size_of::<Pos<Grapheme>>(),
            std::mem::size_of::<Pos<Utf16>>()
        );
        let pos = Pos::<Grapheme>::new(4, Affinity::After);
        assert_eq!(pos.offset(), 4);
    }
}
