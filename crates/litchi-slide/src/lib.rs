//! Format-neutral presentation vocabulary.
//!
//! Concrete slide, shape, master, and timing handles remain in PPT/PPTX or
//! other presentation format crates.

#![forbid(unsafe_code)]

use litchi_core::Selector;

/// Semantic layout roles used as the primary slide-creation selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LayoutRole {
    Title,
    TitleAndContent,
    SectionHeader,
    TwoContent,
    Comparison,
    TitleOnly,
    Blank,
    ContentWithCaption,
    PictureWithCaption,
    Custom,
    Unknown,
}

/// Explicit review-history handling during duplication or transfer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Review {
    Keep,
    Drop,
}

/// Appearance policy for cross-presentation transfer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Look {
    Source,
    Destination,
}

/// Convenient selector used by concrete presentation crates.
pub type SlideSelector<'a, Id> = Selector<'a, Id>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn semantic_layout_roles_are_small_values() {
        assert_eq!(std::mem::size_of::<LayoutRole>(), 1);
        assert_ne!(LayoutRole::Blank, LayoutRole::TitleAndContent);
    }
}
