//! Archive-free paragraph base-writing direction.

/// Base direction used to lay out bidirectional paragraph text.
///
/// `Natural` lets iWork infer the direction from the paragraph contents. The
/// explicit variants are useful for neutral or mixed-direction content.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum WritingDirection {
    /// Infer direction from the paragraph contents.
    #[default]
    Natural,
    /// Lay out the paragraph from left to right.
    LeftToRight,
    /// Lay out the paragraph from right to left.
    RightToLeft,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn direction_is_a_closed_semantic_value() {
        assert_eq!(WritingDirection::default(), WritingDirection::Natural);
        assert_ne!(WritingDirection::LeftToRight, WritingDirection::RightToLeft);
    }
}
