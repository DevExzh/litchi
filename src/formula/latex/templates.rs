// Template conversion helpers for LaTeX

use crate::formula::ast::MathNode;

/// Check if base needs grouping for scripts (subscript/superscript)
#[inline]
pub fn needs_grouping_for_scripts(nodes: &[MathNode]) -> bool {
    nodes.len() > 1
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::borrow::Cow;

    #[test]
    fn test_needs_grouping_for_scripts() {
        // Empty slice - no grouping needed
        let empty: &[MathNode] = &[];
        assert!(!needs_grouping_for_scripts(empty));

        // Single element - no grouping needed
        let single: &[MathNode<'_>] = &[MathNode::Text(Cow::Borrowed("x"))];
        assert!(!needs_grouping_for_scripts(single));

        // Multiple elements - grouping needed
        let multiple: &[MathNode<'_>] = &[
            MathNode::Text(Cow::Borrowed("x")),
            MathNode::Operator(crate::formula::ast::Operator::Plus),
            MathNode::Text(Cow::Borrowed("y")),
        ];
        assert!(needs_grouping_for_scripts(multiple));
    }
}
