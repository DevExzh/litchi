// Template conversion helpers for LaTeX

use crate::formula::ast::MathNode;

/// Check if base needs grouping for scripts (subscript/superscript)
#[inline]
pub fn needs_grouping_for_scripts(nodes: &[MathNode]) -> bool {
    nodes.len() > 1
}
