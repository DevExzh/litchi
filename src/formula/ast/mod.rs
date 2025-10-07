// Abstract Syntax Tree for Mathematical Formulas
//
// This module defines a comprehensive AST for representing mathematical formulas
// that can be parsed from OMML, LaTeX, and MTEF formats and converted between them.
//
// The design is inspired by the plurimath Ruby project but adapted for Rust's
// type system and performance characteristics.

mod types;
mod node;
mod builder;

pub use types::*;
pub use node::MathNode;
pub use builder::FormulaBuilder;

use bumpalo::Bump;

/// Arena-allocated formula AST
///
/// This struct uses bump allocation for efficient memory management when
/// parsing and converting large formulas. All nodes are allocated from
/// the internal arena and have a lifetime tied to the `Formula` instance.
pub struct Formula<'arena> {
    arena: Bump,
    root: Vec<MathNode<'arena>>,
    display_style: bool,
}

impl<'arena> Formula<'arena> {
    /// Create a new formula with default display style
    pub fn new() -> Self {
        Self {
            arena: Bump::new(),
            root: Vec::new(),
            display_style: true,
        }
    }

    /// Create a new formula with specified display style
    pub fn with_display_style(display_style: bool) -> Self {
        Self {
            arena: Bump::new(),
            root: Vec::new(),
            display_style,
        }
    }

    /// Get the arena allocator for this formula
    #[inline]
    pub fn arena(&self) -> &Bump {
        &self.arena
    }

    /// Get the root nodes
    #[inline]
    pub fn root(&self) -> &[MathNode<'arena>] {
        &self.root
    }

    /// Set the root nodes
    pub fn set_root(&mut self, root: Vec<MathNode<'arena>>) {
        self.root = root;
    }

    /// Get display style
    #[inline]
    pub fn display_style(&self) -> bool {
        self.display_style
    }

    /// Set display style
    pub fn set_display_style(&mut self, display_style: bool) {
        self.display_style = display_style;
    }
}

impl Default for Formula<'_> {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_formula_creation() {
        let formula = Formula::new();
        assert!(formula.root().is_empty());
        assert!(formula.display_style());
    }
}

