use super::node::MathNode;
use bumpalo::Bump;
use std::borrow::Cow;

/// Builder for constructing formula nodes efficiently
pub struct FormulaBuilder<'arena> {
    arena: &'arena Bump,
}

impl<'arena> FormulaBuilder<'arena> {
    /// Create a new builder with the given arena
    pub fn new(arena: &'arena Bump) -> Self {
        Self { arena }
    }

    /// Allocate a string in the arena
    pub fn alloc_str(&self, s: &str) -> &'arena str {
        self.arena.alloc_str(s)
    }

    /// Create a text node
    pub fn text(&self, text: impl Into<Cow<'arena, str>>) -> MathNode<'arena> {
        MathNode::Text(text.into())
    }

    /// Create a number node
    pub fn number(&self, num: impl Into<Cow<'arena, str>>) -> MathNode<'arena> {
        MathNode::Number(num.into())
    }

    /// Create a fraction node
    pub fn frac(
        &self,
        numerator: Vec<MathNode<'arena>>,
        denominator: Vec<MathNode<'arena>>,
    ) -> MathNode<'arena> {
        MathNode::Frac {
            numerator,
            denominator,
            line_thickness: None,
            frac_type: None,
        }
    }

    /// Create a square root node
    pub fn sqrt(&self, base: Vec<MathNode<'arena>>) -> MathNode<'arena> {
        MathNode::Root { base, index: None }
    }

    /// Create an nth root node
    pub fn root(
        &self,
        base: Vec<MathNode<'arena>>,
        index: Vec<MathNode<'arena>>,
    ) -> MathNode<'arena> {
        MathNode::Root {
            base,
            index: Some(index),
        }
    }

    /// Create a power node
    pub fn power(
        &self,
        base: Vec<MathNode<'arena>>,
        exponent: Vec<MathNode<'arena>>,
    ) -> MathNode<'arena> {
        MathNode::Power { base, exponent }
    }

    /// Create a subscript node
    pub fn sub(
        &self,
        base: Vec<MathNode<'arena>>,
        subscript: Vec<MathNode<'arena>>,
    ) -> MathNode<'arena> {
        MathNode::Sub { base, subscript }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::Formula;

    #[test]
    fn test_builder() {
        let formula = Formula::new();
        let builder = FormulaBuilder::new(formula.arena());

        let node = builder.text("x");
        match node {
            MathNode::Text(ref text) => assert_eq!(text, "x"),
            _ => panic!("Expected text node"),
        }
    }
}
