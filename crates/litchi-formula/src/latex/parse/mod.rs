// LaTeX -> AST parsing
//
// This is the inverse of `latex::conv`: it reads LaTeX math source and produces
// the same [`MathNode`] tree that the OMML and MTEF front ends build, so every
// format can be reached from every other one.
//
// The parser borrows from its input. Identifiers, numbers, verbatim `\text`
// arguments and unknown command names all become `Cow::Borrowed` slices of the
// original string, so a parse allocates only the node tree itself.

mod atoms;
pub(crate) mod commands;
mod environments;
mod error;
mod fences;
mod lexer;
mod parser;
mod scripts;
mod token;

#[cfg(test)]
mod tests;

pub use error::LatexParseError;

use crate::ast::MathNode;
use parser::Parser;

/// Default limit on how deeply groups, fences and environments may nest.
///
/// Deeply nested input is rejected rather than risking a stack overflow, which
/// matters because the parser is exposed to untrusted document content.
pub const DEFAULT_MAX_DEPTH: usize = 64;

/// Parses LaTeX math source into the formula AST.
///
/// The parser is deliberately forgiving. Unknown control sequences degrade into
/// [`MathNode::Symbol`], stray alignment markers are dropped, and presentation
/// commands such as `\displaystyle` are consumed silently. Only structurally
/// broken input — unbalanced braces, an unterminated environment, runaway
/// nesting — produces a [`LatexParseError`].
///
/// # Example
///
/// ```ignore
/// use litchi_formula::latex::LatexParser;
///
/// let nodes = LatexParser::new().parse("\\frac{a}{b}")?;
/// assert_eq!(nodes.len(), 1);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LatexParser {
    /// Maximum nesting depth accepted by this parser.
    max_depth: usize,
}

impl LatexParser {
    /// Create a parser using [`DEFAULT_MAX_DEPTH`].
    #[inline]
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_depth: DEFAULT_MAX_DEPTH,
        }
    }

    /// Create a parser with a custom nesting limit.
    ///
    /// A limit of zero is raised to one so that a flat expression still parses.
    #[inline]
    #[must_use]
    pub const fn with_max_depth(max_depth: usize) -> Self {
        Self {
            max_depth: if max_depth == 0 { 1 } else { max_depth },
        }
    }

    /// The nesting limit this parser enforces.
    #[inline]
    #[must_use]
    pub const fn max_depth(&self) -> usize {
        self.max_depth
    }

    /// Parse `input` into a list of AST nodes.
    ///
    /// Surrounding math-mode delimiters (`$...$`, `\(...\)`, `\[...\]`) are
    /// accepted and stripped, so the output of
    /// [`LatexConverter`](crate::latex::LatexConverter) can be fed back in.
    ///
    /// # Errors
    ///
    /// Returns a [`LatexParseError`] when the input is structurally broken.
    pub fn parse<'a>(&self, input: &'a str) -> Result<Vec<MathNode<'a>>, LatexParseError> {
        Parser::new(input, self.max_depth)?.parse()
    }
}

impl Default for LatexParser {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}
