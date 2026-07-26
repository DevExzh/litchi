// Sub-, super- and pre-script parsing
//
// Scripts bind to the atom that precedes them, so the sequence loop in
// `parser` hands the already-parsed base over here. Prescripts use the
// `{}_a^b X` spelling, where an empty group marks the scripts as belonging to
// the atom that follows instead.

use super::error::LatexParseError;
use super::parser::{NumberMode, Parser};
use super::token::TokenKind;
use crate::ast::MathNode;

/// The optional scripts attached to a single base.
#[derive(Debug, Default, Clone, PartialEq)]
pub(crate) struct Scripts<'a> {
    /// Argument of the `_` marker, if one was written.
    pub subscript: Option<Vec<MathNode<'a>>>,
    /// Argument of the `^` marker, if one was written.
    pub superscript: Option<Vec<MathNode<'a>>>,
}

impl<'a> Parser<'a> {
    /// Report whether the cursor is at a `{}_`/`{}^` prescript marker.
    pub(crate) fn at_prescript(&self) -> bool {
        matches!(self.peek().map(|t| t.kind), Some(TokenKind::GroupOpen))
            && matches!(self.peek_at(1).map(|t| t.kind), Some(TokenKind::GroupClose))
            && matches!(
                self.peek_at(2).map(|t| t.kind),
                Some(TokenKind::Subscript | TokenKind::Superscript)
            )
    }

    /// Parse `{}_a^b X` into a pre-script node.
    pub(crate) fn parse_prescript(&mut self) -> Result<MathNode<'a>, LatexParseError> {
        self.advance();
        self.advance();

        let Scripts {
            subscript: pre_subscript,
            superscript: pre_superscript,
        } = self.collect_scripts()?;
        self.skip_spaces();
        let base = match self.parse_atom(NumberMode::Run)? {
            Some(MathNode::Row(inner)) => inner,
            Some(node) => vec![node],
            None => Vec::new(),
        };

        Ok(match (pre_subscript, pre_superscript) {
            (Some(pre_subscript), Some(pre_superscript)) => MathNode::PreSubSup {
                base,
                pre_subscript,
                pre_superscript,
            },
            (Some(pre_subscript), None) => MathNode::PreSub {
                base,
                pre_subscript,
            },
            (None, Some(pre_superscript)) => MathNode::PreSup {
                base,
                pre_superscript,
            },
            // `at_prescript` guarantees at least one marker, so this is only a
            // defensive fallback rather than a reachable state.
            (None, None) => MathNode::Row(base),
        })
    }

    /// Attach the `_`/`^` scripts under the cursor to `base`.
    pub(crate) fn parse_scripts(
        &mut self,
        base: Vec<MathNode<'a>>,
    ) -> Result<MathNode<'a>, LatexParseError> {
        let Scripts {
            subscript,
            superscript,
        } = self.collect_scripts()?;

        Ok(match (subscript, superscript) {
            (Some(subscript), Some(superscript)) => MathNode::SubSup {
                base,
                subscript,
                superscript,
            },
            (Some(subscript), None) => MathNode::Sub { base, subscript },
            (None, Some(exponent)) => MathNode::Power { base, exponent },
            (None, None) => MathNode::Row(base),
        })
    }

    /// Collect the subscript and superscript arguments under the cursor.
    ///
    /// Either may be absent; both orders (`x_i^2` and `x^2_i`) are accepted.
    pub(crate) fn collect_scripts(&mut self) -> Result<Scripts<'a>, LatexParseError> {
        let mut scripts = Scripts::default();

        loop {
            self.skip_spaces();
            let Some(token) = self.peek() else {
                break;
            };
            let slot = match token.kind {
                TokenKind::Subscript => &mut scripts.subscript,
                TokenKind::Superscript => &mut scripts.superscript,
                _ => break,
            };
            if slot.is_some() {
                return Err(LatexParseError::DuplicateScript {
                    position: token.start,
                });
            }
            self.advance();
            *slot = Some(self.parse_script_argument()?);
        }

        Ok(scripts)
    }

    /// Parse the argument of a `_` or `^` marker.
    fn parse_script_argument(&mut self) -> Result<Vec<MathNode<'a>>, LatexParseError> {
        self.skip_spaces();
        let position = self.offset();

        match self.peek().map(|token| token.kind) {
            Some(TokenKind::GroupOpen) => self.parse_group_contents(),
            None => Err(LatexParseError::UnexpectedEndOfInput { position }),
            _ => match self.parse_atom(NumberMode::Single)? {
                Some(MathNode::Row(inner)) => Ok(inner),
                Some(node) => Ok(vec![node]),
                None => Ok(Vec::new()),
            },
        }
    }
}
