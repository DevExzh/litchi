// `\begin{...}` / `\end{...}` handling
//
// Matrix-like environments become [`MathNode::Matrix`] with the fence the
// environment name implies; alignment-like environments become
// [`MathNode::EqArray`], whose rows are flat node lists.

use super::commands::{CMD_END, CMD_LINE_BREAK};
use super::error::LatexParseError;
use super::parser::{Limit, Parser};
use super::token::{Token, TokenKind};
use crate::ast::{MathNode, MatrixFence};

/// How an environment body is turned into a node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum EnvKind {
    /// A grid of cells rendered as [`MathNode::Matrix`].
    Matrix {
        /// The fence drawn around the grid.
        fence: MatrixFence,
        /// Whether a `{lcr}` column specification precedes the body.
        column_spec: bool,
    },
    /// A list of aligned rows rendered as [`MathNode::EqArray`].
    EqArray,
}

/// Environment names understood by the parser.
///
/// The matrix entries mirror `latex::matrix::matrix_fence_to_env`, so a matrix
/// node renders and re-parses to the same fence.
static ENVIRONMENTS: phf::Map<&'static str, EnvKind> = phf::phf_map! {
    "matrix"      => EnvKind::Matrix { fence: MatrixFence::None,       column_spec: false },
    "pmatrix"     => EnvKind::Matrix { fence: MatrixFence::Paren,      column_spec: false },
    "bmatrix"     => EnvKind::Matrix { fence: MatrixFence::Bracket,    column_spec: false },
    "Bmatrix"     => EnvKind::Matrix { fence: MatrixFence::Brace,      column_spec: false },
    "vmatrix"     => EnvKind::Matrix { fence: MatrixFence::Pipe,       column_spec: false },
    "Vmatrix"     => EnvKind::Matrix { fence: MatrixFence::DoublePipe, column_spec: false },
    "smallmatrix" => EnvKind::Matrix { fence: MatrixFence::None,       column_spec: false },
    "cases"       => EnvKind::Matrix { fence: MatrixFence::Brace,      column_spec: false },
    "array"       => EnvKind::Matrix { fence: MatrixFence::None,       column_spec: true },

    "aligned"     => EnvKind::EqArray,
    "align"       => EnvKind::EqArray,
    "align*"      => EnvKind::EqArray,
    "alignat"     => EnvKind::EqArray,
    "alignat*"    => EnvKind::EqArray,
    "gathered"    => EnvKind::EqArray,
    "gather"      => EnvKind::EqArray,
    "gather*"     => EnvKind::EqArray,
    "split"       => EnvKind::EqArray,
    "substack"    => EnvKind::EqArray,
    "eqnarray"    => EnvKind::EqArray,
    "eqnarray*"   => EnvKind::EqArray,
};

/// Environment used for names the parser does not recognise.
const UNKNOWN_ENVIRONMENT: EnvKind = EnvKind::EqArray;

impl<'a> Parser<'a> {
    /// Parse a `\begin{...} ... \end{...}` block.
    ///
    /// `begin` is the `\begin` token, used only for error reporting.
    pub(crate) fn parse_environment(
        &mut self,
        begin: Token<'a>,
    ) -> Result<MathNode<'a>, LatexParseError> {
        let name = self.read_environment_name()?;
        let kind = ENVIRONMENTS
            .get(name)
            .copied()
            .unwrap_or(UNKNOWN_ENVIRONMENT);

        if let EnvKind::Matrix {
            column_spec: true, ..
        } = kind
        {
            self.skip_column_specification()?;
        }

        let rows = self.parse_environment_rows(name, begin)?;

        Ok(match kind {
            EnvKind::Matrix { fence, .. } => MathNode::Matrix {
                rows,
                fence_type: fence,
                properties: None,
            },
            EnvKind::EqArray => MathNode::EqArray {
                rows: rows.into_iter().map(join_cells).collect(),
                properties: None,
            },
        })
    }

    /// Read the `{name}` argument of `\begin` or `\end`.
    pub(crate) fn read_environment_name(&mut self) -> Result<&'a str, LatexParseError> {
        self.skip_spaces();
        let position = self.offset();
        if !matches!(
            self.peek().map(|token| token.kind),
            Some(TokenKind::GroupOpen)
        ) {
            return Err(LatexParseError::MissingEnvironmentName { position });
        }
        Ok(self.read_raw_group()?.trim())
    }

    /// Consume and discard the `[pos]{lcr}` preamble of an `array`.
    ///
    /// The AST models column alignment through matrix properties the parser
    /// does not populate, so the specification is accepted and dropped.
    fn skip_column_specification(&mut self) -> Result<(), LatexParseError> {
        self.parse_optional_argument()?;
        self.skip_spaces();
        if matches!(
            self.peek().map(|token| token.kind),
            Some(TokenKind::GroupOpen)
        ) {
            self.read_raw_group()?;
        }
        Ok(())
    }

    /// Parse environment rows up to and including the closing `\end{...}`.
    fn parse_environment_rows(
        &mut self,
        name: &'a str,
        begin: Token<'a>,
    ) -> Result<Vec<Vec<Vec<MathNode<'a>>>>, LatexParseError> {
        self.enter()?;
        let mut rows: Vec<Vec<Vec<MathNode<'a>>>> = Vec::new();
        let mut row: Vec<Vec<MathNode<'a>>> = Vec::new();

        loop {
            let cell = self.parse_sequence(Limit::Cell)?;
            row.push(cell);

            let Some(token) = self.peek() else {
                return Err(LatexParseError::UnclosedEnvironment {
                    name: name.to_string(),
                    position: begin.start,
                });
            };

            match token.kind {
                TokenKind::Align => self.advance(),
                TokenKind::Command(CMD_LINE_BREAK) => {
                    self.advance();
                    // `\\[6pt]` adds row spacing the AST does not model.
                    self.parse_optional_argument()?;
                    rows.push(std::mem::take(&mut row));
                },
                TokenKind::Command(CMD_END) => {
                    self.advance();
                    let found = self.read_environment_name()?;
                    if found != name {
                        return Err(LatexParseError::MismatchedEnvironment {
                            expected: name.to_string(),
                            found: found.to_string(),
                            position: token.start,
                        });
                    }
                    rows.push(std::mem::take(&mut row));
                    break;
                },
                // `parse_sequence` only stops on the three tokens above, so
                // this arm is unreachable in practice.
                _ => {
                    return Err(LatexParseError::UnclosedEnvironment {
                        name: name.to_string(),
                        position: begin.start,
                    });
                },
            }
        }

        self.leave();
        drop_trailing_empty_row(&mut rows);
        Ok(rows)
    }

    /// Parse `\substack{a \\ b}` into a single-column equation array.
    pub(crate) fn parse_substack(&mut self) -> Result<MathNode<'a>, LatexParseError> {
        self.skip_spaces();
        let position = self.offset();
        if !matches!(
            self.peek().map(|token| token.kind),
            Some(TokenKind::GroupOpen)
        ) {
            return Err(LatexParseError::MissingArgument {
                command: "substack".to_string(),
                position,
            });
        }

        let nodes = self.parse_group_contents()?;
        Ok(MathNode::EqArray {
            rows: split_on_line_breaks(nodes),
            properties: None,
        })
    }
}

/// Concatenate the cells of a row into a single node list.
fn join_cells<'a>(row: Vec<Vec<MathNode<'a>>>) -> Vec<MathNode<'a>> {
    if row.len() == 1 {
        // The overwhelmingly common case: no `&` in the row, so no copying.
        return row.into_iter().next().unwrap_or_default();
    }
    row.into_iter().flatten().collect()
}

/// Split a node list on [`MathNode::LineBreak`] separators.
fn split_on_line_breaks<'a>(nodes: Vec<MathNode<'a>>) -> Vec<Vec<MathNode<'a>>> {
    let mut rows: Vec<Vec<MathNode<'a>>> = Vec::new();
    let mut current: Vec<MathNode<'a>> = Vec::new();

    for node in nodes {
        match node {
            MathNode::LineBreak => rows.push(std::mem::take(&mut current)),
            other => current.push(other),
        }
    }
    if !current.is_empty() || rows.is_empty() {
        rows.push(current);
    }
    rows
}

/// Drop the empty row a trailing `\\` leaves behind.
fn drop_trailing_empty_row(rows: &mut Vec<Vec<Vec<MathNode<'_>>>>) {
    let is_empty = rows
        .last()
        .is_some_and(|row| row.iter().all(|cell| cell.is_empty()));
    if is_empty && rows.len() > 1 {
        rows.pop();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn known_environments_map_to_the_matching_fence() {
        assert_eq!(
            ENVIRONMENTS.get("pmatrix"),
            Some(&EnvKind::Matrix {
                fence: MatrixFence::Paren,
                column_spec: false
            })
        );
        assert_eq!(ENVIRONMENTS.get("align*"), Some(&EnvKind::EqArray));
        assert_eq!(ENVIRONMENTS.get("nowhere"), None);
    }

    #[test]
    fn environment_table_inverts_the_rendering_table() {
        use crate::latex::matrix::matrix_fence_to_env;

        for fence in [
            MatrixFence::None,
            MatrixFence::Paren,
            MatrixFence::Bracket,
            MatrixFence::Brace,
            MatrixFence::Pipe,
            MatrixFence::DoublePipe,
        ] {
            let env = matrix_fence_to_env(fence);
            assert_eq!(
                ENVIRONMENTS.get(env),
                Some(&EnvKind::Matrix {
                    fence,
                    column_spec: false
                }),
                "`{env}` does not parse back to its fence"
            );
        }
    }

    #[test]
    fn splits_a_node_list_on_line_breaks() {
        let nodes = vec![
            MathNode::Text("a".into()),
            MathNode::LineBreak,
            MathNode::Text("b".into()),
        ];
        let rows = split_on_line_breaks(nodes);
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0], vec![MathNode::Text("a".into())]);
        assert_eq!(rows[1], vec![MathNode::Text("b".into())]);
    }

    #[test]
    fn keeps_a_single_empty_row_for_an_empty_body() {
        let rows = split_on_line_breaks(Vec::new());
        assert_eq!(rows.len(), 1);
        assert!(rows[0].is_empty());
    }
}
