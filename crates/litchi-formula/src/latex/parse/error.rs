// Error definitions for LaTeX parsing
//
// Every variant carries the byte offset in the source string at which the
// problem was detected so that callers can point at the offending construct.

/// Errors reported while parsing LaTeX math source into the formula AST.
///
/// The parser is deliberately forgiving: unknown control sequences and stray
/// alignment markers degrade into plain AST nodes instead of failing. Only
/// structurally broken input (unbalanced groups, unterminated environments,
/// runaway nesting) produces one of these errors.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum LatexParseError {
    /// The input ended while a construct still expected more tokens.
    UnexpectedEndOfInput {
        /// Byte offset at which the input ran out.
        position: usize,
    },

    /// A `{` was opened but never closed.
    UnmatchedGroupOpen {
        /// Byte offset of the offending `{`.
        position: usize,
    },

    /// A `}` appeared without a matching `{`.
    UnmatchedGroupClose {
        /// Byte offset of the offending `}`.
        position: usize,
    },

    /// A `\left` delimiter was never closed by a `\right`.
    UnmatchedLeft {
        /// Byte offset of the offending `\left`.
        position: usize,
    },

    /// A `\right` appeared without a matching `\left`.
    UnmatchedRight {
        /// Byte offset of the offending `\right`.
        position: usize,
    },

    /// `\left` or `\right` was not followed by a recognised delimiter.
    MissingDelimiter {
        /// Byte offset just after the `\left` or `\right`.
        position: usize,
    },

    /// A command that requires an argument did not receive one.
    MissingArgument {
        /// Name of the command (without the leading backslash).
        command: String,
        /// Byte offset of the command.
        position: usize,
    },

    /// `\begin` or `\end` was not followed by a `{name}` group.
    MissingEnvironmentName {
        /// Byte offset of the `\begin` or `\end`.
        position: usize,
    },

    /// A `\begin{...}` was never closed by a matching `\end{...}`.
    UnclosedEnvironment {
        /// Name of the environment that was left open.
        name: String,
        /// Byte offset of the `\begin`.
        position: usize,
    },

    /// `\end{...}` closed an environment other than the innermost open one.
    MismatchedEnvironment {
        /// Name taken from the open `\begin{...}`.
        expected: String,
        /// Name taken from the offending `\end{...}`.
        found: String,
        /// Byte offset of the `\end`.
        position: usize,
    },

    /// `\end{...}` appeared without a matching `\begin{...}`.
    UnexpectedEnd {
        /// Byte offset of the offending `\end`.
        position: usize,
    },

    /// Two subscripts or two superscripts were attached to the same base.
    DuplicateScript {
        /// Byte offset of the second `_` or `^`.
        position: usize,
    },

    /// A backslash appeared at the very end of the input.
    IncompleteCommand {
        /// Byte offset of the dangling backslash.
        position: usize,
    },

    /// The expression nests more deeply than the configured limit allows.
    NestingTooDeep {
        /// Byte offset at which the limit was exceeded.
        position: usize,
        /// The configured maximum nesting depth.
        limit: usize,
    },
}

impl std::fmt::Display for LatexParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            LatexParseError::UnexpectedEndOfInput { position } => {
                write!(f, "unexpected end of input at byte {position}")
            },
            LatexParseError::UnmatchedGroupOpen { position } => {
                write!(f, "unmatched '{{' at byte {position}")
            },
            LatexParseError::UnmatchedGroupClose { position } => {
                write!(f, "unmatched '}}' at byte {position}")
            },
            LatexParseError::UnmatchedLeft { position } => {
                write!(f, "\\left at byte {position} has no matching \\right")
            },
            LatexParseError::UnmatchedRight { position } => {
                write!(f, "\\right at byte {position} has no matching \\left")
            },
            LatexParseError::MissingDelimiter { position } => {
                write!(f, "expected a delimiter at byte {position}")
            },
            LatexParseError::MissingArgument { command, position } => {
                write!(f, "\\{command} at byte {position} is missing an argument")
            },
            LatexParseError::MissingEnvironmentName { position } => {
                write!(f, "expected an environment name at byte {position}")
            },
            LatexParseError::UnclosedEnvironment { name, position } => {
                write!(
                    f,
                    "environment '{name}' opened at byte {position} was never closed"
                )
            },
            LatexParseError::MismatchedEnvironment {
                expected,
                found,
                position,
            } => {
                write!(
                    f,
                    "\\end{{{found}}} at byte {position} does not close \\begin{{{expected}}}"
                )
            },
            LatexParseError::UnexpectedEnd { position } => {
                write!(f, "\\end at byte {position} has no matching \\begin")
            },
            LatexParseError::DuplicateScript { position } => {
                write!(f, "duplicate script marker at byte {position}")
            },
            LatexParseError::IncompleteCommand { position } => {
                write!(f, "dangling backslash at byte {position}")
            },
            LatexParseError::NestingTooDeep { position, limit } => {
                write!(f, "nesting deeper than {limit} levels at byte {position}")
            },
        }
    }
}

impl std::error::Error for LatexParseError {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_variant_renders_a_human_readable_message() {
        let errors = [
            LatexParseError::UnexpectedEndOfInput { position: 1 },
            LatexParseError::UnmatchedGroupOpen { position: 2 },
            LatexParseError::UnmatchedGroupClose { position: 3 },
            LatexParseError::UnmatchedLeft { position: 4 },
            LatexParseError::UnmatchedRight { position: 5 },
            LatexParseError::MissingDelimiter { position: 6 },
            LatexParseError::MissingArgument {
                command: "frac".to_string(),
                position: 7,
            },
            LatexParseError::MissingEnvironmentName { position: 8 },
            LatexParseError::UnclosedEnvironment {
                name: "matrix".to_string(),
                position: 9,
            },
            LatexParseError::MismatchedEnvironment {
                expected: "matrix".to_string(),
                found: "cases".to_string(),
                position: 10,
            },
            LatexParseError::UnexpectedEnd { position: 11 },
            LatexParseError::DuplicateScript { position: 12 },
            LatexParseError::IncompleteCommand { position: 13 },
            LatexParseError::NestingTooDeep {
                position: 14,
                limit: 64,
            },
        ];

        for error in errors {
            assert!(!error.to_string().is_empty());
        }
    }
}
