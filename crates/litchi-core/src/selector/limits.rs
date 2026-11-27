//! Explicit bounds on path and selector parsing.
//!
//! Paths and selectors arrive from agents, batch files, and command lines, so
//! the parser treats them as adversarial input (ADR 0006): every dimension that
//! could grow work or memory is capped, and exceeding a cap is a typed
//! [`ParseError`](super::ParseError) carrying the [`Limit`](super::Limit) that
//! tripped, never a panic or unbounded recursion.
//!
//! Every recursive production (parenthesised groups, `not(...)`, `:has(...)`,
//! `:not(...)`) increments one shared nesting counter checked against
//! [`MAX_NESTING`], so the parser's stack depth is bounded independently of the
//! input length.

/// Maximum accepted input length in bytes.
///
/// Real-world OfficeCLI paths are well under 200 bytes; the allowance leaves
/// room for long `:contains("...")` arguments and quoted header names.
pub const MAX_INPUT_BYTES: usize = 16 * 1024;

/// Maximum number of steps in one path, including the scope path of a query.
///
/// The deepest OfficeCLI path forms (table cells inside headers, chart series
/// axes) use fewer than ten steps.
pub const MAX_STEPS: usize = 64;

/// Maximum number of bracket predicates plus pseudo-classes attached to one
/// path step or one selector compound.
pub const MAX_PREDICATES: usize = 16;

/// Maximum nesting depth of parenthesised groups, `not(...)`, `:has(...)`, and
/// `:not(...)`, counted together.
pub const MAX_NESTING: usize = 32;

/// Maximum number of leaf conditions in one bracket's boolean expression.
pub const MAX_TERMS: usize = 64;

/// Maximum number of comma-separated alternatives in one selector list.
pub const MAX_ALTERNATIVES: usize = 64;

/// Maximum number of compounds chained by combinators in one complex
/// selector.
pub const MAX_COMPOUNDS: usize = 32;
