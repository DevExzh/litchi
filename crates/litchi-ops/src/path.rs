//! Paths: the `litchi_core::selector` grammar, re-exported, and the
//! resolution of a path against a document (record 0010 §4, M1 step R2).
//!
//! A path arrives as text or as a parsed [`Path`] through [`IntoPath`].
//! Resolution walks it one step at a time through a format adapter:
//!
//! 1. Each step name is normalized through the format family's static alias
//!    table, ASCII case-insensitively (`Paragraph`, `p` and `P` are the same
//!    element); a worksheet name at a workbook's root and a style id under
//!    `/styles` are names instead, and `A1`/`A1:C5` under a worksheet are
//!    cells.
//! 2. The step's predicates are classified: `[N]` (an ordinal, or a number
//!    the element carries such as a row number), `[last()]`, stacked
//!    `[@key=value]` identities, or a bare token such as `col[B]`. Content
//!    filters belong to `query` and are refused.
//! 3. The adapter returns that one child, so only the parts on the path are
//!    read.
//!
//! Common short paths and resolver lookups allocate nothing; escaped sheet
//! names and Unicode normalization may allocate. Failures carry one code each:
//! `invalid_path` for syntax, unknown names (with a "did you mean"
//! suggestion) and bad predicates; `unsupported_path` for a concept the
//! format does not have (`/slide[1]` in a Word document);
//! `unsupported_capability` for one Litchi cannot read yet; and
//! `path_not_found`, with what exists, when nothing is there. The resolver
//! itself is internal; future `File::get` and `File::find` adapters will expose it.
//! This record supplies the shared core and test-only fake adapters.
//!
//! ```
//! use litchi_ops::path::{IntoPath, Path, Predicate, Position};
//!
//! let path = "/body/p[3]".into_path()?;
//! assert_eq!(
//!     path.path().steps()[1].predicates(),
//!     &[Predicate::Index(Position::new(2))]
//! );
//!
//! // Excel notation is rewritten to the workbook path it names.
//! assert_eq!("Sheet1!B:B".into_path()?.to_string(), "/Sheet1/col[B]");
//!
//! // A parsed path is reused without parsing again.
//! let reusable = Path::parse("/slide[1]/shape[2]")?;
//! assert_eq!((&reusable).into_path()?.path(), &reusable);
//! # Ok::<(), litchi_ops::Error>(())
//! ```

// The format adapters of M1 step R4 are the production callers of the
// resolution core; until they land, only the test-only fake formats call it.
#![cfg_attr(
    not(test),
    allow(
        dead_code,
        reason = "M1 step R4's format adapters call the resolution core; until then only the \
                  test-only fake formats do"
    )
)]

pub(crate) mod arg;
#[cfg(any(test, feature = "test-support"))]
mod fake;
pub(crate) mod fence;
pub(crate) mod navigate;
pub(crate) mod report;
pub(crate) mod resolve;
#[cfg(test)]
mod tests;
pub(crate) mod vocabulary;

pub use arg::{IntoPath, PathArg};
pub use litchi_core::selector::{
    Axis, CellRange, CellRef, Combinator, CompareOp, Complex, Compound, Condition, Element,
    ErrorCode, ErrorKind, Expected, Expr, Found, Grammar, Key, Limit, ParseError, Path, Position,
    Predicate, Pseudo, Query, Scope, Step, Value, column_number, limits,
};

/// Test-only access to the resolver over the fake formats, for tests that
/// must live outside the crate (the counting-allocator test). Not public
/// API.
#[cfg(feature = "test-support")]
#[doc(hidden)]
pub mod __fake {
    use super::fake::{Fake, Presentation, Word, Workbook};
    use super::{IntoPath, resolve};
    use crate::Result;

    /// Which fake document to resolve against.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub enum Document {
        /// A word-processing document (the Word vocabulary).
        Word,
        /// A presentation.
        Presentation,
        /// A workbook.
        Workbook,
    }

    /// Resolves `path` with `get`; returns the canonical name of the element
    /// reached.
    ///
    /// # Errors
    ///
    /// As `File::get`.
    pub fn get<'p>(document: Document, path: impl IntoPath<'p>) -> Result<&'static str> {
        let path = path.into_path()?;
        let node = match document {
            Document::Word => resolve::get(&Fake::<Word>::new(), &path)?,
            Document::Presentation => resolve::get(&Fake::<Presentation>::new(), &path)?,
            Document::Workbook => resolve::get(&Fake::<Workbook>::new(), &path)?,
        };
        Ok(node.kind())
    }

    /// Resolves `path` with `find`; returns the canonical name of the element
    /// reached, or `None` when nothing is there.
    ///
    /// # Errors
    ///
    /// As `File::find`.
    pub fn find<'p>(document: Document, path: impl IntoPath<'p>) -> Result<Option<&'static str>> {
        let path = path.into_path()?;
        let node = match document {
            Document::Word => resolve::find(&Fake::<Word>::new(), &path)?,
            Document::Presentation => resolve::find(&Fake::<Presentation>::new(), &path)?,
            Document::Workbook => resolve::find(&Fake::<Workbook>::new(), &path)?,
        };
        Ok(node.map(|node| node.kind()))
    }
}

/// Explicit native-identity path resolution input.
pub mod raw {
    /// Permit native identities in this path. Ordinary paths accept positions
    /// and semantic names; raw identity values are never emitted in diagnostics.
    ///
    /// # Errors
    /// Returns the typed path parsing error when input is malformed.
    pub fn path<'p>(input: impl super::IntoPath<'p>) -> crate::Result<super::PathArg<'p>> {
        let mut path = input.into_path()?;
        path.raw = true;
        Ok(path)
    }
}
