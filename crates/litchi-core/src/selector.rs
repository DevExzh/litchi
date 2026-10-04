//! Concise collection selectors and the unified path/selector grammar.
//!
//! The module has two layers:
//!
//! - The typed vocabulary used by format crates: [`Position`] (a checked
//!   zero-based collection position) and [`Selector`] (name, position, or
//!   opaque stable identity).
//! - The OfficeCLI-aligned string grammar used by the unified operations
//!   layer: [`Path`] for element addressing (`/body/p[3]/r[1]`,
//!   `/Sheet1/A1:C5`, `/slide[2]/notes`) and [`Query`] for CSS-like selection
//!   (`paragraph[style=Heading1] > run[font!=Arial]`,
//!   `row[Salary>5000 and Region=EMEA]`). Both are typed ASTs; strings are an
//!   edge layer parsed once through [`Path::parse`] / [`Query::parse`] (which
//!   borrow from the input) or [`std::str::FromStr`] (which owns). `Display`
//!   renders a canonical string that parses back to an equal value.
//!
//! The grammar is syntax only. Whether a step names an existing element is a
//! resolution-time question answered by the operations layer, which reports
//! [`ErrorCode::PathNotFound`]; this module never produces that code.
//!
//! # Example
//!
//! ```
//! use litchi_core::selector::{Path, Position, Predicate, Query};
//!
//! let path = Path::parse("/body/p[3]/r[1]")?;
//! assert_eq!(path.steps().len(), 3);
//! assert_eq!(path.steps()[1].name(), "p");
//! // String indices are 1-based; the typed value is the zero-based Position.
//! assert_eq!(path.steps()[1].predicates(), &[Predicate::Index(Position::new(2))]);
//! assert_eq!(path.to_string(), "/body/p[3]/r[1]");
//!
//! let query: Query<'_> = Query::parse("row[Salary>5000 and Region=EMEA]")?;
//! assert_eq!(query.to_string(), "row[Salary>5000 and Region=EMEA]");
//! # Ok::<(), litchi_core::selector::ParseError>(())
//! ```

use std::borrow::Cow;

mod axis;
mod cursor;
mod error;
pub mod limits;
mod path;
mod predicate;
mod query;
mod store;
mod text;

#[cfg(test)]
mod grammar_tests;

pub use axis::{Axis, CellRange, CellRef, column_number};
pub use error::{ErrorCode, ErrorKind, Expected, Found, Grammar, Limit, ParseError};
pub use path::{Path, Step};
pub use predicate::{CompareOp, Condition, Expr, Key, Predicate, Value};
pub use query::{Combinator, Complex, Compound, Element, Pseudo, Query, Scope};

/// A checked collection position.
///
/// Positions are always zero-based. Whether a position exists is decided by
/// the selected collection and represented by `Option`, not by this scalar.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Position(usize);

impl Position {
    /// Creates a zero-based position.
    #[inline]
    #[must_use]
    pub const fn new(value: usize) -> Self {
        Self(value)
    }

    /// Returns the zero-based value.
    #[inline]
    #[must_use]
    pub const fn get(self) -> usize {
        self.0
    }

    /// Converts a 1-based string-grammar index (`p[3]`) to a position.
    ///
    /// Returns `None` for `0`, which has no 1-based meaning.
    ///
    /// ```
    /// use litchi_core::Position;
    ///
    /// assert_eq!(Position::from_one_based(3), Some(Position::new(2)));
    /// assert_eq!(Position::from_one_based(0), None);
    /// ```
    #[inline]
    #[must_use]
    pub const fn from_one_based(index: usize) -> Option<Self> {
        match index.checked_sub(1) {
            Some(value) => Some(Self(value)),
            None => None,
        }
    }

    /// Returns the 1-based string-grammar index, or `None` when the position
    /// is `usize::MAX` and has no representable 1-based form.
    ///
    /// ```
    /// use litchi_core::Position;
    ///
    /// assert_eq!(Position::new(2).one_based(), Some(3));
    /// ```
    #[inline]
    #[must_use]
    pub const fn one_based(self) -> Option<usize> {
        self.0.checked_add(1)
    }
}

impl From<usize> for Position {
    fn from(value: usize) -> Self {
        Self::new(value)
    }
}

/// A convenient semantic selector with an advanced stable-identity variant.
///
/// Names and zero-based positions are intended as normal API entry points.
/// Format crates choose their own opaque `Id` type for durable selections and
/// patch plumbing.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Selector<'a, Id> {
    /// Select by a developer-facing name.
    Name(Cow<'a, str>),
    /// Select by a transient zero-based position.
    Position(Position),
    /// Select by an opaque stable identity.
    Id(Id),
}

impl<'a, Id> Selector<'a, Id> {
    /// Creates the advanced stable-identity form explicitly.
    #[inline]
    pub const fn id(id: Id) -> Self {
        Self::Id(id)
    }

    /// Maps only the stable identity while retaining name/position selectors.
    pub fn map_id<Other>(self, map: impl FnOnce(Id) -> Other) -> Selector<'a, Other> {
        match self {
            Self::Name(name) => Selector::Name(name),
            Self::Position(position) => Selector::Position(position),
            Self::Id(id) => Selector::Id(map(id)),
        }
    }
}

impl<'a, Id> From<&'a str> for Selector<'a, Id> {
    fn from(value: &'a str) -> Self {
        Self::Name(Cow::Borrowed(value))
    }
}

impl<Id> From<String> for Selector<'_, Id> {
    fn from(value: String) -> Self {
        Self::Name(Cow::Owned(value))
    }
}

impl<'a, Id> From<&'a String> for Selector<'a, Id> {
    fn from(value: &'a String) -> Self {
        Self::Name(Cow::Borrowed(value.as_str()))
    }
}

impl<Id> From<usize> for Selector<'_, Id> {
    fn from(value: usize) -> Self {
        Self::Position(value.into())
    }
}

impl<Id> From<Position> for Selector<'_, Id> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    struct SheetId(u64);

    #[test]
    fn ergonomic_inputs_do_not_require_raw_ids() {
        let by_name: Selector<'_, SheetId> = "Summary".into();
        let by_position: Selector<'_, SheetId> = 1usize.into();

        assert!(matches!(by_name, Selector::Name(name) if name == "Summary"));
        assert_eq!(by_position, Selector::Position(Position::new(1)));
    }

    #[test]
    fn identity_mapping_does_not_touch_semantic_selectors() {
        let selector = Selector::id(SheetId(7));
        assert_eq!(selector.map_id(|id| id.0), Selector::Id(7));
    }

    #[test]
    fn owned_names_adapt_to_the_callers_scope() {
        fn scoped(name: String, _scope: &()) -> Selector<'_, SheetId> {
            name.into()
        }

        let scope = ();
        assert!(matches!(
            scoped("Summary".to_owned(), &scope),
            Selector::Name(name) if name == "Summary"
        ));
    }
}
