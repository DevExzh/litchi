//! Concise collection selectors that do not overload indexing.

use std::borrow::Cow;

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
    pub const fn new(value: usize) -> Self {
        Self(value)
    }

    /// Returns the zero-based value.
    #[inline]
    pub const fn get(self) -> usize {
        self.0
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

impl<'a, Id> From<String> for Selector<'a, Id> {
    fn from(value: String) -> Self {
        Self::Name(Cow::Owned(value))
    }
}

impl<'a, Id> From<&'a String> for Selector<'a, Id> {
    fn from(value: &'a String) -> Self {
        Self::Name(Cow::Borrowed(value.as_str()))
    }
}

impl<'a, Id> From<usize> for Selector<'a, Id> {
    fn from(value: usize) -> Self {
        Self::Position(value.into())
    }
}

impl<'a, Id> From<Position> for Selector<'a, Id> {
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
        fn scoped<'a>(name: String, _scope: &'a ()) -> Selector<'a, SheetId> {
            name.into()
        }

        let scope = ();
        assert!(matches!(
            scoped("Summary".to_owned(), &scope),
            Selector::Name(name) if name == "Summary"
        ));
    }
}
