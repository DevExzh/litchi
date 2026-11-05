//! [`Placement`]: where `move_to` puts an element and where `add` inserts it.

use litchi_core::Position;

use crate::error::{Code, Error, Result};
use crate::path::{IntoPath, PathArg};

/// Where an element goes among siblings or into another parent.
///
/// Positions are zero-based. For a move, the position is interpreted after
/// the element has been detached, matching OfficeCLI. No placement constructor
/// guesses a position from a path. The public positional value is the checked
/// [`Position`] type; callers with a `usize` may use the explicit
/// [`IntoPlacement`] convenience implementation.
///
/// ```
/// use litchi_ops::path::Position;
/// use litchi_ops::{IntoPlacement, Placement};
///
/// let position = Position::new(2).into_placement()?;
/// assert!(matches!(position, Placement::Position(value) if value == Position::new(2)));
/// let before = Placement::before("/body/p[2]")?;
/// assert!(matches!(before, Placement::Before(_)));
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Placement<'q> {
    /// First among the siblings.
    Start,
    /// Last among the siblings; also the default for `add`.
    End,
    /// At this zero-based sibling position.
    Position(Position),
    /// Immediately before the element at `path`.
    Before(PathArg<'q>),
    /// Immediately after the element at `path`.
    After(PathArg<'q>),
    /// Into the parent at `path`, optionally at a zero-based position.
    Into(PathArg<'q>, Option<Position>),
}

impl<'q> Placement<'q> {
    /// Places before the element at `path`.
    ///
    /// # Errors
    ///
    /// Returns the typed path parsing error.
    pub fn before(path: impl IntoPath<'q>) -> Result<Self> {
        Ok(Self::Before(path.into_path()?))
    }

    /// Places after the element at `path`.
    ///
    /// # Errors
    ///
    /// Returns the typed path parsing error.
    pub fn after(path: impl IntoPath<'q>) -> Result<Self> {
        Ok(Self::After(path.into_path()?))
    }

    /// Places at the end of the parent's children.
    ///
    /// # Errors
    ///
    /// Returns the typed path parsing error.
    pub fn inside(parent: impl IntoPath<'q>) -> Result<Self> {
        Ok(Self::Into(parent.into_path()?, None))
    }

    /// Places inside `parent` at a zero-based position.
    ///
    /// # Errors
    ///
    /// Returns the typed path parsing error.
    pub fn inside_at(parent: impl IntoPath<'q>, position: Position) -> Result<Self> {
        Ok(Self::Into(parent.into_path()?, Some(position)))
    }
}

/// A value accepted by `File::move_to` as its destination.
pub trait IntoPlacement<'q> {
    /// Converts this input into a placement.
    ///
    /// ```
    /// use litchi_ops::path::Position;
    /// use litchi_ops::{IntoPlacement, Placement};
    ///
    /// let placement: Placement<'_> = Position::new(0).into_placement()?;
    /// assert!(matches!(placement, Placement::Position(_)));
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    fn into_placement(self) -> Result<Placement<'q>>;
}

impl<'q> IntoPlacement<'q> for Placement<'q> {
    fn into_placement(self) -> Result<Placement<'q>> {
        validate(self)
    }
}

/// Validate paths stored in a manually constructed placement as well as paths
/// produced by the constructors.  The resolver performs the format-specific
/// identity check later; this boundary only enforces the common path grammar
/// and keeps content filters on the query API.
fn validate<'q>(placement: Placement<'q>) -> Result<Placement<'q>> {
    let path = match &placement {
        Placement::Before(path) | Placement::After(path) | Placement::Into(path, _) => Some(path),
        Placement::Start | Placement::End | Placement::Position(_) => None,
    };
    if let Some(path) = path {
        crate::path::arg::validate(path)?;
        if !path.path().is_structural() {
            return Err(Error::new(
                Code::InvalidPath,
                "placement paths must select by position or identity",
            )
            .with_suggestion("use get/find for content filters, or use an addressing path"));
        }
    }
    Ok(placement)
}

impl<'q> IntoPlacement<'q> for Position {
    fn into_placement(self) -> Result<Placement<'q>> {
        Ok(Placement::Position(self))
    }
}

impl<'q> IntoPlacement<'q> for usize {
    fn into_placement(self) -> Result<Placement<'q>> {
        Ok(Placement::Position(Position::new(self)))
    }
}

#[cfg(test)]
mod tests {
    use litchi_core::Position;

    use crate::Code;
    use crate::path::PathArg;

    use super::{IntoPlacement, Placement};

    #[test]
    fn constructors_preserve_explicit_position_and_paths() {
        assert!(matches!(
            3_usize.into_placement().expect("index"),
            Placement::Position(position) if position == Position::new(3)
        ));
        assert!(matches!(
            Position::new(4).into_placement().expect("position"),
            Placement::Position(position) if position == Position::new(4)
        ));
        assert!(matches!(
            Placement::before("/body/p[2]").expect("before"),
            Placement::Before(_)
        ));
        assert!(matches!(
            Placement::inside_at("/body", Position::new(0)).expect("inside"),
            Placement::Into(_, Some(position)) if position == Position::new(0)
        ));
        assert!(Placement::after("/body/p[").is_err());
    }

    #[test]
    fn placement_debug_does_not_echo_path_text() {
        let placement = Placement::before("/PRIVATE/p[1]").expect("path");
        assert!(!format!("{placement:?}").contains("PRIVATE"));
    }

    #[test]
    fn manually_constructed_content_placement_is_rejected_at_input_boundary() {
        let path = PathArg::parse("/body/p[@text!=PRIVATE]").expect("path");
        let error = Placement::Before(path)
            .into_placement()
            .expect_err("content placement");
        assert_eq!(error.code(), Code::InvalidPath);
    }
}
