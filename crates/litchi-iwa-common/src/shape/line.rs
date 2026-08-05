//! Archive-independent line endpoint values shared by iWork owners.

/// A native endpoint decoration supported by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Endpoint {
    /// No endpoint decoration.
    #[default]
    None,
    /// A compact filled triangular arrowhead.
    SimpleArrow,
    /// A filled circular endpoint.
    FilledCircle,
    /// A filled diamond endpoint.
    FilledDiamond,
    /// An outlined arrowhead with a short center stem.
    OpenArrow,
    /// A broad filled arrowhead with an inset base.
    FilledArrow,
    /// A filled square endpoint.
    FilledSquare,
    /// An outlined square endpoint.
    OpenSquare,
    /// An outlined circular endpoint.
    OpenCircle,
    /// A filled arrowhead pointing toward the line segment.
    InvertedArrow,
    /// A perpendicular bar endpoint.
    Line,
}

/// Decorations at the directed start and end of a straight line.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Endpoints {
    /// Decoration at the directed start point.
    pub start: Endpoint,
    /// Decoration at the directed end point.
    pub end: Endpoint,
}

impl Endpoints {
    /// Construct independently typed start and end decorations.
    #[must_use]
    pub const fn new(start: Endpoint, end: Endpoint) -> Self {
        Self { start, end }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Endpoint, Endpoints};

    #[test]
    fn endpoint_values_are_compact_and_copyable() {
        assert_eq!(size_of::<Endpoint>(), 1);
        assert_eq!(size_of::<Endpoints>(), 2);
        assert_eq!(
            Endpoints::new(Endpoint::None, Endpoint::SimpleArrow).start,
            Endpoint::None
        );
    }
}
