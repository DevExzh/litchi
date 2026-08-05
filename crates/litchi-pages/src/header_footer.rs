//! Archive-free header and footer role vocabulary.
//!
//! Native object identifiers, text storage, and package traversal remain in
//! `litchi-iwa`. This module contains only the semantic roles needed by a
//! Pages document's header and footer views.

/// Which page-template variant owns a header or footer.
#[repr(u8)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Template {
    /// The first page of a section.
    First,
    /// Even-numbered pages in a section.
    Even,
    /// Odd-numbered pages in a section.
    Odd,
}

/// Whether a text region is a header or a footer.
#[repr(u8)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A region above the section body.
    Header,
    /// A region below the section body.
    Footer,
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Kind, Template};

    #[test]
    fn roles_are_compact_closed_values() {
        assert_eq!(size_of::<Template>(), 1);
        assert_eq!(size_of::<Kind>(), 1);
        assert_ne!(Template::First, Template::Even);
        assert_ne!(Kind::Header, Kind::Footer);
    }
}
