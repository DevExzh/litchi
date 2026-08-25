//! Archive-free header and footer roles, selectors, and text values.
//!
//! Native object identifiers, text storage, and package traversal remain in
//! the concrete Pages package adapter. This module contains only semantic
//! values that can safely cross the Pages facade.

use std::fmt;

use litchi_core::Position;

use crate::selector::SectionSelector;

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

/// Selects one logical header/footer slot without exposing native IDs.
///
/// A section can expose the same physical storage through several template
/// roles or slots. The package adapter resolves this selector against the
/// rooted section/template graph and verifies every affected alias after a
/// write. `slot` is a typed zero-based semantic position, never a native
/// text-storage or object identifier.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct HeaderFooterSelector<'a> {
    section: SectionSelector<'a>,
    template: Template,
    kind: Kind,
    slot: Position,
}

impl fmt::Debug for HeaderFooterSelector<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HeaderFooterSelector")
            // Section names are authored content; retain selector shape but
            // do not place their contents into diagnostics.
            .field("section_by_name", &self.section.as_name().is_some())
            .field("section_position", &self.section.as_position())
            .field("template", &self.template)
            .field("kind", &self.kind)
            .field("slot", &self.slot)
            .finish()
    }
}

impl<'a> HeaderFooterSelector<'a> {
    /// Construct a selector from a section selector and a template role.
    #[must_use]
    pub const fn new(
        section: SectionSelector<'a>,
        template: Template,
        kind: Kind,
        slot: Position,
    ) -> Self {
        Self {
            section,
            template,
            kind,
            slot,
        }
    }

    /// Construct a selector using a zero-based section index and slot index.
    #[must_use]
    pub const fn index(
        section_index: usize,
        template: Template,
        kind: Kind,
        slot_index: usize,
    ) -> HeaderFooterSelector<'static> {
        HeaderFooterSelector {
            section: SectionSelector::index(section_index),
            template,
            kind,
            slot: Position::new(slot_index),
        }
    }

    /// Construct a selector from typed section and slot positions.
    #[must_use]
    pub const fn position(
        section: Position,
        template: Template,
        kind: Kind,
        slot: Position,
    ) -> HeaderFooterSelector<'static> {
        HeaderFooterSelector {
            section: SectionSelector::position(section),
            template,
            kind,
            slot,
        }
    }

    /// Construct a selector from a section selector and a template role.
    #[must_use]
    pub const fn from_parts(
        section: SectionSelector<'a>,
        template: Template,
        kind: Kind,
        slot: Position,
    ) -> Self {
        Self::new(section, template, kind, slot)
    }

    /// Borrow the section selector.
    #[must_use]
    pub const fn section(self) -> SectionSelector<'a> {
        self.section
    }

    /// Borrow the section selector.
    #[must_use]
    pub const fn section_selector(self) -> SectionSelector<'a> {
        self.section
    }

    /// Return the requested template role.
    #[must_use]
    pub const fn template(self) -> Template {
        self.template
    }

    /// Return whether this selects a header or footer.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Return the zero-based typed slot position.
    #[must_use]
    pub const fn slot(self) -> Position {
        self.slot
    }

    /// Return the zero-based typed slot position.
    #[must_use]
    pub const fn slot_position(self) -> Position {
        self.slot
    }

    /// Return the zero-based slot index.
    #[must_use]
    pub const fn slot_index(self) -> usize {
        self.slot.get()
    }
}

/// Construct a selector from a section selector, template, role, and slot.
impl<'a> From<(SectionSelector<'a>, Template, Kind, Position)> for HeaderFooterSelector<'a> {
    fn from(
        (section, template, kind, slot): (SectionSelector<'a>, Template, Kind, Position),
    ) -> Self {
        Self::new(section, template, kind, slot)
    }
}

/// One semantic header or footer region.
///
/// The section position and optional visible name identify the owning
/// section, while the template, role, and slot identify the semantic region.
/// Text and the optional name are owned so a package snapshot can return
/// values that outlive lower-level archive views. No native object,
/// text-storage, archive-member, protobuf, or wire state is retained here.
#[derive(Clone, PartialEq, Eq)]
pub struct HeaderFooter {
    section_position: Position,
    section_name: Option<Box<str>>,
    template: Template,
    kind: Kind,
    slot: Position,
    text: Box<str>,
}

impl fmt::Debug for HeaderFooter {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HeaderFooter")
            .field("section_position", &self.section_position)
            // Authored names and text should not leak through debug logs.
            .field("section_name_present", &self.section_name.is_some())
            .field("section_name_bytes", &self.section_name_bytes())
            .field("template", &self.template)
            .field("kind", &self.kind)
            .field("slot", &self.slot)
            .field("text_bytes", &self.text.len())
            .finish()
    }
}

impl HeaderFooter {
    /// Construct a semantic region from borrowed section-name and text input.
    ///
    /// The section name is optional because a native section may omit its
    /// visible name. Text and the name are copied into exact-size boxed
    /// strings; no lower-level package state is retained.
    #[must_use]
    pub fn new(
        section_position: Position,
        section_name: Option<&str>,
        template: Template,
        kind: Kind,
        slot: Position,
        text: impl Into<Box<str>>,
    ) -> Self {
        Self::from_parts(
            section_position,
            section_name.map(Into::into),
            template,
            kind,
            slot,
            text.into(),
        )
    }

    /// Construct a semantic region from owned strings.
    #[must_use]
    pub fn from_parts(
        section_position: Position,
        section_name: Option<Box<str>>,
        template: Template,
        kind: Kind,
        slot: Position,
        text: Box<str>,
    ) -> Self {
        Self {
            section_position,
            section_name,
            template,
            kind,
            slot,
            text,
        }
    }

    /// Return the owning section's zero-based semantic position.
    #[must_use]
    pub const fn section_position(&self) -> Position {
        self.section_position
    }

    /// Return the owning section's optional visible name.
    #[must_use]
    pub fn section_name(&self) -> Option<&str> {
        self.section_name.as_deref()
    }

    /// Return the page-template variant owning this region.
    #[must_use]
    pub const fn template(&self) -> Template {
        self.template
    }

    /// Return whether this region is a header or footer.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return this region's zero-based semantic slot.
    #[must_use]
    pub const fn slot(&self) -> Position {
        self.slot
    }

    /// Return this region's zero-based slot index.
    #[must_use]
    pub const fn slot_index(&self) -> usize {
        self.slot.get()
    }

    /// Return the user-visible text in this region.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return whether this region currently contains no text.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.text.is_empty()
    }

    /// Return the number of UTF-8 bytes in the text.
    #[must_use]
    pub const fn text_bytes(&self) -> usize {
        self.text.len()
    }

    /// Return the number of UTF-8 bytes in the optional section name.
    #[must_use]
    pub fn section_name_bytes(&self) -> usize {
        self.section_name.as_deref().map_or(0, str::len)
    }

    /// Return a stable, position-based selector for this value.
    ///
    /// The selector intentionally uses the section position rather than
    /// borrowing the visible section name, so this returned value has a
    /// `'static` lifetime and remains safe as an immutable transaction key.
    #[must_use]
    pub const fn selector(&self) -> HeaderFooterSelector<'static> {
        HeaderFooterSelector::position(self.section_position, self.template, self.kind, self.slot)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use litchi_core::Position;

    use super::{HeaderFooter, HeaderFooterSelector, Kind, Template};
    use crate::selector::SectionSelector;

    #[test]
    fn roles_are_compact_closed_values() {
        assert_eq!(size_of::<Template>(), 1);
        assert_eq!(size_of::<Kind>(), 1);
        assert_ne!(Template::First, Template::Even);
        assert_ne!(Kind::Header, Kind::Footer);
    }

    #[test]
    fn selector_is_typed_and_borrowed_without_native_identity() {
        let selector = HeaderFooterSelector::new(
            SectionSelector::name("Introduction"),
            Template::First,
            Kind::Header,
            Position::new(1),
        );

        assert_eq!(selector.section(), SectionSelector::name("Introduction"));
        assert_eq!(
            selector.section_selector(),
            SectionSelector::name("Introduction")
        );
        assert_eq!(selector.template(), Template::First);
        assert_eq!(selector.kind(), Kind::Header);
        assert_eq!(selector.slot(), Position::new(1));
        assert_eq!(selector.slot_position(), Position::new(1));
        assert_eq!(selector.slot_index(), 1);

        let indexed = HeaderFooterSelector::index(2, Template::Even, Kind::Footer, 3);
        assert_eq!(indexed.section(), SectionSelector::index(2));
        assert_eq!(indexed.slot_position(), Position::new(3));
        let positioned = HeaderFooterSelector::position(
            Position::new(2),
            Template::Even,
            Kind::Footer,
            Position::new(3),
        );
        assert_eq!(indexed, positioned);
    }

    #[test]
    fn region_accessors_preserve_semantic_shape_and_aliasable_slots() {
        let region = HeaderFooter::new(
            Position::new(2),
            Some("Introduction"),
            Template::Even,
            Kind::Footer,
            Position::new(0),
            "Confidential",
        );

        assert_eq!(region.section_position(), Position::new(2));
        assert_eq!(region.section_name(), Some("Introduction"));
        assert_eq!(region.template(), Template::Even);
        assert_eq!(region.kind(), Kind::Footer);
        assert_eq!(region.slot(), Position::new(0));
        assert_eq!(region.slot_index(), 0);
        assert_eq!(region.text(), "Confidential");
        assert!(!region.is_empty());
        assert_eq!(region.text_bytes(), "Confidential".len());
        assert_eq!(region.section_name_bytes(), "Introduction".len());
        assert_eq!(region.selector().section(), SectionSelector::index(2));
        assert_eq!(region.selector().slot(), Position::new(0));

        let alias = HeaderFooter::new(
            Position::new(2),
            Some("Introduction"),
            Template::Even,
            Kind::Footer,
            Position::new(0),
            "Confidential",
        );
        assert_eq!(region, alias);
    }

    #[test]
    fn debug_output_redacts_authored_names_and_text() {
        let region = HeaderFooter::new(
            Position::new(0),
            Some("private section name"),
            Template::Odd,
            Kind::Header,
            Position::new(3),
            "private header text",
        );
        let debug = format!("{region:?}");

        assert!(debug.contains("HeaderFooter"));
        assert!(debug.contains("section_name_bytes"));
        assert!(debug.contains("text_bytes"));
        assert!(!debug.contains("private section name"));
        assert!(!debug.contains("private header text"));

        let selector = HeaderFooterSelector::new(
            SectionSelector::name("private section name"),
            Template::Odd,
            Kind::Header,
            Position::new(3),
        );
        let selector_debug = format!("{selector:?}");
        assert!(!selector_debug.contains("private section name"));
    }

    #[test]
    fn unnamed_region_uses_position_selector_and_empty_text_is_explicit() {
        let region = HeaderFooter::new(
            Position::new(4),
            None,
            Template::First,
            Kind::Footer,
            Position::new(0),
            "",
        );

        assert_eq!(region.section_name(), None);
        assert_eq!(region.selector().section(), SectionSelector::index(4));
        assert!(region.is_empty());
        assert_eq!(region.text_bytes(), 0);
    }
}
