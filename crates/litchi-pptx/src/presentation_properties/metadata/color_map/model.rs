//! Package-independent `PresentationML` color-map values.

use crate::shape::theme::{Color as ThemeColor, Palette, Slot as ThemeSlot};

/// A role that can be mapped by a `PresentationML` color map.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Slot {
    Background1,
    Text1,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl Slot {
    /// Every color slot in its `PresentationML` source order.
    pub const ALL: [Self; 12] = [
        Self::Background1,
        Self::Text1,
        Self::Background2,
        Self::Text2,
        Self::Accent1,
        Self::Accent2,
        Self::Accent3,
        Self::Accent4,
        Self::Accent5,
        Self::Accent6,
        Self::Hyperlink,
        Self::FollowedHyperlink,
    ];

    /// Return the unqualified `PresentationML` attribute name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Background1 => "bg1",
            Self::Text1 => "tx1",
            Self::Background2 => "bg2",
            Self::Text2 => "tx2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }
}

/// A color role defined by a `DrawingML` theme.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Role {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl Role {
    /// Return the `DrawingML` theme color name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }
}

/// A complete `PresentationML` color map.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Map {
    pub(crate) background1: Role,
    pub(crate) text1: Role,
    pub(crate) background2: Role,
    pub(crate) text2: Role,
    pub(crate) accent1: Role,
    pub(crate) accent2: Role,
    pub(crate) accent3: Role,
    pub(crate) accent4: Role,
    pub(crate) accent5: Role,
    pub(crate) accent6: Role,
    pub(crate) hyperlink: Role,
    pub(crate) followed_hyperlink: Role,
}

impl Map {
    /// Return the theme color role mapped from a `PresentationML` color slot.
    #[must_use]
    pub const fn color(&self, slot: Slot) -> Role {
        match slot {
            Slot::Background1 => self.background1,
            Slot::Text1 => self.text1,
            Slot::Background2 => self.background2,
            Slot::Text2 => self.text2,
            Slot::Accent1 => self.accent1,
            Slot::Accent2 => self.accent2,
            Slot::Accent3 => self.accent3,
            Slot::Accent4 => self.accent4,
            Slot::Accent5 => self.accent5,
            Slot::Accent6 => self.accent6,
            Slot::Hyperlink => self.hyperlink,
            Slot::FollowedHyperlink => self.followed_hyperlink,
        }
    }

    /// Set one mapped color role in this detached typed value.
    pub const fn set_color(&mut self, slot: Slot, role: Role) {
        match slot {
            Slot::Background1 => self.background1 = role,
            Slot::Text1 => self.text1 = role,
            Slot::Background2 => self.background2 = role,
            Slot::Text2 => self.text2 = role,
            Slot::Accent1 => self.accent1 = role,
            Slot::Accent2 => self.accent2 = role,
            Slot::Accent3 => self.accent3 = role,
            Slot::Accent4 => self.accent4 = role,
            Slot::Accent5 => self.accent5 = role,
            Slot::Accent6 => self.accent6 = role,
            Slot::Hyperlink => self.hyperlink = role,
            Slot::FollowedHyperlink => self.followed_hyperlink = role,
        }
    }

    /// Return a copy with one mapped color role changed.
    #[must_use]
    pub const fn with_color(mut self, slot: Slot, role: Role) -> Self {
        self.set_color(slot, role);
        self
    }

    /// Resolve a mapped presentation slot against a typed `DrawingML` palette.
    #[must_use]
    pub fn resolve<'a>(&self, palette: &'a Palette, slot: Slot) -> Option<&'a ThemeColor> {
        palette.color(match self.color(slot) {
            Role::Dark1 => ThemeSlot::Dark1,
            Role::Light1 => ThemeSlot::Light1,
            Role::Dark2 => ThemeSlot::Dark2,
            Role::Light2 => ThemeSlot::Light2,
            Role::Accent1 => ThemeSlot::Accent1,
            Role::Accent2 => ThemeSlot::Accent2,
            Role::Accent3 => ThemeSlot::Accent3,
            Role::Accent4 => ThemeSlot::Accent4,
            Role::Accent5 => ThemeSlot::Accent5,
            Role::Accent6 => ThemeSlot::Accent6,
            Role::Hyperlink => ThemeSlot::Hyperlink,
            Role::FollowedHyperlink => ThemeSlot::FollowedHyperlink,
        })
    }
}

/// The color-map selection declared by a slide or slide layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Override {
    Master,
    Override(Map),
}

/// The typed value selected by a source color-map snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Value {
    /// A complete map declared by a slide master.
    Master(Map),
    /// The optional mapping selected by a slide or layout override.
    Override(Option<Override>),
}
