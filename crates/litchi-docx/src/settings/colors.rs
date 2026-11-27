#![expect(
    clippy::format_push_string,
    reason = "serialization preserves the established byte-emission path"
)]
use crate::Result;

use super::support::invalid;

/// Theme color slot produced by a `w:clrSchemeMapping` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ColorSchemeIndex {
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

impl ColorSchemeIndex {
    /// Parse the schema token.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(value: &str) -> Result<Self> {
        match value {
            "dark1" => Ok(Self::Dark1),
            "light1" => Ok(Self::Light1),
            "dark2" => Ok(Self::Dark2),
            "light2" => Ok(Self::Light2),
            "accent1" => Ok(Self::Accent1),
            "accent2" => Ok(Self::Accent2),
            "accent3" => Ok(Self::Accent3),
            "accent4" => Ok(Self::Accent4),
            "accent5" => Ok(Self::Accent5),
            "accent6" => Ok(Self::Accent6),
            "hyperlink" => Ok(Self::Hyperlink),
            "followedHyperlink" => Ok(Self::FollowedHyperlink),
            _ => Err(invalid(format!(
                "invalid color scheme index value '{value}'"
            ))),
        }
    }

    /// Get the XML value for this theme color slot.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dark1",
            Self::Light1 => "light1",
            Self::Dark2 => "dark2",
            Self::Light2 => "light2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
        }
    }
}

/// A remappable theme color slot on `w:clrSchemeMapping`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ColorSchemeSlot {
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

impl ColorSchemeSlot {
    /// Number of remappable color slots.
    pub const COUNT: usize = 12;

    /// Every slot in schema attribute order.
    pub const ALL: [Self; Self::COUNT] = [
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

    const fn index(self) -> usize {
        self as usize
    }

    /// Get the attribute name carrying this slot.
    #[must_use]
    pub const fn attribute_name(self) -> &'static str {
        match self {
            Self::Background1 => "bg1",
            Self::Text1 => "t1",
            Self::Background2 => "bg2",
            Self::Text2 => "t2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hyperlink",
            Self::FollowedHyperlink => "followedHyperlink",
        }
    }
}

/// Theme color slot remapping from `w:clrSchemeMapping`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ColorSchemeMapping {
    slots: [Option<ColorSchemeIndex>; ColorSchemeSlot::COUNT],
}

impl ColorSchemeMapping {
    /// Create a mapping with every slot left at its default.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            slots: [None; ColorSchemeSlot::COUNT],
        }
    }

    /// Remap a slot to a theme color index.
    pub fn set(&mut self, slot: ColorSchemeSlot, index: ColorSchemeIndex) -> &mut Self {
        self.slots[slot.index()] = Some(index);
        self
    }

    /// Restore a slot to its default mapping.
    pub fn clear(&mut self, slot: ColorSchemeSlot) -> &mut Self {
        self.slots[slot.index()] = None;
        self
    }

    /// Return the theme color index a slot maps to, when remapped.
    #[inline]
    #[must_use]
    pub const fn get(&self, slot: ColorSchemeSlot) -> Option<ColorSchemeIndex> {
        self.slots[slot.index()]
    }

    /// Whether no slot is remapped.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.slots.iter().all(Option::is_none)
    }

    /// Iterate the remapped slots in schema attribute order.
    pub fn iter(&self) -> impl Iterator<Item = (ColorSchemeSlot, ColorSchemeIndex)> + '_ {
        ColorSchemeSlot::ALL
            .into_iter()
            .filter_map(|slot| self.get(slot).map(|index| (slot, index)))
    }

    /// Serialize a standalone `w:clrSchemeMapping` fragment.
    #[must_use]
    pub fn to_xml(&self, prefix: &str) -> String {
        let mut xml = format!("<{prefix}:clrSchemeMapping");
        for (slot, index) in self.iter() {
            xml.push_str(&format!(
                " {prefix}:{}=\"{}\"",
                slot.attribute_name(),
                index.as_str()
            ));
        }
        xml.push_str("/>");
        xml
    }
}
