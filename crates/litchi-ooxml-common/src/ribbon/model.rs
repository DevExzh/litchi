//! Typed, contextual Ribbon vocabulary independent of any OOXML host format.

use crate::{Error, Result};
use litchi_opc::PackURI;

pub(super) const V2007_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/01/customui";
pub(super) const V2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/2009/07/customui";
pub(super) const UI2_NAMESPACE: &str = "http://schemas.microsoft.com/office/2007/10/customui";
pub(super) const LEGACY_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility";
pub(super) const MODERN_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";

/// A package-level Ribbon relationship family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Family {
    /// Office 2007 Custom UI relationship family.
    Legacy,
    /// Office 2010 and `CustomUI2` relationship family.
    Modern,
}

impl Family {
    /// Package relationship type for this family.
    #[must_use]
    pub const fn relationship(self) -> &'static str {
        match self {
            Self::Legacy => LEGACY_RELATIONSHIP,
            Self::Modern => MODERN_RELATIONSHIP,
        }
    }

    pub(super) fn from_relationship(value: &str) -> Option<Self> {
        if value == LEGACY_RELATIONSHIP {
            Some(Self::Legacy)
        } else if value == MODERN_RELATIONSHIP {
            Some(Self::Modern)
        } else {
            None
        }
    }
}

/// Custom UI vocabulary selected by the root namespace.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Version {
    /// Office 2007 Custom UI vocabulary.
    V2007,
    /// Office 2010 Ribbon and Backstage vocabulary.
    V2010,
    /// `CustomUI2` vocabulary documented by the Office extensions.
    Ui2,
}

impl Version {
    /// Root namespace required by this vocabulary.
    #[must_use]
    pub const fn namespace(self) -> &'static str {
        match self {
            Self::V2007 => V2007_NAMESPACE,
            Self::V2010 => V2010_NAMESPACE,
            Self::Ui2 => UI2_NAMESPACE,
        }
    }

    /// Package relationship type required by this vocabulary.
    #[must_use]
    pub const fn relationship(self) -> &'static str {
        self.family().relationship()
    }

    /// Package relationship family containing this vocabulary.
    #[must_use]
    pub const fn family(self) -> Family {
        match self {
            Self::V2007 => Family::Legacy,
            Self::V2010 | Self::Ui2 => Family::Modern,
        }
    }

    pub(super) const fn default_part(self) -> &'static str {
        match self {
            Self::V2007 => "/customUI/customUI.xml",
            Self::V2010 => "/customUI/customUI14.xml",
            Self::Ui2 => "/customUI/customUI2.xml",
        }
    }
}

/// A Ribbon customization borrowing its package-owned XML allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Ui<'a> {
    part: &'a PackURI,
    id: &'a str,
    version: Version,
    xml: &'a [u8],
}

impl<'a> Ui<'a> {
    pub(super) const fn new(
        part: &'a PackURI,
        id: &'a str,
        version: Version,
        xml: &'a [u8],
    ) -> Self {
        Self {
            part,
            id,
            version,
            xml,
        }
    }

    /// Canonical package part containing the Custom UI XML.
    #[must_use]
    #[inline]
    pub const fn part(self) -> &'a PackURI {
        self.part
    }

    /// Low-level package relationship ID.
    ///
    /// Prefer [`Set::effective`] and [`crate::ribbon::remove`] for semantic operations.
    #[must_use]
    #[inline]
    pub const fn id(self) -> &'a str {
        self.id
    }

    /// Custom UI vocabulary identified by the relationship and root namespace.
    #[must_use]
    #[inline]
    pub const fn version(self) -> Version {
        self.version
    }

    /// Original package-owned XML bytes.
    #[must_use]
    #[inline]
    pub const fn xml(self) -> &'a [u8] {
        self.xml
    }
}

/// Fixed Ribbon family slots for one package.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Set<'a> {
    legacy: Option<Ui<'a>>,
    modern: Option<Ui<'a>>,
}

impl<'a> Set<'a> {
    /// Office 2007 relationship-family slot.
    #[must_use]
    #[inline]
    pub const fn legacy(self) -> Option<Ui<'a>> {
        self.legacy
    }

    /// Office 2010 and `CustomUI2` relationship-family slot.
    #[must_use]
    #[inline]
    pub const fn modern(self) -> Option<Ui<'a>> {
        self.modern
    }

    /// Customization used by modern-first consumers.
    #[must_use]
    #[inline]
    pub const fn effective(self) -> Option<Ui<'a>> {
        match self.modern {
            Some(value) => Some(value),
            None => self.legacy,
        }
    }

    /// Present slots in stable legacy-then-modern order.
    #[inline]
    pub fn iter(self) -> impl Iterator<Item = Ui<'a>> {
        [self.legacy, self.modern].into_iter().flatten()
    }

    pub(super) const fn get(self, family: Family) -> Option<Ui<'a>> {
        match family {
            Family::Legacy => self.legacy,
            Family::Modern => self.modern,
        }
    }

    pub(super) fn insert(&mut self, value: Ui<'a>) {
        let slot = match value.version.family() {
            Family::Legacy => &mut self.legacy,
            Family::Modern => &mut self.modern,
        };
        *slot = Some(value);
    }

    pub(super) fn require_empty(self, family: Family) -> Result<()> {
        if self.get(family).is_some() {
            return Err(Error::Relationship(format!(
                "a package may contain at most one {family:?} Ribbon relationship"
            )));
        }
        Ok(())
    }
}

/// Resource ceilings applied while validating Ribbon package data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes in each Custom UI XML part.
    pub xml_bytes: usize,
    /// Maximum Custom UI XML element nesting depth.
    pub depth: usize,
    /// Maximum XML events and attributes in each Custom UI part.
    pub nodes: usize,
    /// Maximum aggregate image relationships across both Ribbon parts.
    pub images: usize,
}

impl Limits {
    /// Conservative defaults for untrusted Office packages.
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            xml_bytes: 4 * 1024 * 1024,
            depth: 128,
            nodes: 262_144,
            images: 4_096,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::standard()
    }
}
