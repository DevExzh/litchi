//! Package-independent values for Word 2010 paragraph extensions.

use crate::error::Result;
use std::fmt::{self, Display, Formatter};
use std::str::FromStr;

use super::validation::{parse_id, validate_ids};

/// Word 2010 WordprocessingML extension namespace.
pub const WORD_2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

/// A checked `ST_LongHexNumber` used by `paraId` and `textId`.
///
/// The wire value is exactly eight hexadecimal digits and is restricted to
/// the non-zero values below `0x80000000` required by `[MS-DOCX]` 2.6.2.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct Id(u32);

impl Id {
    /// Construct an identifier in the Word-defined range.
    #[must_use]
    pub const fn new(value: u32) -> Option<Self> {
        if value != 0 && value < 0x8000_0000 {
            Some(Self(value))
        } else {
            None
        }
    }

    /// Return the numeric value.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }

    /// Parse the exact eight-digit hexadecimal wire form.
    pub fn parse(value: &str) -> Result<Self> {
        parse_id(value, "paragraph extension identifier")
    }
}

impl Display for Id {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        write!(formatter, "{:08x}", self.0)
    }
}

impl FromStr for Id {
    type Err = crate::Error;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::parse(value)
    }
}

/// The optional `paraId`/`textId` pair carried by a paragraph or table row.
///
/// `textId` cannot exist without `paraId`, so the ordinary mutators reject
/// that invalid intermediate state instead of exposing a partially valid
/// struct. The values are copyable and contain no heap storage.
#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq)]
pub struct Ids {
    para_id: Option<Id>,
    text_id: Option<Id>,
}

impl Ids {
    /// Create an empty identifier set.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            para_id: None,
            text_id: None,
        }
    }

    /// Create a set with a paragraph identifier and no text version.
    #[must_use]
    pub const fn with_para_id(para_id: Id) -> Self {
        Self {
            para_id: Some(para_id),
            text_id: None,
        }
    }

    /// Construct a checked pair from optional values.
    pub fn from_parts(para_id: Option<Id>, text_id: Option<Id>) -> Result<Self> {
        let value = Self { para_id, text_id };
        validate_ids(&value)?;
        Ok(value)
    }

    /// Return the paragraph identifier.
    #[must_use]
    pub const fn para_id(self) -> Option<Id> {
        self.para_id
    }

    /// Return the paragraph text-version identifier.
    #[must_use]
    pub const fn text_id(self) -> Option<Id> {
        self.text_id
    }

    /// Set or remove `paraId`.
    ///
    /// Removing `paraId` while `textId` is present is rejected and leaves the
    /// value unchanged, preserving the schema dependency atomically.
    pub fn set_para_id(&mut self, para_id: Option<Id>) -> Result<&mut Self> {
        let candidate = Self {
            para_id,
            text_id: self.text_id,
        };
        validate_ids(&candidate)?;
        *self = candidate;
        Ok(self)
    }

    /// Set or remove `textId`.
    ///
    /// A present `textId` requires a present `paraId`; failed validation does
    /// not modify the current value.
    pub fn set_text_id(&mut self, text_id: Option<Id>) -> Result<&mut Self> {
        let candidate = Self {
            para_id: self.para_id,
            text_id,
        };
        validate_ids(&candidate)?;
        *self = candidate;
        Ok(self)
    }

    /// Validate the complete identifier dependency.
    pub fn validate(self) -> Result<()> {
        validate_ids(&self)
    }
}

/// All modeled Word 2010 extension attributes on a paragraph.
///
/// Table rows use [`Ids`] because `noSpellErr` is paragraph-only. `None` for
/// [`Self::no_spell_err`] means the attribute was absent; `Some(false)` is an
/// explicit `0`/`false` value meaning that no spelling-error result is known.
#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq)]
pub struct Extensions {
    ids: Ids,
    no_spell_err: Option<bool>,
}

impl Extensions {
    /// Create an empty paragraph extension value.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            ids: Ids::new(),
            no_spell_err: None,
        }
    }

    /// Return the paragraph/table identifier pair.
    #[must_use]
    pub const fn ids(self) -> Ids {
        self.ids
    }

    /// Return the optional `noSpellErr` state.
    #[must_use]
    pub const fn no_spell_err(self) -> Option<bool> {
        self.no_spell_err
    }

    /// Set the checked identifier pair.
    pub fn set_ids(&mut self, ids: Ids) -> Result<&mut Self> {
        ids.validate()?;
        self.ids = ids;
        Ok(self)
    }

    /// Set or remove `paraId` while preserving the `textId` dependency.
    pub fn set_para_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.ids.set_para_id(value)?;
        Ok(self)
    }

    /// Set or remove `textId` while preserving the `paraId` dependency.
    pub fn set_text_id(&mut self, value: Option<Id>) -> Result<&mut Self> {
        self.ids.set_text_id(value)?;
        Ok(self)
    }

    /// Set or remove the paragraph-only spelling result.
    pub fn set_no_spell_err(&mut self, value: Option<bool>) -> &mut Self {
        self.no_spell_err = value;
        self
    }

    /// Validate the complete paragraph extension value.
    pub fn validate(self) -> Result<()> {
        self.ids.validate()
    }

    /// Whether this value emits no extension attributes.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.ids.para_id().is_none() && self.ids.text_id().is_none() && self.no_spell_err.is_none()
    }
}
