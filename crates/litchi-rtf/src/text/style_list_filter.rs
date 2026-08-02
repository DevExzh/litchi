use crate::{RtfError, RtfResult};

/// Passive suggested sorting for a host application's document-style list.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum DocumentStyleSortMethod {
    Name = 0,
    #[default]
    HostDefault = 1,
    Font = 2,
    BasedOnStyle = 3,
    StyleType = 4,
}

impl DocumentStyleSortMethod {
    pub fn from_rtf_value(value: i32) -> RtfResult<Self> {
        match value {
            0 => Ok(Self::Name),
            1 => Ok(Self::HostDefault),
            2 => Ok(Self::Font),
            3 => Ok(Self::BasedOnStyle),
            4 => Ok(Self::StyleType),
            _ => Err(RtfError::MalformedDocument(
                "RTF style-sort method must be in the range 0 through 4".to_string(),
            )),
        }
    }

    pub fn rtf_value(self) -> i32 {
        self as i32
    }
}

/// Passive suggested filters for an application's document-style list.
#[derive(Debug, Clone, Copy)]
pub struct DocumentStyleListFilter(u16, bool);

impl PartialEq for DocumentStyleListFilter {
    fn eq(&self, other: &Self) -> bool {
        self.0 == other.0
    }
}

impl Eq for DocumentStyleListFilter {}

impl DocumentStyleListFilter {
    pub const ALL_STYLES: Self = Self(0x0001, false);
    pub const CUSTOM_STYLES: Self = Self(0x0002, false);
    pub const LATENT_STYLES: Self = Self(0x0004, false);
    pub const STYLES_IN_USE: Self = Self(0x0008, false);
    pub const HEADING_STYLES: Self = Self(0x0020, false);
    pub const NUMBERING_STYLES: Self = Self(0x0040, false);
    pub const TABLE_STYLES: Self = Self(0x0080, false);
    pub const DIRECT_RUN_FORMATTING: Self = Self(0x0100, false);
    pub const DIRECT_PARAGRAPH_FORMATTING: Self = Self(0x0200, false);
    pub const DIRECT_NUMBERING_FORMATTING: Self = Self(0x0400, false);
    pub const DIRECT_TABLE_FORMATTING: Self = Self(0x0800, false);
    pub const CLEAR_FORMATTING_STYLE: Self = Self(0x1000, false);
    pub const TOP_LEVEL_HEADING_STYLES: Self = Self(0x2000, false);
    pub const VISIBLE_STYLES: Self = Self(0x4000, false);
    pub const ALTERNATE_STYLE_NAMES: Self = Self(0x8000, false);

    const RESERVED_BITS: u16 = 0x0010;

    pub fn from_bits(bits: u16) -> RtfResult<Self> {
        if bits & Self::RESERVED_BITS != 0 {
            return Err(RtfError::MalformedDocument(
                "RTF style-list filter uses the reserved 0010 bit".to_string(),
            ));
        }
        Ok(Self(bits, false))
    }

    pub(crate) fn from_parsed_bits(bits: u16) -> Self {
        Self(bits, bits & Self::RESERVED_BITS != 0)
    }

    pub fn bits(self) -> u16 {
        self.0
    }
    pub fn is_empty(self) -> bool {
        self.0 == 0
    }
    pub fn contains(self, filter: Self) -> bool {
        self.0 & filter.0 == filter.0
    }
    pub fn union(self, filter: Self) -> Self {
        Self(self.0 | filter.0, false)
    }
    pub fn validate(self) -> RtfResult<()> {
        Self::from_bits(self.0).map(|_| ())
    }

    pub(crate) fn validate_for_write(self) -> RtfResult<()> {
        if self.1 { Ok(()) } else { self.validate() }
    }
}
