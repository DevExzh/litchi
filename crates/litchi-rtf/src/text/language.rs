//! RTF document and character language metadata.

use crate::{RtfError, RtfResult};

/// An RTF language identifier from the specification's standard language table.
///
/// Unknown identifiers remain representable so long as they fit the unsigned
/// 16-bit range used by RTF and Windows language identifiers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LanguageId(u16);

impl LanguageId {
    /// The RTF undefined/no-language identifier used with `\noproof`.
    pub const UNDEFINED: Self = Self(1024);

    pub fn new(value: u32) -> RtfResult<Self> {
        u16::try_from(value).map(Self).map_err(|_| {
            RtfError::MalformedDocument("RTF language ID must be in 0..=65535".to_string())
        })
    }

    pub fn from_rtf(value: i32) -> RtfResult<Self> {
        let value = u32::try_from(value).map_err(|_| {
            RtfError::MalformedDocument("RTF language ID must be in 0..=65535".to_string())
        })?;
        Self::new(value)
    }

    pub const fn value(self) -> u16 {
        self.0
    }

    pub const fn rtf_value(self) -> i32 {
        self.0 as i32
    }
}

/// Default languages declared in the RTF header.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentLanguageDefaults {
    pub primary: Option<LanguageId>,
    pub east_asian: Option<LanguageId>,
    pub complex_script: Option<LanguageId>,
}

impl DocumentLanguageDefaults {
    pub fn new() -> Self {
        Self::default()
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        Ok(())
    }
}
