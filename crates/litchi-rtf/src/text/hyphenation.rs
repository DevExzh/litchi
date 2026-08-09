//! Passive document-level RTF hyphenation settings.

use crate::{RtfError, RtfResult};

/// Maximum retained value for `hyphconsec`.
pub const MAX_HYPHENATION_CONSECUTIVE_LINES: u32 = i32::MAX as u32;
/// Maximum retained `hyphhotz` measurement in twips.
pub const MAX_HYPHENATION_HOT_ZONE_TWIPS: u32 = i32::MAX as u32;

/// Explicit document-level hyphenation controls.
///
/// `None` means that the control was absent and the RTF reader default applies.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentHyphenation {
    /// `hyphauto`: automatic hyphenation, whose RTF default is off.
    pub automatic: Option<bool>,
    /// `hyphcaps`: hyphenation of capitalized words, whose RTF default is on.
    pub capitalized_words: Option<bool>,
    /// `hyphconsec`: maximum consecutive hyphen-ending lines; zero means unlimited.
    pub consecutive_line_limit: Option<u32>,
    /// `hyphhotz`: right-margin hyphenation hot zone in twips.
    pub hot_zone_twips: Option<u32>,
}

impl DocumentHyphenation {
    /// Validate all explicit numeric settings against parser safety bounds.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self
            .consecutive_line_limit
            .is_some_and(|value| value > MAX_HYPHENATION_CONSECUTIVE_LINES)
        {
            return Err(RtfError::MalformedDocument(format!(
                "RTF hyphenation consecutive-line limit exceeds {MAX_HYPHENATION_CONSECUTIVE_LINES}"
            )));
        }
        if self
            .hot_zone_twips
            .is_some_and(|value| value > MAX_HYPHENATION_HOT_ZONE_TWIPS)
        {
            return Err(RtfError::MalformedDocument(format!(
                "RTF hyphenation hot zone exceeds {MAX_HYPHENATION_HOT_ZONE_TWIPS} twips"
            )));
        }
        Ok(())
    }

    /// Return whether any hyphenation control is explicitly present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.automatic.is_none()
            && self.capitalized_words.is_none()
            && self.consecutive_line_limit.is_none()
            && self.hot_zone_twips.is_none()
    }
}
