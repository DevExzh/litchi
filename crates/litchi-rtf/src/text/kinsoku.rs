//! Inert custom kinsoku (East Asian line-breaking) character sets.
//!
//! The RTF specification defines the `\*\fchars` and `\*\lchars` document
//! destinations, which list custom kinsoku characters: *following* kinsoku
//! may not appear at the end of a line, and *leading* kinsoku may not start
//! a line. The companion document control `\ksulangN` declares the language
//! the customized sets belong to.
//!
//! The sets are parsed and stored as passive metadata only; no line-breaking
//! rule is ever evaluated or applied.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_KINSOKU_BYTES: usize = 65_536;

/// Inert custom kinsoku character sets and their language.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentKinsoku<'a> {
    /// Following kinsoku characters (`\*\fchars`), which may not end a line.
    pub following: Option<Cow<'a, str>>,
    /// Leading kinsoku characters (`\*\lchars`), which may not start a line.
    pub leading: Option<Cow<'a, str>>,
    /// Language identifier of the customized sets (`\ksulangN`).
    pub language: Option<u32>,
}

impl DocumentKinsoku<'_> {
    pub(crate) fn validate_characters(kind: &str, characters: &str) -> RtfResult<()> {
        if characters.is_empty() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {kind} kinsoku character set cannot be empty"
            )));
        }
        if characters.len() > MAX_KINSOKU_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {kind} kinsoku character set exceeds the safety limit"
            )));
        }
        if characters.contains(['\0', '\r', '\n']) {
            return Err(RtfError::MalformedDocument(format!(
                "RTF {kind} kinsoku character set contains a forbidden control character"
            )));
        }
        Ok(())
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if let Some(following) = &self.following {
            Self::validate_characters("following", following)?;
        }
        if let Some(leading) = &self.leading {
            Self::validate_characters("leading", leading)?;
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> DocumentKinsoku<'static> {
        DocumentKinsoku {
            following: self.following.map(|value| Cow::Owned(value.into_owned())),
            leading: self.leading.map(|value| Cow::Owned(value.into_owned())),
            language: self.language,
        }
    }
}
