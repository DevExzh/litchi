//! Inert RTF latent-style defaults and exceptions.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_LATENT_STYLE_INDEX: u32 = 65_535;
pub(crate) const MAX_LATENT_STYLE_EXCEPTIONS: usize = 65_536;
pub(crate) const MAX_LATENT_STYLE_NAME_BYTES: usize = 65_536;
pub(crate) const MAX_LATENT_STYLE_TEXT_BYTES: usize = 16 * 1_048_576;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LatentStyleException<'a> {
    pub name: Cow<'a, str>,
    pub locked: Option<bool>,
    pub semi_hidden: Option<bool>,
    pub unhide_when_used: Option<bool>,
    pub quick_format: Option<bool>,
    pub priority: Option<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LatentStyles<'a> {
    pub max_style_index: u32,
    pub locked_default: Option<bool>,
    pub semi_hidden_default: Option<bool>,
    pub unhide_when_used_default: Option<bool>,
    pub quick_format_default: Option<bool>,
    pub priority_default: Option<u8>,
    pub exceptions: Vec<LatentStyleException<'a>>,
}

impl LatentStyleException<'_> {
    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.name.trim().is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF latent-style exception name cannot be empty".to_string(),
            ));
        }
        if self.name.len() > MAX_LATENT_STYLE_NAME_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF latent-style exception name exceeds the safety limit".to_string(),
            ));
        }
        if self.name.contains(['\0', '\r', '\n', ';']) {
            return Err(RtfError::MalformedDocument(
                "RTF latent-style exception name contains a forbidden delimiter/control"
                    .to_string(),
            ));
        }
        if self.priority.is_some_and(|priority| priority > 99) {
            return Err(RtfError::MalformedDocument(
                "RTF latent-style priority must be in 0..=99".to_string(),
            ));
        }
        Ok(())
    }
}

impl LatentStyles<'_> {
    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.max_style_index > MAX_LATENT_STYLE_INDEX {
            return Err(RtfError::MalformedDocument(
                "RTF latent-style maximum index exceeds 65535".to_string(),
            ));
        }
        if self.priority_default.is_some_and(|priority| priority > 99) {
            return Err(RtfError::MalformedDocument(
                "RTF latent-style default priority must be in 0..=99".to_string(),
            ));
        }
        if self.exceptions.len() > MAX_LATENT_STYLE_EXCEPTIONS {
            return Err(RtfError::MalformedDocument(
                "RTF latent-style exception count exceeds the safety limit".to_string(),
            ));
        }
        let mut total = 0usize;
        for (index, exception) in self.exceptions.iter().enumerate() {
            exception.validate()?;
            if self.exceptions[..index]
                .iter()
                .any(|existing| existing.name == exception.name)
            {
                return Err(RtfError::MalformedDocument(
                    "RTF latent-style exception names must be unique".to_string(),
                ));
            }
            total = total.checked_add(exception.name.len()).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF latent-style aggregate text size overflow".to_string(),
                )
            })?;
            if total > MAX_LATENT_STYLE_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF latent-style aggregate text exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> LatentStyles<'static> {
        LatentStyles {
            max_style_index: self.max_style_index,
            locked_default: self.locked_default,
            semi_hidden_default: self.semi_hidden_default,
            unhide_when_used_default: self.unhide_when_used_default,
            quick_format_default: self.quick_format_default,
            priority_default: self.priority_default,
            exceptions: self
                .exceptions
                .into_iter()
                .map(|exception| LatentStyleException {
                    name: Cow::Owned(exception.name.into_owned()),
                    locked: exception.locked,
                    semi_hidden: exception.semi_hidden,
                    unhide_when_used: exception.unhide_when_used,
                    quick_format: exception.quick_format,
                    priority: exception.priority,
                })
                .collect(),
        }
    }
}
