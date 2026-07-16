//! Inert legacy RTF drawing text boxes.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_LEGACY_TEXT_BOXES: usize = 16_384;
pub(crate) const MAX_LEGACY_TEXT_BOX_BYTES: usize = 1_048_576;
pub(crate) const MAX_LEGACY_TEXT_BOX_TOTAL_BYTES: usize = 16 * 1_048_576;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyHorizontalAnchor {
    Page,
    Margin,
    Column,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyVerticalAnchor {
    Page,
    Margin,
    Paragraph,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum LegacyTextDirection {
    #[default]
    LeftToRightTopToBottom,
    LeftToRightTopToBottomVertical,
    TopToBottomRightToLeft,
    TopToBottomRightToLeftVertical,
    BottomToTopLeftToRight,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyTextBox<'a> {
    pub text: Cow<'a, str>,
    pub position: usize,
    pub horizontal_anchor: Option<LegacyHorizontalAnchor>,
    pub vertical_anchor: Option<LegacyVerticalAnchor>,
    pub x: Option<i32>,
    pub y: Option<i32>,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub margin: Option<i32>,
    pub z_order: Option<i32>,
    pub direction: LegacyTextDirection,
}

impl LegacyTextBox<'_> {
    pub fn validate(&self) -> RtfResult<()> {
        if self.text.is_empty() || self.text.len() > MAX_LEGACY_TEXT_BOX_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box text is empty or exceeds the safety limit".to_string(),
            ));
        }
        if self.text.contains('\0') {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box text contains a NUL character".to_string(),
            ));
        }
        if self.width.is_some_and(|value| value <= 0)
            || self.height.is_some_and(|value| value <= 0)
            || self.margin.is_some_and(|value| value < 0)
        {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box size or margin is outside its valid range".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> LegacyTextBox<'static> {
        LegacyTextBox {
            text: Cow::Owned(self.text.into_owned()),
            position: self.position,
            horizontal_anchor: self.horizontal_anchor,
            vertical_anchor: self.vertical_anchor,
            x: self.x,
            y: self.y,
            width: self.width,
            height: self.height,
            margin: self.margin,
            z_order: self.z_order,
            direction: self.direction,
        }
    }
}
