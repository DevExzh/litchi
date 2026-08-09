//! Passive RTF document view and zoom metadata.

use crate::{RtfError, RtfResult};

/// Safety bound for a retained `viewscale` percentage.
pub const MAX_DOCUMENT_VIEW_SCALE_PERCENT: u16 = 10_000;

/// RTF `viewkind` document view mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentViewKind {
    None,
    PageLayout,
    Outline,
    MasterDocument,
    Normal,
    OnlineLayout,
}

impl DocumentViewKind {
    pub(crate) fn from_rtf(value: i32) -> RtfResult<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::PageLayout),
            2 => Ok(Self::Outline),
            3 => Ok(Self::MasterDocument),
            4 => Ok(Self::Normal),
            5 => Ok(Self::OnlineLayout),
            _ => Err(RtfError::MalformedDocument(
                "RTF viewkind must be in 0..=5".to_string(),
            )),
        }
    }

    pub(crate) fn rtf_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::PageLayout => 1,
            Self::Outline => 2,
            Self::MasterDocument => 3,
            Self::Normal => 4,
            Self::OnlineLayout => 5,
        }
    }
}

/// RTF `viewzk` zoom mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentZoomKind {
    None,
    FullPage,
    BestFit,
    TextWidth,
}

impl DocumentZoomKind {
    pub(crate) fn from_rtf(value: i32) -> RtfResult<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::FullPage),
            2 => Ok(Self::BestFit),
            3 => Ok(Self::TextWidth),
            _ => Err(RtfError::MalformedDocument(
                "RTF viewzk must be in 0..=3".to_string(),
            )),
        }
    }

    pub(crate) fn rtf_value(self) -> i32 {
        match self {
            Self::None => 0,
            Self::FullPage => 1,
            Self::BestFit => 2,
            Self::TextWidth => 3,
        }
    }
}

/// Explicit passive document view controls.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentView {
    pub kind: Option<DocumentViewKind>,
    pub scale_percent: Option<u16>,
    pub zoom_kind: Option<DocumentZoomKind>,
    /// `viewbksp`: whether background shapes show in Page Layout view.
    pub background_shapes: Option<bool>,
    /// `viewnobound`: hide white space between pages.
    pub hide_page_boundaries: bool,
}

impl DocumentView {
    /// Validate percentage resource bounds.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self
            .scale_percent
            .is_some_and(|value| value == 0 || value > MAX_DOCUMENT_VIEW_SCALE_PERCENT)
        {
            return Err(RtfError::MalformedDocument(format!(
                "RTF viewscale must be in 1..={MAX_DOCUMENT_VIEW_SCALE_PERCENT}"
            )));
        }
        Ok(())
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.kind.is_none()
            && self.scale_percent.is_none()
            && self.zoom_kind.is_none()
            && self.background_shapes.is_none()
            && !self.hide_page_boundaries
    }
}
