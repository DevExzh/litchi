//! Chart legend models.
//!
//! This module contains structures for representing chart legends
//! and their positioning.

use crate::chart::data::Layout;
use crate::chart::model::{ExtensionList, ShapeProperties, TextProperties};
use crate::chart::types::LegendPosition;

/// Chart legend configuration.
#[derive(Debug, Clone)]
pub struct Legend {
    /// Legend position
    pub position: LegendPosition,
    /// Overlay on chart area
    pub overlay: bool,
    /// Manual layout
    pub layout: Option<Layout>,
    /// Individual legend entries
    pub entries: Vec<LegendEntry>,
    /// `DrawingML` shape properties for the legend container
    pub shape_properties: Option<ShapeProperties>,
    /// `DrawingML` text properties for the legend container
    pub text_properties: Option<TextProperties>,
    /// Legend extension list
    pub extension_list: Option<ExtensionList>,
}

impl Legend {
    /// Create a new legend with default settings.
    #[inline]
    #[must_use]
    pub fn new(position: LegendPosition) -> Self {
        Self {
            position,
            overlay: false,
            layout: None,
            entries: Vec::new(),
            shape_properties: None,
            text_properties: None,
            extension_list: None,
        }
    }

    /// Set whether to overlay on chart.
    #[inline]
    #[must_use]
    pub fn with_overlay(mut self, overlay: bool) -> Self {
        self.overlay = overlay;
        self
    }

    /// Set manual layout.
    #[inline]
    #[must_use]
    pub fn with_layout(mut self, layout: Layout) -> Self {
        self.layout = Some(layout);
        self
    }

    /// Create a default right-positioned legend.
    #[inline]
    #[must_use]
    pub fn default_right() -> Self {
        Self::new(LegendPosition::Right)
    }
}

impl Default for Legend {
    #[inline]
    fn default() -> Self {
        Self::default_right()
    }
}

/// Individual legend entry.
#[derive(Debug, Clone)]
pub struct LegendEntry {
    /// Entry index
    pub index: u32,
    /// Whether entry is deleted
    pub deleted: bool,
    /// `DrawingML` text properties used instead of the delete choice
    pub text_properties: Option<TextProperties>,
    /// Legend-entry extension list
    pub extension_list: Option<ExtensionList>,
}

impl LegendEntry {
    /// Create a new legend entry.
    #[inline]
    #[must_use]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            deleted: false,
            text_properties: None,
            extension_list: None,
        }
    }
}
