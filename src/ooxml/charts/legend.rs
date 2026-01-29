//! Chart legend models.
//!
//! This module contains structures for representing chart legends
//! and their positioning.

use crate::ooxml::charts::models::Layout;
use crate::ooxml::charts::types::LegendPosition;

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
}

impl Legend {
    /// Create a new legend with default settings.
    #[inline]
    pub fn new(position: LegendPosition) -> Self {
        Self {
            position,
            overlay: false,
            layout: None,
            entries: Vec::new(),
        }
    }

    /// Set whether to overlay on chart.
    #[inline]
    pub fn with_overlay(mut self, overlay: bool) -> Self {
        self.overlay = overlay;
        self
    }

    /// Set manual layout.
    #[inline]
    pub fn with_layout(mut self, layout: Layout) -> Self {
        self.layout = Some(layout);
        self
    }

    /// Create a default right-positioned legend.
    #[inline]
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
}

impl LegendEntry {
    /// Create a new legend entry.
    #[inline]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            deleted: false,
        }
    }
}
