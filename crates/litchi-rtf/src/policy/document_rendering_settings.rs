/// Requested document rendering direction retained from the RTF header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentRenderingOrientation {
    /// Horizontal rendering (`horzdoc`).
    Horizontal,
    /// Vertical rendering (`vertdoc`).
    Vertical,
}

/// Requested document justification behavior retained from the RTF header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentJustificationMode {
    /// Compressing justification (`jcompress`).
    Compress,
    /// Expanding justification (`jexpand`).
    Expand,
}

/// Passive document rendering flags from the RTF header.
///
/// This crate retains these values for round-tripping and does not apply them
/// to layout or rendering.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentRenderingSettings {
    /// Explicit rendering orientation, or `None` when omitted.
    pub orientation: Option<DocumentRenderingOrientation>,
    /// Explicit justification mode, or `None` when omitted.
    pub justification_mode: Option<DocumentJustificationMode>,
    /// Whether `lnongrid` was present.
    pub line_based_on_grid: bool,
}

impl DocumentRenderingSettings {
    /// Return whether all rendering controls were omitted.
    pub fn is_empty(&self) -> bool {
        self.orientation.is_none() && self.justification_mode.is_none() && !self.line_based_on_grid
    }

    /// Return the explicit justification mode or the RTF default, compression.
    pub fn effective_justification_mode(&self) -> DocumentJustificationMode {
        self.justification_mode
            .unwrap_or(DocumentJustificationMode::Compress)
    }
}
