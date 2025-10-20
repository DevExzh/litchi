//! Slide information extracted from PPT Slide records.
//!
//! Based on POI's Slide and SlideAtom record parsing.

/// Information extracted from a Slide record.
#[derive(Debug, Clone, Default)]
pub struct SlideInfo {
    /// Layout ID (reference to master slide)
    pub layout_id: u32,
    /// Master slide ID
    pub master_id: u32,
    /// Notes slide ID (0 if no notes)
    pub notes_id: u32,
    /// Whether the slide has drawing data (PPDrawing record)
    pub has_drawing: bool,
    /// Whether the slide has notes
    pub has_notes: bool,
}

