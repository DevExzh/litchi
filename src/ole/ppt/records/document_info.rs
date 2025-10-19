/// Document information extracted from PPT Document records.
///
/// Based on POI's Document and DocumentAtom record parsing.

/// Information extracted from a Document record.
#[derive(Debug, Clone, Default)]
pub struct DocumentInfo {
    /// Slide width in EMUs (English Metric Units)
    pub slide_width: u32,
    /// Slide height in EMUs
    pub slide_height: u32,
    /// Number of slides in the presentation
    pub slide_count: usize,
    /// Number of notes slides
    pub notes_count: usize,
    /// Number of master slides
    pub master_count: usize,
    /// Whether the document has an Environment record
    pub has_environment: bool,
    /// Whether the document has a PPDrawingGroup record
    pub has_drawing_group: bool,
}

