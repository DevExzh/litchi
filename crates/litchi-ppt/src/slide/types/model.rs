//! Stable slide state and typed value objects exposed by the slide facade.

use super::Slide;
use crate::slide::factory::SlideData;
use once_cell::unsync::OnceCell;

impl<'doc> Slide<'doc> {
    /// Create a slide from parsed slide data.
    #[must_use]
    pub fn from_slide_data(data: SlideData<'doc>, slide_number: usize) -> Self {
        let doc_data_ref = data.doc_data();
        Self {
            persist_id: data.persist_id,
            slide_id: data.slide_id,
            slide_list_text: data.slide_list_text,
            outline_text_interactions: data.outline_text_interactions,
            outline_text_refs: data.outline_text_refs,
            slide_number,
            doc_data: doc_data_ref,
            record: data.record,
            shapes: OnceCell::new(),
            text_cache: OnceCell::new(),
            animations: OnceCell::new(),
            animation_extension: OnceCell::new(),
            powerpoint12_extension: OnceCell::new(),
            sync_info: OnceCell::new(),
            round_trip_metadata: OnceCell::new(),
            notes_descriptor: data.note_descriptor,
            speaker_notes: OnceCell::new(),
            record_limits: data.record_limits,
        }
    }
}

/// A parsed comment from a PPT slide.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParsedComment {
    /// Nonnegative comment index.
    pub index: i32,
    /// Author name.
    pub author: String,
    /// Comment text.
    pub text: String,
    /// Author initials.
    pub initials: String,
    /// Year.
    pub year: u16,
    /// Month (1-12).
    pub month: u16,
    /// Day of week (`0` is Sunday).
    pub day_of_week: u16,
    /// Day (1-31).
    pub day: u16,
    /// Hour (0-23).
    pub hour: u16,
    /// Minute (0-59).
    pub minute: u16,
    /// Second (0-59).
    pub second: u16,
    /// Millisecond (0-999).
    pub millisecond: u16,
    /// X position in master units (576/inch).
    pub x: i32,
    /// Y position in master units.
    pub y: i32,
}

/// Parsed per-slide timing information.
#[derive(Debug, Clone)]
pub struct ParsedSlideTiming {
    /// Auto-advance time in milliseconds (0 = no auto-advance).
    pub advance_time_ms: u32,
    /// Whether the slide advances on mouse click.
    pub advance_on_click: bool,
    /// Whether auto-advance is enabled.
    pub auto_advance: bool,
    /// Whether the slide is hidden.
    pub hidden: bool,
}
