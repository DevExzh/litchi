//! Layered, prefix-free slide facade for legacy PowerPoint.
//!
//! The facade keeps the stable [`Slide`] object and its lazy caches together,
//! while delegating value objects, semantic queries, binary decoding, and
//! binary invariants to focused child modules.  Public callers continue to
//! use `slide::types::{Slide, ParsedComment, ParsedSlideTiming}` without
//! knowing the storage layout.

mod codec;
mod model;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

use crate::animation::{ShapeAnimation, SlideAnimationExtension};
use crate::records::Record;
use crate::shapes::ShapeEnum;
use crate::slide::notes::{NoteDescriptor, SpeakerNotes};
use crate::slide_extension::SlideExtension;
use crate::slide_round_trip::SlideRoundTripMetadata12;
use crate::slide_sync::SlideSyncInfo;
use once_cell::unsync::OnceCell;

pub use model::{ParsedComment, ParsedSlideTiming};

/// A slide in a PowerPoint presentation with lazy-loaded shapes and metadata.
///
/// The object owns parsed record state and uses one-time cells for expensive
/// projections.  Shape conversion and text extraction therefore remain
/// demand-driven with no repeated allocation after the first access.
pub struct Slide<'doc> {
    /// Slide persist ID.
    persist_id: u32,
    /// Stable SlideId from the live SlidePersistAtom.
    slide_id: u32,
    slide_list_text: String,
    outline_text_interactions: Vec<crate::TextBodyInteractions>,
    outline_text_refs: Vec<crate::OutlineTextRef>,
    /// Slide number (1-based for display).
    slide_number: usize,
    /// Slide record.
    record: Record,
    /// Reference to document data for lazy speaker-notes parsing.
    #[allow(dead_code)]
    doc_data: &'doc [u8],
    /// Lazily-loaded shapes stored as owned values.
    shapes: OnceCell<Vec<ShapeEnum<'static>>>,
    /// Cached text content.
    text_cache: OnceCell<String>,
    /// Lazily parsed, inert shape animation metadata.
    animations: OnceCell<Vec<ShapeAnimation>>,
    /// Lazily parsed PowerPoint 2002 slide animation extension.
    animation_extension: OnceCell<Option<SlideAnimationExtension>>,
    /// Lazily parsed PowerPoint 12 slide/master round-trip metadata.
    powerpoint12_extension: OnceCell<SlideExtension>,
    /// Lazily parsed, inert slide-library synchronization metadata.
    sync_info: OnceCell<Option<SlideSyncInfo>>,
    /// Lazily parsed direct PowerPoint 12 slide round-trip metadata.
    round_trip_metadata: OnceCell<SlideRoundTripMetadata12>,
    notes_descriptor: Result<Option<NoteDescriptor>, String>,
    speaker_notes: OnceCell<Option<SpeakerNotes>>,
}
