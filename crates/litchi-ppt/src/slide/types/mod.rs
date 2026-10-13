//! Layered, prefix-free slide facade for legacy `PowerPoint`.
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
use crate::slide_sync::Synchronization;
use once_cell::unsync::OnceCell;
use std::sync::Arc;

pub use model::{ParsedComment, ParsedSlideTiming};
pub(crate) use validation::is_isolatable_shape_error;

/// A slide in a `PowerPoint` presentation with lazy-loaded shapes and metadata.
///
/// The object owns parsed record state and uses one-time cells for expensive
/// projections.  Shape conversion and text extraction therefore remain
/// demand-driven with no repeated allocation after the first access.
#[allow(
    clippy::struct_field_names,
    reason = "the `slide_id`, `slide_list_text`, and `slide_number` fields mirror the MS-PPT `SlidePersistAtom` terminology; they are private and renaming them would churn sibling modules without changing the public API"
)]
pub struct Slide<'doc> {
    /// Slide persist ID.
    persist_id: u32,
    /// Stable `SlideId` from the live `SlidePersistAtom`.
    slide_id: u32,
    slide_list_text: String,
    outline_text_interactions: Vec<crate::TextBodyInteractions>,
    outline_text_refs: Vec<crate::OutlineTextRef>,
    /// The slide-list text bodies its shapes' outline references select.
    outline_text_bodies: &'doc [crate::OutlineTextBody],
    /// Slide number (1-based for display).
    slide_number: usize,
    /// Slide record.
    record: Record,
    /// Where the document stream lives, for lazy speaker-notes parsing.
    #[allow(dead_code, reason = "reserved for lazy speaker-notes parsing")]
    source: crate::records::PayloadStore<'doc>,
    /// Lazily-loaded shapes stored as owned values, together with the
    /// non-structural defects repaired while they were parsed.
    shapes: OnceCell<(Arc<[ShapeEnum<'static>]>, crate::leniency::ToleranceReport)>,
    /// The owning presentation's one-slide memo, when the slide was selected
    /// through a presentation (record 0260).
    shape_memo: Option<&'doc ShapeMemo>,
    /// Cached text content.
    text_cache: OnceCell<String>,
    /// Lazily parsed, inert shape animation metadata.
    animations: OnceCell<Vec<ShapeAnimation>>,
    /// Lazily parsed `PowerPoint` 2002 slide animation extension.
    animation_extension: OnceCell<Option<SlideAnimationExtension>>,
    /// Lazily parsed `PowerPoint` 12 slide/master round-trip metadata.
    powerpoint12_extension: OnceCell<SlideExtension>,
    /// Lazily parsed, inert slide-library synchronization metadata.
    sync_info: OnceCell<Option<Synchronization>>,
    /// Lazily parsed direct `PowerPoint` 12 slide round-trip metadata.
    round_trip_metadata: OnceCell<SlideRoundTripMetadata12>,
    notes_descriptor: Result<Option<NoteDescriptor>, String>,
    speaker_notes: OnceCell<Option<SpeakerNotes>>,
    source_limits: crate::SourceLimits,
}

/// Parsed shapes of one slide.
type ParsedShapes = (Arc<[ShapeEnum<'static>]>, crate::leniency::ToleranceReport);

/// A presentation-scoped memo of the most recently parsed slide's shapes
/// (record 0260).
///
/// Every [`crate::Presentation::slide_at`] or [`crate::Presentation::slides`]
/// call returns new [`Slide`] handles, each with its own lazily parsed shape
/// array. A unified write resolves its path through one handle and the
/// native edit then resolves the same target through another handle of the
/// same immutable presentation, so the slide was parsed twice. The memo
/// keeps the one most recently parsed slide's shape array (an `Arc` shared
/// with the handle that parsed it) and its repairs, keyed by the slide's
/// persist identifier within this immutable presentation, whose bytes never
/// change. A hit returns exactly what parsing would; a miss parses and
/// replaces the entry; a parse failure is never stored. The resident cost is
/// one slide's shapes. A poisoned lock is treated as a miss.
///
/// A committed text edit may instead seed the memo with [`PatchedShapes`]
/// that it proved equal to a parse of the published slide (record 0264).
/// They are materialized on the first hit, so a publication that nobody
/// reads again pays nothing for the seed.
#[derive(Default)]
pub(crate) struct ShapeMemo(std::sync::Mutex<Option<(u32, MemoEntry)>>);

enum MemoEntry {
    Parsed(ParsedShapes),
    Patched(Arc<PatchedShapes>),
}

/// A slide's shapes as a source array with some positions replaced: the
/// value a parse of the slide returns, as proven by an incremental readback
/// (record 0264). Parsing such a slide reports no repairs.
pub(crate) struct PatchedShapes {
    base: Arc<[ShapeEnum<'static>]>,
    /// Replaced positions, strictly increasing.
    replacements: Vec<(usize, ShapeEnum<'static>)>,
}

impl PatchedShapes {
    /// Combine a source array with replacements at strictly increasing
    /// in-bounds positions; `None` otherwise.
    pub(crate) fn new(
        base: Arc<[ShapeEnum<'static>]>,
        replacements: Vec<(usize, ShapeEnum<'static>)>,
    ) -> Option<Self> {
        let ordered = replacements.windows(2).all(|pair| pair[0].0 < pair[1].0);
        let bounded = replacements
            .last()
            .is_none_or(|(position, _)| *position < base.len());
        (ordered && bounded).then_some(Self { base, replacements })
    }

    /// The patched array: unchanged shapes cloned, replacements in place.
    pub(crate) fn materialize(&self) -> Arc<[ShapeEnum<'static>]> {
        let mut replacements = self.replacements.iter().peekable();
        let mut shapes = Vec::with_capacity(self.base.len());
        for (position, before) in self.base.iter().enumerate() {
            match replacements.next_if(|(replaced, _)| *replaced == position) {
                Some((_, after)) => shapes.push(after.clone()),
                None => shapes.push(before.clone()),
            }
        }
        Arc::from(shapes)
    }
}

impl ShapeMemo {
    fn get(&self, persist_id: u32) -> Option<ParsedShapes> {
        let mut entry = self.0.lock().ok()?;
        let (key, value) = entry.as_mut()?;
        if *key != persist_id {
            return None;
        }
        if let MemoEntry::Patched(patched) = value {
            *value = MemoEntry::Parsed((
                patched.materialize(),
                crate::leniency::ToleranceReport::default(),
            ));
        }
        match value {
            MemoEntry::Parsed(parsed) => Some((Arc::clone(&parsed.0), parsed.1.clone())),
            MemoEntry::Patched(_) => None,
        }
    }

    /// Seed the memo with a slide's shapes that were proven equal to what
    /// parsing this presentation's slide `persist_id` returns (record 0264:
    /// a committed text edit hands its incrementally verified shapes to the
    /// published package, so a following edit does not parse them again).
    /// Callers must hold that proof; the memo cannot check it.
    pub(crate) fn seed(&self, persist_id: u32, shapes: Arc<PatchedShapes>) {
        if let Ok(mut entry) = self.0.lock() {
            *entry = Some((persist_id, MemoEntry::Patched(shapes)));
        }
    }

    /// Whether the memo holds slide `persist_id` (tests only).
    #[cfg(test)]
    pub(crate) fn holds(&self, persist_id: u32) -> bool {
        self.0
            .lock()
            .is_ok_and(|entry| entry.as_ref().is_some_and(|(key, _)| *key == persist_id))
    }

    fn put(&self, persist_id: u32, value: &ParsedShapes) {
        if let Ok(mut entry) = self.0.lock() {
            *entry = Some((
                persist_id,
                MemoEntry::Parsed((Arc::clone(&value.0), value.1.clone())),
            ));
        }
    }
}

impl std::fmt::Debug for ShapeMemo {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("ShapeMemo")
    }
}
