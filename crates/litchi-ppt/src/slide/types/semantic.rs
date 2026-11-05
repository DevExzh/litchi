//! Contextual, caller-facing slide queries and lazy facade methods.

use super::Slide;
use super::model::{ParsedComment, ParsedSlideTiming};
use crate::animation::{ShapeAnimation, SlideAnimationExtension};
use crate::consts::RecordType;
use crate::odraw::ShapeLimitsExt as _;
use crate::package::{Error, Result};
use crate::records::Record;
use crate::shapes::ShapeEnum;
use crate::slide::notes::SpeakerNotes;
use crate::slide_extension::SlideExtension;
use crate::slide_round_trip::SlideRoundTripMetadata12;
use crate::slide_sync::Synchronization;
use crate::transition::{TransitionInfo, parse_transition};
use std::sync::Arc;

impl Slide<'_> {
    /// Get the slide number (1-based).
    #[inline]
    pub fn slide_number(&self) -> usize {
        self.slide_number
    }

    /// Get the persist ID.
    #[inline]
    pub fn persist_id(&self) -> u32 {
        self.persist_id
    }

    /// Get the stable presentation `SlideId`.
    #[inline]
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Get shapes on this slide (lazy-loaded).
    ///
    /// A shape whose text box references slide-list text (an
    /// `OutlineTextRefAtom`, MS-PPT 2.9.78) carries the text of the body the
    /// reference selects (L-PPT-R1); see [`Self::outline_text_bodies`].
    ///
    /// Non-structural text-style defects repaired while the shapes are parsed
    /// (see [`crate::TextPropDefect`]) are recorded alongside the cached
    /// shapes; enumerate them with [`Self::tolerance_report`] or thread a
    /// report through [`Self::shapes_with_report`].
    ///
    /// # Performance
    ///
    /// - Shapes are parsed only on first call
    /// - Subsequent calls return cached reference
    /// - Zero allocation after first parse
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes(&self) -> Result<&[ShapeEnum<'static>]> {
        self.parsed_shapes().map(|(shapes, _)| shapes.as_ref())
    }

    /// Share the cached native shape array without cloning shapes or text.
    ///
    /// The array owns its payloads and can outlive this slide. All handles
    /// returned by this slide share the same immutable allocation.
    ///
    /// # Errors
    /// Returns the same typed parse errors as [`Self::shapes`].
    ///
    /// ```
    /// let mut builder = litchi_ppt::Builder::new();
    /// builder.add_slide()?;
    /// let package = builder.build()?;
    /// let presentation = package.presentation()?;
    /// let slide = presentation.slide_at(0)?.expect("authored slide");
    /// let shapes = slide.shared_shapes()?;
    /// assert!(shapes.is_empty());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn shared_shapes(&self) -> Result<Arc<[ShapeEnum<'static>]>> {
        self.parsed_shapes().map(|(shapes, _)| Arc::clone(shapes))
    }

    /// Get shapes on this slide (lazy-loaded), appending every non-structural
    /// defect repaired while they were parsed to `report`.
    ///
    /// Records the [`crate::TextPropDefect`] classes repaired while shape
    /// text is parsed: clamped `StyleTextPropAtom` run coverage (F5.1) and
    /// masked reserved `BulletFlags`/`PFWrapFlags` bits (F5.2), and the
    /// [`crate::ShapeDefect`] classes of per-shape isolation (F5.11): an
    /// unreadable shape is dropped while its siblings are kept. Each entry is
    /// located by this slide's persist ID and the shape's ID.
    ///
    /// The shapes and their repairs are parsed once and cached together, so
    /// every call appends the same entries, whichever accessor triggered the
    /// parse.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shapes_with_report(
        &self,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<&[ShapeEnum<'static>]> {
        let (shapes, tolerance) = self.parsed_shapes()?;
        report.append(tolerance);
        Ok(shapes.as_ref())
    }

    /// Every non-structural defect repaired while this slide's shapes were
    /// parsed (triggers parsing if not yet loaded).
    ///
    /// An empty report means the slide's text styling conformed.
    ///
    /// # Errors
    ///
    /// Returns an error if the shapes cannot be parsed.
    pub fn tolerance_report(&self) -> Result<&crate::leniency::ToleranceReport> {
        self.parsed_shapes().map(|(_, tolerance)| tolerance)
    }

    /// The payload of this slide's first `PPDrawing` child, exactly the
    /// bytes [`Self::shapes`] parses (record 0264 incremental readback).
    pub(crate) fn drawing_payload(&self) -> Option<&[u8]> {
        self.record
            .find_child(RecordType::PPDrawing)
            .map(|drawing| drawing.data.as_slice())
    }

    fn parsed_shapes(
        &self,
    ) -> Result<&(Arc<[ShapeEnum<'static>]>, crate::leniency::ToleranceReport)> {
        self.shapes.get_or_try_init(|| {
            if let Some(hit) = self.shape_memo.and_then(|memo| memo.get(self.persist_id)) {
                return Ok(hit);
            }
            let (shapes, report) = self.parse_shapes()?;
            let parsed = (Arc::from(shapes), report);
            if let Some(memo) = self.shape_memo {
                memo.put(self.persist_id, &parsed);
            }
            Ok(parsed)
        })
    }

    /// Find a shape by its `OfficeArt` shape identifier (`spid`, see
    /// [`ShapeEnum::id`]), searching group children depth-first in drawing
    /// order (L-PPT-R4).
    ///
    /// MS-ODRAW requires identifiers to be unique within a drawing; when a
    /// file repeats one, the first shape in drawing order is returned. `None`
    /// when no returned shape has the identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the shapes cannot be parsed.
    pub fn shape_by_id(&self, id: u32) -> Result<Option<&ShapeEnum<'static>>> {
        let mut pending = self.shapes()?.iter().rev().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            if shape.id() == id {
                return Ok(Some(shape));
            }
            if let ShapeEnum::Group(group) = shape {
                pending.extend(group.children().iter().rev());
            }
        }
        Ok(None)
    }

    /// Get the number of shapes (triggers parsing if not yet loaded).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
    }

    /// Return every shape that has a click or mouse-over interaction.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_interactions(&self) -> Result<Vec<crate::ShapeInteractionEntry>> {
        self.shape_interactions_with_limits(crate::text_edit::runs::interaction_limits_for_source(
            self.source_limits,
        ))
    }

    /// Return shape interactions with caller-supplied record and name limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn shape_interactions_with_limits(
        &self,
        limits: crate::InteractionLimits,
    ) -> Result<Vec<crate::ShapeInteractionEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(
            &ppdrawing.data,
            &crate::odraw::limits_for_source(self.source_limits),
        )?;
        let mut result = Vec::new();
        let mut pending = escher_shapes.iter().rev().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            let interactions = shape.interactions_with_limits(limits)?;
            if !interactions.is_empty() {
                result.push(crate::ShapeInteractionEntry {
                    shape_id: shape.id(),
                    interactions,
                });
            }
            pending.extend(shape.children().iter().rev());
        }
        Ok(result)
    }

    /// Return every shape that has a range-anchored text action.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_text_interactions(&self) -> Result<Vec<crate::ShapeTextInteractionEntry>> {
        self.shape_text_interactions_with_limits(crate::text_interaction::limits_for_source(
            self.source_limits,
        ))
    }

    /// Return shape text actions with caller-supplied resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn shape_text_interactions_with_limits(
        &self,
        limits: crate::TextInteractionLimits,
    ) -> Result<Vec<crate::ShapeTextInteractionEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(
            &ppdrawing.data,
            &crate::odraw::limits_for_source(self.source_limits),
        )?;
        let mut result = Vec::new();
        let mut pending = escher_shapes.iter().rev().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            let interactions = shape.text_interactions_with_limits(limits)?;
            if !interactions.is_empty() {
                result.push(crate::ShapeTextInteractionEntry {
                    shape_id: shape.id(),
                    interactions,
                });
            }
            pending.extend(shape.children().iter().rev());
        }
        Ok(result)
    }

    /// Range-anchored actions stored with outline/placeholder text.
    pub fn outline_text_interactions(&self) -> &[crate::TextBodyInteractions] {
        &self.outline_text_interactions
    }

    /// Validated outline text references (`OutlineTextRefAtom`, MS-PPT 2.9.78)
    /// found among this slide's slide-list records.
    ///
    /// Conforming files keep these references in the shapes' text boxes,
    /// not in the slide list, so this is normally empty; the slide reader
    /// resolves the references the shapes hold (see
    /// [`Self::outline_text_bodies`]).
    pub fn outline_text_refs(&self) -> &[crate::OutlineTextRef] {
        &self.outline_text_refs
    }

    /// This slide's slide-list text bodies, in slide-list order (MS-PPT
    /// 2.4.14.3; L-PPT-R1).
    ///
    /// `PowerPoint` keeps the text of title and body placeholders here: the
    /// placeholder's text box holds an `OutlineTextRefAtom` whose index
    /// selects one of these bodies. [`Self::shapes`] and [`Self::text`]
    /// already read such placeholders from the body they select. Empty when
    /// the slide has no slide-list text or it could not be read
    /// ([`crate::SlideListDefect::UnreadableSlideListText`]).
    pub fn outline_text_bodies(&self) -> &[crate::OutlineTextBody] {
        self.outline_text_bodies
    }

    /// Return every shape-scoped programmable-tag container on this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_programmable_tags(&self) -> Result<Vec<crate::ShapeProgrammableTagsEntry>> {
        self.shape_programmable_tags_with_limits(crate::ShapeProgrammableTagLimits::default())
    }

    /// Return shape programmable tags with caller-supplied resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn shape_programmable_tags_with_limits(
        &self,
        limits: crate::ShapeProgrammableTagLimits,
    ) -> Result<Vec<crate::ShapeProgrammableTagsEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(
            &ppdrawing.data,
            &crate::odraw::limits_for_source(self.source_limits),
        )?;
        let mut result = Vec::new();
        for shape in &escher_shapes {
            if let Some(programmable_tags) = shape.programmable_tags_with_limits(limits)? {
                result.push(crate::ShapeProgrammableTagsEntry {
                    shape_id: shape.id(),
                    programmable_tags,
                });
            }
        }
        Ok(result)
    }

    /// Return this slide's typed slide-level programmable tags (MS-PPT 2.5.19),
    /// when the slide carries a `SlideProgTagsContainer`.
    ///
    /// Tag payloads are inert: they are parsed and preserved, never executed,
    /// loaded, or resolved. Use
    /// [`crate::ProgTags::slide_extensions`] to decode the
    /// versioned binary-tag payloads into typed extension structs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn programmable_tags(&self) -> Result<Option<crate::ProgTags>> {
        self.programmable_tags_with_limits(crate::ProgTagLimits::default())
    }

    /// Return slide-level programmable tags with caller-supplied resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn programmable_tags_with_limits(
        &self,
        limits: crate::ProgTagLimits,
    ) -> Result<Option<crate::ProgTags>> {
        crate::ProgTags::parse_slide(&self.record, limits)
    }

    /// Return every typed shape-flag projection on this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_flags(&self) -> Result<Vec<crate::ShapeFlagEntry>> {
        self.shape_flags_with_limits(crate::ShapeFlagLimits::default())
    }

    /// Return shape flags with caller-supplied client-data resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn shape_flags_with_limits(
        &self,
        limits: crate::ShapeFlagLimits,
    ) -> Result<Vec<crate::ShapeFlagEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(
            &ppdrawing.data,
            &crate::odraw::limits_for_source(self.source_limits),
        )?;
        let mut result = Vec::new();
        for shape in &escher_shapes {
            if let Some(projection) = shape.ppt_flags_with(limits)? {
                result.push(crate::ShapeFlagEntry {
                    shape_id: shape.id(),
                    projection,
                });
            }
        }
        Ok(result)
    }

    /// Return context-validated placeholders on this presentation slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn placeholder_atoms(&self) -> Result<Vec<crate::PlaceholderEntry>> {
        self.placeholder_atoms_with_limits(crate::PlaceholderLimits::default())
    }

    /// Return placeholders with caller-supplied client-data limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn placeholder_atoms_with_limits(
        &self,
        limits: crate::PlaceholderLimits,
    ) -> Result<Vec<crate::PlaceholderEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(
            &ppdrawing.data,
            &crate::odraw::limits_for_source(self.source_limits),
        )?;
        let mut positions = std::collections::HashSet::new();
        let mut result = Vec::new();
        for shape in &escher_shapes {
            if let Some(placeholder) = shape.placeholder_atom_with_limits(
                crate::PlaceholderContext::PresentationSlide,
                limits,
            )? {
                if placeholder.position != -1 && !positions.insert(placeholder.position) {
                    return Err(Error::Corrupted(
                        "Presentation slide contains duplicate placeholder positions".to_string(),
                    ));
                }
                result.push(crate::PlaceholderEntry {
                    shape_id: shape.id(),
                    placeholder,
                });
            }
        }
        Ok(result)
    }

    /// Return this slide's speaker-notes page, if one exists.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn speaker_notes(&self) -> Result<Option<&SpeakerNotes>> {
        self.speaker_notes
            .get_or_try_init(|| match &self.notes_descriptor {
                Ok(None) => Ok(None),
                Ok(Some(descriptor)) => {
                    SpeakerNotes::parse_with_limits(*descriptor, self.source, self.source_limits)
                        .map(Some)
                },
                Err(error) => Err(Error::Corrupted(error.clone())),
            })
            .map(Option::as_ref)
    }

    /// Return inert `PowerPoint` 97 animation metadata keyed by shape ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn animations(&self) -> Result<&[ShapeAnimation]> {
        self.animations
            .get_or_try_init(|| self.parse_animations())
            .map(Vec::as_slice)
    }
    /// Return inert `PowerPoint` 2002 timing and build metadata from `___PPT10`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn animation_extension(&self) -> Result<Option<&SlideAnimationExtension>> {
        self.animation_extension
            .get_or_try_init(|| self.parse_animation_extension())
            .map(Option::as_ref)
    }
    /// Return `PowerPoint` 12 slide/master round-trip metadata from `___PPT12`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn powerpoint12_extension(&self) -> Result<&SlideExtension> {
        self.powerpoint12_extension
            .get_or_try_init(|| SlideExtension::parse(&self.record, &self.source_limits))
    }

    /// Return inert `PowerPoint` 12 slide-library synchronization metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn sync_info(&self) -> Result<Option<&Synchronization>> {
        self.sync_info
            .get_or_try_init(|| Synchronization::parse(&self.record))
            .map(Option::as_ref)
    }

    /// Return inert `PowerPoint` 12 metadata stored directly on this slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn powerpoint12_round_trip_metadata(&self) -> Result<&SlideRoundTripMetadata12> {
        self.round_trip_metadata
            .get_or_try_init(|| SlideRoundTripMetadata12::parse(&self.record))
    }

    /// Extract all text from this slide (lazy-loaded).
    ///
    /// The text is the slide's direct text records, then the text of every
    /// text box in drawing order. A text box that references slide-list
    /// text (an `OutlineTextRefAtom`, as title and body placeholders do)
    /// contributes the body it selects (L-PPT-R1). When none of that yields
    /// text, the slide's slide-list text is returned instead.
    ///
    /// # Performance
    ///
    /// - Text is extracted and cached on first call
    /// - Includes text from:
    ///   * Direct text records in the slide
    ///   * Shapes (via PPDrawing/Escher), with outline references resolved
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(&self) -> Result<&str> {
        self.text_cache
            .get_or_try_init(|| {
                let text = self.extract_all_text()?;
                if text.is_empty() {
                    Ok(self.slide_list_text.clone())
                } else {
                    Ok(text)
                }
            })
            .map(String::as_str)
    }
    /// Check if this slide has a `PPDrawing` record (shapes).
    #[inline]
    pub fn has_drawing(&self) -> bool {
        self.record.find_child(RecordType::PPDrawing).is_some()
    }

    /// Get raw slide record for advanced use cases.
    #[inline]
    pub fn record(&self) -> &Record {
        &self.record
    }

    /// Parse comments from this slide's `BinaryTagData`.
    ///
    /// Comments are stored inside `ProgTags/ProgBinaryTag/BinaryTagData`
    /// as `Comment2000` (type=12000) containers.
    ///
    /// # Returns
    ///
    /// A vector of parsed comments (author, text, initials, position, date).
    /// Returns an empty vector if no comments are found.
    ///
    /// # Errors
    ///
    /// Returns an error when the `PowerPoint` 10 programmable tag or a comment record is malformed.
    /// The text, character runs and text hyperlinks of the shape with
    /// `OfficeArt` ID `shape_id` (record 0323); an outline-backed placeholder
    /// reads the slide-list body it selects. `None` for a shape without a
    /// text body.
    ///
    /// # Errors
    ///
    /// Returns an error if the shape's text records are malformed.
    pub fn shape_text_runs(&self, shape_id: u32) -> Result<Option<crate::text_edit::TextRuns>> {
        let Some(drawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(None);
        };
        match crate::text_edit::runs::read(
            &drawing.data,
            self.outline_text_bodies,
            shape_id,
            self.source_limits,
        ) {
            Ok(runs) => Ok(runs),
            // A body the run model does not cover (several text atoms, an
            // ambiguous textbox) has no run projection.
            Err(crate::text_edit::Error::Refused(_)) => Ok(None),
            Err(crate::text_edit::Error::Package(error)) => Err(error),
            Err(crate::text_edit::Error::Source(error)) => Err(Error::Corrupted(error.to_string())),
            Err(
                error @ (crate::text_edit::Error::PatchConflict | crate::text_edit::Error::Signed),
            ) => Err(Error::Corrupted(error.to_string())),
        }
    }

    pub fn comments(&self) -> Result<Vec<ParsedComment>> {
        crate::comments::parse_slide_comments(&self.record, &self.source_limits)
    }

    /// Get the slide transition from the `SSSlideInfoAtom` record.
    ///
    /// The transition describes the visual effect (type, direction, speed),
    /// the advance mode (on click, automatic, or both), and an optional
    /// sound played when the slide is shown.
    ///
    /// # Returns
    ///
    /// `Ok(None)` when the slide has no `SSSlideInfoAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error when the `SSSlideInfoAtom` record is truncated.
    pub fn transition(&self) -> Result<Option<TransitionInfo>> {
        match self.record.find_child(RecordType::SSSlideInfoAtom) {
            Some(info) => Ok(Some(parse_transition(info)?)),
            None => Ok(None),
        }
    }

    /// Get the slide timing from the `SSSlideInfoAtom` record.
    ///
    /// Returns `None` if the slide has no timing record.
    pub fn timing(&self) -> Option<ParsedSlideTiming> {
        // SSSlideInfoAtom (type=1017) is a direct child of the Slide container
        let info = self.record.find_child(RecordType::SSSlideInfoAtom)?;

        if info.data.len() < 16 {
            return None;
        }

        let d = &info.data;
        let slide_time_ms = u32::from_le_bytes([d[0], d[1], d[2], d[3]]);
        // d[4..8] soundIdRef, d[8] effect direction, d[9] effect type, and
        // d[12] speed are part of the atom layout but not yet exposed.
        let flags = u16::from_le_bytes([d[10], d[11]]);

        Some(ParsedSlideTiming {
            advance_time_ms: slide_time_ms,
            advance_on_click: (flags & (1 << 0)) != 0,
            auto_advance: (flags & (1 << 10)) != 0,
            hidden: (flags & (1 << 2)) != 0,
        })
    }
}
