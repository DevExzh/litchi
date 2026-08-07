//! Contextual, caller-facing slide queries and lazy facade methods.

use super::Slide;
use super::model::{ParsedComment, ParsedSlideTiming};
use crate::animation::{ShapeAnimation, SlideAnimationExtension};
use crate::consts::RecordType;
use crate::odraw::ShapeExt as _;
use crate::package::{Error, Result};
use crate::records::Record;
use crate::shapes::ShapeEnum;
use crate::slide::notes::SpeakerNotes;
use crate::slide_extension::SlideExtension;
use crate::slide_round_trip::SlideRoundTripMetadata12;
use crate::slide_sync::Synchronization;
use crate::transition::{TransitionInfo, parse_transition};

impl<'doc> Slide<'doc> {
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

    /// Get the stable presentation SlideId.
    #[inline]
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Get shapes on this slide (lazy-loaded).
    ///
    /// # Performance
    ///
    /// - Shapes are parsed only on first call
    /// - Subsequent calls return cached reference
    /// - Zero allocation after first parse
    pub fn shapes(&self) -> Result<&[ShapeEnum<'static>]> {
        self.shapes
            .get_or_try_init(|| self.parse_shapes())
            .map(|v| v.as_slice())
    }

    /// Get the number of shapes (triggers parsing if not yet loaded).
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
    }

    /// Return every shape that has a click or mouse-over interaction.
    pub fn shape_interactions(&self) -> Result<Vec<crate::ShapeInteractionEntry>> {
        self.shape_interactions_with_limits(crate::InteractionLimits::default())
    }

    /// Return shape interactions with caller-supplied record and name limits.
    pub fn shape_interactions_with_limits(
        &self,
        limits: crate::InteractionLimits,
    ) -> Result<Vec<crate::ShapeInteractionEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(&ppdrawing.data)?;
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
    pub fn shape_text_interactions(&self) -> Result<Vec<crate::ShapeTextInteractionEntry>> {
        self.shape_text_interactions_with_limits(crate::TextInteractionLimits::default())
    }

    /// Return shape text actions with caller-supplied resource limits.
    pub fn shape_text_interactions_with_limits(
        &self,
        limits: crate::TextInteractionLimits,
    ) -> Result<Vec<crate::ShapeTextInteractionEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(&ppdrawing.data)?;
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
    /// tying this slide's shapes to outline text bodies.
    pub fn outline_text_refs(&self) -> &[crate::OutlineTextRef] {
        &self.outline_text_refs
    }

    /// Return every shape-scoped programmable-tag container on this slide.
    pub fn shape_programmable_tags(&self) -> Result<Vec<crate::ShapeProgrammableTagsEntry>> {
        self.shape_programmable_tags_with_limits(crate::ShapeProgrammableTagLimits::default())
    }

    /// Return shape programmable tags with caller-supplied resource limits.
    pub fn shape_programmable_tags_with_limits(
        &self,
        limits: crate::ShapeProgrammableTagLimits,
    ) -> Result<Vec<crate::ShapeProgrammableTagsEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(&ppdrawing.data)?;
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
    pub fn programmable_tags(&self) -> Result<Option<crate::ProgTags>> {
        self.programmable_tags_with_limits(crate::ProgTagLimits::default())
    }

    /// Return slide-level programmable tags with caller-supplied resource limits.
    pub fn programmable_tags_with_limits(
        &self,
        limits: crate::ProgTagLimits,
    ) -> Result<Option<crate::ProgTags>> {
        crate::ProgTags::parse_slide(&self.record, limits)
    }

    /// Return every typed shape-flag projection on this slide.
    pub fn shape_flags(&self) -> Result<Vec<crate::ShapeFlagEntry>> {
        self.shape_flags_with_limits(crate::ShapeFlagLimits::default())
    }

    /// Return shape flags with caller-supplied client-data resource limits.
    pub fn shape_flags_with_limits(
        &self,
        limits: crate::ShapeFlagLimits,
    ) -> Result<Vec<crate::ShapeFlagEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(&ppdrawing.data)?;
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
    pub fn placeholder_atoms(&self) -> Result<Vec<crate::PlaceholderEntry>> {
        self.placeholder_atoms_with_limits(crate::PlaceholderLimits::default())
    }

    /// Return placeholders with caller-supplied client-data limits.
    pub fn placeholder_atoms_with_limits(
        &self,
        limits: crate::PlaceholderLimits,
    ) -> Result<Vec<crate::PlaceholderEntry>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::odraw::parse(&ppdrawing.data)?;
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
    pub fn speaker_notes(&self) -> Result<Option<&SpeakerNotes>> {
        self.speaker_notes
            .get_or_try_init(|| match &self.notes_descriptor {
                Ok(None) => Ok(None),
                Ok(Some(descriptor)) => {
                    SpeakerNotes::parse_with_limits(*descriptor, self.doc_data, self.record_limits)
                        .map(Some)
                },
                Err(error) => Err(Error::Corrupted(error.clone())),
            })
            .map(Option::as_ref)
    }

    /// Return inert PowerPoint 97 animation metadata keyed by shape ID.
    pub fn animations(&self) -> Result<&[ShapeAnimation]> {
        self.animations
            .get_or_try_init(|| self.parse_animations())
            .map(Vec::as_slice)
    }
    /// Return inert PowerPoint 2002 timing and build metadata from `___PPT10`.
    pub fn animation_extension(&self) -> Result<Option<&SlideAnimationExtension>> {
        self.animation_extension
            .get_or_try_init(|| self.parse_animation_extension())
            .map(Option::as_ref)
    }
    /// Return PowerPoint 12 slide/master round-trip metadata from `___PPT12`.
    pub fn powerpoint12_extension(&self) -> Result<&SlideExtension> {
        self.powerpoint12_extension
            .get_or_try_init(|| SlideExtension::parse(&self.record))
    }

    /// Return inert PowerPoint 12 slide-library synchronization metadata.
    pub fn sync_info(&self) -> Result<Option<&Synchronization>> {
        self.sync_info
            .get_or_try_init(|| Synchronization::parse(&self.record))
            .map(Option::as_ref)
    }

    /// Return inert PowerPoint 12 metadata stored directly on this slide.
    pub fn powerpoint12_round_trip_metadata(&self) -> Result<&SlideRoundTripMetadata12> {
        self.round_trip_metadata
            .get_or_try_init(|| SlideRoundTripMetadata12::parse(&self.record))
    }

    /// Extract all text from this slide (lazy-loaded).
    ///
    /// # Performance
    ///
    /// - Text is extracted and cached on first call
    /// - Includes text from:
    ///   * Direct text records in the slide
    ///   * Shapes (via PPDrawing/Escher)
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
            .map(|s| s.as_str())
    }
    /// Check if this slide has a PPDrawing record (shapes).
    #[inline]
    pub fn has_drawing(&self) -> bool {
        self.record.find_child(RecordType::PPDrawing).is_some()
    }

    /// Get raw slide record for advanced use cases.
    #[inline]
    pub fn record(&self) -> &Record {
        &self.record
    }

    /// Parse comments from this slide's BinaryTagData.
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
    /// Returns an error when the PowerPoint 10 programmable tag or a comment record is malformed.
    pub fn comments(&self) -> Result<Vec<ParsedComment>> {
        crate::comments::parse_slide_comments(&self.record)
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

    /// Get the slide timing from the SSSlideInfoAtom record.
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
        let _sound_id_ref = u32::from_le_bytes([d[4], d[5], d[6], d[7]]);
        let _effect_direction = d[8];
        let _effect_type = d[9];
        let flags = u16::from_le_bytes([d[10], d[11]]);
        let _speed = d[12];

        Some(ParsedSlideTiming {
            advance_time_ms: slide_time_ms,
            advance_on_click: (flags & (1 << 0)) != 0,
            auto_advance: (flags & (1 << 10)) != 0,
            hidden: (flags & (1 << 2)) != 0,
        })
    }
}
