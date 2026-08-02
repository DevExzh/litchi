//! SlideAtomsSet - groups text records associated with a single slide.
//!
//! Based on POI's SlideListWithText.SlideAtomsSet inner class.

use super::record::PptRecord;
use crate::package::Result;

/// A set of records associated with a single slide's text.
///
/// This groups together a SlidePersistAtom (which identifies the slide)
/// with all its associated text records (TextHeaderAtom, TextCharsAtom/TextBytesAtom, etc.)
#[derive(Debug, Clone)]
pub struct SlideAtomsSet<'a> {
    /// The SlidePersistAtom that identifies which slide this text belongs to
    pub slide_persist_atom: &'a PptRecord,
    /// All text-related records for this slide (TextHeaderAtom, TextCharsAtom, etc.)
    pub slide_records: Vec<&'a PptRecord>,
}

impl<'a> SlideAtomsSet<'a> {
    /// Extract all text from this SlideAtomsSet.
    /// Based on POI's text extraction logic for slide atom sets.
    pub fn extract_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        for record in &self.slide_records {
            if let Ok(text) = record.extract_text()
                && !text.is_empty()
            {
                text_parts.push(text);
            }
        }

        Ok(text_parts.join("\n"))
    }

    /// Get the slide ID from the SlidePersistAtom.
    pub fn get_slide_id(&self) -> Option<u32> {
        self.slide_persist_atom.get_slide_id()
    }

    /// Validated outline text references (`OutlineTextRefAtom`, MS-PPT
    /// 2.9.78) tying this slide's shapes to outline text bodies.
    pub fn outline_text_refs(&self) -> Result<Vec<crate::PowerPointOutlineTextRef>> {
        self.slide_records
            .iter()
            .filter(|record| record.record_type == crate::consts::PptRecordType::OutlineTextRefAtom)
            .map(|record| crate::PowerPointOutlineTextRef::parse_record(record))
            .collect()
    }

    /// Parse range-anchored actions from every text body in this slide set.
    pub fn text_interactions(&self) -> Result<Vec<crate::PowerPointTextBodyInteractions>> {
        self.text_interactions_with_limits(crate::PowerPointTextInteractionLimits::default())
    }

    /// Parse text actions with caller-supplied resource limits.
    pub fn text_interactions_with_limits(
        &self,
        limits: crate::PowerPointTextInteractionLimits,
    ) -> Result<Vec<crate::PowerPointTextBodyInteractions>> {
        crate::text_interaction::parse_text_bodies(&self.slide_records, limits)
    }
}
