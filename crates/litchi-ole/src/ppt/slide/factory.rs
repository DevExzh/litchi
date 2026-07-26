use super::directory::{SlideDirectory, SlideDirectoryEntry};
use super::notes::{NoteDescriptor, NotesIndex};
use crate::consts::PptRecordType;
/// SlideFactory - Creates slides from persist mapping with zero-copy parsing.
///
/// High-performance implementation using lifetimes to avoid data copying.
use crate::ppt::package::{PptError, Result};
use crate::ppt::persist::PersistMapping;
use crate::ppt::records::PptRecord;
use once_cell::unsync::OnceCell;

/// Factory for creating slides from document data using persist mapping.
///
/// # Performance
///
/// - Zero-copy: Borrows from original document data
/// - Lazy evaluation: Slides created on-demand via iterator
/// - Minimal allocations: Direct slice access
pub struct SlideFactory<'doc> {
    /// Reference to the complete document stream data
    doc_data: &'doc [u8],
    /// Persist ID to byte offset mapping
    persist_mapping: &'doc PersistMapping,
    notes_index: OnceCell<NotesIndex>,
    slide_directory: &'doc SlideDirectory,
}

impl<'doc> SlideFactory<'doc> {
    /// Create a new slide factory.
    #[inline]
    pub fn new(
        doc_data: &'doc [u8],
        persist_mapping: &'doc PersistMapping,
        slide_directory: &'doc SlideDirectory,
    ) -> Self {
        Self {
            doc_data,
            persist_mapping,
            notes_index: OnceCell::new(),
            slide_directory,
        }
    }

    /// Get all slide persist IDs in sorted order (filtered to only Slide records).
    pub fn slide_ids(&self) -> Vec<u32> {
        self.slide_directory
            .entries()
            .iter()
            .map(|entry| entry.persist_id())
            .collect()
    }

    /// Parse a slide at the given persist ID.
    ///
    /// # Performance
    ///
    /// - Zero-copy: Returns SlideData borrowing from doc_data
    /// - No intermediate buffers
    /// - Direct record parsing at offset
    pub fn parse_slide(&self, persist_id: u32) -> Result<SlideData<'doc>> {
        let offset = self.persist_mapping.get_offset(persist_id).ok_or_else(|| {
            PptError::InvalidFormat(format!("No offset found for persist_id {}", persist_id))
        })?;

        let entry = self
            .slide_directory
            .get_by_persist_id(persist_id)
            .ok_or_else(|| {
                PptError::InvalidFormat(format!(
                    "persist ID {persist_id} is not a logical presentation slide"
                ))
            })?;
        self.parse_slide_at_offset(offset, entry)
    }

    /// Parse slide record at specific byte offset.
    fn parse_slide_at_offset(
        &self,
        offset: u32,
        entry: &SlideDirectoryEntry,
    ) -> Result<SlideData<'doc>> {
        let offset = offset as usize;

        if offset + 8 > self.doc_data.len() {
            return Err(PptError::Corrupted(format!(
                "Offset {} exceeds document length",
                offset
            )));
        }

        // Parse the Slide record at this offset
        let (record, _consumed) = PptRecord::parse(self.doc_data, offset)?;

        if record.record_type != PptRecordType::Slide {
            return Err(PptError::InvalidFormat(format!(
                "Expected Slide record, got {:?}",
                record.record_type
            )));
        }

        let note_descriptor = self
            .notes_index
            .get_or_init(|| NotesIndex::build(self.doc_data, self.slide_directory))
            .descriptor(&record, entry.persist_id(), self.persist_mapping);

        Ok(SlideData {
            persist_id: entry.persist_id(),
            slide_id: entry.slide_id(),
            slide_list_text: entry.list_text().to_string(),
            outline_text_interactions: entry.outline_text_interactions().to_vec(),
            offset,
            record,
            doc_data: self.doc_data,
            note_descriptor,
        })
    }

    /// Create iterator over all slides.
    ///
    /// # Performance
    ///
    /// - Lazy: Slides parsed only when iterated
    /// - No allocation until iteration begins
    /// - Short-circuits on error
    pub fn slides(&self) -> impl Iterator<Item = Result<SlideData<'doc>>> + '_ {
        self.slide_ids()
            .into_iter()
            .map(move |persist_id| self.parse_slide(persist_id))
    }
}

/// Parsed slide data with zero-copy references.
///
/// # Lifetimes
///
/// Borrows from the original document data ('doc lifetime).
#[derive(Debug)]
pub struct SlideData<'doc> {
    /// Persist ID for this slide
    pub persist_id: u32,
    /// Stable SlideId from SlidePersistAtom.
    pub slide_id: u32,
    /// Text records associated with this slide in SlideListWithText.
    pub(crate) slide_list_text: String,
    /// Range-anchored actions from SlideListWithText text bodies.
    pub(crate) outline_text_interactions: Vec<crate::ppt::PowerPointTextBodyInteractions>,
    /// Byte offset in document stream
    pub offset: usize,
    /// Parsed Slide record
    pub record: PptRecord,
    /// Reference to complete document data (for lazy shape parsing)
    doc_data: &'doc [u8],
    pub(crate) note_descriptor: std::result::Result<Option<NoteDescriptor>, String>,
}

impl<'doc> SlideData<'doc> {
    /// Get the SlideAtom child record containing layout/master info.
    #[inline]
    pub fn slide_atom(&self) -> Option<&PptRecord> {
        self.record.find_child(PptRecordType::SlideAtom)
    }

    /// Get the PPDrawing child record containing shapes.
    #[inline]
    pub fn ppdrawing(&self) -> Option<&PptRecord> {
        self.record.find_child(PptRecordType::PPDrawing)
    }

    /// Check if this slide has drawing data (shapes).
    #[inline]
    pub fn has_shapes(&self) -> bool {
        self.ppdrawing().is_some()
    }

    /// Get reference to document data for advanced parsing.
    #[inline]
    pub fn doc_data(&self) -> &'doc [u8] {
        self.doc_data
    }

    /// Create a SlideData instance for testing purposes.
    ///
    /// # Note
    ///
    /// This is only available in test builds.
    #[cfg(test)]
    pub fn new_for_test(
        persist_id: u32,
        offset: usize,
        record: PptRecord,
        doc_data: &'doc [u8],
    ) -> Self {
        Self {
            persist_id,
            slide_id: persist_id,
            slide_list_text: String::new(),
            outline_text_interactions: Vec::new(),
            offset,
            record,
            doc_data,
            note_descriptor: Ok(None),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_slide_factory_creation() {
        let doc_data = vec![0u8; 1024];
        let mapping = PersistMapping::new();

        let directory = SlideDirectory::new_for_test(0);
        let factory = SlideFactory::new(&doc_data, &mapping, &directory);
        assert_eq!(factory.slide_ids().len(), 0);
    }
}
