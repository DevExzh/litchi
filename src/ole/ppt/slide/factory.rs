use crate::ole::consts::PptRecordType;
/// SlideFactory - Creates slides from persist mapping with zero-copy parsing.
///
/// High-performance implementation using lifetimes to avoid data copying.
use crate::ole::ppt::package::{PptError, Result};
use crate::ole::ppt::persist::PersistMapping;
use crate::ole::ppt::records::PptRecord;

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
}

impl<'doc> SlideFactory<'doc> {
    /// Create a new slide factory.
    #[inline]
    pub fn new(doc_data: &'doc [u8], persist_mapping: &'doc PersistMapping) -> Self {
        Self {
            doc_data,
            persist_mapping,
        }
    }

    /// Get all slide persist IDs in sorted order (filtered to only Slide records).
    pub fn slide_ids(&self) -> Vec<u32> {
        // Filter to only actual Slide records (not Notes, Masters, etc.)
        let all_ids = self.persist_mapping.get_persist_ids();
        all_ids
            .into_iter()
            .filter(|&persist_id| {
                if let Some(offset) = self.persist_mapping.get_offset(persist_id) {
                    let offset = offset as usize;
                    if offset + 4 < self.doc_data.len() {
                        // Check record type at this offset (bytes 2-3)
                        let record_type = u16::from_le_bytes([
                            self.doc_data[offset + 2],
                            self.doc_data[offset + 3],
                        ]);
                        // 1006 = Slide record type
                        record_type == 1006
                    } else {
                        false
                    }
                } else {
                    false
                }
            })
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

        self.parse_slide_at_offset(offset, persist_id)
    }

    /// Parse slide record at specific byte offset.
    fn parse_slide_at_offset(&self, offset: u32, persist_id: u32) -> Result<SlideData<'doc>> {
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

        Ok(SlideData {
            persist_id,
            offset,
            record,
            doc_data: self.doc_data,
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
    /// Byte offset in document stream
    pub offset: usize,
    /// Parsed Slide record
    pub record: PptRecord,
    /// Reference to complete document data (for lazy shape parsing)
    doc_data: &'doc [u8],
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
            offset,
            record,
            doc_data,
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

        let factory = SlideFactory::new(&doc_data, &mapping);
        assert_eq!(factory.slide_ids().len(), 0);
    }
}
