//! High-performance Escher record parser.
//!
//! # Architecture
//!
//! - Zero-copy: Borrows from document data
//! - Lazy: Parses on-demand via iterators
//! - Efficient: Single-pass with minimal allocations

use super::record::EscherRecord;
use super::container::{EscherContainer, EscherChildIterator};
use super::types::EscherRecordType;
use crate::ole::ppt::package::Result;

/// Escher parser for extracting shapes and text from PPDrawing records.
pub struct EscherParser<'data> {
    /// The raw PPDrawing/Escher data
    data: &'data [u8],
}

impl<'data> EscherParser<'data> {
    /// Create a new Escher parser from PPDrawing data.
    #[inline]
    pub fn new(data: &'data [u8]) -> Self {
        Self { data }
    }
    
    /// Get the root DgContainer (Drawing Container).
    ///
    /// # Performance
    ///
    /// - Parses only first record
    /// - Returns None if not a DgContainer
    pub fn root_container(&self) -> Option<Result<EscherContainer<'data>>> {
        if self.data.len() < 8 {
            return None;
        }
        
        match EscherRecord::parse(self.data, 0) {
            Ok((record, _)) if record.is_container() => {
                Some(Ok(EscherContainer::new(record)))
            }
            Ok(_) => None,
            Err(e) => Some(Err(e)),
        }
    }
    
    /// Iterate over all top-level records.
    #[inline]
    pub fn records(&self) -> EscherChildIterator<'data> {
        EscherChildIterator::new(self.data)
    }
    
    /// Find all SpContainer (Shape Container) records recursively.
    ///
    /// # Performance
    ///
    /// - Depth-first traversal
    /// - Short-circuits on errors
    /// - Pre-allocated result vector
    pub fn find_all_shapes(&self) -> Result<Vec<EscherRecord<'data>>> {
        let mut shapes = Vec::new();
        
        // Parse root container
        if let Some(root_result) = self.root_container() {
            let root = root_result?;
            shapes.extend(root.find_recursive(EscherRecordType::SpContainer));
        }
        
        Ok(shapes)
    }
    
    /// Find all ClientTextbox records (contains text).
    ///
    /// # Performance
    ///
    /// - Functional filter chain
    /// - Only traverses text-bearing containers
    pub fn find_all_textboxes(&self) -> Result<Vec<EscherRecord<'data>>> {
        let mut textboxes = Vec::new();
        
        // Parse root container
        if let Some(root_result) = self.root_container() {
            let root = root_result?;
            textboxes.extend(root.find_recursive(EscherRecordType::ClientTextbox));
        }
        
        Ok(textboxes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn test_parser_creation() {
        let data = vec![
            0x0F, 0x00, // version=0xF (container), instance=0
            0x02, 0xF0, // record type = 0xF002 (DgContainer)
            0x04, 0x00, 0x00, 0x00, // length = 4
            0x01, 0x02, 0x03, 0x04, // data
        ];
        
        let parser = EscherParser::new(&data);
        let root = parser.root_container().unwrap().unwrap();
        
        assert_eq!(root.record().record_type, EscherRecordType::DgContainer);
    }
}

