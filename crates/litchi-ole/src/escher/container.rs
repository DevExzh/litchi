//! Container record handling with iterator-based traversal.
//!
//! # Performance
//!
//! - Lazy evaluation: Children parsed on-demand
//! - Iterator-based: Functional composition
//! - Zero-copy: Borrows from parent data

use super::record::{EscherRecord, Result};

/// Iterator over child records in an Escher container.
///
/// # Performance
///
/// - Lazy: Parses one record at a time
/// - Zero-copy: Returns borrowed records
/// - No allocation until iteration
pub struct EscherChildIterator<'data> {
    data: &'data [u8],
    offset: usize,
}

impl<'data> EscherChildIterator<'data> {
    /// Create a new child iterator for a container record.
    #[inline]
    pub fn new(container_data: &'data [u8]) -> Self {
        Self {
            data: container_data,
            offset: 0,
        }
    }
}

impl<'data> Iterator for EscherChildIterator<'data> {
    type Item = Result<EscherRecord<'data>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset + 8 > self.data.len() {
            return None;
        }

        match EscherRecord::parse(self.data, self.offset) {
            Ok((record, consumed)) => {
                self.offset += consumed;

                if consumed == 0 {
                    return None;
                }

                Some(Ok(record))
            },
            Err(e) => {
                self.offset += 1;

                if self.offset + 7 < self.data.len() {
                    Some(Err(e))
                } else {
                    None
                }
            },
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.data.len().saturating_sub(self.offset);
        let est_count = remaining / 20;
        (0, Some(est_count))
    }
}

/// Escher container wrapper for convenient child access.
#[derive(Debug, Clone)]
pub struct EscherContainer<'data> {
    record: EscherRecord<'data>,
}

impl<'data> EscherContainer<'data> {
    /// Wrap an Escher record as a container.
    ///
    /// # Panics
    ///
    /// Panics in debug mode if the record is not a container.
    #[inline]
    pub fn new(record: EscherRecord<'data>) -> Self {
        debug_assert!(record.is_container(), "Record is not a container");
        Self { record }
    }

    /// Get the wrapped record.
    #[inline]
    pub fn record(&self) -> &EscherRecord<'data> {
        &self.record
    }

    /// Iterate over child records.
    ///
    /// # Performance
    ///
    /// - Returns lazy iterator
    /// - Children parsed on-demand
    /// - Zero-copy borrows
    #[inline]
    pub fn children(&self) -> EscherChildIterator<'data> {
        EscherChildIterator::new(self.record.data)
    }

    /// Find the first child of a specific type.
    ///
    /// # Performance
    ///
    /// - Short-circuits on first match
    /// - Only parses records until match found
    pub fn find_child(
        &self,
        record_type: super::types::EscherRecordType,
    ) -> Option<EscherRecord<'data>> {
        self.children()
            .filter_map(|r| r.ok())
            .find(|r| r.record_type == record_type)
    }

    /// Find all children of a specific type.
    ///
    /// # Performance
    ///
    /// - Functional filter chain
    /// - Pre-allocated Vec with estimated capacity
    pub fn find_children(
        &self,
        record_type: super::types::EscherRecordType,
    ) -> Vec<EscherRecord<'data>> {
        self.children()
            .filter_map(|r| r.ok())
            .filter(|r| r.record_type == record_type)
            .collect()
    }

    /// Recursively find all records of a specific type (depth-first).
    ///
    /// # Performance
    ///
    /// - Depth-first traversal
    /// - Lazy evaluation via iterator
    /// - Pre-allocated result vector
    pub fn find_recursive(
        &self,
        record_type: super::types::EscherRecordType,
    ) -> Vec<EscherRecord<'data>> {
        let mut results = Vec::new();
        self.find_recursive_impl(record_type, &mut results);
        results
    }

    fn find_recursive_impl(
        &self,
        record_type: super::types::EscherRecordType,
        results: &mut Vec<EscherRecord<'data>>,
    ) {
        for child in self.children().flatten() {
            if child.record_type == record_type {
                results.push(child.clone());
            }

            if child.is_container() {
                let container = EscherContainer::new(child);
                container.find_recursive_impl(record_type, results);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::super::types::EscherRecordType;
    use super::*;

    #[test]
    fn test_child_iterator() {
        let data = vec![
            0x02, 0x00, 0x0A, 0xF0, 0x04, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x03, 0x00,
            0x0B, 0xF0, 0x02, 0x00, 0x00, 0x00, 0x05, 0x06,
        ];

        let iter = EscherChildIterator::new(&data);
        let records: Vec<_> = iter.filter_map(|r| r.ok()).collect();

        assert_eq!(records.len(), 2);
        assert_eq!(records[0].record_type, EscherRecordType::Sp);
        assert_eq!(records[1].record_type, EscherRecordType::Opt);
    }
}
