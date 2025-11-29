//! Shared byte slice for zero-copy element storage.
//!
//! This module provides `XmlSlice`, a shared reference to a portion of data
//! stored in a contiguous arena buffer. This eliminates per-element allocations
//! by storing all element data in a single buffer and referencing slices.
//!
//! While named `XmlSlice` for historical reasons, this type is generic and can
//! be used for any byte data that benefits from shared arena-based storage.

use std::sync::Arc;

/// A shared slice of data from an arena buffer.
///
/// This struct provides zero-copy access to element data by storing
/// a reference to a shared buffer and the byte range within it.
///
/// # Performance
///
/// - Clone is O(1) - just increments Arc refcount and copies two u32s
/// - No per-element heap allocation - all data lives in the shared arena
/// - Memory-efficient: 24 bytes per slice (Arc pointer + start + len)
#[derive(Debug, Clone)]
pub struct XmlSlice {
    /// Shared reference to the arena buffer containing all data
    arena: Arc<Vec<u8>>,
    /// Start offset in the arena
    start: u32,
    /// Length of the slice
    len: u32,
}

impl XmlSlice {
    /// Create a new XmlSlice from an arena and byte range.
    #[inline]
    pub fn new(arena: Arc<Vec<u8>>, start: u32, len: u32) -> Self {
        Self { arena, start, len }
    }

    /// Get the data bytes as a slice.
    #[inline]
    pub fn as_bytes(&self) -> &[u8] {
        let start = self.start as usize;
        let end = start + self.len as usize;
        &self.arena[start..end]
    }

    /// Get the length of the data.
    #[inline]
    pub fn len(&self) -> usize {
        self.len as usize
    }

    /// Check if the slice is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Get a clone of the underlying Arc (for creating sub-slices).
    #[inline]
    pub fn arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.arena)
    }

    /// Get the start offset.
    #[inline]
    pub fn start(&self) -> u32 {
        self.start
    }
}

/// Builder for creating multiple XmlSlices from a single arena.
///
/// This collects all data into a contiguous buffer during parsing,
/// then converts to a shared Arc for creating slices.
#[derive(Debug)]
pub struct XmlArenaBuilder {
    /// The buffer being built
    buffer: Vec<u8>,
    /// Recorded (start, len) positions for each element
    positions: Vec<(u32, u32)>,
}

impl XmlArenaBuilder {
    /// Create a new arena builder with estimated capacity.
    #[inline]
    pub fn with_capacity(buffer_capacity: usize, element_count: usize) -> Self {
        Self {
            buffer: Vec::with_capacity(buffer_capacity),
            positions: Vec::with_capacity(element_count),
        }
    }

    /// Get a mutable reference to the current write position in the buffer.
    /// Returns the start position for this element.
    #[inline]
    pub fn start_element(&self) -> u32 {
        self.buffer.len() as u32
    }

    /// Get mutable access to the buffer for writing.
    #[inline]
    pub fn buffer_mut(&mut self) -> &mut Vec<u8> {
        &mut self.buffer
    }

    /// Finish writing an element and record its position.
    /// Returns the index of this element.
    #[inline]
    pub fn finish_element(&mut self, start: u32) -> usize {
        let len = self.buffer.len() as u32 - start;
        let idx = self.positions.len();
        self.positions.push((start, len));
        idx
    }

    /// Convert the builder into a shared arena and return slice creators.
    #[inline]
    pub fn build(self) -> (Arc<[u8]>, Vec<(u32, u32)>) {
        let arena: Arc<[u8]> = self.buffer.into();
        (arena, self.positions)
    }

    /// Get the number of elements recorded.
    #[inline]
    pub fn element_count(&self) -> usize {
        self.positions.len()
    }
}
