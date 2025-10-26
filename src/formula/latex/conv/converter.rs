// LaTeX Converter Implementation
//
// This module contains the LatexConverter struct and its core methods.

use super::error::LatexError;
use super::utils::estimate_nodes_size;
use super::utils::extend_buffer_with_capacity;
use crate::formula::ast::{Formula, MathNode};
use crate::formula::latex::{LatexConversionStats, LatexStringCache};
use smallvec::SmallVec;

/// LaTeX converter that converts formula AST to LaTeX strings
///
/// Uses optimized string building and memory management for high performance.
/// Includes SIMD optimizations and efficient buffer management.
pub struct LatexConverter {
    /// Buffer for building the LaTeX output with pre-allocated capacity
    pub(super) buffer: String,
    /// Temporary buffer for complex conversions to avoid allocations
    pub(super) temp_buffer: SmallVec<[String; 4]>,
    /// String cache for repeated LaTeX commands
    pub(super) string_cache: LatexStringCache,
    /// Performance statistics
    pub(super) stats: LatexConversionStats,
}

impl LatexConverter {
    /// Create a new LaTeX converter with optimized initial capacity
    pub fn new() -> Self {
        Self {
            buffer: String::with_capacity(2048), // Larger initial capacity for better performance
            temp_buffer: SmallVec::new(),
            string_cache: LatexStringCache::new(),
            stats: LatexConversionStats::default(),
        }
        // NOTE: Cache initialization removed - it was O(n²) and never used.
        // The cache will be populated lazily during conversion as needed.
    }

    /// Create a new LaTeX converter with custom initial capacity
    pub fn with_capacity(capacity: usize) -> Self {
        Self {
            buffer: String::with_capacity(capacity),
            temp_buffer: SmallVec::new(),
            string_cache: LatexStringCache::new(),
            stats: LatexConversionStats::default(),
        }
        // NOTE: Cache initialization removed - lazy population is more efficient
    }

    // NOTE: initialize_cache() removed - was O(n²) complexity with 150+ string allocations.
    // The cache now populated lazily during conversion, which is more efficient since:
    // 1. Avoids upfront cost when converter is created but not used
    // 2. Only caches strings that are actually needed
    // 3. Eliminates wasteful linear search through growing cache (O(n²) → O(1) with lazy)

    /// Convert a formula to LaTeX
    ///
    /// Returns a reference to avoid unnecessary string cloning.
    ///
    /// # Example
    /// ```ignore
    /// let converter = LatexConverter::new();
    /// let latex = converter.convert(&formula)?;
    /// ```
    pub fn convert(&mut self, formula: &Formula) -> Result<&str, LatexError> {
        self.reset();

        // Reserve additional capacity based on estimated formula size
        let estimated_size = super::utils::estimate_formula_size(formula.root());
        extend_buffer_with_capacity(&mut self.buffer, "", estimated_size);

        // Add display style delimiters
        if formula.display_style() {
            self.buffer.push_str("\\[");
        } else {
            self.buffer.push_str("\\(");
        }

        // Convert all root nodes
        for node in formula.root() {
            self.convert_node(node)?;
        }

        // Close delimiters
        if formula.display_style() {
            self.buffer.push_str("\\]");
        } else {
            self.buffer.push_str("\\)");
        }

        Ok(&self.buffer)
    }

    /// Convert nodes without wrapping delimiters
    ///
    /// Returns a reference to avoid unnecessary string cloning.
    pub fn convert_nodes(&mut self, nodes: &[MathNode]) -> Result<&str, LatexError> {
        self.reset();

        // Reserve capacity
        let estimated_size = estimate_nodes_size(nodes);
        extend_buffer_with_capacity(&mut self.buffer, "", estimated_size);

        for node in nodes {
            self.convert_node(node)?;
        }

        Ok(&self.buffer)
    }

    /// Get the current buffer content without clearing
    #[inline]
    pub fn buffer(&self) -> &str {
        &self.buffer
    }

    /// Reset the converter state for a new conversion
    #[inline]
    pub fn reset(&mut self) {
        self.buffer.clear();
        self.temp_buffer.clear();
        self.string_cache.clear();
        // Keep stats for performance monitoring
    }

    /// Clear the internal buffers (legacy method)
    #[inline]
    pub fn clear(&mut self) {
        self.reset();
    }

    /// Get performance statistics
    #[inline]
    pub fn stats(&self) -> &LatexConversionStats {
        &self.stats
    }

    /// Get a cached LaTeX command string to avoid repeated allocations
    #[inline]
    #[allow(dead_code)]
    fn get_cached_command(&mut self, cmd: &str) -> &str {
        let index = self.string_cache.get_or_insert(cmd);
        self.string_cache.get(index)
    }

    /// Efficiently append a cached LaTeX command to the buffer
    #[inline]
    pub fn append_cached_command(&mut self, cmd: &str) {
        let index = self.string_cache.get_or_insert(cmd);
        let cached = self.string_cache.get(index);
        self.buffer.push_str(cached);
    }
}

impl Default for LatexConverter {
    fn default() -> Self {
        Self::new()
    }
}

impl AsRef<str> for LatexConverter {
    fn as_ref(&self) -> &str {
        &self.buffer
    }
}
