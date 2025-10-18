mod operators;
mod templates;
mod matrix;
mod symbols;
mod utils;
mod conv;

pub use conv::converter::LatexConverter;
pub use conv::error::LatexError;

/// Efficient string interning for repeated LaTeX commands
/// Uses SmallVec to avoid allocations for common cases
#[derive(Debug)]
pub struct LatexStringCache {
    cache: smallvec::SmallVec<[String; 16]>,
}

impl LatexStringCache {
    pub fn new() -> Self {
        Self {
            cache: smallvec::SmallVec::new(),
        }
    }

    pub fn get_or_insert(&mut self, s: &str) -> usize {
        // Simple linear search for small cache
        for (i, cached) in self.cache.iter().enumerate() {
            if cached == s {
                return i;
            }
        }

        // Not found, insert
        let index = self.cache.len();
        self.cache.push(s.to_string());
        index
    }

    pub fn get(&self, index: usize) -> &str {
        &self.cache[index]
    }

    pub fn clear(&mut self) {
        self.cache.clear();
    }
}

impl Default for LatexStringCache {
    fn default() -> Self {
        Self::new()
    }
}

/// Performance statistics for LaTeX conversion
#[derive(Debug, Default)]
pub struct LatexConversionStats {
    pub nodes_processed: usize,
    pub strings_allocated: usize,
    pub buffer_resizes: usize,
    pub total_chars_written: usize,
}

impl LatexConversionStats {
    pub fn record_node(&mut self) {
        self.nodes_processed += 1;
    }

    pub fn record_allocation(&mut self, size: usize) {
        self.strings_allocated += 1;
        self.total_chars_written += size;
    }

    pub fn record_resize(&mut self) {
        self.buffer_resizes += 1;
    }
}

// Performance optimization imports
// Note: SIMD features can be enabled by adding appropriate crates
