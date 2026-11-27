//! Text extraction utilities for PPT records.

pub mod extractor;

// Re-export commonly used items
pub use extractor::{
    from_utf16le_lossy, parse_cstring, parse_text_bytes_atom, parse_text_chars_atom,
};
