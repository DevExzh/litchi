//! Text Extraction Utilities for iWork Documents
//!
//! This module provides shared text extraction functionality used across
//! Pages, Numbers, and Keynote documents.

pub mod extractor;
pub mod storage;
pub mod style;

pub use extractor::TextExtractor;
pub use storage::{TextFragment, TextRun, TextStorage};
pub use style::{ParagraphStyle, TextStyle};
