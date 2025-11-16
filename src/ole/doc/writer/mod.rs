//! DOC file writing module
//!
//! This module provides comprehensive support for creating and modifying
//! Microsoft Word documents in the legacy binary format (.doc files).

/// Core DOC writer implementation
mod core;

/// FIB (File Information Block) generation
pub mod fib;

/// Piece table for text storage
pub mod piece_table;

/// FKP (Formatted Disk Pages) structures
pub mod fkp;

/// SPRM (Single Property Modifier) generation
pub mod sprm;

/// TAP (Table Properties) generation
pub mod tap;

/// StyleSheet generation
pub mod stylesheet;

/// DocumentProperties generation
pub mod dop;

/// Section table generation
pub mod section;

/// Bin table (plcfbte) generation
pub mod bin_table;

/// Font table generation
pub mod font_table;

/// OLE metadata streams (CompObj, Ole)
pub mod ole_metadata;

// Re-export public types
pub use core::{CharacterFormatting, DocWriteError, DocWriter, ParagraphFormatting};
pub use fib::FibBuilder;
pub use fkp::{ChpxFkpBuilder, PapxFkpBuilder};
pub use piece_table::{Piece, PieceTableBuilder};
pub use sprm::SprmBuilder;
pub use tap::{TableCell, TableRow, TapBuilder};
