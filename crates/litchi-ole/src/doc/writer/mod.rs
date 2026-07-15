//! DOC file writing module
//!
//! This module provides comprehensive support for creating and modifying
//! Microsoft Word documents in the legacy binary format (.doc files).

/// Core DOC writer implementation
mod core;

/// Comments writer input types
pub mod bookmarks;
pub mod comments;

/// FIB (File Information Block) generation
pub mod fib;

/// Piece table for text storage
pub mod piece_table;

/// Tracked revision writer input types
pub mod revisions;

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

/// Headers and footers writer
pub mod headers;

/// Footnotes and endnotes writer
pub mod footnotes;

/// Hyperlinks writer
pub mod hyperlinks;

/// List numbering writer
pub mod numbering;

// Re-export public types
pub use crate::doc::parts::pap::{
    Border as ParagraphBorder, BorderStyle as ParagraphBorderStyle, Borders as ParagraphBorders,
    Shading as ParagraphShading, TextBoxTightWrap,
};
pub use crate::doc::parts::tap::{
    CellBorderTypes, CellShading, CellSpacing, CellSpacingSource, ShadingPattern,
    TableConditionalFormatting, TableHorizontalAnchor, TableHorizontalPosition, TableJustification,
    TableLook, TableLookFlags, TablePositioning, TableStyleBorder, TableStyleCondition,
    TableStyleDefaults, TableStyleShading, TableVerticalAnchor, TableVerticalPosition, TableWidth,
    WidthType,
};
pub use bookmarks::BookmarkEntry;
pub use comments::CommentEntry;
pub use core::{CharacterFormatting, DocWriteError, DocWriter, LineSpacing, ParagraphFormatting};
pub use fib::FibBuilder;
pub use fkp::{ChpxFkpBuilder, PapxFkpBuilder};
pub use footnotes::{EndnotesWriter, FootnoteEntry, FootnotesWriter};
pub use headers::{HeaderFooterEntry, HeaderFooterType, HeadersWriter};
pub use hyperlinks::{HyperlinkEntry, HyperlinkType, HyperlinksWriter};
pub use numbering::{ListFormatOverride, ListLevel, ListStructure, NumberFormat, NumberingWriter};
pub use piece_table::{Piece, PieceTableBuilder};
pub use revisions::{DisplayFieldRevision, FormattingRevision, NumberingRevision, TextRevision};
pub use sprm::SprmBuilder;
pub use tap::{
    TableBorders, TableCell, TableRevisionMark, TableRow, TapBuildError, TapBuilder,
    generate_table_style_sprms, generate_table_style_sprms_with_conditionals,
};
