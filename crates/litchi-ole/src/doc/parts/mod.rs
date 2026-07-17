pub mod bookmarks;
pub mod associated_strings;
/// Internal parts for parsing DOC file structures.
///
/// This module contains parsers for the binary structures used in
/// legacy Word documents, including:
/// - FIB (File Information Block)
/// - Text extraction
/// - Character and paragraph properties
/// - Style definitions
/// - Table structures
/// - Headers/footers, footnotes/endnotes, comments, hyperlinks, numbering/lists
pub mod chp;
pub mod chp_bin_table;
pub mod comments;
pub mod fib;
pub mod fields;
pub mod fkp;
pub mod footnotes;
pub mod headers;
pub mod hyperlinks;
pub mod numbering;
pub mod pap;
pub mod pap_bin_table;
pub mod paragraph_extractor;
pub mod piece_table;
pub mod list_names;
pub mod revisions;
pub mod saved_by;
pub mod sections;
pub mod styles;
pub mod tap;
pub mod tap_parser;
pub mod text;
