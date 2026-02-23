/// Internal parts for parsing DOC file structures.
///
/// This module contains parsers for the binary structures used in
/// legacy Word documents, including:
/// - FIB (File Information Block)
/// - Text extraction
/// - Character and paragraph properties
/// - Style definitions
/// - Table structures
/// - Headers/footers, footnotes/endnotes, hyperlinks, numbering/lists
pub mod chp;
pub mod chp_bin_table;
pub mod fib;
pub mod fields;
pub mod fkp;
pub mod footnotes;
pub mod headers;
pub mod hyperlinks;
pub mod numbering;
pub mod pap;
pub mod paragraph_extractor;
pub mod piece_table;
pub mod tap;
pub mod tap_parser;
pub mod text;
