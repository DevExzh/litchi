/// Internal parts for parsing DOC file structures.
///
/// This module contains parsers for the binary structures used in
/// legacy Word documents, including:
/// - FIB (File Information Block)
/// - Text extraction
/// - Character and paragraph properties
/// - Style definitions
/// - Table structures
pub mod chp;
pub mod fib;
pub mod pap;
pub mod paragraph_extractor;
pub mod tap;
pub mod text;
pub mod fields;
pub mod fkp;
pub mod chp_bin_table;
pub mod piece_table;

