pub mod associated_strings;
pub mod bookmarks;
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
pub mod document_properties;
pub mod document_properties_2000;
pub mod document_properties_2002;
pub mod document_properties_2003;
pub mod document_properties_97;
pub mod fib;
pub mod fields;
pub mod fkp;
pub mod footnotes;
pub mod form_fields;
pub mod glossary;
pub mod headers;
pub mod hyperlinks;
pub mod images;
pub mod list_names;
pub mod list_templates;
pub mod mail_merge;
pub mod numbering;
pub mod pap;
pub mod pap_bin_table;
pub mod paragraph_extractor;
pub mod piece_table;
pub mod proofing;
pub mod protection;
pub mod revisions;
pub mod rsids;
pub mod saved_by;
pub mod sections;
pub mod smart_tags;
pub mod spa;
pub mod styles;
pub mod tap;
pub mod tap_parser;
pub mod text;
pub mod textbox;
