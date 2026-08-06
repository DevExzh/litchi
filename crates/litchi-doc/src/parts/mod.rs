pub mod associated_strings;
pub mod auto_summary;
pub mod bookmarks;
pub mod captions;
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
pub mod command_bars;
pub mod comments;
pub mod document_properties;
pub mod document_properties_2000;
pub mod document_properties_2002;
pub mod document_properties_2003;
pub mod document_properties_97;
pub mod embedded_fonts;
pub mod envelope;
pub mod fib;
pub mod fields;
pub mod fkp;
pub mod footnotes;
pub mod form_fields;
pub mod format_consistency;
pub mod glossary;
pub mod grammar_cookies;
pub mod headers;
pub mod hyperlinks;
pub mod images;
pub mod list_names;
pub mod list_templates;
pub mod mail_merge;
pub mod numbering;
pub mod ole_controls;
pub mod pap;
pub mod pap_bin_table;
pub mod paragraph_extractor;
pub mod piece_table;
pub mod proofing;
pub mod protection;
pub mod repair_bookmarks;
pub mod revisions;
pub mod rmd_threading;
pub mod route_slip;
pub mod rsids;
pub mod saved_by;
pub mod sections;
pub mod smart_tags;
pub mod spa;
pub mod structured_tags;
pub mod styles;
pub mod subdocuments;
pub mod table_char_cache;
pub mod tap;
pub mod tap_parser;
pub mod text;
pub mod text_services;
pub mod textbox;
pub mod textbox_breaks;
pub mod xml_schemas;
