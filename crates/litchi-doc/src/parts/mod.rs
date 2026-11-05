pub mod annotation_bookmarks;
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
pub mod document_properties_2007;
pub mod document_properties_2010;
pub mod document_properties_2013;
pub mod document_properties_97;
pub mod dofr;
pub mod embedded_fonts;
pub mod envelope;
pub mod fib;
pub mod fields;
pub mod fkp;
pub mod fonts;
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
pub mod paragraph_groups;
pub mod piece_table;
pub mod print_environment;
pub mod proofing;
pub mod protection;
pub mod repair_bookmarks;
pub mod revisions;
pub mod rmd_threading;
pub mod route_slip;
pub mod rsids;
pub mod saved_by;
pub mod saved_selection;
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
pub mod vba_signature;
pub mod xml_schemas;

/// Charge one variable-size DOC fragment against the retained source-entry
/// policy before its codec allocates from a count or length field.
///
/// Fragment codecs receive slices of a CFB stream, so the source-entry
/// dimension is the applicable physical budget.  The package ingress has
/// already charged the complete stream aggregate; this helper prevents a
/// public raw boundary from silently selecting a larger per-fragment budget.
pub(crate) fn enforce_source_entry_limit(
    data: &[u8],
    limits: &crate::package::SourceLimits,
) -> crate::package::Result<()> {
    enforce_source_entry_size(data.len(), limits)
}

/// Charge an already computed fragment size before reserving its output.
///
/// DOC's raw part writers use the same finite per-entry policy as their input
/// codecs.  Keeping the check separate from the slice helper lets writers
/// reject an exact over-budget size before allocating their output buffer.
pub(crate) fn enforce_source_entry_size(
    size: usize,
    limits: &crate::package::SourceLimits,
) -> crate::package::Result<()> {
    let observed = u64::try_from(size).unwrap_or(u64::MAX);
    let maximum = u64::try_from(limits.max_entry_bytes()).unwrap_or(u64::MAX);
    if observed > maximum {
        return Err(crate::package::Error::Limit {
            kind: crate::package::SourceLimitKind::EntryBytes,
            observed,
            maximum,
        });
    }
    Ok(())
}
