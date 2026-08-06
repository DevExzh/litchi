use crate::package::Result;
use crate::parts::associated_strings::DocumentAssociatedStrings;
use crate::parts::auto_summary::DocumentAutoSummary;
use crate::parts::bookmarks::BookmarksTable;
use crate::parts::captions::CaptionTables;
use crate::parts::chp_bin_table::ChpBinTable;
use crate::parts::comments::CommentsTable;
use crate::parts::embedded_fonts::DocumentEmbeddedFonts;
use crate::parts::fib::FileInformationBlock;
use crate::parts::fields::FieldsTable;
use crate::parts::footnotes::{EndnotesTable, FootnotesTable};
use crate::parts::format_consistency::DocumentFormatConsistencyMarks;
use crate::parts::glossary::{AttachedGlossary, GlossaryMetadata};
use crate::parts::grammar_cookies::GrammarCookieTables;
use crate::parts::headers::HeadersTable;
use crate::parts::hyperlinks::HyperlinksTable;
use crate::parts::list_names::ListNamesTable;
use crate::parts::list_templates::ListTemplateTable;
use crate::parts::mail_merge::DocumentMailMerge;
use crate::parts::numbering::ListTables;
use crate::parts::ole::controls::Controls;
use crate::parts::pap_bin_table::PapBinTable;
use crate::parts::proofing::ProofingTables;
use crate::parts::protection::Ranges;
use crate::parts::repair_bookmarks::DocumentRepairBookmarks;
use crate::parts::revisions::RevisionAuthorTable;
use crate::parts::rmd_threading::DocumentRmdThreading;
use crate::parts::rsids::DocumentRsids;
use crate::parts::saved_by::SavedByTable;
use crate::parts::sections::SectionsTable;
use crate::parts::smart_tags::DocumentSmartTags;
use crate::parts::structured_tags::DocumentStructuredTags;
use crate::parts::styles::StyleSheet;
use crate::parts::subdocuments::Collection;
use crate::parts::table_char_cache::TableCharacterCache;
use crate::parts::text::TextExtractor;
use crate::parts::text_services::TextServicesTables;
use crate::parts::textbox_breaks::TextBoxBreakTables;
use std::collections::HashMap;
use std::sync::Arc;

/// A Word document (.doc).
///
/// This is the main API for reading and manipulating legacy Word document
/// content. Its storage is kept separate from the semantic query implementation
/// so binary streams and parsed tables remain owned here while the facade can
/// expose borrowed views without cloning document state.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_doc::Package;
///
/// let mut pkg = Package::open("document.doc")?;
/// let doc = pkg.document()?;
/// let text = doc.text()?;
/// println!("Document text: {}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Document {
    /// File Information Block from WordDocument stream
    pub(in crate::document) fib: FileInformationBlock,
    /// The WordDocument stream - main document binary data
    /// Used during initialization for TextExtractor and ChpBinTable parsing
    #[allow(dead_code)]
    pub(in crate::document) word_document: Vec<u8>,
    /// The Data stream - contains embedded objects, pictures, etc.
    /// According to Apache POI, pictures are stored here, not in WordDocument stream.
    pub(in crate::document) data_stream: Option<Vec<u8>>,
    /// Text extractor - holds the extracted document text
    pub(in crate::document) text_extractor: TextExtractor,
    /// Character property bin table - parsed once and shared across all paragraph extractors
    pub(in crate::document) chp_bin_table: Option<ChpBinTable>,
    /// Paragraph property bin table - parsed once and shared across all paragraph extractors
    pub(in crate::document) pap_bin_table: Option<PapBinTable>,
    /// Fields table - contains field information (embedded equations, hyperlinks, etc.)
    /// Used during initialization for hyperlink extraction; exposed via `fields_table()` accessor.
    pub(in crate::document) fields_table: Option<FieldsTable>,
    /// Headers and footers table
    pub(in crate::document) headers_table: Option<HeadersTable>,
    /// Footnotes table
    pub(in crate::document) footnotes_table: Option<FootnotesTable>,
    /// Endnotes table
    pub(in crate::document) endnotes_table: Option<EndnotesTable>,
    /// Comments table
    pub(in crate::document) comments_table: CommentsTable,
    /// Standard bookmark tables
    pub(in crate::document) bookmarks_table: BookmarksTable,
    /// Legacy Word smart-tag bookmarks, property bags, and recognizer ranges.
    pub(in crate::document) smart_tags: Option<DocumentSmartTags>,
    /// Revision-save identifiers assigned in the document.
    pub(in crate::document) rsids: Option<DocumentRsids>,
    /// E-mail review threading data parallel to the revision-author table.
    pub(in crate::document) rmd_threading: Option<DocumentRmdThreading>,
    /// Embedded TrueType font descriptions.
    pub(in crate::document) embedded_fonts: Option<DocumentEmbeddedFonts>,
    /// AutoSummary priority ranges for the main document.
    pub(in crate::document) auto_summary: Option<DocumentAutoSummary>,
    /// Word 2003 range-level protection ("editable ranges") metadata.
    pub(in crate::document) protected_ranges: Option<Ranges>,
    /// Format consistency-checker marks.
    pub(in crate::document) format_consistency_marks: Option<DocumentFormatConsistencyMarks>,
    /// Word 2003 structured document tag bookmarks.
    pub(in crate::document) structured_tags: Option<DocumentStructuredTags>,
    /// Word 2003 XML schema definition references (`Hplxsdr`).
    pub(in crate::document) xml_schemas: Option<crate::parts::xml_schemas::Collection>,
    /// Custom XML save transform path (`fcCustomXForm`).
    pub(in crate::document) custom_xml_transform_path: Option<String>,
    /// OLE controls recorded in the document.
    pub(in crate::document) ole_controls: Option<Controls>,
    /// Mail-merge data-source state (`Pms` and the ODSO property set).
    pub(in crate::document) mail_merge: Option<DocumentMailMerge>,
    /// Master-document subdocument directory and referenced-file name table.
    pub(in crate::document) subdocuments: Option<Collection>,
    /// Revision-mark authors
    pub(in crate::document) revision_authors: RevisionAuthorTable,
    /// Fixed associated-document strings
    pub(in crate::document) associated_strings: Option<DocumentAssociatedStrings>,
    /// Names parallel to list definitions for LISTNUM fields
    pub(in crate::document) list_names: Option<ListNamesTable>,
    /// List-level template codes parallel to list definitions
    pub(in crate::document) list_templates: Option<ListTemplateTable>,
    /// Deferred strict spelling/grammar proofing metadata parse
    pub(in crate::document) proofing_tables: Result<ProofingTables>,
    /// Deferred strict grammar-checker cookie metadata parse
    pub(in crate::document) grammar_cookies: Result<GrammarCookieTables>,
    /// Deferred strict deprecated table-character cache parse
    pub(in crate::document) table_char_cache: Result<Option<TableCharacterCache>>,
    /// Deferred strict textbox break-table metadata parse
    pub(in crate::document) textbox_breaks: Result<TextBoxBreakTables>,
    /// Deferred strict Text Services Framework metadata parse
    pub(in crate::document) text_services: Result<TextServicesTables>,
    /// Deferred strict Word 97/2000 save-history metadata parse
    pub(in crate::document) saved_by_table: Result<SavedByTable>,
    /// Deferred strict caption label and AutoCaption metadata parse
    pub(in crate::document) caption_tables: Result<CaptionTables>,
    /// Deferred strict repair-bookmark metadata parse
    pub(in crate::document) repair_bookmarks: Result<Option<DocumentRepairBookmarks>>,
    /// Deferred strict glossary-only AutoText metadata parse
    pub(in crate::document) glossary_metadata: Result<Option<GlossaryMetadata>>,
    /// Deferred strict secondary-FIB glossary parse for templates
    pub(in crate::document) attached_glossary: Result<Option<AttachedGlossary>>,
    /// Section ranges, layout, and property revision marks
    pub(in crate::document) sections: SectionsTable,
    /// Floating-shape anchors from the Main Document PlcfSpa (empty when the
    /// document has no floating shapes in the main story).
    pub(in crate::document) shape_anchors: Vec<crate::parts::spa::ShapeAnchor>,
    /// Floating-shape anchors from the Header Document PlcfSpa (empty when
    /// the document has no floating shapes in the header story).
    pub(in crate::document) header_shape_anchors: Vec<crate::parts::spa::ShapeAnchor>,
    /// Text box entries from the PlcftxbxTxt (empty when the document has no
    /// textbox story).
    pub(in crate::document) textbox_entries: Vec<crate::parts::textbox::TextBoxEntry>,
    /// Text box entries from the PlcfHdrtxbxTxt (empty when the document has no
    /// header textbox story).
    pub(in crate::document) header_textbox_entries: Vec<crate::parts::textbox::TextBoxEntry>,
    /// Hyperlinks table
    pub(in crate::document) hyperlinks_table: Option<HyperlinksTable>,
    /// List/numbering tables
    pub(in crate::document) list_tables: Option<ListTables>,
    /// Word 97+ stylesheet, including raw style UPX property sets.
    pub(in crate::document) stylesheet: Option<StyleSheet>,
    /// Extracted MTEF data from OLE streams (stream_name -> mtef_data)
    #[allow(dead_code)]
    pub(in crate::document) mtef_data: HashMap<String, Vec<u8>>,
    /// Parsed MTEF formulas rendered while their temporary parser arena is alive.
    /// Owned strings avoid a self-referential document and remain cheap to share.
    pub(in crate::document) parsed_mtef: HashMap<String, Arc<str>>,
}
