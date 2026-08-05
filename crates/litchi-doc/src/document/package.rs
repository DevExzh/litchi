//! OLE package orchestration for the typed Word document model.

use super::model::Document;
use crate::encryption::decrypt_document_streams;
use crate::package::{Error as PackageError, OpenOptions, Result};
use crate::parts::associated_strings::DocumentAssociatedStrings;
use crate::parts::auto_summary::DocumentAutoSummary;
use crate::parts::bookmarks::BookmarksTable;
use crate::parts::captions::CaptionTables;
use crate::parts::chp_bin_table::ChpBinTable;
use crate::parts::comments::CommentsTable;
use crate::parts::embedded_fonts::DocumentEmbeddedFonts;
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};
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
use litchi_cfb::OleFile;
use std::io::{Read, Seek};

impl Document {
    /// Create a new Document from an OLE file.
    ///
    /// This is typically called internally by `Package::document()`.
    pub(crate) fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self> {
        Self::from_ole_with_options(ole, OpenOptions::default())
    }

    /// Create a new Document from an OLE file with password-to-open options.
    pub(crate) fn from_ole_with_options<R: Read + Seek>(
        ole: &mut OleFile<R>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        // Read the WordDocument stream (main document stream)
        let mut word_document = ole
            .open_stream(&["WordDocument"])
            .map_err(|_| PackageError::StreamNotFound("WordDocument".to_string()))?;

        // Parse the File Information Block (FIB) from the start of WordDocument
        let mut fib = FileInformationBlock::parse(&word_document)?;

        // Determine which table stream to use (0Table or 1Table)
        let table_stream_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };

        // Read the table stream. MS-DOC 2.1 requires one of `0Table`/`1Table`
        // to be present, so a file without it is either damaged or predates the
        // format: Word 6.0 and Word 95 keep those structures inside
        // `WordDocument`. Naming the version is far more actionable for the
        // caller than reporting a missing stream.
        let mut table_stream = ole.open_stream(&[table_stream_name]).map_err(|_| {
            if fib.version() < WORD_97_NFIB {
                PackageError::UnsupportedVersion {
                    nfib: fib.version(),
                    name: fib.version_name(),
                }
            } else {
                PackageError::StreamNotFound(table_stream_name.to_string())
            }
        })?;

        // Read the Data stream (optional - contains embedded pictures and objects)
        // According to Apache POI, pictures are stored in Data stream, not WordDocument stream
        let mut data_stream = ole.open_stream(&["Data"]).ok();

        if fib.is_encrypted() {
            decrypt_document_streams(
                &fib,
                &mut word_document,
                &mut table_stream,
                data_stream.as_deref_mut(),
                options.password,
            )?;
            // FibBase is clear, but the rest of the FIB was encrypted and must be
            // reparsed before any offsets or character counts are consulted.
            fib = FileInformationBlock::parse(&word_document)?;
        }

        // Create text extractor
        let text_extractor = TextExtractor::new(&fib, &word_document, &table_stream)?;

        // Parse fields table to identify embedded equations and hyperlinks
        let fields_table = Some(FieldsTable::parse(&fib, &table_stream)?);

        // Parse headers/footers table
        let headers_table = HeadersTable::parse(&fib, &table_stream).ok();

        // Parse footnotes and endnotes tables
        let footnotes_table = FootnotesTable::parse(&fib, &table_stream).ok();
        let endnotes_table = EndnotesTable::parse(&fib, &table_stream).ok();
        let comments_table = CommentsTable::parse(&fib, &table_stream)?;
        let bookmarks_table = BookmarksTable::parse(&fib, &table_stream)?;
        let smart_tags = DocumentSmartTags::parse(&fib, &table_stream)?;
        let rsids = DocumentRsids::parse(&fib, &table_stream)?;
        let rmd_threading = DocumentRmdThreading::parse(&fib, &table_stream)?;
        let embedded_fonts = DocumentEmbeddedFonts::parse(&fib, &table_stream)?;
        let auto_summary = DocumentAutoSummary::parse(&fib, &table_stream)?;
        let protected_ranges = Ranges::parse(&fib, &table_stream)?;
        let format_consistency_marks = DocumentFormatConsistencyMarks::parse(&fib, &table_stream)?;
        let structured_tags = DocumentStructuredTags::parse(&fib, &table_stream)?;
        let xml_schemas = crate::parts::xml_schemas::Collection::parse(&fib, &table_stream)?;
        let custom_xml_transform_path =
            crate::parts::xml_schemas::parse_custom_xml_transform(&fib, &table_stream)?;
        let ole_controls = Controls::parse(&fib, &table_stream)?;
        let mail_merge = DocumentMailMerge::parse(&fib, &table_stream)?;
        let subdocuments = Collection::parse(&fib, &table_stream)?;
        let revision_authors = RevisionAuthorTable::parse(&fib, &table_stream)?;
        let associated_strings = DocumentAssociatedStrings::parse(&fib, &table_stream)?;
        let list_names = ListNamesTable::parse(&fib, &table_stream)?;
        let list_templates = ListTemplateTable::parse(&fib, &table_stream)?;
        let proofing_tables = ProofingTables::parse(&fib, &table_stream);
        let grammar_cookies = GrammarCookieTables::parse(&fib, &table_stream);
        let table_char_cache = TableCharacterCache::parse(&fib, &table_stream);
        let textbox_breaks = TextBoxBreakTables::parse(&fib, &table_stream);
        let text_services = TextServicesTables::parse(&fib, &table_stream);
        let saved_by_table = SavedByTable::parse(&fib, &table_stream);
        let caption_tables = CaptionTables::parse(&fib, &table_stream);
        let repair_bookmarks = DocumentRepairBookmarks::parse(&fib, &table_stream);
        let glossary_metadata = GlossaryMetadata::parse(&fib, &table_stream).and_then(|metadata| {
            if let Some(metadata) = &metadata {
                metadata.validate_text_boundaries(&text_extractor)?;
            }
            Ok(metadata)
        });
        let attached_glossary =
            AttachedGlossary::parse(&fib, &word_document, &table_stream, data_stream.as_deref());
        let sections =
            SectionsTable::parse(&fib, &table_stream, &word_document, &revision_authors)?;
        let shape_anchors = Self::parse_shape_anchors(&fib, &table_stream);
        let header_shape_anchors = Self::parse_header_shape_anchors(&fib, &table_stream);
        let textbox_entries = Self::parse_textbox_entries(
            &fib,
            &table_stream,
            crate::parts::textbox::FIB_INDEX_PLCF_TXBX_TXT,
        );
        let header_textbox_entries = Self::parse_textbox_entries(
            &fib,
            &table_stream,
            crate::parts::textbox::FIB_INDEX_PLCF_HDR_TXBX_TXT,
        );

        // Parse hyperlinks from fields table
        let hyperlinks_table = fields_table.as_ref().and_then(|ft| {
            HyperlinksTable::from_fields(ft, |start, end| {
                Ok(text_extractor.text_at_range(start, end).to_string())
            })
            .ok()
        });

        // Parse list/numbering tables
        let list_tables = ListTables::parse(&fib, &table_stream).ok();

        // Word 97+ files are required to carry a stylesheet. Older Word files use
        // a different FIB and stylesheet representation that is not interpreted here.
        let mut stylesheet = (fib.version() >= 0x00C1)
            .then(|| StyleSheet::parse_with_leniency(&fib, &table_stream, options.leniency))
            .transpose()?;
        if let Some(stylesheet) = &mut stylesheet {
            stylesheet.resolve_revision_authors(&revision_authors)?;
        }

        // Extract MTEF data from OLE streams
        let mtef_data = Self::extract_mtef_data(ole)?;

        // Parse MTEF data into AST nodes
        #[cfg(feature = "formula")]
        let parsed_mtef = Self::parse_all_mtef_data(&mtef_data)?;
        #[cfg(not(feature = "formula"))]
        let parsed_mtef = Self::parse_all_mtef_data(&mtef_data)?;

        // Reconstruct both property bin tables from one shared piece-table parse.
        let piece_table = Self::parse_piece_table(&fib, &table_stream);
        let chp_bin_table = piece_table.as_ref().and_then(|piece_table| {
            Self::table_slice(&fib, &table_stream, 12)
                .and_then(|data| ChpBinTable::parse(data, &word_document, piece_table))
        });
        let pap_bin_table = if let (Some(piece_table), Some(data)) = (
            piece_table.as_ref(),
            Self::table_slice(&fib, &table_stream, 13),
        ) {
            PapBinTable::parse(
                data,
                &word_document,
                data_stream.as_deref(),
                piece_table,
                stylesheet.as_ref(),
            )?
        } else {
            None
        };

        Ok(Self {
            fib,
            word_document,
            data_stream,
            text_extractor,
            chp_bin_table,
            pap_bin_table,
            fields_table,
            headers_table,
            footnotes_table,
            endnotes_table,
            comments_table,
            bookmarks_table,
            smart_tags,
            rsids,
            rmd_threading,
            embedded_fonts,
            auto_summary,
            protected_ranges,
            format_consistency_marks,
            structured_tags,
            xml_schemas,
            custom_xml_transform_path,
            ole_controls,
            mail_merge,
            subdocuments,
            revision_authors,
            associated_strings,
            list_names,
            list_templates,
            proofing_tables,
            grammar_cookies,
            table_char_cache,
            textbox_breaks,
            text_services,
            saved_by_table,
            caption_tables,
            repair_bookmarks,
            glossary_metadata,
            attached_glossary,
            sections,
            shape_anchors,
            header_shape_anchors,
            textbox_entries,
            header_textbox_entries,
            hyperlinks_table,
            list_tables,
            stylesheet,
            mtef_data,
            parsed_mtef,
        })
    }
}
