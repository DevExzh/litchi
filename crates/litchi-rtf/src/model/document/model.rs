#![cfg_attr(not(test), deny(clippy::indexing_slicing))]

//! RTF document representation.

use super::super::error::{RtfError, RtfResult, try_reserve_one};
use super::super::lexer::Lexer;
use super::super::limits::ParseLimits;
use super::super::parser::Parser;
use super::super::types::{ColorTable, FontTable, Paragraph as RtfParagraph, Run, StyleBlock};
use bumpalo::Bump;
use std::borrow::Cow;
use std::collections::HashSet;
use std::fs::File;
use std::io::Read;
use std::path::Path;

const MAX_ROOT_SHAPES: usize = 65_536;
const MAX_ROOT_SHAPE_GROUPS: usize = 16_384;

fn read_file_with_limit(path: &Path, limit: usize) -> RtfResult<Vec<u8>> {
    let file = File::open(path)
        .map_err(|error| RtfError::ParserError(format!("Failed to open file: {error}")))?;

    if let Ok(metadata) = file.metadata() {
        let limit_u64 = u64::try_from(limit).unwrap_or(u64::MAX);
        if metadata.len() > limit_u64 {
            return Err(RtfError::LimitExceeded {
                resource: "source bytes",
                observed: usize::try_from(metadata.len()).unwrap_or(usize::MAX),
                limit,
            });
        }
    }

    // Read at most one byte beyond the configured ceiling so special files,
    // concurrent growth, and inaccurate metadata cannot bypass the budget.
    let read_ceiling = limit.checked_add(1).unwrap_or(limit);
    let mut reader = file.take(u64::try_from(read_ceiling).unwrap_or(u64::MAX));
    let mut bytes = Vec::new();
    reader
        .read_to_end(&mut bytes)
        .map_err(|error| RtfError::ParserError(format!("Failed to read file: {error}")))?;
    if bytes.len() > limit {
        return Err(RtfError::LimitExceeded {
            resource: "source bytes",
            observed: bytes.len(),
            limit,
        });
    }
    Ok(bytes)
}

fn owned_table(
    table: &super::super::table::Table<'_>,
) -> RtfResult<super::super::table::Table<'static>> {
    let mut output = super::super::table::Table::new();
    output.set_direction(table.direction());
    for row in table.rows() {
        let mut owned_row = super::super::table::Row::new();
        owned_row.set_table_style(row.table_style());
        owned_row.set_table_rsid(row.table_rsid());
        owned_row.set_direction(row.direction());
        owned_row.set_layout(*row.layout());
        owned_row.set_borders(row.borders().clone());
        owned_row.set_shading(row.shading());
        owned_row.set_geometry(row.geometry());
        owned_row.set_autoformat_flags(row.autoformat_flags());
        owned_row.set_banding(row.banding());
        owned_row.set_revision(row.revision());
        for cell in row.cells() {
            let mut owned_cell = super::super::table::Cell::with_distances(
                Cow::Owned(cell.text().to_string()),
                cell.padding().clone(),
                cell.spacing().clone(),
            );
            owned_cell.set_layout(*cell.layout());
            owned_cell.set_merge(cell.merge());
            owned_cell.set_right_boundary(cell.right_boundary());
            owned_cell.set_preferred_width(cell.preferred_width());
            owned_cell.set_revision(cell.revision());
            owned_cell.set_borders(cell.borders().clone());
            owned_cell.set_shading(cell.shading());
            for nested in cell.nested_tables() {
                owned_cell.add_nested_table(nested.text_offset, owned_table(&nested.table)?)?;
            }
            owned_cell.set_story_content(
                cell.shapes()
                    .iter()
                    .cloned()
                    .map(crate::Shape::into_owned)
                    .collect(),
                cell.shape_groups()
                    .iter()
                    .cloned()
                    .map(crate::ShapeGroup::into_owned)
                    .collect(),
                cell.drawing_order().to_vec(),
                cell.story_events().to_vec(),
            )?;
            owned_row.add_cell(owned_cell);
        }
        owned_row.set_padding(row.padding().clone());
        owned_row.set_spacing(row.spacing().clone());
        owned_row.set_cell_defaults(row.cell_defaults().clone());
        owned_row.set_positioning(row.positioning().clone());
        output.add_row(owned_row);
    }
    Ok(output)
}

/// RTF Document.
///
/// This is the main entry point for parsing RTF documents.
/// It provides access to the document's text content, paragraphs, runs, and tables.
pub struct RtfDocument<'a> {
    /// Font table
    font_table: FontTable<'a>,
    /// Optional external-file metadata table.
    file_table: Option<crate::FileTable<'a>>,
    /// Color table
    color_table: ColorTable,
    /// Style blocks
    blocks: Vec<StyleBlock<'a>>,
    /// Extracted tables
    tables: Vec<super::super::table::Table<'a>>,
    /// Extracted pictures
    pictures: Vec<super::super::picture::Picture<'a>>,
    /// Positional body picture wrappers referencing `pictures`.
    picture_compatibility_records: Vec<crate::PictureCompatibilityRecord>,
    /// Extracted fields
    fields: Vec<super::super::field::Field<'a>>,
    /// Ordered positional legacy form fields.
    form_fields: Vec<super::super::form_field::FormField<'a>>,
    /// Inert producer provenance from the generator destination.
    generator: Option<crate::DocumentGenerator<'a>>,
    /// Ordered revision-save/session provenance.
    revision_save: Option<crate::RevisionSaveMetadata>,
    /// Ordered inert XML namespace table; `Some([])` preserves an empty table.
    xml_namespaces: Option<Vec<crate::XmlNamespace<'a>>>,
    /// Ordered inert custom XML markup tags spanning body text.
    custom_xml_tags: Vec<crate::CustomXmlTag<'a>>,
    /// Ordered inert math zones anchored in the body story.
    math_zones: Vec<crate::MathZone<'a>>,
    /// Ordered inert protection-exception ranges spanning body text.
    protection_ranges: Vec<crate::ProtectionRange<'a>>,
    /// Ordered inert editable regions spanning body text.
    editable_regions: Vec<crate::EditableRegion<'a>>,
    /// Ordered inert usernames used by range-level protection.
    protection_user_table: Option<crate::ProtectionUserTable<'a>>,
    /// Explicit passive document hyphenation properties.
    hyphenation: crate::DocumentHyphenation,
    /// Inert names from `nextfile` and `template` document destinations.
    external_references: crate::DocumentExternalReferences<'a>,
    /// Passive document view and zoom metadata.
    document_view: crate::DocumentView,
    /// Passive review-display suppression preferences.
    review_display: crate::DocumentReviewDisplay,
    /// Inert document-window caption text.
    window_caption: Option<crate::DocumentWindowCaption<'a>>,
    /// Inert custom kinsoku character sets and their language.
    kinsoku: crate::DocumentKinsoku<'a>,
    /// Inert custom XSL transform location metadata.
    xsl_transform: Option<crate::DocumentXslTransform<'a>>,
    /// Passive requested intent from the `usexform` flag.
    xsl_transform_usage: crate::DocumentXslTransformUsage,
    /// Passive suggested filters for a host application's style list.
    style_list_filter: Option<crate::DocumentStyleListFilter>,
    /// Passive suggested sorting for a host application's style list.
    style_sort_method: Option<crate::DocumentStyleSortMethod>,
    /// Passive read-only and thumbnail-generation save preferences.
    save_preferences: crate::DocumentSavePreferences,
    /// Opaque, inert write-reservation metadata.
    write_reservations: crate::DocumentWriteReservations<'a>,
    /// Passive source and AutoFormat classification metadata.
    origin_metadata: crate::DocumentOriginMetadata,
    /// Passive backup, storage-format, and template flags.
    file_settings: crate::DocumentFileSettings,
    /// Passive compatibility and output-request flags.
    output_settings: crate::DocumentOutputSettings,
    /// Passive document rendering flags.
    rendering_settings: crate::DocumentRenderingSettings,
    /// Passive printing, cleanup, and event-mask properties.
    processing_settings: crate::DocumentProcessingSettings,
    /// Passive document-level drawing-grid properties.
    drawing_grid: crate::DocumentDrawingGrid,
    /// Passive document print-layout settings.
    print_layout_settings: crate::DocumentPrintLayoutSettings,
    /// Passive theme font-resolution language identifiers.
    theme_languages: crate::DocumentThemeLanguages,
    /// Passive web-save and custom-XML policies.
    xml_policies: crate::DocumentXmlPolicies,
    /// Passive system-font and linguistic-data embedding policies.
    embedding_policies: crate::DocumentEmbeddingPolicies,
    /// Passive move and formatting revision policies.
    revision_policies: crate::DocumentRevisionPolicies,
    /// Passive theme and style-application policies.
    style_policies: crate::DocumentStylePolicies,
    /// Passive legacy style and formatting restriction declarations.
    style_restrictions: crate::DocumentStyleRestrictions,
    /// Passive booklet-printing requests.
    booklet_printing: crate::DocumentBookletPrinting,
    /// Passive privacy-removal requests.
    privacy_policies: crate::DocumentPrivacyPolicies,
    /// Passive legacy extra-line-spacing compatibility requests.
    line_spacing_compatibility: crate::DocumentLineSpacingCompatibility,
    /// Passive Word 6-era East Asian typography compatibility requests.
    east_asian_compatibility: crate::DocumentEastAsianCompatibility,
    /// Passive legacy table-layout compatibility requests.
    table_layout_compatibility: crate::DocumentTableLayoutCompatibility,
    /// Passive legacy automatic-layout compatibility requests.
    legacy_layout_compatibility: crate::DocumentLegacyLayoutCompatibility,
    /// Passive Asian character-grid and line-breaking compatibility requests.
    asian_grid_compatibility: crate::DocumentAsianGridCompatibility,
    /// Passive compatibility reset, UI-throttling, and upgrade requests.
    compatibility_policy: crate::DocumentCompatibilityPolicy,
    /// Passive Word 2003-era compatibility requests.
    word_2003_compatibility: crate::DocumentWord2003Compatibility,
    /// Inert Office theme package and optional color-scheme mapping bytes.
    theme: Option<crate::DocumentTheme<'a>>,
    /// Inert latent-style defaults and ordered exceptions.
    latent_styles: Option<crate::LatentStyles<'a>>,
    /// Inert custom XML data-store bytes.
    data_store: Option<crate::DocumentDataStore<'a>>,
    /// Inert mail-merge connection, query, and data-source metadata.
    mail_merge: Option<crate::MailMerge<'a>>,
    /// Document-level defaults for mathematical layout.
    math_properties: Option<crate::DocumentMathProperties>,
    /// Default primary, East Asian, and complex-script languages.
    language_defaults: crate::DocumentLanguageDefaults,
    /// Passive root default-font selectors and default property destinations.
    default_formatting: crate::DocumentDefaultFormatting,
    /// Explicit source `deftab` width; omission is preserved as `None`.
    default_tab_width_twips: Option<u32>,
    /// Explicit default bidirectional precedence for document text.
    document_direction: Option<crate::TextDirection>,
    /// Whether the document gutter is positioned on the right.
    gutter_on_right: bool,
    /// Embedded and linked objects
    objects: Vec<super::super::object::EmbeddedObject<'a>>,
    /// Ordered inert document-variable metadata
    document_variables: Vec<super::super::document_variable::DocumentVariable<'a>>,
    /// Ordered inert user-defined document properties
    user_properties: Vec<super::super::user_property::UserProperty<'a>>,
    /// Ordered inert index and table-of-contents source marks
    navigation_entries: Vec<super::super::navigation_entry::NavigationEntry<'a>>,
    /// Ordered inert generated list-marker destinations.
    generated_list_markers: Vec<crate::GeneratedListMarker<'a>>,
    /// List table
    list_table: super::super::list::ListTable<'a>,
    /// List override table
    list_override_table: super::super::list::ListOverrideTable,
    /// Ordered inert legacy section-numbering defaults.
    legacy_section_numbering: crate::LegacySectionNumbering<'a>,
    /// Ordered inert legacy paragraph-numbering records referenced by paragraphs.
    legacy_paragraph_numbering: Vec<crate::LegacyParagraphNumbering<'a>>,
    /// Optional paragraph-group property table.
    paragraph_group_table: Option<crate::ParagraphGroupPropertyTable>,
    /// Sections
    sections: Vec<super::super::section::Section<'a>>,
    /// Bookmarks
    bookmarks: super::super::bookmark::BookmarkTable<'a>,
    /// Shapes
    pub(super) shapes: Vec<super::super::shape::Shape<'a>>,
    /// Exact source order of non-background root drawings in the body story.
    pub(super) drawing_order: Vec<crate::StoryDrawing>,
    /// Structural paragraph and line boundaries in flattened body text.
    body_boundaries: Vec<crate::story::Boundary>,
    pub(super) body_story_events: Vec<crate::BodyStoryEvent>,
    /// Index in `shapes` owned by the unique document-background destination.
    pub(super) background_shape_index: Option<usize>,
    /// Inert positional legacy drawing text boxes.
    legacy_text_boxes: Vec<crate::LegacyTextBox<'a>>,
    /// Inert positional legacy drawing primitives.
    legacy_drawings: Vec<crate::LegacyDrawing<'a>>,
    /// Shape groups
    shape_groups: Vec<super::super::shape::ShapeGroup<'a>>,
    /// Stylesheet
    stylesheet: super::super::stylesheet::StyleSheet<'a>,
    /// Document information
    info: super::super::info::DocumentInfo<'a>,
    /// Annotations
    annotations: Vec<super::super::annotation::Annotation<'a>>,
    /// Footnotes and endnotes
    notes: Vec<super::super::section::Note<'a>>,
    /// Explicit document-level footnote and endnote configuration.
    note_options: crate::NoteOptions,
    /// Ordered inert footnote/endnote separator destinations.
    note_separators: crate::NoteSeparatorTable<'a>,
    /// Track changes/revisions
    revisions: Vec<super::super::annotation::Revision<'a>>,
    /// Ordered revision-author table referenced by revision author indices.
    revision_authors: Vec<super::super::annotation::RevisionAuthor<'a>>,
}

impl<'a> RtfDocument<'a> {
    /// Parse an RTF document from a string.
    ///
    /// This method automatically detects and decompresses compressed RTF data.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::raw::Document;
    ///
    /// let rtf = r#"{\rtf1\ansi Hello World!\par}"#;
    /// let doc = Document::parse(rtf)?;
    /// let text = doc.text();
    /// # Ok::<(), litchi_rtf::Error>(())
    /// ```
    pub fn parse(input: &str) -> RtfResult<RtfDocument<'static>> {
        Self::parse_with_limits(input, ParseLimits::default())
    }

    /// Parse an RTF string with an explicit finite resource profile.
    pub fn parse_with_limits(input: &str, limits: ParseLimits) -> RtfResult<RtfDocument<'static>> {
        Self::parse_internal(input.as_bytes(), limits)
    }

    /// Parse RTF from its original byte representation.
    ///
    /// Use this entry point when the document can contain `bin` destinations or
    /// legacy-code-page bytes that are not valid UTF-8.
    pub fn parse_bytes(input: &[u8]) -> RtfResult<RtfDocument<'static>> {
        Self::parse_bytes_with_limits(input, ParseLimits::default())
    }

    /// Parse original RTF bytes with an explicit finite resource profile.
    pub fn parse_bytes_with_limits(
        input: &[u8],
        limits: ParseLimits,
    ) -> RtfResult<RtfDocument<'static>> {
        Self::parse_internal(input, limits)
    }

    /// Parse RTF from bytes (handles both compressed and uncompressed)
    fn parse_internal(bytes: &[u8], limits: ParseLimits) -> RtfResult<RtfDocument<'static>> {
        if bytes.len() > limits.max_source_bytes() {
            return Err(RtfError::LimitExceeded {
                resource: "source bytes",
                observed: bytes.len(),
                limit: limits.max_source_bytes(),
            });
        }

        // Check if it's compressed RTF
        let input_bytes = if super::super::compressed::is_compressed_rtf(bytes) {
            // Decompress first
            Cow::Owned(super::super::compressed::decompress_with_limits(
                bytes,
                super::super::compressed::DecompressionLimits::new(limits.max_decompressed_bytes()),
            )?)
        } else {
            Cow::Borrowed(bytes)
        };

        // RTF files are NOT UTF-8. They contain bytes in whatever code page is
        // specified by \ansicpg (e.g., Windows-1252, GB2312, etc.).
        //
        // We use Latin-1 (ISO-8859-1) encoding for initial parsing because:
        // 1. It provides 1:1 byte-to-character mapping (byte 0xNN -> U+00NN)
        // 2. Control words (ASCII) parse correctly
        // 3. We can recover original bytes and decode them with correct encoding later
        //
        // The parser will detect \ansicpg and use the proper encoding for text.
        let input_str = if input_bytes.is_ascii() {
            Cow::Borrowed(std::str::from_utf8(input_bytes.as_ref()).map_err(|error| {
                RtfError::InvalidUnicode(format!("ASCII RTF transport conversion failed: {error}"))
            })?)
        } else {
            Cow::Owned(input_bytes.iter().map(|byte| char::from(*byte)).collect())
        };

        Self::parse_string(input_str.as_ref(), limits)
    }

    /// Parse an RTF document from a UTF-8 string (internal)
    fn parse_string(input: &str, limits: ParseLimits) -> RtfResult<RtfDocument<'static>> {
        // Create arena for temporary allocations during parsing
        let arena = Bump::new();

        // Lexer phase
        let mut lexer = Lexer::new_with_limits(input, &arena, limits);
        let tokens = lexer.tokenize()?;

        // Parser phase
        let parser = Parser::new(&tokens, &arena);
        let parsed = parser.parse()?;

        // Convert parsed document to owned document
        // We need to convert Cow::Borrowed to Cow::Owned to detach from input lifetime
        let owned_blocks: Vec<StyleBlock<'static>> = parsed
            .blocks
            .into_iter()
            .map(|block| StyleBlock {
                text: Cow::Owned(block.text.into_owned()),
                formatting: block.formatting,
                paragraph: block.paragraph,
            })
            .collect();

        // Convert font table to owned
        let owned_font_table = parsed.font_table.into_owned();

        // Convert tables to owned
        let owned_tables: Vec<super::super::table::Table<'static>> = parsed
            .tables
            .into_iter()
            .map(|table| owned_table(&table))
            .collect::<RtfResult<_>>()?;

        // Convert pictures to owned
        let owned_pictures: Vec<super::super::picture::Picture<'static>> = parsed
            .pictures
            .into_iter()
            .map(super::super::picture::Picture::into_owned)
            .collect();

        // Convert fields to owned
        let owned_fields: Vec<super::super::field::Field<'static>> = parsed
            .fields
            .into_iter()
            .map(|field| super::super::field::Field {
                field_type: field.field_type,
                instruction: Cow::Owned(field.instruction.into_owned()),
                result: Cow::Owned(field.result.into_owned()),
                status: field.status,
                shapes: field
                    .shapes
                    .into_iter()
                    .map(crate::Shape::into_owned)
                    .collect(),
                shape_groups: field
                    .shape_groups
                    .into_iter()
                    .map(crate::ShapeGroup::into_owned)
                    .collect(),
                drawing_order: field.drawing_order,
                result_events: field.result_events,
                owner: field.owner,
                position: field.position,
                range_end: field.range_end,
            })
            .collect();

        let owned_objects = parsed
            .objects
            .into_iter()
            .map(|object| super::super::object::EmbeddedObject {
                position: object.position,
                kind: object.kind,
                link_self: object.link_self,
                class_name: Cow::Owned(object.class_name.into_owned()),
                name: Cow::Owned(object.name.into_owned()),
                alias: object.alias.map(|alias| Cow::Owned(alias.into_owned())),
                section: object
                    .section
                    .map(|section| Cow::Owned(section.into_owned())),
                time: object.time,
                class_id: Cow::Owned(object.class_id.into_owned()),
                width: object.width,
                height: object.height,
                alignment: object.alignment,
                translation_y: object.translation_y,
                crop_top: object.crop_top,
                crop_bottom: object.crop_bottom,
                crop_left: object.crop_left,
                crop_right: object.crop_right,
                scale_x: object.scale_x,
                scale_y: object.scale_y,
                locked: object.locked,
                update_requested: object.update_requested,
                set_size: object.set_size,
                merge_result: object.merge_result,
                result_kind: object.result_kind,
                result_text: Cow::Owned(object.result_text.into_owned()),
                result_picture_indices: object.result_picture_indices,
                data: object.data,
            })
            .collect();

        // Convert all borrowed data to owned
        Ok(RtfDocument {
            font_table: owned_font_table,
            file_table: parsed.file_table.map(crate::FileTable::into_owned),
            color_table: parsed.color_table,
            blocks: owned_blocks,
            tables: owned_tables,
            pictures: owned_pictures,
            picture_compatibility_records: parsed.picture_compatibility_records,
            fields: owned_fields,
            form_fields: parsed
                .form_fields
                .into_iter()
                .map(super::super::form_field::FormField::into_owned)
                .collect(),
            generator: parsed.generator.map(crate::DocumentGenerator::into_owned),
            revision_save: parsed.revision_save,
            xml_namespaces: parsed.saw_xml_namespace_table.then(|| {
                parsed
                    .xml_namespaces
                    .into_iter()
                    .map(crate::XmlNamespace::into_owned)
                    .collect()
            }),
            custom_xml_tags: parsed
                .custom_xml_tags
                .into_iter()
                .map(crate::CustomXmlTag::into_owned)
                .collect(),
            math_zones: parsed
                .math_zones
                .into_iter()
                .map(crate::MathZone::into_owned)
                .collect(),
            protection_ranges: parsed
                .protection_ranges
                .into_iter()
                .map(crate::ProtectionRange::into_owned)
                .collect(),
            editable_regions: parsed
                .editable_regions
                .into_iter()
                .map(crate::EditableRegion::into_owned)
                .collect(),
            protection_user_table: parsed
                .protection_user_table
                .map(crate::ProtectionUserTable::into_owned),
            hyphenation: parsed.hyphenation,
            external_references: parsed.external_references.into_owned(),
            document_view: parsed.document_view,
            review_display: parsed.review_display,
            window_caption: parsed
                .window_caption
                .map(crate::DocumentWindowCaption::into_owned),
            kinsoku: parsed.kinsoku.into_owned(),
            xsl_transform: parsed
                .xsl_transform
                .map(crate::DocumentXslTransform::into_owned),
            xsl_transform_usage: parsed.xsl_transform_usage,
            style_list_filter: parsed.style_list_filter,
            style_sort_method: parsed.style_sort_method,
            save_preferences: parsed.save_preferences,
            write_reservations: parsed.write_reservations.into_owned(),
            origin_metadata: parsed.origin_metadata,
            file_settings: parsed.file_settings,
            output_settings: parsed.output_settings,
            rendering_settings: parsed.rendering_settings,
            processing_settings: parsed.processing_settings,
            drawing_grid: parsed.drawing_grid,
            print_layout_settings: parsed.print_layout_settings,
            theme_languages: parsed.theme_languages,
            xml_policies: parsed.xml_policies,
            embedding_policies: parsed.embedding_policies,
            revision_policies: parsed.revision_policies,
            style_policies: parsed.style_policies,
            style_restrictions: parsed.style_restrictions,
            booklet_printing: parsed.booklet_printing,
            privacy_policies: parsed.privacy_policies,
            line_spacing_compatibility: parsed.line_spacing_compatibility,
            east_asian_compatibility: parsed.east_asian_compatibility,
            table_layout_compatibility: parsed.table_layout_compatibility,
            legacy_layout_compatibility: parsed.legacy_layout_compatibility,
            asian_grid_compatibility: parsed.asian_grid_compatibility,
            compatibility_policy: parsed.compatibility_policy,
            word_2003_compatibility: parsed.word_2003_compatibility,
            theme: parsed.theme.map(crate::DocumentTheme::into_owned),
            latent_styles: parsed.latent_styles.map(crate::LatentStyles::into_owned),
            data_store: parsed.data_store.map(crate::DocumentDataStore::into_owned),
            mail_merge: parsed.mail_merge.map(crate::MailMerge::into_owned),
            math_properties: parsed.math_properties,
            language_defaults: parsed.language_defaults,
            default_formatting: parsed.default_formatting,
            default_tab_width_twips: parsed.default_tab_width_twips,
            document_direction: parsed.document_direction,
            gutter_on_right: parsed.gutter_on_right,
            objects: owned_objects,
            document_variables: parsed
                .document_variables
                .into_iter()
                .map(super::super::document_variable::DocumentVariable::into_owned)
                .collect(),
            user_properties: parsed
                .user_properties
                .into_iter()
                .map(super::super::user_property::UserProperty::into_owned)
                .collect(),
            navigation_entries: parsed
                .navigation_entries
                .into_iter()
                .map(super::super::navigation_entry::NavigationEntry::into_owned)
                .collect(),
            generated_list_markers: parsed
                .generated_list_markers
                .into_iter()
                .map(crate::GeneratedListMarker::into_owned)
                .collect(),
            list_table: Self::convert_list_table_to_owned(parsed.list_table)?,
            list_override_table: parsed.list_override_table,
            legacy_section_numbering: parsed.legacy_section_numbering.into_owned(),
            legacy_paragraph_numbering: parsed
                .legacy_paragraph_numbering
                .into_iter()
                .map(crate::LegacyParagraphNumbering::into_owned)
                .collect(),
            paragraph_group_table: parsed
                .paragraph_group_table
                .map(crate::ParagraphGroupPropertyTable::into_owned),
            sections: Self::convert_sections_to_owned(parsed.sections),
            bookmarks: Self::convert_bookmarks_to_owned(parsed.bookmarks),
            shapes: Self::convert_shapes_to_owned(parsed.shapes),
            drawing_order: parsed.drawing_order,
            body_boundaries: parsed.body_boundaries,
            body_story_events: parsed.body_story_events,
            background_shape_index: parsed.background_shape_index,
            legacy_text_boxes: parsed
                .legacy_text_boxes
                .into_iter()
                .map(crate::LegacyTextBox::into_owned)
                .collect(),
            legacy_drawings: parsed
                .legacy_drawings
                .into_iter()
                .map(crate::LegacyDrawing::into_owned)
                .collect(),
            shape_groups: Self::convert_shape_groups_to_owned(parsed.shape_groups),
            stylesheet: Self::convert_stylesheet_to_owned(parsed.stylesheet),
            info: Self::convert_info_to_owned(parsed.info),
            annotations: Self::convert_annotations_to_owned(parsed.annotations),
            notes: Self::convert_notes_to_owned(parsed.notes),
            note_options: parsed.note_options,
            note_separators: parsed.note_separators.into_owned(),
            revisions: Self::convert_revisions_to_owned(parsed.revisions),
            revision_authors: parsed
                .revision_authors
                .into_iter()
                .map(super::super::annotation::RevisionAuthor::into_owned)
                .collect(),
        })
    }

    /// Parse an RTF document from a file.
    ///
    /// This method automatically detects and handles compressed RTF files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::raw::Document;
    ///
    /// let doc = Document::open("document.rtf")?;
    /// let text = doc.text();
    /// # Ok::<(), litchi_rtf::Error>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> RtfResult<RtfDocument<'static>> {
        Self::open_with_limits(path, ParseLimits::default())
    }

    /// Open an RTF file with an explicit finite resource profile.
    pub fn open_with_limits<P: AsRef<Path>>(
        path: P,
        limits: ParseLimits,
    ) -> RtfResult<RtfDocument<'static>> {
        let bytes = read_file_with_limit(path.as_ref(), limits.max_source_bytes())?;
        Self::parse_internal(&bytes, limits)
    }

    /// Parse an RTF document from bytes.
    ///
    /// This method automatically detects and decompresses compressed RTF data.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::raw::Document;
    ///
    /// let bytes = std::fs::read("document.rtf").map_err(|e| format!("IO error: {}", e))?;
    /// let doc = Document::from_bytes(&bytes)?;
    /// let text = doc.text();
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> RtfResult<RtfDocument<'static>> {
        Self::parse_bytes(bytes)
    }

    /// Parse RTF bytes with an explicit finite resource profile.
    pub fn from_bytes_with_limits(
        bytes: &[u8],
        limits: ParseLimits,
    ) -> RtfResult<RtfDocument<'static>> {
        Self::parse_bytes_with_limits(bytes, limits)
    }

    /// Get all text content from the document.
    ///
    /// This concatenates all text blocks with their natural separators.
    pub fn text(&self) -> String {
        self.blocks
            .iter()
            .map(|block| block.text.as_ref())
            .collect::<Vec<&str>>()
            .join("")
    }

    /// Get the number of paragraphs in the document.
    ///
    /// Paragraphs are determined by paragraph breaks in the RTF source.
    pub fn paragraph_count(&self) -> usize {
        self.paragraphs().len()
    }

    /// Get all paragraphs in the document.
    ///
    /// This groups style blocks into paragraphs based on newline characters.
    pub fn paragraphs(&self) -> Vec<RtfParagraph> {
        let mut paragraphs = Vec::new();
        let mut current_para = RtfParagraph::default();
        let mut has_content = false;

        for block in &self.blocks {
            let text = block.text.as_ref();

            // Split on newlines to detect paragraph boundaries
            let parts: Vec<&str> = text.split('\n').collect();

            for (i, part) in parts.iter().enumerate() {
                if !part.is_empty() {
                    // Inherit paragraph properties from the style block
                    current_para = block.paragraph;
                    has_content = true;
                }

                // If this is not the last part, we have a paragraph break
                if i < parts.len() - 1 && has_content {
                    paragraphs.push(current_para);
                    current_para = RtfParagraph::default();
                    has_content = false;
                }
            }
        }

        // Add final paragraph if it has content
        if has_content {
            paragraphs.push(current_para);
        }

        paragraphs
    }

    /// Get all paragraphs with their content (runs).
    ///
    /// This groups style blocks into paragraphs based on newline characters,
    /// and returns each paragraph with its associated runs.
    pub fn paragraphs_with_content(&self) -> Vec<super::super::types::ParagraphContent<'_>> {
        use std::borrow::Cow;

        let mut paragraphs = Vec::new();
        let mut current_para_props = RtfParagraph::default();
        let mut current_runs: Vec<Run<'_>> = Vec::new();
        let mut has_content = false;

        for block in &self.blocks {
            let text = block.text.as_ref();

            // Split on newlines to detect paragraph boundaries
            let parts: Vec<&str> = text.split('\n').collect();

            for (i, part) in parts.iter().enumerate() {
                if !part.is_empty() {
                    // Inherit paragraph properties from the style block
                    current_para_props = block.paragraph;
                    has_content = true;

                    // Add run for this part
                    current_runs.push(Run::new(Cow::Borrowed(part), block.formatting));
                }

                // If this is not the last part, we have a paragraph break
                if i < parts.len() - 1 && has_content {
                    paragraphs.push(super::super::types::ParagraphContent::new(
                        current_para_props,
                        current_runs.clone(),
                    ));
                    current_runs.clear();
                    current_para_props = RtfParagraph::default();
                    has_content = false;
                }
            }
        }

        // Add final paragraph if it has content
        if has_content {
            paragraphs.push(super::super::types::ParagraphContent::new(
                current_para_props,
                current_runs,
            ));
        }

        paragraphs
    }

    /// Get all runs in the document.
    ///
    /// A run is a contiguous block of text with the same formatting.
    pub fn runs(&self) -> Vec<Run<'_>> {
        self.blocks
            .iter()
            .map(|block| Run::new(block.text.clone(), block.formatting))
            .collect()
    }

    /// Get all tables in the document.
    ///
    /// Returns all tables extracted from the RTF document.
    pub fn tables(&self) -> &[super::super::table::Table<'_>] {
        &self.tables
    }

    /// Mutably access document tables, including their nested cell stories.
    pub fn tables_mut(&mut self) -> &mut [super::super::table::Table<'a>] {
        &mut self.tables
    }

    fn table_cell_mut(&mut self, path: &crate::TableCellPath) -> RtfResult<&mut crate::Cell<'a>> {
        let root = path.root;
        let mut cell = self
            .tables
            .get_mut(root.table_index)
            .and_then(|table| table.rows_mut().get_mut(root.row_index))
            .and_then(|row| row.cells_mut().get_mut(root.cell_index))
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF table-cell path is outside the document".to_string(),
                )
            })?;
        for coordinate in &path.nested {
            cell = cell
                .nested_tables_mut()
                .get_mut(coordinate.table_index)
                .and_then(|nested| nested.table.rows_mut().get_mut(coordinate.row_index))
                .and_then(|row| row.cells_mut().get_mut(coordinate.cell_index))
                .ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF nested table-cell path is outside the document".to_string(),
                    )
                })?;
        }
        Ok(cell)
    }

    /// Get all document elements (paragraphs and tables) in approximate document order.
    ///
    /// Note: Due to RTF's structure, tables are extracted separately from paragraph flow.
    /// This method returns paragraphs first, followed by tables. For most use cases this
    /// is sufficient. If you need precise positional information, work with `blocks()` directly.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::raw::{Document, DocumentElement};
    ///
    /// let doc = Document::open("document.rtf")?;
    /// for element in doc.elements() {
    ///     match element {
    ///         DocumentElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text());
    ///         }
    ///         DocumentElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count());
    ///         }
    ///     }
    /// }
    /// # Ok::<(), litchi_rtf::Error>(())
    /// ```
    pub fn elements(&self) -> Vec<super::super::DocumentElement<'_>> {
        let mut elements = Vec::new();

        // Add all paragraphs first
        for para in self.paragraphs_with_content() {
            elements.push(super::super::DocumentElement::Paragraph(para));
        }

        // Add all tables
        for table in &self.tables {
            elements.push(super::super::DocumentElement::Table(table.clone()));
        }

        elements
    }

    /// Get the font table.
    pub fn font_table(&self) -> &FontTable<'a> {
        &self.font_table
    }

    /// Get the external-file metadata table, if present.
    pub fn file_table(&self) -> Option<&crate::FileTable<'_>> {
        self.file_table.as_ref()
    }

    /// Replace the external-file metadata table after validating it.
    pub fn set_file_table(&mut self, table: crate::FileTable<'a>) -> RtfResult<()> {
        table.validate()?;
        self.file_table = Some(table);
        Ok(())
    }

    /// Remove the external-file metadata table.
    pub fn clear_file_table(&mut self) {
        self.file_table = None;
    }

    /// Get the color table.
    pub fn color_table(&self) -> &ColorTable {
        &self.color_table
    }

    /// Get all style blocks.
    pub fn blocks(&self) -> &[StyleBlock<'_>] {
        &self.blocks
    }

    pub(crate) fn retained_blocks(&self) -> &[StyleBlock<'a>] {
        &self.blocks
    }

    /// Get all pictures in the document.
    ///
    /// Returns all embedded images extracted from the RTF document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::raw::Document;
    ///
    /// let doc = Document::open("document.rtf")?;
    /// for (i, picture) in doc.pictures().iter().enumerate() {
    ///     println!("Picture {}: {:?}, {} bytes", i, picture.image_type, picture.data().len());
    /// }
    /// # Ok::<(), litchi_rtf::Error>(())
    /// ```
    pub fn pictures(&self) -> &[super::super::picture::Picture<'_>] {
        &self.pictures
    }

    /// Replace typed `picprop` metadata on one existing picture without cloning image bytes.
    pub fn set_picture_shape_properties(
        &mut self,
        picture_index: usize,
        properties: Option<crate::PictureShapeProperties<'a>>,
    ) -> RtfResult<Option<crate::PictureShapeProperties<'a>>> {
        if let Some(properties) = &properties {
            properties.validate()?;
        }
        let picture = self.pictures.get_mut(picture_index).ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF picture shape-property mutation references a missing picture".to_string(),
            )
        })?;
        Ok(std::mem::replace(&mut picture.shape_properties, properties))
    }

    /// Return positional body `shppict` and `nonshppict` wrapper records.
    pub fn picture_compatibility_records(&self) -> &[crate::PictureCompatibilityRecord] {
        &self.picture_compatibility_records
    }

    /// Resolve a wrapper record to its shared picture payload.
    pub fn picture_for_compatibility_record(
        &self,
        record: &crate::PictureCompatibilityRecord,
    ) -> Option<&super::super::picture::Picture<'_>> {
        self.pictures.get(record.picture_index)
    }

    /// Append a validated positional wrapper without cloning picture bytes.
    pub fn push_picture_compatibility_record(
        &mut self,
        record: crate::PictureCompatibilityRecord,
    ) -> RtfResult<()> {
        let body = self.text();
        record.validate(&body, self.pictures.len())?;
        if self.picture_compatibility_records.len() >= crate::MAX_PICTURE_COMPATIBILITY_RECORDS {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility record count exceeds the safety limit".to_string(),
            ));
        }
        if self
            .picture_compatibility_records
            .last()
            .is_some_and(|previous| {
                previous.position > record.position
                    || (previous.position == record.position && previous.kind == record.kind)
            })
        {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility records are duplicated or out of body order".to_string(),
            ));
        }
        let index = self.picture_compatibility_records.len();
        self.picture_compatibility_records.push(record);
        self.insert_body_story_event(crate::BodyStoryEvent::PictureCompatibility(index))?;
        Ok(())
    }

    /// Clear wrapper provenance without deleting shared picture payloads.
    pub fn clear_picture_compatibility_records(&mut self) {
        self.picture_compatibility_records.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::PictureCompatibility(_)));
    }

    /// Get all fields in the document.
    ///
    /// Returns all fields (hyperlinks, cross-references, etc.) from the RTF document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::{field::FieldType, raw::Document};
    ///
    /// let doc = Document::open("document.rtf")?;
    /// for field in doc.fields() {
    ///     if field.field_type == FieldType::Hyperlink {
    ///         if let Some(url) = field.extract_url() {
    ///             println!("Hyperlink: {}", url);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), litchi_rtf::Error>(())
    /// ```
    pub fn fields(&self) -> &[super::super::field::Field<'_>] {
        &self.fields
    }

    /// Return typed, inert `HYPERLINK` fields in document field order.
    ///
    /// Targets, bookmarks, display options, cached results, and state are
    /// stored metadata only. This method never resolves, opens, fetches,
    /// validates, or activates a target; changes the insertion point; or
    /// refreshes a field.
    pub fn hyperlinks(&self) -> Vec<crate::HyperlinkField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::hyperlink)
            .collect()
    }

    /// Return the number of typed, inert `HYPERLINK` fields in the document.
    pub fn hyperlink_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.hyperlink().is_some())
            .count()
    }

    /// Return typed, inert `REF`, `PAGEREF`, and `NOTEREF` fields in document
    /// field order.
    ///
    /// Bookmark names, switches, cached results, and state are stored metadata
    /// only. This method never looks up a bookmark, calculates a page or note
    /// number, inserts text, changes layout, or refreshes a field.
    pub fn reference_fields(&self) -> Vec<crate::ReferenceField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::reference_field)
            .collect()
    }

    /// Return the number of typed, inert cross-reference fields in the document.
    pub fn reference_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.reference_field().is_some())
            .count()
    }

    /// Return typed, inert legacy `EQ` fields in document field order.
    ///
    /// Expressions and cached results are exposed as stored metadata only. This
    /// method never parses, calculates, formats, or renders equations.
    pub fn equations(&self) -> Vec<crate::EquationField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::equation)
            .collect()
    }

    /// Return the number of typed inert `EQ` fields in the document.
    pub fn equation_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.equation().is_some())
            .count()
    }

    /// Return typed, inert `MACROBUTTON` fields in document field order.
    ///
    /// Macro names and display text are exposed as stored metadata only. This
    /// method never resolves, loads, invokes, or executes a macro.
    pub fn macro_buttons(&self) -> Vec<crate::MacroButtonField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::macro_button)
            .collect()
    }

    /// Return the number of typed, inert `MACROBUTTON` fields in the document.
    pub fn macro_button_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.macro_button().is_some())
            .count()
    }

    /// Return typed, inert `GOTOBUTTON` fields in document field order.
    ///
    /// Destinations, button text, cached results, and state are stored metadata
    /// only. This method never resolves a destination, changes the insertion
    /// point, activates a jump, or refreshes a field.
    pub fn go_to_buttons(&self) -> Vec<crate::GoToButtonField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::go_to_button)
            .collect()
    }

    /// Return the number of typed, inert `GOTOBUTTON` fields in the document.
    pub fn go_to_button_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.go_to_button().is_some())
            .count()
    }

    /// Return typed, inert `PRINT` fields in document field order.
    ///
    /// Stored printer-instruction text, cached results, and field state are
    /// opaque metadata only. This method never interprets control codes, opens
    /// a printer, sends output, changes print settings, or refreshes a field.
    pub fn print_fields(&self) -> Vec<crate::PrintField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::print_field)
            .collect()
    }

    /// Return the number of typed, inert `PRINT` fields in the document.
    pub fn print_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.print_field().is_some())
            .count()
    }

    /// Return typed, inert `EMBED` fields in document field order.
    ///
    /// Stored opaque object instructions, cached results, and state are exposed
    /// solely as metadata. This method never loads, inspects, deserializes,
    /// activates, renders, or executes an embedded object, accesses an external
    /// resource, or refreshes a field.
    pub fn embed_fields(&self) -> Vec<crate::EmbedField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::embed_field)
            .collect()
    }

    /// Return the number of typed, inert `EMBED` fields in the document.
    pub fn embed_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.embed_field().is_some())
            .count()
    }

    /// Return typed, inert `BARCODE` fields in document field order.
    ///
    /// Stored opaque barcode instructions, cached results, and state are
    /// metadata only. This method never parses or validates barcode data or
    /// symbology, generates or renders a barcode, accesses an external
    /// resource, or refreshes a field.
    pub fn barcode_fields(&self) -> Vec<crate::BarcodeField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::barcode_field)
            .collect()
    }

    /// Return the number of typed, inert `BARCODE` fields in the document.
    pub fn barcode_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.barcode_field().is_some())
            .count()
    }

    /// Return typed, inert `DISPLAYBARCODE` and `MERGEBARCODE` fields in document
    /// field order.
    ///
    /// Stored data arguments, barcode types, switches, cached results, and
    /// state are metadata only. This method never validates barcode data or
    /// symbology; resolves a mail-merge data field; generates or renders a
    /// barcode; or refreshes a field.
    pub fn barcode_display_fields(&self) -> Vec<crate::BarcodeDisplayField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::barcode_display_field)
            .collect()
    }

    /// Return the number of typed, inert barcode display fields in the document.
    pub fn barcode_display_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.barcode_display_field().is_some())
            .count()
    }

    /// Return typed, inert `BIDIOUTLINE` fields in document field order.
    ///
    /// Stored opaque instructions, cached results, and state are metadata only.
    /// This method never reads right-to-left language, paragraph outline, or
    /// layout state; chooses a numbering system; calculates a result; or
    /// refreshes a field.
    pub fn bidi_outline_fields(&self) -> Vec<crate::BidiOutlineField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::bidi_outline_field)
            .collect()
    }

    /// Return the number of typed, inert `BIDIOUTLINE` fields in the document.
    pub fn bidi_outline_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.bidi_outline_field().is_some())
            .count()
    }

    /// Return typed, inert `SHAPE` drawing-canvas anchor fields in document field order.
    ///
    /// Stored opaque instructions, cached results, and state are metadata only.
    /// This method never locates, links, loads, positions, lays out, or renders
    /// a drawing or canvas, or refreshes a field.
    pub fn shape_fields(&self) -> Vec<crate::ShapeField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::shape_field)
            .collect()
    }

    /// Return the number of typed, inert `SHAPE` drawing-canvas anchor fields.
    pub fn shape_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.shape_field().is_some())
            .count()
    }

    /// Return typed, inert legacy form-code fields in document field order.
    ///
    /// Stored text/checkbox/drop-down kind, opaque instructions, cached
    /// results, and state are metadata only. This method does not read or
    /// reconcile the separate RTF `\formfield` destination. It never fills a form,
    /// changes a selection or checkbox state, invokes entry or exit macros, or
    /// refreshes a field.
    pub fn legacy_form_fields(&self) -> Vec<crate::LegacyFormField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::legacy_form_field)
            .collect()
    }

    /// Return the number of typed, inert legacy form-code fields.
    pub fn legacy_form_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.legacy_form_field().is_some())
            .count()
    }

    /// Return typed, inert `PRIVATE` conversion-data fields in document field order.
    ///
    /// Stored opaque instructions, cached results, and state are metadata only.
    /// This method never converts a document, interprets or reveals hidden
    /// content, changes layout, or refreshes a field. `PRIVATE` is not treated
    /// as a confidentiality mechanism.
    pub fn private_fields(&self) -> Vec<crate::PrivateField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::private_field)
            .collect()
    }

    /// Return the number of typed, inert `PRIVATE` conversion-data fields.
    pub fn private_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.private_field().is_some())
            .count()
    }

    /// Return typed, inert `ADDIN`, `CONTROL`, and `HTMLCONTROL` fields in
    /// document field order.
    ///
    /// Stored instructions, cached results, and state are opaque metadata only.
    /// This method never loads an add-in, instantiates a control, invokes code,
    /// executes script, renders content, accesses an external resource, or
    /// refreshes a field.
    pub fn active_content_fields(&self) -> Vec<crate::ActiveContentField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::active_content_field)
            .collect()
    }

    /// Return the number of typed, inert active-content fields in the document.
    pub fn active_content_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.active_content_field().is_some())
            .count()
    }

    /// Return typed, inert `GLOSSARY` and `AUTOTEXT` fields in document field order.
    ///
    /// Stored entry names, switches, cached results, and state are metadata
    /// only. This method never looks up a building block, reads a template,
    /// inserts content, changes bookmarks, accesses an external resource, or
    /// refreshes a field.
    pub fn auto_text_fields(&self) -> Vec<crate::AutoTextField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::auto_text_field)
            .collect()
    }

    /// Return the number of typed, inert building-block fields in the document.
    pub fn auto_text_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.auto_text_field().is_some())
            .count()
    }

    /// Return typed, inert `AUTOTEXTLIST` fields in document field order.
    ///
    /// Stored display text, style/tip options, switches, cached results, and
    /// state are metadata only. This method never shows a selection UI, looks
    /// up eligible building blocks, reads a template, inserts content, changes
    /// bookmarks, accesses an external resource, or refreshes a field.
    pub fn auto_text_list_fields(&self) -> Vec<crate::AutoTextListField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::auto_text_list_field)
            .collect()
    }

    /// Return the number of typed, inert `AUTOTEXTLIST` fields in the document.
    pub fn auto_text_list_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.auto_text_list_field().is_some())
            .count()
    }

    /// Return typed, inert `DDE` and `DDEAUTO` fields in document field order.
    ///
    /// Application, source, item, representation, and storage metadata are
    /// exposed solely as stored text. This method never launches an
    /// application, initiates a DDE conversation, opens a source, requests
    /// data, refreshes, evaluates, converts, or executes anything.
    pub fn dde_links(&self) -> Vec<crate::DdeField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::dde_link)
            .collect()
    }

    /// Return the number of typed, inert `DDE` and `DDEAUTO` fields.
    pub fn dde_link_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.dde_link().is_some())
            .count()
    }

    /// Return typed, inert `LINK` fields in document field order.
    ///
    /// Application type, source, item, result, and formatting metadata are
    /// exposed solely as stored text. This method never activates an OLE
    /// server, launches an application, opens a source, requests data,
    /// refreshes, evaluates, converts, or executes anything.
    pub fn link_fields(&self) -> Vec<crate::LinkField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::link_field)
            .collect()
    }

    /// Return the number of typed, inert `LINK` fields.
    pub fn link_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.link_field().is_some())
            .count()
    }

    /// Return typed, inert external include fields in document field order.
    ///
    /// Sources, converter names, and XML options are exposed as stored metadata
    /// only. This method never resolves, opens, fetches, transforms, converts,
    /// updates, or writes to an external source.
    pub fn external_includes(&self) -> Vec<crate::ExternalIncludeField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::external_include)
            .collect()
    }

    /// Return the number of typed, inert external include fields in the document.
    pub fn external_include_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.external_include().is_some())
            .count()
    }

    /// Return typed, inert `RD` referenced-document fields in document field order.
    ///
    /// Stored paths, relative-path requests, switches, cached results, and
    /// field state are metadata only. This method never opens, resolves, reads,
    /// imports, refreshes, evaluates, or executes a referenced document.
    pub fn referenced_documents(&self) -> Vec<crate::ReferencedDocumentField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::referenced_document)
            .collect()
    }

    /// Return the number of typed, inert `RD` referenced-document fields.
    pub fn referenced_document_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.referenced_document().is_some())
            .count()
    }

    /// Return typed, inert CITATION fields in document field order.
    ///
    /// Source tags, switches, and cached results are exposed solely as stored
    /// metadata. This method never loads bibliography sources, resolves tags,
    /// applies a style, or formats a citation.
    pub fn citations(&self) -> Vec<crate::CitationField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::citation)
            .collect()
    }

    /// Return the number of typed, inert CITATION fields in the document.
    pub fn citation_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.citation().is_some())
            .count()
    }

    /// Return typed, inert BIBLIOGRAPHY fields in document field order.
    ///
    /// Switches and cached results are exposed solely as stored metadata. This
    /// method never loads bibliography sources, filters records, applies a
    /// style, sorts entries, or generates bibliography content.
    pub fn bibliographies(&self) -> Vec<crate::BibliographyField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::bibliography)
            .collect()
    }

    /// Return the number of typed, inert BIBLIOGRAPHY fields in the document.
    pub fn bibliography_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.bibliography().is_some())
            .count()
    }

    /// Return typed, inert DOCVARIABLE fields in document field order.
    ///
    /// Variable names, switches, and cached results are exposed solely as
    /// stored metadata. This method never reads document-variable destinations,
    /// resolves values, evaluates fields, or refreshes results.
    pub fn document_variable_fields(&self) -> Vec<crate::DocumentVariableField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::document_variable)
            .collect()
    }

    /// Return the number of typed, inert DOCVARIABLE fields in the document.
    pub fn document_variable_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.document_variable().is_some())
            .count()
    }

    /// Return typed, inert `DOCPROPERTY` fields in document field order.
    ///
    /// Property names, switches, cached results, and state are exposed solely
    /// as stored metadata. This method never reads core, extended, or custom
    /// document properties, resolves values, or refreshes a field.
    pub fn document_property_fields(&self) -> Vec<crate::DocumentPropertyField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::document_property)
            .collect()
    }

    /// Return the number of typed, inert `DOCPROPERTY` fields in the document.
    pub fn document_property_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.document_property().is_some())
            .count()
    }

    /// Return typed, inert explicit legacy `INFO` fields in document field order.
    ///
    /// Property selectors, optional replacement values, switches, cached
    /// results, and state are exposed solely as stored metadata. This method
    /// never reads, resolves, modifies, or writes document or template
    /// properties, or refreshes a field.
    pub fn info_fields(&self) -> Vec<crate::InfoField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::info_field)
            .collect()
    }

    /// Return the number of typed, inert explicit legacy `INFO` fields.
    pub fn info_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.info_field().is_some())
            .count()
    }

    /// Return typed, inert document-information fields in document field order.
    ///
    /// Kinds, switches, cached results, and state are exposed solely as stored
    /// metadata. This method never reads document properties or host identity
    /// data, calculates dates, revisions, or statistics, resolves values, or
    /// refreshes a field.
    pub fn document_information_fields(&self) -> Vec<crate::DocumentInformationField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::document_information)
            .collect()
    }

    /// Return the number of typed, inert document-information fields.
    pub fn document_information_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.document_information().is_some())
            .count()
    }

    /// Return typed, inert document-context and runtime fields in document
    /// field order.
    ///
    /// Kinds, switches, cached results, and state are exposed solely as stored
    /// metadata. This method never reads a document path, attached template,
    /// host filesystem state or file size, current clock, or page and section
    /// layout, resolves values, or refreshes a field.
    pub fn document_context_fields(&self) -> Vec<crate::DocumentContextField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::document_context)
            .collect()
    }

    /// Return the number of typed, inert document-context and runtime fields.
    pub fn document_context_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.document_context().is_some())
            .count()
    }

    /// Return typed, inert `MERGEFIELD` fields in document field order.
    ///
    /// Field names, switches, and cached results are exposed solely as stored
    /// metadata. This method never opens a data source, resolves records,
    /// performs a merge, or refreshes a field result.
    pub fn merge_fields(&self) -> Vec<crate::MergeField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::merge_field)
            .collect()
    }

    /// Return the number of typed, inert `MERGEFIELD` fields in the document.
    pub fn merge_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.merge_field().is_some())
            .count()
    }

    /// Return typed, inert `DATABASE` query fields in document field order.
    ///
    /// Stored opaque instructions, cached results, and field state are
    /// metadata only. This method never opens a data source or database, uses
    /// connection information, executes SQL, generates or inserts a table,
    /// changes layout, or refreshes a field.
    pub fn database_fields(&self) -> Vec<crate::DatabaseField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::database_field)
            .collect()
    }

    /// Return the number of typed, inert `DATABASE` query fields.
    pub fn database_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.database_field().is_some())
            .count()
    }

    /// Return typed, inert `DATA` mail-merge source fields in document order.
    ///
    /// Data-source and header-source identifiers, switches, cached results, and
    /// field state are exposed solely as stored metadata. This method never
    /// opens, reads, connects to, resolves, or modifies either source; it never
    /// selects a record, performs a merge, or refreshes a field result.
    pub fn mail_merge_data_fields(&self) -> Vec<crate::MailMergeDataField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::mail_merge_data)
            .collect()
    }

    /// Return the number of typed, inert `DATA` mail-merge source fields.
    pub fn mail_merge_data_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.mail_merge_data().is_some())
            .count()
    }

    /// Return typed, inert `MERGEREC` and `MERGESEQ` fields in document order.
    ///
    /// Stored kinds and cached results are exposed solely as metadata. This
    /// method never selects or counts records, opens a data source, performs a
    /// merge, or refreshes field results.
    pub fn mail_merge_counters(&self) -> Vec<crate::MailMergeCounterField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::mail_merge_counter)
            .collect()
    }

    /// Return the number of typed, inert mail-merge counter fields in the document.
    pub fn mail_merge_counter_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.mail_merge_counter().is_some())
            .count()
    }

    /// Return typed, inert `NEXT` mail-merge control fields in document order.
    ///
    /// Stored cached results and field state are exposed solely as metadata.
    /// This method never advances a record, opens a data source, performs a
    /// merge, or refreshes field results.
    pub fn mail_merge_next_fields(&self) -> Vec<crate::MailMergeNextField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::mail_merge_next)
            .collect()
    }

    /// Return the number of typed, inert `NEXT` mail-merge control fields.
    pub fn mail_merge_next_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.mail_merge_next().is_some())
            .count()
    }

    /// Return typed, inert `NEXTIF` and `SKIPIF` fields in document order.
    ///
    /// Stored comparison text, cached results, and field state are exposed
    /// solely as metadata. This method never evaluates a comparison, changes
    /// record selection, opens a data source, performs a merge, or refreshes
    /// field results.
    pub fn mail_merge_conditional_controls(
        &self,
    ) -> Vec<crate::MailMergeConditionalControlField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::mail_merge_conditional_control)
            .collect()
    }

    /// Return the number of typed, inert conditional mail-merge control fields.
    pub fn mail_merge_conditional_control_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.mail_merge_conditional_control().is_some())
            .count()
    }

    /// Return typed, inert `IF` fields in document order.
    ///
    /// Stored expression text, cached results, and field state are exposed
    /// solely as metadata. This method never parses or evaluates an expression,
    /// resolves field values, or refreshes field results.
    pub fn if_fields(&self) -> Vec<crate::IfField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::if_field)
            .collect()
    }

    /// Return the number of typed, inert `IF` fields.
    pub fn if_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.if_field().is_some())
            .count()
    }

    /// Return typed, inert `COMPARE` fields in document order.
    ///
    /// Stored comparisons, cached results, and field state are exposed solely
    /// as metadata. This method never parses or evaluates a comparison,
    /// resolves nested field values, or refreshes field results.
    pub fn compare_fields(&self) -> Vec<crate::CompareField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::compare_field)
            .collect()
    }

    /// Return the number of typed, inert `COMPARE` fields.
    pub fn compare_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.compare_field().is_some())
            .count()
    }

    /// Return typed, inert `QUOTE` fields in document order.
    ///
    /// Stored text arguments, switches, cached results, and field state are
    /// exposed solely as metadata. This method never interprets character codes,
    /// expands nested fields, inserts text, or refreshes a field result.
    pub fn quote_fields(&self) -> Vec<crate::QuoteField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::quote_field)
            .collect()
    }

    /// Return the number of typed, inert `QUOTE` fields.
    pub fn quote_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.quote_field().is_some())
            .count()
    }

    /// Return typed, inert `SYMBOL` fields in document order.
    ///
    /// Stored character arguments, switches, cached results, and field state
    /// are exposed solely as metadata. This method never maps a character code,
    /// looks up a font, inserts a glyph, changes formatting or layout, or
    /// refreshes a field result.
    pub fn symbol_fields(&self) -> Vec<crate::SymbolField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::symbol_field)
            .collect()
    }

    /// Return the number of typed, inert `SYMBOL` fields.
    pub fn symbol_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.symbol_field().is_some())
            .count()
    }

    /// Return typed, inert legacy automatic-numbering fields in document order.
    ///
    /// Stored kinds, switches, cached results, and field state are exposed
    /// solely as metadata. This method never calculates paragraph numbers,
    /// reads heading or style state, changes paragraphs or layout, or refreshes
    /// a field result.
    pub fn auto_number_fields(&self) -> Vec<crate::AutoNumberField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::auto_number_field)
            .collect()
    }

    /// Return the number of typed, inert legacy automatic-numbering fields.
    pub fn auto_number_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.auto_number_field().is_some())
            .count()
    }

    /// Return typed, inert `LISTNUM` fields in document order.
    ///
    /// Stored optional list names, switches, cached results, and field state
    /// are exposed solely as metadata. This method never looks up a list,
    /// determines a level or start value, calculates a number, changes layout,
    /// or refreshes a field result.
    pub fn list_number_fields(&self) -> Vec<crate::ListNumberField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::list_number_field)
            .collect()
    }

    /// Return the number of typed, inert `LISTNUM` fields.
    pub fn list_number_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.list_number_field().is_some())
            .count()
    }

    /// Return typed, inert `ASK` and `FILLIN` fields in document order.
    ///
    /// Stored prompt, bookmark, default-response, cached results, and field
    /// state are exposed solely as metadata. This method never displays a
    /// prompt, captures a response, creates or updates a bookmark, performs a
    /// merge, or refreshes field results.
    pub fn prompt_fields(&self) -> Vec<crate::PromptField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::prompt_field)
            .collect()
    }

    /// Return the number of typed, inert `ASK` and `FILLIN` fields.
    pub fn prompt_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.prompt_field().is_some())
            .count()
    }

    /// Return typed, inert user-identity fields in document order.
    ///
    /// Stored kind, override, formatting, cached result, and field state are
    /// exposed solely as metadata. This method never reads or modifies a host
    /// user's identity, applies formatting, or refreshes a field.
    pub fn user_identity_fields(&self) -> Vec<crate::UserIdentityField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::user_identity_field)
            .collect()
    }

    /// Return the number of typed, inert user-identity fields.
    pub fn user_identity_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.user_identity_field().is_some())
            .count()
    }

    /// Return typed, inert `ADVANCE` fields in document order.
    ///
    /// Stored point adjustments, cached results, and field state are exposed
    /// solely as metadata. This method never moves text, changes layout,
    /// reflows content, or refreshes field results.
    pub fn advance_fields(&self) -> Vec<crate::AdvanceField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::advance_field)
            .collect()
    }

    /// Return the number of typed, inert `ADVANCE` fields.
    pub fn advance_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.advance_field().is_some())
            .count()
    }

    /// Return typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields in document
    /// order.
    ///
    /// Stored recipient layout, locale, country, fallback, cached-result, and
    /// field state are exposed solely as metadata. This method never opens a
    /// data source, selects a record, performs a merge, expands placeholders,
    /// generates text, or refreshes a field result.
    pub fn mail_merge_recipient_fields(&self) -> Vec<crate::MailMergeRecipientField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::mail_merge_recipient_field)
            .collect()
    }

    /// Return the number of typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields.
    pub fn mail_merge_recipient_field_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.mail_merge_recipient_field().is_some())
            .count()
    }

    /// Return typed, inert TOC fields in document field order.
    ///
    /// Stored options and cached results are exposed solely as metadata. This
    /// method never scans entries, reads bookmarks, resolves hyperlinks,
    /// calculates page numbers, regenerates a table of contents, or refreshes
    /// a field.
    pub fn table_of_contents(&self) -> Vec<crate::TableOfContentsField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::table_of_contents)
            .collect()
    }

    /// Return the number of typed, inert TOC fields in the document.
    pub fn table_of_contents_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.table_of_contents().is_some())
            .count()
    }

    /// Return typed, inert TC entry fields in document field order.
    ///
    /// Stored entry text and options are exposed solely as metadata. This
    /// method never changes hidden text, calculates page numbers, generates a
    /// table of contents, or refreshes a field.
    pub fn table_of_contents_entries(&self) -> Vec<crate::TableOfContentsEntryField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::table_of_contents_entry)
            .collect()
    }

    /// Return the number of typed, inert TC entry fields in the document.
    pub fn table_of_contents_entry_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.table_of_contents_entry().is_some())
            .count()
    }

    /// Return typed, inert TA entry fields in document field order.
    ///
    /// Stored citation options and cached results are exposed solely as
    /// metadata. This method never finds cited text, follows bookmarks,
    /// calculates page numbers, generates a table of authorities, or refreshes
    /// a field.
    pub fn table_of_authorities_entries(&self) -> Vec<crate::TableOfAuthoritiesEntryField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::table_of_authorities_entry)
            .collect()
    }

    /// Return the number of typed, inert TA entry fields in the document.
    pub fn table_of_authorities_entry_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.table_of_authorities_entry().is_some())
            .count()
    }

    /// Return typed, inert TOA fields in document field order.
    ///
    /// Stored options and cached results are exposed solely as metadata. This
    /// method never finds citations, follows bookmarks, calculates page
    /// numbers, paginates the document, generates a table of authorities, or
    /// refreshes a field.
    pub fn tables_of_authorities(&self) -> Vec<crate::TableOfAuthoritiesField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::table_of_authorities)
            .collect()
    }

    /// Return the number of typed, inert TOA fields in the document.
    pub fn table_of_authorities_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.table_of_authorities().is_some())
            .count()
    }

    /// Return typed, inert INDEX fields in document field order.
    ///
    /// Stored configuration and cached results are exposed solely as metadata.
    /// This method never scans XE markers, follows bookmarks, calculates page
    /// numbers, paginates the document, generates an index, or refreshes a
    /// field.
    pub fn indexes(&self) -> Vec<crate::IndexField<'_>> {
        self.fields.iter().filter_map(crate::Field::index).collect()
    }

    /// Return the number of typed, inert INDEX fields in the document.
    pub fn index_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.index().is_some())
            .count()
    }

    /// Return typed, inert XE index-entry fields in document field order.
    ///
    /// Stored entry text and options are exposed solely as metadata. This
    /// method never changes hidden text, follows bookmarks, calculates pages,
    /// generates an index, or refreshes a field.
    pub fn index_entries(&self) -> Vec<crate::IndexEntryField<'_>> {
        self.fields
            .iter()
            .filter_map(crate::Field::index_entry)
            .collect()
    }

    /// Return the number of typed, inert XE index-entry fields in the document.
    pub fn index_entry_count(&self) -> usize {
        self.fields
            .iter()
            .filter(|field| field.index_entry().is_some())
            .count()
    }

    pub fn push_field(&mut self, field: super::super::field::Field<'a>) -> RtfResult<()> {
        if !matches!(field.owner, crate::FieldOwner::Body) {
            return Err(RtfError::MalformedDocument(
                "document-level generic fields must be owned by the body story".to_string(),
            ));
        }
        if self.fields.len() >= crate::field::MAX_GENERIC_FIELDS {
            return Err(RtfError::MalformedDocument(
                "RTF generic field count exceeds the safety limit".to_string(),
            ));
        }
        field.validate()?;
        let body = self.text();
        if body.get(field.position..field.position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF generic field position is not a UTF-8 body boundary".to_string(),
            ));
        }
        if self
            .last_body_story_position()?
            .is_some_and(|position| position > field.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF body story order moves backwards".to_string(),
            ));
        }
        let index = self.fields.len();
        self.fields.push(field);
        self.insert_body_story_event(crate::BodyStoryEvent::Field(index))?;
        Ok(())
    }

    pub fn clear_fields(&mut self) {
        self.fields.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::Field(_)));
    }

    /// Return ordered positional legacy form fields.
    pub fn form_fields(&self) -> &[super::super::form_field::FormField<'_>] {
        &self.form_fields
    }

    /// Return inert producer provenance from the RTF generator destination.
    pub fn generator(&self) -> Option<&crate::DocumentGenerator<'_>> {
        self.generator.as_ref()
    }

    /// Set validated inert producer provenance.
    pub fn set_generator(&mut self, generator: crate::DocumentGenerator<'a>) -> RtfResult<()> {
        generator.validate()?;
        self.generator = Some(generator);
        Ok(())
    }

    /// Remove producer provenance metadata.
    pub fn clear_generator(&mut self) {
        self.generator = None;
    }

    /// Return ordered revision-save/session provenance.
    pub fn revision_save_metadata(&self) -> Option<&crate::RevisionSaveMetadata> {
        self.revision_save.as_ref()
    }

    /// Replace revision-save/session provenance after full validation.
    pub fn set_revision_save_metadata(
        &mut self,
        metadata: crate::RevisionSaveMetadata,
    ) -> RtfResult<()> {
        metadata.validate()?;
        self.revision_save = Some(metadata);
        Ok(())
    }

    /// Remove revision-save/session provenance.
    pub fn clear_revision_save_metadata(&mut self) {
        self.revision_save = None;
    }

    /// Return the ordered inert XML namespace table, preserving empty-table presence.
    pub fn xml_namespaces(&self) -> Option<&[crate::XmlNamespace<'_>]> {
        self.xml_namespaces.as_deref()
    }

    /// Return the custom XML markup tags in body source order.
    ///
    /// The tags are inert `\xmlopen`/`\xmlclose` markup metadata (RTF 1.9.1
    /// custom XML markup): names and attributes are stored verbatim and are
    /// never resolved against a schema.
    pub fn custom_xml_tags(&self) -> &[crate::CustomXmlTag<'_>] {
        &self.custom_xml_tags
    }

    /// Return the math zones anchored in the body story, in source order.
    ///
    /// Math zones are inert `\mmath`/`\mmathPara` destinations (RTF 1.9.1
    /// mathematics): their typed trees are stored verbatim and are never
    /// evaluated, laid out, or rendered. Math run text does not enter the
    /// document body text (like field results).
    pub fn math_zones(&self) -> &[crate::MathZone<'_>] {
        &self.math_zones
    }

    /// Return the protection-exception ranges in `\*\protstart` source order.
    ///
    /// The ranges are inert Word 2003 document-protection metadata: their
    /// identifiers are stored verbatim and no editing restriction is ever
    /// evaluated or enforced.
    pub fn protection_ranges(&self) -> &[crate::ProtectionRange<'_>] {
        &self.protection_ranges
    }

    /// Return the editable regions in `\ebcstart` source order.
    ///
    /// The regions are inert boundary marks from protected documents: no
    /// editing restriction is ever evaluated or enforced.
    pub fn editable_regions(&self) -> &[crate::EditableRegion<'_>] {
        &self.editable_regions
    }

    /// Replace the XML namespace table after full validation.
    pub fn set_xml_namespaces(
        &mut self,
        namespaces: Vec<crate::XmlNamespace<'a>>,
    ) -> RtfResult<()> {
        Self::validate_xml_namespaces(&namespaces)?;
        self.xml_namespaces = Some(namespaces);
        Ok(())
    }

    /// Append one inert XML namespace entry, creating the table if absent.
    pub fn push_xml_namespace(&mut self, namespace: crate::XmlNamespace<'a>) -> RtfResult<()> {
        namespace.validate()?;
        let was_present = self.xml_namespaces.is_some();
        let mut namespaces = self.xml_namespaces.take().unwrap_or_default();
        namespaces.push(namespace);
        if let Err(error) = Self::validate_xml_namespaces(&namespaces) {
            namespaces.pop();
            self.xml_namespaces = was_present.then_some(namespaces);
            return Err(error);
        }
        self.xml_namespaces = Some(namespaces);
        Ok(())
    }

    /// Remove the XML namespace table entirely.
    pub fn clear_xml_namespaces(&mut self) {
        self.xml_namespaces = None;
    }

    /// Return inert range-protection usernames in their source order.
    pub fn protection_user_table(&self) -> Option<&crate::ProtectionUserTable<'_>> {
        self.protection_user_table.as_ref()
    }

    /// Replace the inert range-protection username table after full validation.
    pub fn set_protection_user_table(
        &mut self,
        table: crate::ProtectionUserTable<'a>,
    ) -> RtfResult<()> {
        table.validate()?;
        self.protection_user_table = Some(table);
        Ok(())
    }

    /// Remove the range-protection username table entirely.
    pub fn clear_protection_user_table(&mut self) {
        self.protection_user_table = None;
    }

    /// Return explicit document-level hyphenation properties.
    pub fn hyphenation(&self) -> &crate::DocumentHyphenation {
        &self.hyphenation
    }

    /// Replace document-level hyphenation properties after bounds validation.
    pub fn set_hyphenation(&mut self, hyphenation: crate::DocumentHyphenation) -> RtfResult<()> {
        hyphenation.validate()?;
        self.hyphenation = hyphenation;
        Ok(())
    }

    /// Remove all explicit hyphenation controls and restore RTF defaults.
    pub fn clear_hyphenation(&mut self) {
        self.hyphenation = crate::DocumentHyphenation::default();
    }

    /// Return inert external document/template names without resolving them.
    pub fn external_references(&self) -> &crate::DocumentExternalReferences<'_> {
        &self.external_references
    }

    /// Replace inert external document/template names after full validation.
    pub fn set_external_references(
        &mut self,
        references: crate::DocumentExternalReferences<'a>,
    ) -> RtfResult<()> {
        references.validate()?;
        self.external_references = references;
        Ok(())
    }

    /// Remove both external document-reference destinations.
    pub fn clear_external_references(&mut self) {
        self.external_references = crate::DocumentExternalReferences::default();
    }

    /// Return explicit passive document view and zoom settings.
    pub fn document_view(&self) -> &crate::DocumentView {
        &self.document_view
    }

    /// Replace passive document view and zoom settings after validation.
    pub fn set_document_view(&mut self, view: crate::DocumentView) -> RtfResult<()> {
        view.validate()?;
        self.document_view = view;
        Ok(())
    }

    /// Remove explicit document view controls.
    pub fn clear_document_view(&mut self) {
        self.document_view = crate::DocumentView::default();
    }

    /// Return passive review-display preferences.
    pub fn review_display(&self) -> &crate::DocumentReviewDisplay {
        &self.review_display
    }

    /// Replace passive review-display preferences.
    pub fn set_review_display(&mut self, display: crate::DocumentReviewDisplay) {
        self.review_display = display;
    }

    /// Remove all review-display suppression flags.
    pub fn clear_review_display(&mut self) {
        self.review_display = crate::DocumentReviewDisplay::default();
    }

    /// Return passive document-window caption metadata.
    pub fn window_caption(&self) -> Option<&crate::DocumentWindowCaption<'a>> {
        self.window_caption.as_ref()
    }

    /// Return the custom kinsoku (East Asian line-breaking) character sets.
    ///
    /// The sets are inert typography metadata: no line-breaking rule is ever
    /// evaluated or applied.
    pub fn kinsoku(&self) -> &crate::DocumentKinsoku<'_> {
        &self.kinsoku
    }

    /// Replace passive document-window caption metadata.
    pub fn set_window_caption(
        &mut self,
        caption: crate::DocumentWindowCaption<'a>,
    ) -> RtfResult<()> {
        caption.validate()?;
        self.window_caption = Some(caption);
        Ok(())
    }

    /// Remove document-window caption metadata.
    pub fn clear_window_caption(&mut self) {
        self.window_caption = None;
    }

    /// Return the inert custom XSL transform location.
    pub fn xsl_transform(&self) -> Option<&crate::DocumentXslTransform<'a>> {
        self.xsl_transform.as_ref()
    }

    /// Replace the inert custom XSL transform location.
    pub fn set_xsl_transform(
        &mut self,
        transform: crate::DocumentXslTransform<'a>,
    ) -> RtfResult<()> {
        transform.validate()?;
        self.xsl_transform = Some(transform);
        Ok(())
    }

    /// Remove custom XSL transform location metadata.
    pub fn clear_xsl_transform(&mut self) {
        self.xsl_transform = None;
    }

    /// Return the passive requested transform-usage intent.
    pub fn xsl_transform_usage(&self) -> crate::DocumentXslTransformUsage {
        self.xsl_transform_usage
    }

    /// Replace the passive requested transform-usage intent.
    pub fn set_xsl_transform_usage(&mut self, usage: crate::DocumentXslTransformUsage) {
        self.xsl_transform_usage = usage;
    }

    /// Clear requested transform usage without changing the stored location.
    pub fn clear_xsl_transform_usage(&mut self) {
        self.xsl_transform_usage = crate::DocumentXslTransformUsage::NotRequested;
    }

    /// Return passive style-list filter suggestions.
    pub fn style_list_filter(&self) -> Option<crate::DocumentStyleListFilter> {
        self.style_list_filter
    }

    /// Replace passive style-list filter suggestions.
    pub fn set_style_list_filter(
        &mut self,
        filter: crate::DocumentStyleListFilter,
    ) -> RtfResult<()> {
        filter.validate()?;
        self.style_list_filter = Some(filter);
        Ok(())
    }

    /// Remove style-list filter suggestions.
    pub fn clear_style_list_filter(&mut self) {
        self.style_list_filter = None;
    }

    /// Return an explicitly stored style-list sorting suggestion.
    pub fn style_sort_method(&self) -> Option<crate::DocumentStyleSortMethod> {
        self.style_sort_method
    }

    /// Return the stored suggestion or the specification default when omitted.
    pub fn effective_style_sort_method(&self) -> crate::DocumentStyleSortMethod {
        self.style_sort_method.unwrap_or_default()
    }

    /// Replace the passive style-list sorting suggestion.
    pub fn set_style_sort_method(&mut self, method: crate::DocumentStyleSortMethod) {
        self.style_sort_method = Some(method);
    }

    /// Remove the explicit suggestion, restoring the effective host default.
    pub fn clear_style_sort_method(&mut self) {
        self.style_sort_method = None;
    }

    /// Return passive save-related document preferences.
    pub fn save_preferences(&self) -> &crate::DocumentSavePreferences {
        &self.save_preferences
    }

    /// Replace passive save-related document preferences.
    pub fn set_save_preferences(&mut self, preferences: crate::DocumentSavePreferences) {
        self.save_preferences = preferences;
    }

    /// Remove explicit save-related preferences.
    pub fn clear_save_preferences(&mut self) {
        self.save_preferences = crate::DocumentSavePreferences::default();
    }

    /// Return opaque write-reservation metadata without authenticating it.
    pub fn write_reservations(&self) -> &crate::DocumentWriteReservations<'a> {
        &self.write_reservations
    }

    /// Replace opaque write-reservation metadata without authenticating it.
    pub fn set_write_reservations(
        &mut self,
        reservations: crate::DocumentWriteReservations<'a>,
    ) -> RtfResult<()> {
        reservations.validate()?;
        self.write_reservations = reservations;
        Ok(())
    }

    /// Remove all write-reservation metadata.
    pub fn clear_write_reservations(&mut self) {
        self.write_reservations = crate::DocumentWriteReservations::default();
    }

    pub fn origin_metadata(&self) -> &crate::DocumentOriginMetadata {
        &self.origin_metadata
    }

    pub fn set_origin_metadata(&mut self, metadata: crate::DocumentOriginMetadata) {
        self.origin_metadata = metadata;
    }

    pub fn clear_origin_metadata(&mut self) {
        self.origin_metadata = crate::DocumentOriginMetadata::default();
    }

    /// Return passive file and template settings.
    pub fn file_settings(&self) -> &crate::DocumentFileSettings {
        &self.file_settings
    }

    /// Replace passive file and template settings.
    pub fn set_file_settings(&mut self, settings: crate::DocumentFileSettings) {
        self.file_settings = settings;
    }

    /// Remove explicit file and template settings.
    pub fn clear_file_settings(&mut self) {
        self.file_settings = crate::DocumentFileSettings::default();
    }

    /// Return passive compatibility and output-request flags.
    pub fn output_settings(&self) -> &crate::DocumentOutputSettings {
        &self.output_settings
    }

    /// Replace passive compatibility and output-request flags.
    pub fn set_output_settings(&mut self, settings: crate::DocumentOutputSettings) {
        self.output_settings = settings;
    }

    /// Remove explicit compatibility and output-request flags.
    pub fn clear_output_settings(&mut self) {
        self.output_settings = crate::DocumentOutputSettings::default();
    }

    /// Return passive document rendering flags.
    pub fn rendering_settings(&self) -> &crate::DocumentRenderingSettings {
        &self.rendering_settings
    }

    /// Replace passive document rendering flags.
    pub fn set_rendering_settings(&mut self, settings: crate::DocumentRenderingSettings) {
        self.rendering_settings = settings;
    }

    /// Remove explicit document rendering flags.
    pub fn clear_rendering_settings(&mut self) {
        self.rendering_settings = crate::DocumentRenderingSettings::default();
    }

    /// Return passive printing, cleanup, and event-mask properties.
    pub fn processing_settings(&self) -> &crate::DocumentProcessingSettings {
        &self.processing_settings
    }

    /// Replace passive printing, cleanup, and event-mask properties.
    pub fn set_processing_settings(&mut self, settings: crate::DocumentProcessingSettings) {
        self.processing_settings = settings;
    }

    /// Remove explicit printing, cleanup, and event-mask properties.
    pub fn clear_processing_settings(&mut self) {
        self.processing_settings = crate::DocumentProcessingSettings::default();
    }

    /// Return passive document-level drawing-grid properties.
    pub fn drawing_grid(&self) -> &crate::DocumentDrawingGrid {
        &self.drawing_grid
    }

    /// Replace passive document-level drawing-grid properties.
    pub fn set_drawing_grid(&mut self, drawing_grid: crate::DocumentDrawingGrid) {
        self.drawing_grid = drawing_grid;
    }

    /// Remove all explicit document-level drawing-grid properties.
    pub fn clear_drawing_grid(&mut self) {
        self.drawing_grid = crate::DocumentDrawingGrid::default();
    }

    /// Return passive print-layout settings.
    pub fn print_layout_settings(&self) -> &crate::DocumentPrintLayoutSettings {
        &self.print_layout_settings
    }

    /// Atomically replace passive print-layout settings.
    pub fn set_print_layout_settings(
        &mut self,
        settings: crate::DocumentPrintLayoutSettings,
    ) -> RtfResult<()> {
        settings.validate()?;
        self.print_layout_settings = settings;
        Ok(())
    }

    /// Atomically replace the document-wide gutter width in twips.
    pub fn set_document_gutter_twips(&mut self, value: Option<u32>) -> RtfResult<()> {
        let mut candidate = self.print_layout_settings;
        candidate.set_document_gutter_twips(value)?;
        self.print_layout_settings = candidate;
        Ok(())
    }

    /// Remove explicit print-layout settings.
    pub fn clear_print_layout_settings(&mut self) {
        self.print_layout_settings = crate::DocumentPrintLayoutSettings::default();
    }

    /// Return passive theme font-resolution language identifiers.
    pub fn theme_languages(&self) -> &crate::DocumentThemeLanguages {
        &self.theme_languages
    }

    /// Replace passive theme font-resolution language identifiers.
    pub fn set_theme_languages(&mut self, languages: crate::DocumentThemeLanguages) {
        self.theme_languages = languages;
    }

    /// Remove explicit theme font-resolution language identifiers.
    pub fn clear_theme_languages(&mut self) {
        self.theme_languages = crate::DocumentThemeLanguages::default();
    }

    /// Return passive web-save and custom-XML policies.
    pub fn xml_policies(&self) -> &crate::DocumentXmlPolicies {
        &self.xml_policies
    }

    /// Replace passive web-save and custom-XML policies.
    pub fn set_xml_policies(&mut self, policies: crate::DocumentXmlPolicies) {
        self.xml_policies = policies;
    }

    /// Remove all explicit web-save and custom-XML policies.
    pub fn clear_xml_policies(&mut self) {
        self.xml_policies = crate::DocumentXmlPolicies::default();
    }

    /// Return passive system-font and linguistic-data embedding policies.
    pub fn embedding_policies(&self) -> &crate::DocumentEmbeddingPolicies {
        &self.embedding_policies
    }

    /// Replace passive system-font and linguistic-data embedding policies.
    pub fn set_embedding_policies(&mut self, policies: crate::DocumentEmbeddingPolicies) {
        self.embedding_policies = policies;
    }

    /// Remove all explicit embedding policies.
    pub fn clear_embedding_policies(&mut self) {
        self.embedding_policies = crate::DocumentEmbeddingPolicies::default();
    }

    /// Return passive move and formatting revision policies.
    pub fn revision_policies(&self) -> &crate::DocumentRevisionPolicies {
        &self.revision_policies
    }

    /// Replace passive move and formatting revision policies.
    pub fn set_revision_policies(&mut self, policies: crate::DocumentRevisionPolicies) {
        self.revision_policies = policies;
    }

    /// Remove all explicit revision-policy controls.
    pub fn clear_revision_policies(&mut self) {
        self.revision_policies = crate::DocumentRevisionPolicies::default();
    }

    /// Return passive theme and style-application policies.
    pub fn style_policies(&self) -> &crate::DocumentStylePolicies {
        &self.style_policies
    }

    /// Replace passive theme and style-application policies.
    pub fn set_style_policies(&mut self, policies: crate::DocumentStylePolicies) {
        self.style_policies = policies;
    }

    /// Remove all explicit theme and style-application policies.
    pub fn clear_style_policies(&mut self) {
        self.style_policies = crate::DocumentStylePolicies::default();
    }

    /// Return passive legacy style and formatting restriction declarations.
    pub fn style_restrictions(&self) -> &crate::DocumentStyleRestrictions {
        &self.style_restrictions
    }

    /// Replace passive legacy style and formatting restriction declarations.
    pub fn set_style_restrictions(&mut self, restrictions: crate::DocumentStyleRestrictions) {
        self.style_restrictions = restrictions;
    }

    /// Remove all legacy style and formatting restriction declarations.
    pub fn clear_style_restrictions(&mut self) {
        self.style_restrictions = crate::DocumentStyleRestrictions::default();
    }

    /// Return passive booklet-printing metadata.
    pub fn booklet_printing(&self) -> &crate::DocumentBookletPrinting {
        &self.booklet_printing
    }

    /// Replace passive booklet-printing metadata.
    pub fn set_booklet_printing(&mut self, printing: crate::DocumentBookletPrinting) {
        self.booklet_printing = printing;
    }

    /// Remove all booklet-printing metadata.
    pub fn clear_booklet_printing(&mut self) {
        self.booklet_printing = crate::DocumentBookletPrinting::default();
    }

    /// Return passive privacy-removal requests.
    pub fn privacy_policies(&self) -> &crate::DocumentPrivacyPolicies {
        &self.privacy_policies
    }

    /// Replace passive privacy-removal requests.
    pub fn set_privacy_policies(&mut self, policies: crate::DocumentPrivacyPolicies) {
        self.privacy_policies = policies;
    }

    /// Remove all privacy-removal requests without changing document metadata.
    pub fn clear_privacy_policies(&mut self) {
        self.privacy_policies = crate::DocumentPrivacyPolicies::default();
    }

    /// Return passive legacy extra-line-spacing compatibility requests.
    pub fn line_spacing_compatibility(&self) -> &crate::DocumentLineSpacingCompatibility {
        &self.line_spacing_compatibility
    }

    /// Replace passive legacy extra-line-spacing compatibility requests.
    pub fn set_line_spacing_compatibility(
        &mut self,
        compatibility: crate::DocumentLineSpacingCompatibility,
    ) {
        self.line_spacing_compatibility = compatibility;
    }

    /// Clear legacy extra-line-spacing requests without changing layout.
    pub fn clear_line_spacing_compatibility(&mut self) {
        self.line_spacing_compatibility = crate::DocumentLineSpacingCompatibility::default();
    }

    /// Return passive Word 6-era East Asian typography compatibility requests.
    pub fn east_asian_compatibility(&self) -> &crate::DocumentEastAsianCompatibility {
        &self.east_asian_compatibility
    }

    /// Replace passive East Asian compatibility requests without applying them.
    pub fn set_east_asian_compatibility(
        &mut self,
        compatibility: crate::DocumentEastAsianCompatibility,
    ) {
        self.east_asian_compatibility = compatibility;
    }

    /// Clear East Asian compatibility requests without changing document content.
    pub fn clear_east_asian_compatibility(&mut self) {
        self.east_asian_compatibility = crate::DocumentEastAsianCompatibility::default();
    }

    /// Return passive legacy table-layout compatibility requests.
    pub fn table_layout_compatibility(&self) -> &crate::DocumentTableLayoutCompatibility {
        &self.table_layout_compatibility
    }

    /// Replace passive table-layout compatibility requests without applying them.
    pub fn set_table_layout_compatibility(
        &mut self,
        compatibility: crate::DocumentTableLayoutCompatibility,
    ) {
        self.table_layout_compatibility = compatibility;
    }

    /// Clear table-layout compatibility requests without changing any table.
    pub fn clear_table_layout_compatibility(&mut self) {
        self.table_layout_compatibility = crate::DocumentTableLayoutCompatibility::default();
    }

    /// Return passive legacy automatic-layout compatibility requests.
    pub fn legacy_layout_compatibility(&self) -> &crate::DocumentLegacyLayoutCompatibility {
        &self.legacy_layout_compatibility
    }

    /// Replace passive legacy layout requests without applying them.
    pub fn set_legacy_layout_compatibility(
        &mut self,
        compatibility: crate::DocumentLegacyLayoutCompatibility,
    ) {
        self.legacy_layout_compatibility = compatibility;
    }

    /// Clear legacy layout requests without changing document content or layout.
    pub fn clear_legacy_layout_compatibility(&mut self) {
        self.legacy_layout_compatibility = crate::DocumentLegacyLayoutCompatibility::default();
    }

    /// Return passive Asian character-grid and line-breaking requests.
    pub fn asian_grid_compatibility(&self) -> &crate::DocumentAsianGridCompatibility {
        &self.asian_grid_compatibility
    }

    /// Replace passive Asian grid requests without applying them.
    pub fn set_asian_grid_compatibility(
        &mut self,
        compatibility: crate::DocumentAsianGridCompatibility,
    ) {
        self.asian_grid_compatibility = compatibility;
    }

    /// Clear Asian grid requests without changing text or layout.
    pub fn clear_asian_grid_compatibility(&mut self) {
        self.asian_grid_compatibility = crate::DocumentAsianGridCompatibility::default();
    }

    /// Return passive compatibility reset, UI-throttling, and upgrade requests.
    pub fn compatibility_policy(&self) -> &crate::DocumentCompatibilityPolicy {
        &self.compatibility_policy
    }

    /// Replace passive document compatibility policy declarations.
    pub fn set_compatibility_policy(&mut self, policy: crate::DocumentCompatibilityPolicy) {
        self.compatibility_policy = policy;
    }

    /// Clear compatibility policy declarations without changing document content.
    pub fn clear_compatibility_policy(&mut self) {
        self.compatibility_policy = crate::DocumentCompatibilityPolicy::default();
    }

    /// Return passive Word 2003-era compatibility requests.
    pub fn word_2003_compatibility(&self) -> &crate::DocumentWord2003Compatibility {
        &self.word_2003_compatibility
    }

    /// Replace passive Word 2003 compatibility requests without applying them.
    pub fn set_word_2003_compatibility(
        &mut self,
        compatibility: crate::DocumentWord2003Compatibility,
    ) {
        self.word_2003_compatibility = compatibility;
    }

    /// Clear Word 2003 compatibility requests without changing document layout.
    pub fn clear_word_2003_compatibility(&mut self) {
        self.word_2003_compatibility = crate::DocumentWord2003Compatibility::default();
    }

    /// Return inert Office theme bytes without interpreting their contents.
    pub fn theme(&self) -> Option<&crate::DocumentTheme<'_>> {
        self.theme.as_ref()
    }

    /// Replace inert Office theme bytes after bounds validation.
    pub fn set_theme(&mut self, theme: crate::DocumentTheme<'a>) -> RtfResult<()> {
        theme.validate()?;
        self.theme = Some(theme);
        Ok(())
    }

    /// Remove theme and color-scheme mapping payloads.
    pub fn clear_theme(&mut self) {
        self.theme = None;
    }

    /// Return inert latent-style defaults and ordered exceptions.
    pub fn latent_styles(&self) -> Option<&crate::LatentStyles<'_>> {
        self.latent_styles.as_ref()
    }

    /// Replace latent-style metadata after full validation.
    pub fn set_latent_styles(&mut self, styles: crate::LatentStyles<'a>) -> RtfResult<()> {
        styles.validate()?;
        self.latent_styles = Some(styles);
        Ok(())
    }

    /// Remove latent-style metadata.
    pub fn clear_latent_styles(&mut self) {
        self.latent_styles = None;
    }

    /// Return inert custom XML data-store bytes without interpreting them.
    pub fn data_store(&self) -> Option<&crate::DocumentDataStore<'_>> {
        self.data_store.as_ref()
    }

    /// Replace inert data-store bytes after bounds validation.
    pub fn set_data_store(&mut self, data_store: crate::DocumentDataStore<'a>) -> RtfResult<()> {
        data_store.validate()?;
        self.data_store = Some(data_store);
        Ok(())
    }

    /// Remove the custom XML data-store payload.
    pub fn clear_data_store(&mut self) {
        self.data_store = None;
    }

    /// Return inert mail-merge metadata without opening sources or evaluating queries.
    pub fn mail_merge(&self) -> Option<&crate::MailMerge<'_>> {
        self.mail_merge.as_ref()
    }

    /// Replace inert mail-merge metadata after complete bounds validation.
    pub fn set_mail_merge(&mut self, mail_merge: crate::MailMerge<'a>) -> RtfResult<()> {
        mail_merge.validate()?;
        self.mail_merge = Some(mail_merge);
        Ok(())
    }

    /// Remove all mail-merge metadata.
    pub fn clear_mail_merge(&mut self) {
        self.mail_merge = None;
    }

    /// Return document-level mathematical layout defaults.
    pub fn math_properties(&self) -> Option<&crate::DocumentMathProperties> {
        self.math_properties.as_ref()
    }

    /// Replace document-level mathematical layout defaults after validation.
    pub fn set_math_properties(
        &mut self,
        properties: crate::DocumentMathProperties,
    ) -> RtfResult<()> {
        properties.validate()?;
        self.math_properties = Some(properties);
        Ok(())
    }

    /// Remove document-level mathematical layout defaults.
    pub fn clear_math_properties(&mut self) {
        self.math_properties = None;
    }

    /// Return language defaults declared by the RTF header.
    pub fn language_defaults(&self) -> &crate::DocumentLanguageDefaults {
        &self.language_defaults
    }

    pub fn default_formatting(&self) -> &crate::DocumentDefaultFormatting {
        &self.default_formatting
    }
    pub fn set_default_formatting(
        &mut self,
        value: crate::DocumentDefaultFormatting,
    ) -> RtfResult<()> {
        value.validate()?;
        self.default_formatting = value;
        Ok(())
    }
    pub fn clear_default_formatting(&mut self) {
        self.default_formatting = crate::DocumentDefaultFormatting::default();
    }

    /// Replace document language defaults.
    pub fn set_language_defaults(
        &mut self,
        defaults: crate::DocumentLanguageDefaults,
    ) -> RtfResult<()> {
        defaults.validate()?;
        self.language_defaults = defaults;
        Ok(())
    }

    /// Remove all document language defaults.
    pub fn clear_language_defaults(&mut self) {
        self.language_defaults = crate::DocumentLanguageDefaults::default();
    }

    /// Return the explicitly declared `deftab` width in twips.
    ///
    /// `None` means the source omitted `deftab`; it does not mean zero.
    pub fn default_tab_width_twips(&self) -> Option<u32> {
        self.default_tab_width_twips
    }

    /// Return the explicit width or the RTF 1.9.1 default of 720 twips.
    pub fn effective_default_tab_width_twips(&self) -> u32 {
        self.default_tab_width_twips
            .unwrap_or(crate::DEFAULT_TAB_WIDTH_TWIPS)
    }

    /// Set an explicit default tab width without creating paragraph tab stops.
    pub fn set_default_tab_width_twips(&mut self, width: u32) -> RtfResult<()> {
        if width > crate::MAX_DEFAULT_TAB_WIDTH_TWIPS {
            return Err(RtfError::MalformedDocument(format!(
                "RTF deftab width {width} exceeds {}",
                crate::MAX_DEFAULT_TAB_WIDTH_TWIPS
            )));
        }
        self.default_tab_width_twips = Some(width);
        Ok(())
    }

    /// Remove the explicit width so serialization preserves omission.
    pub fn clear_default_tab_width(&mut self) {
        self.default_tab_width_twips = None;
    }

    /// Return the explicit document-wide bidirectional precedence.
    pub fn document_direction(&self) -> Option<crate::TextDirection> {
        self.document_direction
    }

    /// Set the explicit document-wide bidirectional precedence.
    pub fn set_document_direction(&mut self, direction: crate::TextDirection) {
        self.document_direction = Some(direction);
    }

    /// Remove the explicit document-wide bidirectional precedence.
    pub fn clear_document_direction(&mut self) {
        self.document_direction = None;
    }

    /// Return whether the document gutter is positioned on the right.
    pub fn gutter_on_right(&self) -> bool {
        self.gutter_on_right
    }

    /// Position the document gutter on the right when `enabled` is true.
    pub fn set_gutter_on_right(&mut self, enabled: bool) {
        self.gutter_on_right = enabled;
    }

    fn validate_xml_namespaces(namespaces: &[crate::XmlNamespace<'_>]) -> RtfResult<()> {
        if namespaces.len() > crate::xml_namespace::MAX_XML_NAMESPACES {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace count exceeds the safety limit".to_string(),
            ));
        }
        let mut ids = HashSet::new();
        ids.try_reserve(namespaces.len())
            .map_err(|_| RtfError::AllocationFailed {
                resource: "XML namespace IDs",
                requested: namespaces.len().saturating_mul(std::mem::size_of::<u32>()),
            })?;
        let mut total = 0usize;
        for namespace in namespaces {
            namespace.validate()?;
            if !ids.insert(namespace.id) {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace IDs must be unique".to_string(),
                ));
            }
            total = total
                .checked_add(namespace.namespace.len())
                .ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF XML namespace aggregate size overflow".to_string(),
                    )
                })?;
            if total > crate::xml_namespace::MAX_XML_NAMESPACE_TOTAL_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace aggregate text exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(())
    }

    /// Append inert form-field metadata at a valid visible body range.
    pub fn push_form_field(
        &mut self,
        field: super::super::form_field::FormField<'a>,
    ) -> RtfResult<()> {
        field.validate()?;
        if self.form_fields.len() >= super::super::form_field::MAX_FORM_FIELDS {
            return Err(RtfError::MalformedDocument(
                "RTF form-field count exceeds the safety limit".to_string(),
            ));
        }
        let body = self.text();
        let result = body.get(field.position..field.range_end).ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF form-field range is outside body text or splits a character".to_string(),
            )
        })?;
        if result != field.result_text {
            return Err(RtfError::MalformedDocument(
                "RTF form-field result does not match its visible body range".to_string(),
            ));
        }
        if field.position != field.range_end
            && self.form_fields.iter().any(|existing| {
                existing.position != existing.range_end
                    && field.position < existing.range_end
                    && existing.position < field.range_end
            })
        {
            return Err(RtfError::MalformedDocument(
                "RTF form-field result ranges cannot overlap".to_string(),
            ));
        }
        let total = self
            .form_fields
            .iter()
            .try_fold(
                field.text_bytes().unwrap_or(usize::MAX),
                |total, existing| total.checked_add(existing.text_bytes()?),
            )
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF form-field aggregate size overflow".to_string())
            })?;
        if total > super::super::form_field::MAX_FORM_FIELD_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF form-field aggregate text exceeds the safety limit".to_string(),
            ));
        }
        let index = self.form_fields.len();
        self.form_fields.push(field);
        self.insert_body_story_event(crate::BodyStoryEvent::FormFieldStart(index))?;
        self.insert_body_story_event(crate::BodyStoryEvent::FormFieldEnd(index))?;
        Ok(())
    }

    /// Remove all legacy form-field metadata without changing visible body text.
    pub fn clear_form_fields(&mut self) {
        self.form_fields.clear();
        self.body_story_events.retain(|event| {
            !matches!(
                event,
                crate::BodyStoryEvent::FormFieldStart(_) | crate::BodyStoryEvent::FormFieldEnd(_)
            )
        });
    }

    /// Return embedded and linked object records without activating their content.
    pub fn objects(&self) -> &[super::super::object::EmbeddedObject<'_>] {
        &self.objects
    }

    /// Resolve one object result-picture reference without cloning picture bytes.
    pub fn picture_for_object_result(
        &self,
        object: &super::super::object::EmbeddedObject<'_>,
        result_index: usize,
    ) -> Option<&super::super::picture::Picture<'_>> {
        object
            .result_picture_indices
            .get(result_index)
            .and_then(|index| self.pictures.get(*index))
    }

    /// Append a validated inert object destination at its body position.
    pub fn push_object(
        &mut self,
        object: super::super::object::EmbeddedObject<'a>,
    ) -> RtfResult<()> {
        if self.objects.len() >= super::super::object::MAX_EMBEDDED_OBJECTS {
            return Err(RtfError::MalformedDocument(
                "RTF embedded object count exceeds the safety limit".to_string(),
            ));
        }
        let body = self.text();
        object.validate(&body, self.pictures.len())?;
        if self
            .objects
            .last()
            .is_some_and(|previous| previous.position > object.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF embedded objects are not ordered by body position".to_string(),
            ));
        }
        let index = self.objects.len();
        self.objects.push(object);
        self.insert_body_story_event(crate::BodyStoryEvent::Object(index))?;
        Ok(())
    }

    /// Remove all inert object destinations without removing shared result pictures.
    pub fn clear_objects(&mut self) {
        self.objects.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::Object(_)));
    }

    /// Return ordered inert RTF document-variable name/value pairs.
    pub fn document_variables(&self) -> &[super::super::document_variable::DocumentVariable<'_>] {
        &self.document_variables
    }

    /// Append a document variable without evaluating or resolving it.
    pub fn push_document_variable(
        &mut self,
        variable: super::super::document_variable::DocumentVariable<'a>,
    ) -> RtfResult<()> {
        variable.validate()?;
        if self.document_variables.len() >= super::super::document_variable::MAX_DOCUMENT_VARIABLES
        {
            return Err(RtfError::MalformedDocument(
                "RTF document-variable count limit exceeded".to_string(),
            ));
        }
        let aggregate = self.document_variables.iter().try_fold(
            variable.name.len() + variable.value.len(),
            |size, existing| {
                size.checked_add(existing.name.len())
                    .and_then(|size| size.checked_add(existing.value.len()))
            },
        );
        if aggregate.is_none_or(|size| {
            size > super::super::document_variable::MAX_DOCUMENT_VARIABLE_TEXT_BYTES
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF document-variable aggregate text limit exceeded".to_string(),
            ));
        }
        self.document_variables.push(variable);
        Ok(())
    }

    /// Remove all document variables.
    pub fn clear_document_variables(&mut self) {
        self.document_variables.clear();
    }

    /// Return ordered, inert RTF user-defined document properties.
    pub fn user_properties(&self) -> &[super::super::user_property::UserProperty<'_>] {
        &self.user_properties
    }

    /// Append a unique user-defined property without evaluating its value or link.
    pub fn push_user_property(
        &mut self,
        property: super::super::user_property::UserProperty<'a>,
    ) -> RtfResult<()> {
        property.validate()?;
        if self.user_properties.len() >= super::super::user_property::MAX_USER_PROPERTIES {
            return Err(RtfError::MalformedDocument(
                "RTF user-property count limit exceeded".to_string(),
            ));
        }
        if self
            .user_properties
            .iter()
            .any(|existing| existing.name == property.name)
        {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF user-property name: {}",
                property.name
            )));
        }
        let aggregate = property.text_bytes().and_then(|initial| {
            self.user_properties
                .iter()
                .try_fold(initial, |size, existing| {
                    size.checked_add(existing.text_bytes()?)
                })
        });
        if aggregate
            .is_none_or(|size| size > super::super::user_property::MAX_USER_PROPERTY_TEXT_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "RTF user-property aggregate text limit exceeded".to_string(),
            ));
        }
        self.user_properties.push(property);
        Ok(())
    }

    /// Remove all user-defined properties.
    pub fn clear_user_properties(&mut self) {
        self.user_properties.clear();
    }

    /// Return ordered, inert index and table-of-contents source marks.
    pub fn navigation_entries(&self) -> &[super::super::navigation_entry::NavigationEntry<'_>] {
        &self.navigation_entries
    }

    /// Return ordered inert generated list markers.
    pub fn generated_list_markers(&self) -> &[crate::GeneratedListMarker<'_>] {
        &self.generated_list_markers
    }

    /// Append a generated list marker at a valid UTF-8 body position.
    pub fn push_generated_list_marker(
        &mut self,
        marker: crate::GeneratedListMarker<'a>,
    ) -> RtfResult<()> {
        marker.validate()?;
        if self.generated_list_markers.len()
            >= crate::generated_list_marker::MAX_GENERATED_LIST_MARKERS
        {
            return Err(RtfError::MalformedDocument(
                "RTF generated list-marker count exceeds the safety limit".to_string(),
            ));
        }
        let body = self.text();
        if body.get(marker.position..marker.position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF generated list-marker position is not a UTF-8 body boundary".to_string(),
            ));
        }
        if self.generated_list_markers.last().is_some_and(|previous| {
            previous.position > marker.position
                || (previous.position == marker.position && previous.kind == marker.kind)
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF generated list markers are duplicated or out of body order".to_string(),
            ));
        }
        let total = self
            .generated_list_markers
            .iter()
            .try_fold(marker.text.len(), |total, entry| {
                total.checked_add(entry.text.len())
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF generated list-marker text size overflow".to_string(),
                )
            })?;
        if total > crate::generated_list_marker::MAX_GENERATED_LIST_MARKER_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF generated list-marker text exceeds the aggregate safety limit".to_string(),
            ));
        }
        let index = self.generated_list_markers.len();
        self.generated_list_markers.push(marker);
        self.insert_body_story_event(crate::BodyStoryEvent::GeneratedListMarker(index))?;
        Ok(())
    }

    pub fn clear_generated_list_markers(&mut self) {
        self.generated_list_markers.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::GeneratedListMarker(_)));
    }

    /// Append an inert source mark at a valid UTF-8 body position.
    pub fn push_navigation_entry(
        &mut self,
        entry: super::super::navigation_entry::NavigationEntry<'a>,
    ) -> RtfResult<()> {
        entry.validate()?;
        let body = self.text();
        if body.get(entry.position()..entry.position()).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry position is outside body text or splits a character"
                    .to_string(),
            ));
        }
        if self.navigation_entries.len() >= super::super::navigation_entry::MAX_NAVIGATION_ENTRIES {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry count limit exceeded".to_string(),
            ));
        }
        let aggregate = entry.text_bytes().and_then(|initial| {
            self.navigation_entries
                .iter()
                .try_fold(initial, |size, existing| {
                    size.checked_add(existing.text_bytes()?)
                })
        });
        if aggregate.is_none_or(|size| {
            size > super::super::navigation_entry::MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry aggregate text limit exceeded".to_string(),
            ));
        }
        let index = self.navigation_entries.len();
        self.navigation_entries.push(entry);
        self.insert_body_story_event(crate::BodyStoryEvent::NavigationEntry(index))?;
        Ok(())
    }

    /// Append navigation metadata for ownership by a table-cell story.
    pub fn push_cell_navigation_entry_metadata(
        &mut self,
        entry: super::super::navigation_entry::NavigationEntry<'a>,
    ) -> RtfResult<usize> {
        entry.validate()?;
        if self.navigation_entries.len() >= super::super::navigation_entry::MAX_NAVIGATION_ENTRIES {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry count limit exceeded".to_string(),
            ));
        }
        let aggregate = entry.text_bytes().and_then(|initial| {
            self.navigation_entries
                .iter()
                .try_fold(initial, |size, existing| {
                    size.checked_add(existing.text_bytes()?)
                })
        });
        if aggregate.is_none_or(|size| {
            size > super::super::navigation_entry::MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry aggregate text limit exceeded".to_string(),
            ));
        }
        let index = self.navigation_entries.len();
        self.navigation_entries.push(entry);
        Ok(index)
    }

    /// Atomically append navigation metadata and attach it to one cell story.
    pub fn push_navigation_entry_for_cell(
        &mut self,
        path: &crate::TableCellPath,
        entry: super::super::navigation_entry::NavigationEntry<'a>,
    ) -> RtfResult<usize> {
        let position = entry.position();
        if self
            .table_cell_mut(path)?
            .text()
            .get(position..position)
            .is_none()
        {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry position is outside its table-cell story".to_string(),
            ));
        }
        let index = self.push_cell_navigation_entry_metadata(entry)?;
        if let Err(error) = self
            .table_cell_mut(path)?
            .push_navigation_entry_reference(index, position)
        {
            self.navigation_entries.pop();
            return Err(error);
        }
        Ok(index)
    }

    /// Remove all index and table-of-contents source marks.
    pub fn clear_navigation_entries(&mut self) {
        self.navigation_entries.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::NavigationEntry(_)));
        for table in &mut self.tables {
            table.clear_navigation_entry_references();
        }
    }

    /// Get the list table.
    ///
    /// Returns all list definitions (for bulleted and numbered lists) in the document.
    pub fn list_table(&self) -> &super::super::list::ListTable<'_> {
        &self.list_table
    }

    /// Resolve ordered list-picture records without cloning their image payloads.
    pub fn list_picture_bullets(
        &self,
    ) -> impl ExactSizeIterator<Item = Option<&super::super::picture::Picture<'_>>> {
        (0..self.list_table.picture_bullet_count as usize).map(|slot| {
            self.list_table
                .picture_bullet_picture_indices()
                .get(slot)
                .copied()
                .flatten()
                .and_then(|index| self.pictures.get(index))
        })
    }

    /// Replace list-picture records with validated references into `pictures()`.
    pub fn set_list_picture_bullet_indices(
        &mut self,
        indices: Vec<Option<usize>>,
    ) -> RtfResult<()> {
        if indices
            .iter()
            .flatten()
            .any(|index| *index >= self.pictures.len())
        {
            return Err(RtfError::MalformedDocument(
                "RTF list-picture index is outside the document picture store".to_string(),
            ));
        }
        let old_count = self.list_table.picture_bullet_count;
        let old_indices = self.list_table.picture_bullet_picture_indices().to_vec();
        self.list_table
            .set_picture_bullet_picture_indices(indices)?;
        if let Err(error) = self.list_table.validate() {
            self.list_table.picture_bullet_count = old_count;
            self.list_table
                .set_picture_bullet_picture_indices(old_indices)?;
            self.list_table.picture_bullet_count = old_count;
            return Err(error);
        }
        Ok(())
    }

    /// Clear list-picture references without deleting shared picture payloads.
    pub fn clear_list_picture_bullets(&mut self) -> RtfResult<()> {
        self.set_list_picture_bullet_indices(Vec::new())
    }

    /// Get the list override table.
    ///
    /// Returns list instances that override base list definitions.
    pub fn list_override_table(&self) -> &super::super::list::ListOverrideTable {
        &self.list_override_table
    }

    /// Return ordered legacy `pnseclvl` section-numbering defaults.
    pub fn legacy_section_numbering(&self) -> &crate::LegacySectionNumbering<'_> {
        &self.legacy_section_numbering
    }

    /// Return legacy `pn` records in exact source order.
    pub fn legacy_paragraph_numbering_records(&self) -> &[crate::LegacyParagraphNumbering<'_>] {
        &self.legacy_paragraph_numbering
    }

    /// Resolve the inert `pn` record owned by a paragraph snapshot.
    pub fn legacy_paragraph_numbering(
        &self,
        paragraph: &crate::Paragraph,
    ) -> Option<&crate::LegacyParagraphNumbering<'_>> {
        self.legacy_paragraph_numbering
            .get(paragraph.legacy_numbering? as usize)
    }

    /// Return the inert paragraph-group property table.
    pub fn paragraph_group_table(&self) -> Option<&crate::ParagraphGroupPropertyTable> {
        self.paragraph_group_table.as_ref()
    }

    /// Replace the paragraph-group property table after validation.
    pub fn set_paragraph_group_table(
        &mut self,
        table: crate::ParagraphGroupPropertyTable,
    ) -> RtfResult<()> {
        table.validate()?;
        self.paragraph_group_table = Some(table);
        Ok(())
    }

    /// Remove the paragraph-group property table.
    pub fn clear_paragraph_group_table(&mut self) {
        self.paragraph_group_table = None;
    }

    /// Replace legacy section-numbering defaults after full validation.
    pub fn set_legacy_section_numbering(
        &mut self,
        numbering: crate::LegacySectionNumbering<'a>,
    ) -> RtfResult<()> {
        numbering.validate()?;
        self.legacy_section_numbering = numbering;
        Ok(())
    }

    /// Remove all legacy section-numbering defaults.
    pub fn clear_legacy_section_numbering(&mut self) {
        self.legacy_section_numbering = crate::LegacySectionNumbering::new();
    }

    /// Resolve a paragraph's list override and effective level definition.
    ///
    /// The returned start value is the per-level override when present, followed
    /// by the legacy first-level override used by older producers.
    pub fn resolve_paragraph_list<'s>(
        &'s self,
        paragraph: &crate::Paragraph,
    ) -> Option<(
        &'s super::super::list::ListOverride,
        &'s super::super::list::ListLevel<'s>,
        Option<i32>,
    )> {
        let override_index = paragraph.list_override?;
        let level_index = paragraph.list_level.unwrap_or(0);
        let (list_override, list) = self
            .list_override_table
            .resolve(override_index, &self.list_table)?;
        let level = list
            .levels
            .iter()
            .find(|candidate| candidate.level == level_index)?;
        let start_at = list_override
            .levels
            .iter()
            .find(|candidate| candidate.level == level_index)
            .and_then(|candidate| candidate.start_at)
            .or_else(|| {
                (level_index == 0)
                    .then_some(list_override.start_at_override)
                    .flatten()
            });
        Some((list_override, level, start_at))
    }

    /// Get all sections in the document.
    ///
    /// Returns section information including page layout, headers, and footers.
    pub fn sections(&self) -> &[super::super::section::Section<'_>] {
        &self.sections
    }

    /// Get the bookmark table.
    ///
    /// Returns all bookmarks defined in the document.
    pub fn bookmarks(&self) -> &super::super::bookmark::BookmarkTable<'_> {
        &self.bookmarks
    }

    /// Get all shapes in the document.
    ///
    /// Returns drawing objects, text boxes, and other shapes.
    pub fn shapes(&self) -> &[super::super::shape::Shape<'_>] {
        &self.shapes
    }

    /// Recursively find a body, background, grouped, or text-story shape by name.
    pub fn find_shape_by_name(&self, name: &str) -> Option<&super::super::shape::Shape<'_>> {
        self.shapes
            .iter()
            .find_map(|shape| shape.find_by_name(name))
            .or_else(|| {
                self.shape_groups
                    .iter()
                    .find_map(|group| group.find_shape_by_name(name))
            })
    }

    /// Recursively find a body, background, grouped, or text-story shape by `shplid`.
    pub fn find_shape_by_id(&self, id: i32) -> Option<&super::super::shape::Shape<'_>> {
        self.shapes
            .iter()
            .find_map(|shape| shape.find_by_id(id))
            .or_else(|| {
                self.shape_groups
                    .iter()
                    .find_map(|group| group.find_shape_by_id(id))
            })
    }

    /// Exact source order of root shapes and shape groups in the body story.
    pub fn drawing_order(&self) -> &[crate::StoryDrawing] {
        &self.drawing_order
    }

    pub fn body_story_events(&self) -> &[crate::BodyStoryEvent] {
        &self.body_story_events
    }

    pub(crate) fn body_boundaries(&self) -> &[crate::story::Boundary] {
        &self.body_boundaries
    }

    pub fn page_breaks(&self) -> impl Iterator<Item = &crate::PageBreak> {
        self.body_story_events
            .iter()
            .filter_map(|event| match event {
                crate::BodyStoryEvent::PageBreak(page_break) => Some(page_break),
                _ => None,
            })
    }

    /// Nonrequired (soft) break markers in body source order.
    ///
    /// The markers are passive Galley-view layout metadata; no pagination is
    /// computed from them.
    pub fn soft_breaks(&self) -> impl Iterator<Item = &crate::SoftBreak> {
        self.body_story_events
            .iter()
            .filter_map(|event| match event {
                crate::BodyStoryEvent::SoftBreak(soft_break) => Some(soft_break),
                _ => None,
            })
    }

    /// Explicit main-story section boundaries in source order.
    ///
    /// A boundary with `next_section == None` starts an inherited section that
    /// has no separately retained section definition.
    pub fn section_breaks(&self) -> impl Iterator<Item = &crate::SectionBreak> {
        self.body_story_events
            .iter()
            .filter_map(|event| match event {
                crate::BodyStoryEvent::SectionBreak(section_break) => Some(section_break),
                _ => None,
            })
    }

    pub fn push_page_break(&mut self, position: usize) -> RtfResult<()> {
        let body = self.text();
        if body.get(position..position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF page-break position is not a UTF-8 body boundary".to_string(),
            ));
        }
        self.insert_body_story_event(crate::BodyStoryEvent::PageBreak(crate::PageBreak::new(
            position,
        )))
    }

    pub fn clear_page_breaks(&mut self) {
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::PageBreak(_)));
    }

    pub fn column_breaks(&self) -> impl Iterator<Item = &crate::ColumnBreak> {
        self.body_story_events
            .iter()
            .filter_map(|event| match event {
                crate::BodyStoryEvent::ColumnBreak(column_break) => Some(column_break),
                _ => None,
            })
    }

    pub fn push_column_break(&mut self, position: usize) -> RtfResult<()> {
        let body = self.text();
        if body.get(position..position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF column-break position is not a UTF-8 body boundary".to_string(),
            ));
        }
        self.insert_body_story_event(crate::BodyStoryEvent::ColumnBreak(crate::ColumnBreak::new(
            position,
        )))
    }

    pub fn clear_column_breaks(&mut self) {
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::ColumnBreak(_)));
    }

    /// Append a validated standalone shape at its UTF-8 body position.
    pub fn push_shape(&mut self, shape: super::super::shape::Shape<'a>) -> RtfResult<()> {
        if shape.is_background {
            return Err(RtfError::MalformedDocument(
                "RTF background shapes must use set_background_shape".to_string(),
            ));
        }
        if self.shapes.len() >= MAX_ROOT_SHAPES {
            return Err(RtfError::MalformedDocument(
                "RTF shape count exceeds the safety limit".to_string(),
            ));
        }
        shape.validate()?;
        let body = self.text();
        if body.get(shape.position..shape.position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF shape position is not a UTF-8 body boundary".to_string(),
            ));
        }
        if self
            .shapes
            .iter()
            .rev()
            .find(|shape| !shape.is_background)
            .is_some_and(|previous| previous.position > shape.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF shapes are out of body order".to_string(),
            ));
        }
        if self
            .last_drawing_position()?
            .is_some_and(|position| position > shape.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF body drawing order moves backwards".to_string(),
            ));
        }
        if self
            .last_body_story_position()?
            .is_some_and(|position| position > shape.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF body story order moves backwards".to_string(),
            ));
        }
        let drawing = crate::StoryDrawing::Shape(self.shapes.len());
        let event_at = self.checked_body_story_event_insertion_index(shape.position)?;
        try_reserve_one(&mut self.shapes, "root shapes")?;
        try_reserve_one(&mut self.drawing_order, "root drawing order")?;
        try_reserve_one(&mut self.body_story_events, "body story events")?;
        self.shapes.push(shape);
        self.drawing_order.push(drawing);
        self.body_story_events
            .insert(event_at, crate::BodyStoryEvent::Drawing(drawing));
        Ok(())
    }

    /// Ergonomic alias for appending a validated body shape.
    pub fn add_shape(&mut self, shape: super::super::shape::Shape<'a>) -> RtfResult<usize> {
        let index = self.shapes.len();
        self.push_shape(shape)?;
        Ok(index)
    }

    /// Atomically replace a non-background root shape without relocating its body anchor.
    pub fn replace_shape(
        &mut self,
        index: usize,
        replacement: super::super::shape::Shape<'a>,
    ) -> RtfResult<super::super::shape::Shape<'a>> {
        let current_position = self
            .shapes
            .get(index)
            .ok_or_else(|| {
                RtfError::MalformedDocument(format!("RTF shape index {index} is out of bounds"))
            })?
            .position;
        if self.background_shape_index == Some(index) || replacement.is_background {
            return Err(RtfError::MalformedDocument(
                "RTF background shape must use set_background_shape".to_string(),
            ));
        }
        if replacement.position != current_position {
            return Err(RtfError::MalformedDocument(
                "RTF shape replacement cannot relocate its body anchor".to_string(),
            ));
        }
        replacement.validate()?;
        let current = self.shapes.get_mut(index).ok_or_else(|| {
            RtfError::MalformedDocument(format!("RTF shape index {index} is out of bounds"))
        })?;
        Ok(std::mem::replace(current, replacement))
    }

    /// Atomically remove one non-background root shape and repair every stored index.
    pub fn remove_shape(&mut self, index: usize) -> RtfResult<super::super::shape::Shape<'a>> {
        if index >= self.shapes.len() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF shape index {index} is out of bounds"
            )));
        }
        if self.background_shape_index == Some(index) {
            return Err(RtfError::MalformedDocument(
                "RTF background shape must use clear_background_shape".to_string(),
            ));
        }
        let removed = self.shapes.remove(index);
        self.drawing_order.retain(
            |drawing| !matches!(drawing, crate::StoryDrawing::Shape(value) if *value == index),
        );
        self.body_story_events.retain(|event| !matches!(event, crate::BodyStoryEvent::Drawing(crate::StoryDrawing::Shape(value)) if *value == index));
        for drawing in &mut self.drawing_order {
            if let crate::StoryDrawing::Shape(value) = drawing
                && *value > index
            {
                *value -= 1;
            }
        }
        for event in &mut self.body_story_events {
            if let crate::BodyStoryEvent::Drawing(crate::StoryDrawing::Shape(value)) = event
                && *value > index
            {
                *value -= 1;
            }
        }
        if let Some(background) = &mut self.background_shape_index
            && *background > index
        {
            *background -= 1;
        }
        Ok(removed)
    }

    /// Remove all standalone shapes while preserving the document background.
    pub fn clear_shapes(&mut self) {
        self.drawing_order
            .retain(|drawing| !matches!(drawing, crate::StoryDrawing::Shape(_)));
        self.body_story_events.retain(|event| {
            !matches!(
                event,
                crate::BodyStoryEvent::Drawing(crate::StoryDrawing::Shape(_))
            )
        });
        let background = self
            .background_shape_index
            .filter(|index| self.shapes.get(*index).is_some())
            .map(|index| self.shapes.swap_remove(index));
        self.shapes.clear();
        self.background_shape_index = background.as_ref().map(|_| 0);
        self.shapes.extend(background);
    }

    /// Return the typed shape in the document `background` destination.
    #[must_use]
    pub fn background_shape(&self) -> Option<&super::super::shape::Shape<'_>> {
        self.background_shape_index
            .and_then(|index| self.shapes.get(index))
    }

    /// Set the unique document-background destination shape.
    pub fn set_background_shape(
        &mut self,
        mut shape: super::super::shape::Shape<'a>,
    ) -> RtfResult<()> {
        Self::validate_background_shape(&shape)?;
        shape.is_background = true;
        if let Some(index) = self.background_shape_index {
            let current = self.shapes.get_mut(index).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF background shape index is out of bounds".to_string(),
                )
            })?;
            *current = shape;
        } else {
            if self.shapes.len() >= MAX_ROOT_SHAPES {
                return Err(RtfError::MalformedDocument(
                    "RTF shape count exceeds the safety limit".to_string(),
                ));
            }
            try_reserve_one(&mut self.shapes, "root shapes")?;
            let index = self.shapes.len();
            self.shapes.push(shape);
            self.background_shape_index = Some(index);
        }
        Ok(())
    }

    /// Remove only the destination-owned background shape.
    pub fn clear_background_shape(&mut self) -> Option<super::super::shape::Shape<'a>> {
        let index = self.background_shape_index.take()?;
        self.shapes.get(index)?;
        let removed = self.shapes.remove(index);
        for drawing in &mut self.drawing_order {
            if let crate::StoryDrawing::Shape(shape_index) = drawing
                && *shape_index > index
            {
                *shape_index -= 1;
            }
        }
        for event in &mut self.body_story_events {
            if let crate::BodyStoryEvent::Drawing(crate::StoryDrawing::Shape(shape_index)) = event
                && *shape_index > index
            {
                *shape_index -= 1;
            }
        }
        Some(removed)
    }

    fn validate_background_shape(shape: &super::super::shape::Shape<'_>) -> RtfResult<()> {
        if shape.properties.len() > 65_536 {
            return Err(RtfError::MalformedDocument(
                "RTF background shape property count exceeds the safety limit".to_string(),
            ));
        }
        if shape.text.len() > 16 * 1_048_576 {
            return Err(RtfError::MalformedDocument(
                "RTF background shape text exceeds the safety limit".to_string(),
            ));
        }
        for property in &shape.properties {
            property.validate()?;
            if property.name.len().saturating_add(property.value.len()) > 1_048_576 {
                return Err(RtfError::MalformedDocument(
                    "RTF background shape property exceeds the safety limit".to_string(),
                ));
            }
        }
        if let Some(result) = &shape.result {
            result.validate()?;
        }
        shape
            .geometry
            .x
            .checked_add(shape.geometry.width)
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF background shape horizontal geometry overflows".to_string(),
                )
            })?;
        shape
            .geometry
            .y
            .checked_add(shape.geometry.height)
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF background shape vertical geometry overflows".to_string(),
                )
            })?;
        Ok(())
    }

    /// Return inert positional legacy drawing text boxes.
    pub fn legacy_text_boxes(&self) -> &[crate::LegacyTextBox<'_>] {
        &self.legacy_text_boxes
    }

    pub fn push_legacy_text_box(&mut self, text_box: crate::LegacyTextBox<'a>) -> RtfResult<()> {
        text_box.validate()?;
        if self.legacy_text_boxes.len() >= crate::legacy_text_box::MAX_LEGACY_TEXT_BOXES {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box count exceeds the safety limit".to_string(),
            ));
        }
        let body = self.text();
        if body.get(text_box.position..text_box.position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box position is not a UTF-8 body boundary".to_string(),
            ));
        }
        if self
            .legacy_text_boxes
            .last()
            .is_some_and(|previous| previous.position > text_box.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text boxes are out of body order".to_string(),
            ));
        }
        let total = self
            .legacy_text_boxes
            .iter()
            .try_fold(text_box.text.len(), |total, entry| {
                total.checked_add(entry.text.len())
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF legacy text-box size overflow".to_string())
            })?;
        if total > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box text exceeds the aggregate safety limit".to_string(),
            ));
        }
        let index = self.legacy_text_boxes.len();
        self.legacy_text_boxes.push(text_box);
        self.insert_body_story_event(crate::BodyStoryEvent::LegacyTextBox(index))?;
        Ok(())
    }

    pub fn clear_legacy_text_boxes(&mut self) {
        self.legacy_text_boxes.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::LegacyTextBox(_)));
    }

    /// Return inert positional legacy drawing primitives other than top-level text boxes.
    pub fn legacy_drawings(&self) -> &[crate::LegacyDrawing<'_>] {
        &self.legacy_drawings
    }

    pub fn push_legacy_drawing(&mut self, drawing: crate::LegacyDrawing<'a>) -> RtfResult<()> {
        drawing.validate()?;
        if self.legacy_drawings.len() >= crate::MAX_LEGACY_DRAWINGS {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing count exceeds the safety limit".to_string(),
            ));
        }
        let body = self.text();
        if body.get(drawing.position..drawing.position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawing position is not a UTF-8 body boundary".to_string(),
            ));
        }
        if self
            .legacy_drawings
            .last()
            .is_some_and(|previous| previous.position > drawing.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF legacy drawings are out of body order".to_string(),
            ));
        }
        let index = self.legacy_drawings.len();
        self.legacy_drawings.push(drawing);
        self.insert_body_story_event(crate::BodyStoryEvent::LegacyDrawing(index))?;
        Ok(())
    }

    pub fn clear_legacy_drawings(&mut self) {
        self.legacy_drawings.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::LegacyDrawing(_)));
    }

    /// Get all shape groups in the document.
    ///
    /// Returns grouped shapes.
    pub fn shape_groups(&self) -> &[super::super::shape::ShapeGroup<'_>] {
        &self.shape_groups
    }

    /// Append a validated root shape group.
    pub fn push_shape_group(
        &mut self,
        group: super::super::shape::ShapeGroup<'a>,
    ) -> RtfResult<()> {
        if self.shape_groups.len() >= MAX_ROOT_SHAPE_GROUPS {
            return Err(RtfError::MalformedDocument(
                "RTF shape group count exceeds the safety limit".to_string(),
            ));
        }
        group.validate()?;
        let body = self.text();
        if body.get(group.position..group.position).is_none()
            || self
                .shape_groups
                .last()
                .is_some_and(|previous| previous.position > group.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF shape groups are outside or out of body order".to_string(),
            ));
        }
        if self
            .last_drawing_position()?
            .is_some_and(|position| position > group.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF body drawing order moves backwards".to_string(),
            ));
        }
        if self
            .last_body_story_position()?
            .is_some_and(|position| position > group.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF body story order moves backwards".to_string(),
            ));
        }
        let drawing = crate::StoryDrawing::ShapeGroup(self.shape_groups.len());
        let event_at = self.checked_body_story_event_insertion_index(group.position)?;
        try_reserve_one(&mut self.shape_groups, "root shape groups")?;
        try_reserve_one(&mut self.drawing_order, "root drawing order")?;
        try_reserve_one(&mut self.body_story_events, "body story events")?;
        self.shape_groups.push(group);
        self.drawing_order.push(drawing);
        self.body_story_events
            .insert(event_at, crate::BodyStoryEvent::Drawing(drawing));
        Ok(())
    }

    /// Ergonomic alias for appending a validated root shape group.
    pub fn add_shape_group(
        &mut self,
        group: super::super::shape::ShapeGroup<'a>,
    ) -> RtfResult<usize> {
        let index = self.shape_groups.len();
        self.push_shape_group(group)?;
        Ok(index)
    }

    /// Atomically replace a root group without relocating its body anchor.
    pub fn replace_shape_group(
        &mut self,
        index: usize,
        replacement: super::super::shape::ShapeGroup<'a>,
    ) -> RtfResult<super::super::shape::ShapeGroup<'a>> {
        let current_position = self
            .shape_groups
            .get(index)
            .ok_or_else(|| {
                RtfError::MalformedDocument(format!(
                    "RTF shape-group index {index} is out of bounds"
                ))
            })?
            .position;
        if replacement.position != current_position {
            return Err(RtfError::MalformedDocument(
                "RTF shape-group replacement cannot relocate its body anchor".to_string(),
            ));
        }
        replacement.validate()?;
        let current = self.shape_groups.get_mut(index).ok_or_else(|| {
            RtfError::MalformedDocument(format!("RTF shape-group index {index} is out of bounds"))
        })?;
        Ok(std::mem::replace(current, replacement))
    }

    /// Atomically remove one root group and repair every stored index.
    pub fn remove_shape_group(
        &mut self,
        index: usize,
    ) -> RtfResult<super::super::shape::ShapeGroup<'a>> {
        if index >= self.shape_groups.len() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF shape-group index {index} is out of bounds"
            )));
        }
        let removed = self.shape_groups.remove(index);
        self.drawing_order.retain(
            |drawing| !matches!(drawing, crate::StoryDrawing::ShapeGroup(value) if *value == index),
        );
        self.body_story_events.retain(|event| !matches!(event, crate::BodyStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(value)) if *value == index));
        for drawing in &mut self.drawing_order {
            if let crate::StoryDrawing::ShapeGroup(value) = drawing
                && *value > index
            {
                *value -= 1;
            }
        }
        for event in &mut self.body_story_events {
            if let crate::BodyStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(value)) = event
                && *value > index
            {
                *value -= 1;
            }
        }
        Ok(removed)
    }

    /// Reorder root drawings at the same body anchor without moving unrelated story content.
    pub fn move_drawing(&mut self, from: usize, to: usize) -> RtfResult<()> {
        if from >= self.drawing_order.len() || to >= self.drawing_order.len() {
            return Err(RtfError::MalformedDocument(
                "RTF drawing reorder index is out of bounds".to_string(),
            ));
        }
        if from == to {
            return Ok(());
        }
        let start = from.min(to);
        let end = from.max(to);
        let drawing = self.drawing_order.get(from).copied().ok_or_else(|| {
            RtfError::MalformedDocument("RTF drawing reorder index is out of bounds".to_string())
        })?;
        let anchor = self.root_drawing_position(drawing)?;
        let range = self.drawing_order.get(start..=end).ok_or_else(|| {
            RtfError::MalformedDocument("RTF drawing reorder range is out of bounds".to_string())
        })?;
        for &drawing in range {
            if self.root_drawing_position(drawing)? != anchor {
                return Err(RtfError::MalformedDocument(
                    "RTF drawings at different body anchors cannot be reordered".to_string(),
                ));
            }
        }
        let event_count = self
            .body_story_events
            .iter()
            .filter(|event| matches!(event, crate::BodyStoryEvent::Drawing(_)))
            .count();
        if event_count != self.drawing_order.len() {
            return Err(RtfError::MalformedDocument(
                "RTF drawing event order is incomplete or contains extra events".to_string(),
            ));
        }
        let drawing = self.drawing_order.remove(from);
        self.drawing_order.insert(to, drawing);
        let drawing_events = self.body_story_events.iter_mut().filter_map(|event| {
            if let crate::BodyStoryEvent::Drawing(drawing) = event {
                Some(drawing)
            } else {
                None
            }
        });
        for (event, drawing) in drawing_events.zip(self.drawing_order.iter().copied()) {
            *event = drawing;
        }
        Ok(())
    }

    fn root_drawing_position(&self, drawing: crate::StoryDrawing) -> RtfResult<usize> {
        match drawing {
            crate::StoryDrawing::Shape(index) => self
                .shapes
                .get(index)
                .map(|shape| shape.position)
                .ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF drawing order references a missing shape".to_string(),
                    )
                }),
            crate::StoryDrawing::ShapeGroup(index) => self
                .shape_groups
                .get(index)
                .map(|group| group.position)
                .ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF drawing order references a missing shape group".to_string(),
                    )
                }),
        }
    }

    /// Remove all root shape groups.
    pub fn clear_shape_groups(&mut self) {
        self.shape_groups.clear();
        self.drawing_order
            .retain(|drawing| !matches!(drawing, crate::StoryDrawing::ShapeGroup(_)));
        self.body_story_events.retain(|event| {
            !matches!(
                event,
                crate::BodyStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(_))
            )
        });
    }

    fn last_drawing_position(&self) -> RtfResult<Option<usize>> {
        self.drawing_order
            .last()
            .copied()
            .map(|drawing| self.root_drawing_position(drawing))
            .transpose()
    }

    fn last_body_story_position(&self) -> RtfResult<Option<usize>> {
        for event in self.body_story_events.iter().rev() {
            if matches!(
                *event,
                crate::BodyStoryEvent::Drawing(_)
                    | crate::BodyStoryEvent::Field(_)
                    | crate::BodyStoryEvent::PageBreak(_)
                    | crate::BodyStoryEvent::ColumnBreak(_)
                    | crate::BodyStoryEvent::SectionBreak(_)
            ) {
                return self
                    .body_story_event_position(*event)
                    .map(Some)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF body story event references missing metadata".to_string(),
                        )
                    });
            }
        }
        Ok(None)
    }

    fn body_story_event_position(&self, event: crate::BodyStoryEvent) -> Option<usize> {
        Some(match event {
            crate::BodyStoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                self.shapes.get(index)?.position
            },
            crate::BodyStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                self.shape_groups.get(index)?.position
            },
            crate::BodyStoryEvent::Field(index) => self.fields.get(index)?.position,
            crate::BodyStoryEvent::PageBreak(page_break) => page_break.position,
            crate::BodyStoryEvent::ColumnBreak(column_break) => column_break.position,
            crate::BodyStoryEvent::SectionBreak(section_break) => section_break.position,
            crate::BodyStoryEvent::BookmarkStart(index) => {
                self.bookmarks.bookmarks().get(index)?.position
            },
            crate::BodyStoryEvent::BookmarkEnd(index) => {
                let bookmark = self.bookmarks.bookmarks().get(index)?;
                bookmark.position.checked_add(bookmark.content.len())?
            },
            crate::BodyStoryEvent::AnnotationStart(index) => self.annotations.get(index)?.position,
            crate::BodyStoryEvent::AnnotationEnd(index) => self.annotations.get(index)?.range_end,
            crate::BodyStoryEvent::Note(index) => self.notes.get(index)?.position,
            crate::BodyStoryEvent::Object(index) => self.objects.get(index)?.position,
            crate::BodyStoryEvent::PictureCompatibility(index) => {
                self.picture_compatibility_records.get(index)?.position
            },
            crate::BodyStoryEvent::FormFieldStart(index) => self.form_fields.get(index)?.position,
            crate::BodyStoryEvent::FormFieldEnd(index) => self.form_fields.get(index)?.range_end,
            crate::BodyStoryEvent::RevisionStart(index) => self.revisions.get(index)?.position,
            crate::BodyStoryEvent::RevisionEnd(index) => self.revisions.get(index)?.range_end,
            crate::BodyStoryEvent::RevisionDeletion(index) => self.revisions.get(index)?.position,
            crate::BodyStoryEvent::GeneratedListMarker(index) => {
                self.generated_list_markers.get(index)?.position
            },
            crate::BodyStoryEvent::LegacyTextBox(index) => {
                self.legacy_text_boxes.get(index)?.position
            },
            crate::BodyStoryEvent::LegacyDrawing(index) => {
                self.legacy_drawings.get(index)?.position
            },
            crate::BodyStoryEvent::NavigationEntry(index) => {
                self.navigation_entries.get(index)?.position()
            },
            crate::BodyStoryEvent::CustomXmlOpen(index) => {
                self.custom_xml_tags.get(index)?.position
            },
            crate::BodyStoryEvent::CustomXmlClose(index) => {
                let tag = self.custom_xml_tags.get(index)?;
                tag.position.checked_add(tag.content.len())?
            },
            crate::BodyStoryEvent::MathZone(index) => self.math_zones.get(index)?.position,
            crate::BodyStoryEvent::SoftBreak(soft_break) => soft_break.position,
            crate::BodyStoryEvent::ProtectionRangeStart(index) => {
                self.protection_ranges.get(index)?.position
            },
            crate::BodyStoryEvent::ProtectionRangeEnd(index) => {
                let range = self.protection_ranges.get(index)?;
                range.position.checked_add(range.content.len())?
            },
            crate::BodyStoryEvent::EditableRegionStart(index) => {
                self.editable_regions.get(index)?.position
            },
            crate::BodyStoryEvent::EditableRegionEnd(index) => {
                let region = self.editable_regions.get(index)?;
                region.position.checked_add(region.content.len())?
            },
        })
    }

    fn checked_body_story_event_insertion_index(&self, position: usize) -> RtfResult<usize> {
        let mut at = 0usize;
        let mut previous_position = None;
        for (index, event) in self.body_story_events.iter().enumerate() {
            let event_position = self.body_story_event_position(*event).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF body story event references missing metadata".to_string(),
                )
            })?;
            if previous_position.is_some_and(|previous| previous > event_position) {
                return Err(RtfError::MalformedDocument(
                    "RTF body story events are out of order".to_string(),
                ));
            }
            if event_position <= position {
                at = index.checked_add(1).ok_or_else(|| {
                    RtfError::MalformedDocument("RTF body story event index overflows".to_string())
                })?;
            }
            previous_position = Some(event_position);
        }
        Ok(at)
    }

    fn insert_body_story_event(&mut self, event: crate::BodyStoryEvent) -> RtfResult<()> {
        let position = self.body_story_event_position(event).ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF body story event references missing metadata".to_string(),
            )
        })?;
        let at = self
            .body_story_events
            .iter()
            .rposition(|existing| {
                self.body_story_event_position(*existing)
                    .is_some_and(|value| value <= position)
            })
            .map_or(0, |index| index + 1);
        self.body_story_events.insert(at, event);
        Ok(())
    }

    /// Get the stylesheet.
    ///
    /// Returns style definitions for paragraphs and characters.
    pub fn stylesheet(&self) -> &super::super::stylesheet::StyleSheet<'_> {
        &self.stylesheet
    }

    /// Get document information/metadata.
    ///
    /// Returns document properties like title, author, subject, etc.
    pub fn info(&self) -> &super::super::info::DocumentInfo<'_> {
        &self.info
    }

    /// Return inert document and revision-protection metadata.
    pub fn protection(&self) -> &crate::DocumentProtection<'_> {
        &self.info.protection
    }

    /// Replace inert document-protection metadata.
    pub fn set_protection(&mut self, protection: crate::DocumentProtection<'a>) -> RtfResult<()> {
        protection.validate()?;
        self.info.protection = protection;
        Ok(())
    }

    /// Remove all document-protection metadata.
    pub fn clear_protection(&mut self) {
        self.info.protection = crate::DocumentProtection::default();
    }

    /// Get all annotations (comments) in the document.
    ///
    /// Returns document annotations and revisions.
    pub fn annotations(&self) -> &[super::super::annotation::Annotation<'_>] {
        &self.annotations
    }

    /// Append an inert comment annotation after validating its body range.
    pub fn push_annotation(
        &mut self,
        annotation: super::super::annotation::Annotation<'a>,
    ) -> RtfResult<()> {
        annotation.validate()?;
        let body = self.text();
        if body
            .get(annotation.position..annotation.range_end)
            .is_none()
        {
            return Err(RtfError::MalformedDocument(
                "RTF annotation range is outside body text or splits a character".to_string(),
            ));
        }
        if self.annotations.len() >= super::super::annotation::MAX_ANNOTATIONS {
            return Err(RtfError::MalformedDocument(
                "RTF annotation count limit exceeded".to_string(),
            ));
        }
        if annotation.has_reference
            && self
                .annotations
                .iter()
                .any(|existing| existing.has_reference && existing.id == annotation.id)
        {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF annotation reference".to_string(),
            ));
        }
        let aggregate = annotation.text_bytes().and_then(|initial| {
            self.annotations.iter().try_fold(initial, |size, existing| {
                size.checked_add(existing.text_bytes()?)
            })
        });
        if aggregate
            .is_none_or(|size| size > super::super::annotation::MAX_ANNOTATION_TEXT_TOTAL_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "RTF annotation aggregate text limit exceeded".to_string(),
            ));
        }
        let index = self.annotations.len();
        self.annotations.push(annotation);
        self.insert_body_story_event(crate::BodyStoryEvent::AnnotationStart(index))?;
        self.insert_body_story_event(crate::BodyStoryEvent::AnnotationEnd(index))?;
        Ok(())
    }

    /// Remove all comment annotations.
    pub fn clear_annotations(&mut self) {
        self.annotations.clear();
        self.body_story_events.retain(|event| {
            !matches!(
                event,
                crate::BodyStoryEvent::AnnotationStart(_) | crate::BodyStoryEvent::AnnotationEnd(_)
            )
        });
    }

    // Helper methods to convert borrowed data to owned
    //
    // These methods are used internally during parsing to convert borrowed data
    // (tied to the input lifetime) to owned data (with 'static lifetime).
    // This allows the parsed document to outlive the input string.

    /// Convert list table to owned
    fn convert_list_table_to_owned(
        table: super::super::list::ListTable<'_>,
    ) -> RtfResult<super::super::list::ListTable<'static>> {
        let mut owned = super::super::list::ListTable::new();
        for list in table.lists() {
            owned.add(super::super::list::List {
                id: list.id,
                template_id: list.template_id,
                simple: list.simple,
                hybrid: list.hybrid,
                name: Cow::Owned(list.name.clone().into_owned()),
                style_name: Cow::Owned(list.style_name.clone().into_owned()),
                style_priority: list.style_priority,
                levels: list
                    .levels
                    .iter()
                    .map(|level| super::super::list::ListLevel {
                        level: level.level,
                        level_type: level.level_type,
                        number_text: Cow::Owned(level.number_text.clone().into_owned()),
                        number_positions: Cow::Owned(level.number_positions.clone().into_owned()),
                        start_at: level.start_at,
                        justification: level.justification,
                        follow_previous: level.follow_previous,
                        follow: level.follow,
                        font_ref: level.font_ref,
                        indent: level.indent,
                        space: level.space,
                        left_indent: level.left_indent,
                        first_line_indent: level.first_line_indent,
                        tabs: level.tabs.clone(),
                        picture_index: level.picture_index,
                        tentative: level.tentative,
                        legal_format: level.legal_format,
                        no_restart: level.no_restart,
                        legacy: level.legacy,
                        include_previous: level.include_previous,
                        include_previous_space: level.include_previous_space,
                        template_id: level.template_id,
                    })
                    .collect(),
            });
        }
        owned.picture_bullet_count = table.picture_bullet_count;
        owned
            .set_picture_bullet_picture_indices(table.picture_bullet_picture_indices().to_vec())?;
        owned.picture_bullet_count = table.picture_bullet_count;
        Ok(owned)
    }

    /// Convert sections to owned
    fn convert_sections_to_owned(
        sections: Vec<super::super::section::Section<'_>>,
    ) -> Vec<super::super::section::Section<'static>> {
        sections
            .into_iter()
            .map(|section| super::super::section::Section {
                properties: section.properties,
                headers_footers: section
                    .headers_footers
                    .into_iter()
                    .map(|header_footer| super::super::section::HeaderFooter {
                        header_type: header_footer.header_type,
                        paragraphs: header_footer
                            .paragraphs
                            .into_iter()
                            .map(|paragraph| super::super::section::HeaderFooterParagraph {
                                text: Cow::Owned(paragraph.text.into_owned()),
                                formatting: paragraph.formatting,
                                paragraph: paragraph.paragraph,
                            })
                            .collect(),
                        shapes: Self::convert_shapes_to_owned(header_footer.shapes),
                        shape_groups: Self::convert_shape_groups_to_owned(
                            header_footer.shape_groups,
                        ),
                        drawing_order: header_footer.drawing_order,
                        story_events: header_footer.story_events,
                    })
                    .collect(),
            })
            .collect()
    }

    /// Convert bookmarks to owned
    fn convert_bookmarks_to_owned(
        bookmarks: super::super::bookmark::BookmarkTable<'_>,
    ) -> super::super::bookmark::BookmarkTable<'static> {
        let mut owned = super::super::bookmark::BookmarkTable::new();
        for bookmark in bookmarks.bookmarks() {
            owned.add(super::super::bookmark::Bookmark {
                name: Cow::Owned(bookmark.name.clone().into_owned()),
                position: bookmark.position,
                content: Cow::Owned(bookmark.content.clone().into_owned()),
                first_column: bookmark.first_column,
                last_column: bookmark.last_column,
                is_public: bookmark.is_public,
            });
        }
        owned
    }

    /// Convert shapes to owned
    fn convert_shapes_to_owned(
        shapes: Vec<super::super::shape::Shape<'_>>,
    ) -> Vec<super::super::shape::Shape<'static>> {
        shapes
            .into_iter()
            .map(Self::convert_shape_to_owned)
            .collect()
    }

    /// Convert shape groups to owned
    fn convert_shape_groups_to_owned(
        groups: Vec<super::super::shape::ShapeGroup<'_>>,
    ) -> Vec<super::super::shape::ShapeGroup<'static>> {
        groups
            .into_iter()
            .map(Self::convert_shape_group_to_owned)
            .collect()
    }

    fn convert_shape_group_to_owned(
        group: super::super::shape::ShapeGroup<'_>,
    ) -> super::super::shape::ShapeGroup<'static> {
        super::super::shape::ShapeGroup {
            position: group.position,
            name: Cow::Owned(group.name.into_owned()),
            shapes: group
                .shapes
                .into_iter()
                .map(Self::convert_shape_to_owned)
                .collect(),
            groups: group
                .groups
                .into_iter()
                .map(Self::convert_shape_group_to_owned)
                .collect(),
            child_order: group.child_order,
            info: group.info,
            geometry: group.geometry,
            properties: group
                .properties
                .into_iter()
                .map(Self::convert_shape_property_to_owned)
                .collect(),
            result: group.result.map(crate::ShapeResult::into_owned),
        }
    }

    fn convert_shape_to_owned(
        shape: super::super::shape::Shape<'_>,
    ) -> super::super::shape::Shape<'static> {
        super::super::shape::Shape {
            position: shape.position,
            instruction_present: shape.instruction_present,
            shape_type: shape.shape_type,
            geometry: shape.geometry,
            fill: shape.fill,
            border: shape.border,
            line: shape.line,
            text: Cow::Owned(shape.text.into_owned()),
            text_destination_present: shape.text_destination_present,
            text_formatting: shape.text_formatting,
            text_shapes: Self::convert_shapes_to_owned(shape.text_shapes),
            text_shape_groups: Self::convert_shape_groups_to_owned(shape.text_shape_groups),
            text_drawing_order: shape.text_drawing_order,
            text_story_events: shape.text_story_events,
            wrap_mode: shape.wrap_mode,
            behind_doc: shape.behind_doc,
            is_background: shape.is_background,
            locked: shape.locked,
            name: Cow::Owned(shape.name.into_owned()),
            properties: shape
                .properties
                .into_iter()
                .map(Self::convert_shape_property_to_owned)
                .collect(),
            result: shape
                .result
                .map(super::super::shape::ShapeResult::into_owned),
            info: shape.info,
        }
    }

    fn convert_shape_property_to_owned(
        property: super::super::shape::ShapeProperty<'_>,
    ) -> super::super::shape::ShapeProperty<'static> {
        super::super::shape::ShapeProperty {
            name: Cow::Owned(property.name.into_owned()),
            value: Cow::Owned(property.value.into_owned()),
            binary_value: property
                .binary_value
                .map(|value| Cow::Owned(value.into_owned())),
            theme_value: property.theme_value,
            hyperlink: property.hyperlink.map(crate::ShapeHyperlink::into_owned),
        }
    }

    /// Convert stylesheet to owned
    fn convert_stylesheet_to_owned(
        stylesheet: super::super::stylesheet::StyleSheet<'_>,
    ) -> super::super::stylesheet::StyleSheet<'static> {
        let mut owned = super::super::stylesheet::StyleSheet::new();
        for style in stylesheet.styles() {
            owned.add(super::super::stylesheet::Style {
                id: style.id,
                name: Cow::Owned(style.name.clone().into_owned()),
                style_type: style.style_type,
                based_on: style.based_on,
                next_style: style.next_style,
                linked_style: style.linked_style,
                formatting: style.formatting,
                paragraph: style.paragraph,
                table_conditional: style.table_conditional,
                builtin: style.builtin,
                hidden: style.hidden,
                additive: style.additive,
                auto_update: style.auto_update,
                locked: style.locked,
                semi_hidden: style.semi_hidden,
                unhide_when_used: style.unhide_when_used,
                quick_format: style.quick_format,
                priority: style.priority,
                revision_id: style.revision_id,
                personal: style.personal,
                compose: style.compose,
                reply: style.reply,
            });
        }
        owned
    }

    /// Convert document info to owned
    fn convert_info_to_owned(
        info: super::super::info::DocumentInfo<'_>,
    ) -> super::super::info::DocumentInfo<'static> {
        super::super::info::DocumentInfo {
            title: info.title.map(|value| Cow::Owned(value.into_owned())),
            subject: info.subject.map(|value| Cow::Owned(value.into_owned())),
            author: info.author.map(|value| Cow::Owned(value.into_owned())),
            manager: info.manager.map(|value| Cow::Owned(value.into_owned())),
            company: info.company.map(|value| Cow::Owned(value.into_owned())),
            operator: info.operator.map(|value| Cow::Owned(value.into_owned())),
            category: info.category.map(|value| Cow::Owned(value.into_owned())),
            keywords: info.keywords.map(|value| Cow::Owned(value.into_owned())),
            comment: info.comment.map(|value| Cow::Owned(value.into_owned())),
            document_comment: info
                .document_comment
                .map(|value| Cow::Owned(value.into_owned())),
            hyperlink_base: info
                .hyperlink_base
                .map(|value| Cow::Owned(value.into_owned())),
            version: info.version,
            revision: info.revision,
            creation_time: info
                .creation_time
                .map(|value| Cow::Owned(value.into_owned())),
            creation_timestamp: info.creation_timestamp,
            revision_time: info
                .revision_time
                .map(|value| Cow::Owned(value.into_owned())),
            revision_timestamp: info.revision_timestamp,
            print_time: info.print_time.map(|value| Cow::Owned(value.into_owned())),
            print_timestamp: info.print_timestamp,
            backup_time: info.backup_time.map(|value| Cow::Owned(value.into_owned())),
            backup_timestamp: info.backup_timestamp,
            editing_time: info.editing_time,
            pages: info.pages,
            words: info.words,
            characters: info.characters,
            characters_with_spaces: info.characters_with_spaces,
            id: info.id,
            protection: info.protection.into_owned(),
        }
    }

    /// Convert annotations to owned
    fn convert_annotations_to_owned(
        annotations: Vec<super::super::annotation::Annotation<'_>>,
    ) -> Vec<super::super::annotation::Annotation<'static>> {
        annotations
            .into_iter()
            .map(|annotation| super::super::annotation::Annotation {
                annotation_type: annotation.annotation_type,
                id: annotation.id,
                has_reference: annotation.has_reference,
                author: Cow::Owned(annotation.author.into_owned()),
                initials: Cow::Owned(annotation.initials.into_owned()),
                date: annotation.date.map(|value| Cow::Owned(value.into_owned())),
                text: Cow::Owned(annotation.text.into_owned()),
                shapes: Self::convert_shapes_to_owned(annotation.shapes),
                shape_groups: Self::convert_shape_groups_to_owned(annotation.shape_groups),
                drawing_order: annotation.drawing_order,
                story_events: annotation.story_events,
                position: annotation.position,
                range_end: annotation.range_end,
                parent_id: annotation
                    .parent_id
                    .map(|value| Cow::Owned(value.into_owned())),
                icon: annotation.icon.map(|value| Cow::Owned(value.into_owned())),
                time: annotation.time.map(|value| Cow::Owned(value.into_owned())),
            })
            .collect()
    }

    /// Convert notes to owned
    fn convert_notes_to_owned(
        notes: Vec<super::super::section::Note<'_>>,
    ) -> Vec<super::super::section::Note<'static>> {
        notes
            .into_iter()
            .map(|note| super::super::section::Note {
                position: note.position,
                is_footnote: note.is_footnote,
                reference: Cow::Owned(note.reference.into_owned()),
                content: Cow::Owned(note.content.into_owned()),
                formatting: note.formatting,
                shapes: Self::convert_shapes_to_owned(note.shapes),
                shape_groups: Self::convert_shape_groups_to_owned(note.shape_groups),
                drawing_order: note.drawing_order,
                story_events: note.story_events,
            })
            .collect()
    }

    /// Convert revisions to owned
    fn convert_revisions_to_owned(
        revisions: Vec<super::super::annotation::Revision<'_>>,
    ) -> Vec<super::super::annotation::Revision<'static>> {
        revisions
            .into_iter()
            .map(|rev| super::super::annotation::Revision {
                revision_type: rev.revision_type,
                author: Cow::Owned(rev.author.into_owned()),
                date: rev.date.map(|d| Cow::Owned(d.into_owned())),
                id: rev.id,
                content: Cow::Owned(rev.content.into_owned()),
                position: rev.position,
                range_end: rev.range_end,
            })
            .collect()
    }

    /// Get all footnotes and endnotes in the document.
    pub fn notes(&self) -> &[super::super::section::Note<'_>] {
        &self.notes
    }

    /// Append a validated footnote or endnote at a UTF-8 main-story boundary.
    pub fn push_note(&mut self, note: super::super::section::Note<'a>) -> RtfResult<()> {
        note.validate()?;
        if self.text().get(note.position..note.position).is_none()
            || self
                .notes
                .last()
                .is_some_and(|previous| previous.position > note.position)
        {
            return Err(RtfError::MalformedDocument(
                "RTF notes are outside or out of main-story order".to_string(),
            ));
        }
        if self.notes.len() >= super::super::section::MAX_NOTES {
            return Err(RtfError::MalformedDocument(
                "RTF note count exceeds the safety limit".to_string(),
            ));
        }
        let aggregate = note.text_bytes().and_then(|initial| {
            self.notes.iter().try_fold(initial, |size, existing| {
                size.checked_add(existing.text_bytes()?)
            })
        });
        if aggregate.is_none_or(|size| size > super::super::section::MAX_NOTE_TEXT_TOTAL_BYTES) {
            return Err(RtfError::MalformedDocument(
                "RTF note aggregate text exceeds the safety limit".to_string(),
            ));
        }
        let index = self.notes.len();
        self.notes.push(note);
        self.insert_body_story_event(crate::BodyStoryEvent::Note(index))?;
        Ok(())
    }

    /// Remove all footnotes and endnotes.
    pub fn clear_notes(&mut self) {
        self.notes.clear();
        self.body_story_events
            .retain(|event| !matches!(event, crate::BodyStoryEvent::Note(_)));
    }

    /// Return explicit document-level footnote and endnote configuration.
    pub fn note_options(&self) -> &crate::NoteOptions {
        &self.note_options
    }

    /// Replace document-level footnote and endnote configuration after validation.
    pub fn set_note_options(&mut self, options: crate::NoteOptions) -> RtfResult<()> {
        options.validate()?;
        self.note_options = options;
        Ok(())
    }

    /// Return ordered footnote and endnote separator destinations.
    pub fn note_separators(&self) -> &crate::NoteSeparatorTable<'_> {
        &self.note_separators
    }

    /// Replace note-separator destinations after validation.
    pub fn set_note_separators(
        &mut self,
        separators: crate::NoteSeparatorTable<'a>,
    ) -> RtfResult<()> {
        separators.validate()?;
        self.note_separators = separators;
        Ok(())
    }

    /// Remove all note-separator destinations.
    pub fn clear_note_separators(&mut self) {
        self.note_separators = crate::NoteSeparatorTable::new();
    }

    /// Get all footnotes in the document.
    pub fn footnotes(&self) -> Vec<&super::super::section::Note<'_>> {
        self.notes.iter().filter(|n| n.is_footnote).collect()
    }

    /// Get all endnotes in the document.
    pub fn endnotes(&self) -> Vec<&super::super::section::Note<'_>> {
        self.notes.iter().filter(|n| !n.is_footnote).collect()
    }

    /// Get all track changes/revisions in the document.
    pub fn revisions(&self) -> &[super::super::annotation::Revision<'_>] {
        &self.revisions
    }

    /// Get the ordered revision-author table.
    pub fn revision_authors(&self) -> &[super::super::annotation::RevisionAuthor<'_>] {
        &self.revision_authors
    }

    /// Append an entry to the ordered revision-author table.
    pub fn push_revision_author(
        &mut self,
        author: super::super::annotation::RevisionAuthor<'a>,
    ) -> RtfResult<()> {
        author.validate()?;
        if self.revision_authors.len() >= super::super::annotation::MAX_REVISION_AUTHORS {
            return Err(RtfError::MalformedDocument(
                "RTF revision author count exceeds the safety limit".to_string(),
            ));
        }
        let total = self
            .revision_authors
            .iter()
            .try_fold(author.name.len(), |total, existing| {
                total.checked_add(existing.name.len())
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF aggregate revision-author size overflow".to_string(),
                )
            })?;
        if total > super::super::annotation::MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision-author text exceeds the safety limit".to_string(),
            ));
        }
        self.revision_authors.push(author);
        Ok(())
    }

    /// Remove the revision-author table when no revision still references it.
    pub fn clear_revision_authors(&mut self) -> RtfResult<()> {
        if !self.revisions.is_empty() {
            return Err(RtfError::MalformedDocument(
                "cannot clear an RTF revision-author table while revisions reference it"
                    .to_string(),
            ));
        }
        self.revision_authors.clear();
        Ok(())
    }

    /// Append a validated tracked change.
    pub fn push_revision(
        &mut self,
        revision: super::super::annotation::Revision<'a>,
    ) -> RtfResult<()> {
        revision.validate()?;
        if self.revisions.len() >= super::super::annotation::MAX_REVISIONS {
            return Err(RtfError::MalformedDocument(
                "RTF revision count exceeds the safety limit".to_string(),
            ));
        }
        let author_index = usize::try_from(revision.id).map_err(|_| {
            RtfError::MalformedDocument("RTF revision author index cannot be negative".to_string())
        })?;
        let author = self.revision_authors.get(author_index).ok_or_else(|| {
            RtfError::MalformedDocument("RTF revision author index is outside revtbl".to_string())
        })?;
        if author.name != revision.author {
            return Err(RtfError::MalformedDocument(
                "RTF revision author does not match its revtbl entry".to_string(),
            ));
        }

        let body = self.text();
        match revision.revision_type {
            super::super::annotation::RevisionType::Insertion => {
                let content = body
                    .get(revision.position..revision.range_end)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF insertion range is outside body text or splits a character"
                                .to_string(),
                        )
                    })?;
                if content != revision.content {
                    return Err(RtfError::MalformedDocument(
                        "RTF insertion content does not match its visible body range".to_string(),
                    ));
                }
                if self.revisions.iter().any(|existing| {
                    existing.revision_type == super::super::annotation::RevisionType::Insertion
                        && revision.position < existing.range_end
                        && existing.position < revision.range_end
                }) {
                    return Err(RtfError::MalformedDocument(
                        "RTF insertion ranges cannot overlap".to_string(),
                    ));
                }
            },
            super::super::annotation::RevisionType::Deletion => {
                if body.get(..revision.position).is_none() {
                    return Err(RtfError::MalformedDocument(
                        "RTF deletion position is outside body text or splits a character"
                            .to_string(),
                    ));
                }
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "this RTF revision kind has no lossless scoped-run representation".to_string(),
                ));
            },
        }
        let total = self
            .revisions
            .iter()
            .try_fold(revision.content.len(), |total, existing| {
                total.checked_add(existing.content.len())
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF aggregate revision size overflow".to_string())
            })?;
        if total > super::super::annotation::MAX_REVISION_TEXT_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision text exceeds the safety limit".to_string(),
            ));
        }
        let index = self.revisions.len();
        let kind = revision.revision_type;
        self.revisions.push(revision);
        match kind {
            super::super::annotation::RevisionType::Insertion => {
                self.insert_body_story_event(crate::BodyStoryEvent::RevisionStart(index))?;
                self.insert_body_story_event(crate::BodyStoryEvent::RevisionEnd(index))?;
            },
            super::super::annotation::RevisionType::Deletion => {
                self.insert_body_story_event(crate::BodyStoryEvent::RevisionDeletion(index))?;
            },
            _ => {
                self.revisions.pop();
                return Err(RtfError::MalformedDocument(
                    "this RTF revision kind has no lossless scoped-run representation".to_string(),
                ));
            },
        }
        Ok(())
    }

    /// Append tracked-change metadata for ownership by a table-cell story.
    pub fn push_cell_revision_metadata(
        &mut self,
        revision: super::super::annotation::Revision<'a>,
    ) -> RtfResult<usize> {
        revision.validate()?;
        if self.revisions.len() >= super::super::annotation::MAX_REVISIONS {
            return Err(RtfError::MalformedDocument(
                "RTF revision count exceeds the safety limit".to_string(),
            ));
        }
        let author_index = usize::try_from(revision.id).map_err(|_| {
            RtfError::MalformedDocument("RTF revision author index cannot be negative".to_string())
        })?;
        if self
            .revision_authors
            .get(author_index)
            .is_none_or(|author| author.name != revision.author)
        {
            return Err(RtfError::MalformedDocument(
                "RTF revision author is missing from or does not match revtbl".to_string(),
            ));
        }
        let total = self
            .revisions
            .iter()
            .try_fold(revision.content.len(), |total, existing| {
                total.checked_add(existing.content.len())
            });
        if total.is_none_or(|total| total > super::super::annotation::MAX_REVISION_TEXT_TOTAL_BYTES)
        {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision text exceeds the safety limit".to_string(),
            ));
        }
        let index = self.revisions.len();
        self.revisions.push(revision);
        Ok(index)
    }

    /// Atomically append tracked-change metadata and attach its event(s) to one cell story.
    pub fn push_revision_for_cell(
        &mut self,
        path: &crate::TableCellPath,
        revision: super::super::annotation::Revision<'a>,
    ) -> RtfResult<usize> {
        let kind = revision.revision_type;
        let position = revision.position;
        let range_end = revision.range_end;
        let cell = self.table_cell_mut(path)?;
        match kind {
            super::super::annotation::RevisionType::Insertion
                if cell.text().get(position..range_end) == Some(revision.content.as_ref()) => {},
            super::super::annotation::RevisionType::Deletion
                if cell.text().get(position..position).is_some() => {},
            super::super::annotation::RevisionType::Insertion => {
                return Err(RtfError::MalformedDocument(
                    "RTF insertion revision does not match its table-cell range".to_string(),
                ));
            },
            super::super::annotation::RevisionType::Deletion => {
                return Err(RtfError::MalformedDocument(
                    "RTF deletion revision is outside its table-cell story".to_string(),
                ));
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "this RTF revision kind has no lossless scoped-run representation".to_string(),
                ));
            },
        }
        let index = self.push_cell_revision_metadata(revision)?;
        let result = match kind {
            super::super::annotation::RevisionType::Insertion => self
                .table_cell_mut(path)?
                .push_insertion_revision_reference(index, position, range_end),
            super::super::annotation::RevisionType::Deletion => self
                .table_cell_mut(path)?
                .push_deletion_revision_reference(index, position),
            _ => {
                self.revisions.pop();
                return Err(RtfError::MalformedDocument(
                    "this RTF revision kind has no lossless scoped-run representation".to_string(),
                ));
            },
        };
        if let Err(error) = result {
            self.revisions.pop();
            return Err(error);
        }
        Ok(index)
    }

    /// Remove all tracked changes while retaining the ordered author table.
    pub fn clear_revisions(&mut self) {
        self.revisions.clear();
        self.body_story_events.retain(|event| {
            !matches!(
                event,
                crate::BodyStoryEvent::RevisionStart(_)
                    | crate::BodyStoryEvent::RevisionEnd(_)
                    | crate::BodyStoryEvent::RevisionDeletion(_)
            )
        });
        for table in &mut self.tables {
            table.clear_revision_references();
        }
    }
}
