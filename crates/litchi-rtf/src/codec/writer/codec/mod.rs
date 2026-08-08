#![cfg_attr(not(test), deny(clippy::indexing_slicing))]

//! RTF document writer/serializer.
//!
//! This module provides functionality to write RTF documents from structured data.
//! It supports all RTF features including formatting, tables, pictures, fields, lists, and more.

use super::super::*;
use litchi_codepage::Mbcs;
use std::collections::HashSet;
use std::io::{self, Write};

mod output;
mod semantic;

/// RTF 1.9.1 effective default tab width when `deftab` is omitted.
pub const DEFAULT_TAB_WIDTH_TWIPS: u32 = 720;

/// Largest default tab width representable by an RTF numeric parameter.
pub const MAX_DEFAULT_TAB_WIDTH_TWIPS: u32 = i32::MAX as u32;

/// First non-control ASCII code point; anything below it is escaped in body text.
const ASCII_CONTROL_LIMIT: u32 = 0x20;

/// Controls how a writer resolves document `deftab` metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DefaultTabWidthPolicy {
    /// Preserve an explicit source value or preserve source omission.
    PreserveDocument,
    /// Emit this value, replacing either a source value or source omission.
    Override(u32),
}

/// RTF writer options
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriterOptions {
    /// Checked legacy character set declared by the document header.
    pub charset: Charset,
    /// Indent RTF output for readability
    pub indent: bool,
    /// Default font index
    pub default_font: u16,
    /// Explicit precedence policy for the default tab width.
    pub default_tab_width: DefaultTabWidthPolicy,
}

/// Legacy character set declared by an RTF document header.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Charset {
    /// `\ansi` with an exact byte-stream code page.
    Ansi(Mbcs),
    /// Macintosh Roman (`\mac`).
    Mac,
    /// IBM PC code page 437 (`\pc`).
    Pc,
    /// IBM PC code page 850 (`\pca`).
    Pca,
}

impl Charset {
    /// Windows-1252 ANSI declaration.
    pub const WINDOWS_1252: Self = Self::Ansi(Mbcs::WINDOWS_1252);

    /// Validate a raw byte-stream page for an ANSI declaration.
    pub fn ansi(page: u32) -> std::result::Result<Self, litchi_codepage::Error> {
        Mbcs::require(page).map(Self::Ansi)
    }
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self {
            charset: Charset::WINDOWS_1252,
            indent: false,
            default_font: 0,
            default_tab_width: DefaultTabWidthPolicy::PreserveDocument,
        }
    }
}

/// RTF document writer
///
/// Provides functionality to serialize RTF documents to a writer.
/// All fields are used internally during the writing process.
pub struct RtfWriter<W: Write> {
    /// Output writer
    writer: W,
    /// Writer options
    options: WriterOptions,
    /// Current indentation level (reserved for formatted output)
    #[allow(dead_code)]
    indent_level: usize,
    /// Font table
    font_table: FontTable<'static>,
    /// Color table
    color_table: ColorTable,
    legacy_paragraph_numbering: Vec<crate::LegacyParagraphNumbering<'static>>,
}

#[derive(Clone, Copy)]
enum BodyEventKind<'b, 'a> {
    Shape(&'b crate::Shape<'a>),
    ShapeGroup(&'b crate::ShapeGroup<'a>),
    Object(&'b crate::EmbeddedObject<'a>, &'b [crate::Picture<'a>]),
    PictureCompatibility(
        &'b crate::PictureCompatibilityRecord,
        &'b crate::Picture<'a>,
    ),
    GeneratedListMarker(&'b crate::GeneratedListMarker<'a>),
    LegacyTextBox(&'b crate::LegacyTextBox<'a>),
    LegacyDrawing(&'b crate::LegacyDrawing<'a>),
    NavigationEntry(&'b crate::NavigationEntry<'a>),
    BookmarkStart(&'b Bookmark<'a>),
    BookmarkEnd(&'b Bookmark<'a>),
    CustomXmlOpen(&'b crate::CustomXmlTag<'a>),
    CustomXmlClose(&'b crate::CustomXmlTag<'a>),
    MathZone(&'b crate::MathZone<'a>),
    ProtectionRangeStart(&'b crate::ProtectionRange<'a>),
    ProtectionRangeEnd(&'b crate::ProtectionRange<'a>),
    EditableRegionStart(&'b crate::EditableRegion<'a>),
    EditableRegionEnd(&'b crate::EditableRegion<'a>),
    SoftBreak(crate::SoftBreak),
    AnnotationStart(&'b Annotation<'a>),
    AnnotationEnd(&'b Annotation<'a>),
    Note(&'b Note<'a>),
    RevisionStart(&'b Revision<'a>),
    RevisionEnd,
    RevisionDeletion(&'b Revision<'a>),
    FormFieldStart(&'b crate::FormField<'a>),
    FormFieldEnd,
    GenericField(&'b crate::Field<'a>),
    PageBreak,
    ColumnBreak,
    SectionBreak(Option<&'b Section<'a>>),
    Opaque(&'b crate::opaque::Node),
}

#[derive(Clone, Copy)]
struct BodyEvent<'b, 'a> {
    offset: usize,
    order: u8,
    kind: BodyEventKind<'b, 'a>,
}

#[derive(Clone, Copy)]
enum DrawingStoryTextMode {
    Destination,
    Note,
    ShapeText,
}

fn invalid_story_reference() -> io::Error {
    io::Error::new(
        io::ErrorKind::InvalidInput,
        "RTF story order has an invalid or duplicate reference",
    )
}

/// Resolve and claim an indexed story resource without trusting parallel-vector invariants.
fn take_story_item<'a, T>(items: &'a [T], seen: &mut [bool], index: usize) -> io::Result<&'a T> {
    let item = items.get(index).ok_or_else(invalid_story_reference)?;
    let seen = seen.get_mut(index).ok_or_else(invalid_story_reference)?;
    if std::mem::replace(seen, true) {
        return Err(invalid_story_reference());
    }
    Ok(item)
}

impl<W: Write> RtfWriter<W> {
    /// Create a new RTF writer
    pub fn new(writer: W) -> Self {
        Self::with_options(writer, WriterOptions::default())
    }

    /// Create a new RTF writer with options
    pub fn with_options(writer: W, options: WriterOptions) -> Self {
        Self {
            writer,
            options,
            indent_level: 0,
            font_table: FontTable::new(),
            color_table: ColorTable::new(),
            legacy_paragraph_numbering: Vec::new(),
        }
    }

    /// Serialize an immutable document snapshot.
    pub fn write(&mut self, document: &crate::Document) -> io::Result<()> {
        if self.options == WriterOptions::default()
            && let Some(source) = document.model().preserved_source()
        {
            return self.writer.write_all(source);
        }
        self.write_document(document.model())
    }

    /// Write a complete RTF document
    pub fn write_document<'a>(&mut self, doc: &RtfDocument<'a>) -> io::Result<()> {
        if doc
            .opaque_nodes()
            .iter()
            .any(|node| matches!(node.anchor(), crate::opaque::Anchor::Structural { .. }))
        {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF canonical rewrite cannot preserve a structurally scoped opaque node",
            ));
        }
        Self::validate_table_story_metadata_ownership(doc)?;
        Self::validate_generic_field_ownership(doc)?;
        Self::validate_section_boundary_mapping(doc.sections(), doc.body_story_events())?;
        crate::story::validate_boundaries(doc.blocks(), doc.body_boundaries())
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        // Collect font and color tables from document by cloning them
        // We need to convert the lifetime to 'static for storage
        let font_table: FontTable<'static> = FontTable {
            fonts: doc
                .font_table()
                .fonts()
                .iter()
                .map(|f| Font {
                    name: std::borrow::Cow::Owned(f.name.to_string()),
                    family: f.family,
                    theme: f.theme,
                    charset: f.charset,
                    alternate_name: f
                        .alternate_name
                        .as_ref()
                        .map(|name| std::borrow::Cow::Owned(name.to_string())),
                    non_tagged_name: f
                        .non_tagged_name
                        .as_ref()
                        .map(|name| std::borrow::Cow::Owned(name.to_string())),
                    panose: f.panose,
                    pitch: f.pitch,
                    code_page: f.code_page,
                    embedded: f.embedded.clone().map(crate::EmbeddedFont::into_owned),
                    bidi: f.bidi,
                })
                .collect(),
            defined: doc.font_table().defined.clone(),
        };
        let color_table = doc.color_table().clone();

        self.font_table = font_table;
        self.color_table = color_table;
        self.legacy_paragraph_numbering = doc
            .legacy_paragraph_numbering_records()
            .iter()
            .cloned()
            .map(crate::LegacyParagraphNumbering::into_owned)
            .collect();

        // Write document header
        let default_tab_width = match self.options.default_tab_width {
            DefaultTabWidthPolicy::PreserveDocument => doc.default_tab_width_twips(),
            DefaultTabWidthPolicy::Override(width) => Some(width),
        };
        self.write_document_header_with_origin(
            doc.origin_metadata().origin,
            default_tab_width,
            Some(&doc.default_formatting().fonts),
        )?;

        self.write_language_defaults(doc.language_defaults())?;

        self.write_document_direction(doc)?;

        self.write_document_hyphenation(doc.hyphenation())?;

        self.write_document_external_references(doc.external_references())?;

        self.write_document_file_settings(doc.file_settings())?;

        self.write_document_output_settings(doc.output_settings())?;

        self.write_document_rendering_settings(doc.rendering_settings())?;

        self.write_document_processing_settings(doc.processing_settings())?;

        self.write_document_drawing_grid(doc.drawing_grid())?;

        self.write_document_print_layout_settings(doc.print_layout_settings())?;

        self.write_document_theme_languages(doc.theme_languages())?;
        self.write_document_booklet_printing(doc.booklet_printing())?;
        self.write_document_privacy_policies(doc.privacy_policies())?;
        self.write_document_line_spacing_compatibility(doc.line_spacing_compatibility())?;
        self.write_document_east_asian_compatibility(doc.east_asian_compatibility())?;
        self.write_document_table_layout_compatibility(doc.table_layout_compatibility())?;
        self.write_document_legacy_layout_compatibility(doc.legacy_layout_compatibility())?;
        self.write_document_asian_grid_compatibility(doc.asian_grid_compatibility())?;
        self.write_document_compatibility_policy(doc.compatibility_policy())?;
        self.write_document_word_2003_compatibility(doc.word_2003_compatibility())?;

        self.write_document_xml_policies(doc.xml_policies())?;

        self.write_document_embedding_policies(doc.embedding_policies())?;

        self.write_document_revision_policies(doc.revision_policies())?;

        self.write_document_style_policies(doc.style_policies())?;

        self.write_document_style_restrictions(doc.style_restrictions())?;

        self.write_document_view(doc.document_view())?;

        self.write_review_display(doc.review_display())?;

        self.write_window_caption(doc.window_caption())?;

        self.write_kinsoku(doc.kinsoku())?;

        self.write_document_auto_format_type(doc.origin_metadata().auto_format_type)?;

        self.write_xsl_transform(doc.xsl_transform())?;

        self.write_xsl_transform_usage(doc.xsl_transform_usage())?;

        self.write_style_list_filter(doc.style_list_filter())?;

        self.write_style_sort_method(doc.style_sort_method())?;

        self.write_document_write_reservations(doc.write_reservations())?;

        self.write_document_save_preferences(doc.save_preferences())?;

        // Write font table
        self.write_font_table()?;

        // Write inert external-file metadata without resolving any names.
        self.write_file_table(doc.file_table())?;

        // Write color table
        self.write_color_table()?;

        self.write_default_formatting_destinations(doc.default_formatting())?;

        // Write named paragraph, character, section, and table styles.
        self.write_stylesheet(doc.stylesheet())?;
        self.write_note_options(doc.note_options())?;
        self.write_note_separators(doc.note_separators())?;

        // Write list definitions before body paragraphs reference them.
        doc.list_override_table()
            .validate(doc.list_table())
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_list_table_with_pictures(doc.list_table(), doc.pictures())?;
        self.write_list_override_table(doc.list_override_table())?;
        self.write_paragraph_group_table(doc.paragraph_group_table())?;
        self.write_legacy_section_numbering(doc.legacy_section_numbering())?;

        // Revision controls reference this author table by numeric index.
        self.write_revision_table(doc.revision_authors(), doc.revisions())?;

        self.write_revision_save_metadata(doc.revision_save_metadata())?;

        self.write_xml_namespace_table(doc.xml_namespaces())?;

        self.write_protection_user_table(doc.protection_user_table())?;

        self.write_theme(doc.theme())?;

        self.write_latent_styles(doc.latent_styles())?;

        self.write_data_store(doc.data_store())?;

        self.write_mail_merge(doc.mail_merge())?;

        self.write_math_properties(doc.math_properties())?;

        self.write_document_background(doc.background_shape())?;

        // Producer provenance is inert header metadata.
        self.write_generator(doc.generator())?;

        // Write document properties before body content.
        self.write_document_info(doc.info())?;

        // User-defined properties are header-level inert metadata.
        self.write_user_properties(doc.user_properties())?;

        // Document variables are header-level inert metadata.
        self.write_document_variables(doc.document_variables())?;

        // The first explicit section definition precedes body text unless it is
        // introduced later by a retained `\sect` boundary.
        let first_section_is_boundary_scoped = doc.body_story_events().iter().any(|event| {
            matches!(
                event,
                crate::BodyStoryEvent::SectionBreak(section_break)
                    if section_break.next_section == Some(0)
            )
        });
        if !first_section_is_boundary_scoped && let Some(section) = doc.sections().first() {
            self.write_section_with_fields(section, doc.fields())?;
        }

        // Write document content and reinsert positional bookmark/comment markers.
        self.write_blocks_with_markup(
            doc.blocks(),
            doc.body_boundaries(),
            doc.bookmarks(),
            doc.custom_xml_tags(),
            doc.math_zones(),
            doc.protection_ranges(),
            doc.editable_regions(),
            doc.annotations(),
            doc.notes(),
            doc.revisions(),
            doc.navigation_entries(),
            doc.generated_list_markers(),
            doc.shapes(),
            doc.shape_groups(),
            doc.drawing_order(),
            doc.picture_compatibility_records(),
            doc.pictures(),
            doc.objects(),
            doc.legacy_text_boxes(),
            doc.legacy_drawings(),
            doc.form_fields(),
            doc.fields(),
            doc.sections(),
            doc.body_story_events(),
            doc.opaque_nodes(),
        )?;

        // Write tables
        let mut logical_tables = 0usize;
        for table in doc.tables() {
            Self::validate_table_tree(table, 1, &mut logical_tables)?;
        }
        for (index, table) in doc.tables().iter().enumerate() {
            if index > 0 {
                self.write_control_word("pard", None)?;
                self.write_control_word("par", None)?;
                self.write_str("\n")?;
            }
            self.write_table(
                table,
                doc.fields(),
                doc.navigation_entries(),
                doc.revisions(),
            )?;
        }

        // Close document
        self.write_str("}")?;

        Ok(())
    }
}
