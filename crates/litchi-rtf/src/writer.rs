//! RTF document writer/serializer.
//!
//! This module provides functionality to write RTF documents from structured data.
//! It supports all RTF features including formatting, tables, pictures, fields, lists, and more.

use super::*;
use std::io::{self, Write};

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
#[derive(Debug, Clone)]
pub struct WriterOptions {
    /// Use ANSI code page
    pub use_ansi: bool,
    /// ANSI code page number (default 1252 for Western European)
    pub code_page: u16,
    /// Indent RTF output for readability
    pub indent: bool,
    /// Default font index
    pub default_font: u16,
    /// Explicit precedence policy for the default tab width.
    pub default_tab_width: DefaultTabWidthPolicy,
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self {
            use_ansi: true,
            code_page: 1252,
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

    /// Write a complete RTF document
    pub fn write_document<'a>(&mut self, doc: &RtfDocument<'a>) -> io::Result<()> {
        Self::validate_table_story_metadata_ownership(doc)?;
        Self::validate_generic_field_ownership(doc)?;
        Self::validate_section_boundary_mapping(doc.sections(), doc.body_story_events())?;
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
        if !first_section_is_boundary_scoped
            && let Some(section) = doc.sections().first()
        {
            self.write_section_with_fields(section, doc.fields())?;
        }

        // Write document content and reinsert positional bookmark/comment markers.
        self.write_blocks_with_markup(
            doc.blocks(),
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

    fn validate_section_boundary_mapping(
        sections: &[Section<'_>],
        body_story_events: &[crate::BodyStoryEvent],
    ) -> io::Result<()> {
        let first_section_is_boundary_scoped = body_story_events.iter().any(|event| {
            matches!(
                event,
                crate::BodyStoryEvent::SectionBreak(section_break)
                    if section_break.next_section == Some(0)
            )
        });
        let mut next_section_index = if first_section_is_boundary_scoped {
            0
        } else {
            usize::from(!sections.is_empty())
        };
        for event in body_story_events {
            if let crate::BodyStoryEvent::SectionBreak(section_break) = *event
                && let Some(index) = section_break.next_section
            {
                if index != next_section_index || index >= sections.len() {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF section boundary has an invalid or out-of-order section reference",
                    ));
                }
                next_section_index += 1;
            }
        }
        if next_section_index != sections.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF section definitions are missing main-story boundaries",
            ));
        }
        Ok(())
    }

    fn validate_table_story_metadata_ownership(doc: &RtfDocument<'_>) -> io::Result<()> {
        let mut navigation_owners = vec![0u8; doc.navigation_entries().len()];
        let mut revision_owners = vec![0u8; doc.revisions().len()];
        let mut body_starts = vec![false; doc.revisions().len()];
        let mut body_ends = vec![false; doc.revisions().len()];
        let mut body_deletions = vec![false; doc.revisions().len()];
        for event in doc.body_story_events() {
            match *event {
                crate::BodyStoryEvent::NavigationEntry(index) => {
                    let owner = navigation_owners.get_mut(index).ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "RTF body navigation index is out of range")
                    })?;
                    *owner = owner.checked_add(1).ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "RTF navigation ownership overflow")
                    })?;
                },
                crate::BodyStoryEvent::RevisionStart(index) => {
                    let revision = doc.revisions().get(index).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF body revision index is out of range"))?;
                    if revision.revision_type != RevisionType::Insertion || std::mem::replace(&mut body_starts[index], true) {
                        return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF body revision start has the wrong kind or is duplicated"));
                    }
                },
                crate::BodyStoryEvent::RevisionEnd(index) => {
                    let revision = doc.revisions().get(index).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF body revision index is out of range"))?;
                    if revision.revision_type != RevisionType::Insertion || std::mem::replace(&mut body_ends[index], true) {
                        return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF body revision end has the wrong kind or is duplicated"));
                    }
                },
                crate::BodyStoryEvent::RevisionDeletion(index) => {
                    let revision = doc.revisions().get(index).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF body revision index is out of range"))?;
                    if revision.revision_type != RevisionType::Deletion || std::mem::replace(&mut body_deletions[index], true) {
                        return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF body deletion has the wrong kind or is duplicated"));
                    }
                },
                _ => {},
            }
        }
        for (index, revision) in doc.revisions().iter().enumerate() {
            let owned = match revision.revision_type {
                RevisionType::Insertion => body_starts[index] || body_ends[index],
                RevisionType::Deletion => body_deletions[index],
                _ => false,
            };
            if owned {
                let valid = match revision.revision_type {
                    RevisionType::Insertion => body_starts[index] && body_ends[index] && !body_deletions[index],
                    RevisionType::Deletion => body_deletions[index] && !body_starts[index] && !body_ends[index],
                    _ => false,
                };
                if !valid {
                    return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF body revision ownership is incomplete or conflicting"));
                }
                revision_owners[index] = 1;
            }
        }
        for table in doc.tables() {
            Self::validate_table_metadata_tree(
                table,
                doc.navigation_entries(),
                doc.revisions(),
                &mut navigation_owners,
                &mut revision_owners,
            )?;
        }
        if navigation_owners.iter().any(|owners| *owners != 1) {
            return Err(io::Error::new(io::ErrorKind::InvalidInput, "every RTF navigation entry must be owned by exactly one story"));
        }
        if revision_owners.iter().any(|owners| *owners != 1) {
            return Err(io::Error::new(io::ErrorKind::InvalidInput, "every RTF revision must be owned by exactly one story"));
        }
        Ok(())
    }

    fn validate_table_metadata_tree(
        table: &Table<'_>,
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
        navigation_owners: &mut [u8],
        revision_owners: &mut [u8],
    ) -> io::Result<()> {
        for row in table.rows() {
            for cell in row.cells() {
                let mut starts = vec![false; revisions.len()];
                let mut ends = vec![false; revisions.len()];
                let mut deletions = vec![false; revisions.len()];
                for event in cell.story_events() {
                    match *event {
                        crate::CellStoryEvent::NavigationEntry(reference) => {
                            let entry = navigation_entries.get(reference.index).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell navigation index is out of range"))?;
                            if entry.position() != reference.position || cell.text().get(reference.position..reference.position).is_none() {
                                return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell navigation anchor is invalid"));
                            }
                            let owners = &mut navigation_owners[reference.index];
                            *owners = owners.checked_add(1).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF navigation ownership overflow"))?;
                        },
                        crate::CellStoryEvent::RevisionStart(reference) => {
                            let revision = revisions.get(reference.index).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell revision index is out of range"))?;
                            if revision.revision_type != RevisionType::Insertion
                                || revision.position != reference.position
                                || cell.text().get(revision.position..revision.range_end) != Some(revision.content.as_ref())
                                || std::mem::replace(&mut starts[reference.index], true)
                            {
                                return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell revision start is invalid, duplicated, or has the wrong kind"));
                            }
                        },
                        crate::CellStoryEvent::RevisionEnd(reference) => {
                            let revision = revisions.get(reference.index).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell revision index is out of range"))?;
                            if revision.revision_type != RevisionType::Insertion
                                || revision.range_end != reference.position
                                || std::mem::replace(&mut ends[reference.index], true)
                            {
                                return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell revision end is invalid, duplicated, or has the wrong kind"));
                            }
                        },
                        crate::CellStoryEvent::RevisionDeletion(reference) => {
                            let revision = revisions.get(reference.index).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell revision index is out of range"))?;
                            if revision.revision_type != RevisionType::Deletion
                                || revision.position != reference.position
                                || cell.text().get(reference.position..reference.position).is_none()
                                || std::mem::replace(&mut deletions[reference.index], true)
                            {
                                return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell deletion is invalid, duplicated, or has the wrong kind"));
                            }
                        },
                        _ => {},
                    }
                }
                for (index, revision) in revisions.iter().enumerate() {
                    let touched = starts[index] || ends[index] || deletions[index];
                    if touched {
                        let valid = match revision.revision_type {
                            RevisionType::Insertion => starts[index] && ends[index] && !deletions[index],
                            RevisionType::Deletion => deletions[index] && !starts[index] && !ends[index],
                            _ => false,
                        };
                        if !valid {
                            return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell revision ownership is incomplete or conflicting"));
                        }
                        revision_owners[index] = revision_owners[index].checked_add(1).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF revision ownership overflow"))?;
                    }
                }
                for nested in cell.nested_tables() {
                    Self::validate_table_metadata_tree(
                        &nested.table,
                        navigation_entries,
                        revisions,
                        navigation_owners,
                        revision_owners,
                    )?;
                }
            }
        }
        Ok(())
    }

    fn mark_owned_field(
        reference: crate::StoryField,
        owner: crate::FieldOwner,
        fields: &[crate::Field<'_>],
        seen: &mut [bool],
    ) -> io::Result<()> {
        let field = fields
            .get(reference.field_index)
            .filter(|field| {
                field.owner == owner
                    && field.position == reference.position
                    && field.range_end == reference.position
            })
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF story has an invalid generic-field owner or reference",
                )
            })?;
        field
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if std::mem::replace(&mut seen[reference.field_index], true) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field is referenced by multiple owning stories",
            ));
        }
        Ok(())
    }

    fn mark_table_fields(
        table: &Table<'_>,
        depth: u8,
        fields: &[crate::Field<'_>],
        seen: &mut [bool],
    ) -> io::Result<()> {
        for row in table.rows() {
            for cell in row.cells() {
                for event in cell.story_events() {
                    if let crate::CellStoryEvent::Field(reference) = *event {
                        Self::mark_owned_field(
                            reference,
                            crate::FieldOwner::TableCell(depth),
                            fields,
                            seen,
                        )?;
                    }
                }
                for nested in cell.nested_tables() {
                    Self::mark_table_fields(
                        &nested.table,
                        depth.checked_add(1).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table nesting depth overflow",
                            )
                        })?,
                        fields,
                        seen,
                    )?;
                }
            }
        }
        Ok(())
    }

    fn validate_generic_field_ownership(doc: &RtfDocument<'_>) -> io::Result<()> {
        let fields = doc.fields();
        if fields.len() > crate::field::MAX_GENERIC_FIELDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field count exceeds the safety limit",
            ));
        }
        let mut seen = vec![false; fields.len()];
        for event in doc.body_story_events() {
            if let crate::BodyStoryEvent::Field(index) = *event {
                let position = fields.get(index).map_or(usize::MAX, |field| field.position);
                Self::mark_owned_field(
                    crate::StoryField {
                        field_index: index,
                        position,
                    },
                    crate::FieldOwner::Body,
                    fields,
                    &mut seen,
                )?;
            }
        }
        for section in doc.sections() {
            for hf in &section.headers_footers {
                let owner = match hf.header_type {
                    HeaderFooterType::Header
                    | HeaderFooterType::HeaderFirst
                    | HeaderFooterType::HeaderLeft
                    | HeaderFooterType::HeaderRight => crate::FieldOwner::Header,
                    HeaderFooterType::Footer
                    | HeaderFooterType::FooterFirst
                    | HeaderFooterType::FooterLeft
                    | HeaderFooterType::FooterRight => crate::FieldOwner::Footer,
                };
                for event in &hf.story_events {
                    if let crate::StoryEvent::Field(reference) = *event {
                        Self::mark_owned_field(reference, owner, fields, &mut seen)?;
                    }
                }
            }
        }
        for note in doc.notes() {
            for event in &note.story_events {
                if let crate::StoryEvent::Field(reference) = *event {
                    Self::mark_owned_field(
                        reference,
                        if note.is_footnote {
                            crate::FieldOwner::Footnote
                        } else {
                            crate::FieldOwner::Endnote
                        },
                        fields,
                        &mut seen,
                    )?;
                }
            }
        }
        for table in doc.tables() {
            Self::mark_table_fields(table, 1, fields, &mut seen)?;
        }
        for field in fields {
            for event in &field.result_events {
                if let crate::StoryEvent::Field(reference) = *event {
                    Self::mark_owned_field(
                        reference,
                        crate::FieldOwner::FieldResult,
                        fields,
                        &mut seen,
                    )?;
                }
            }
        }
        if fields.iter().zip(seen).any(|(field, seen)| {
            matches!(
                field.owner,
                crate::FieldOwner::Body
                    | crate::FieldOwner::Header
                    | crate::FieldOwner::Footer
                    | crate::FieldOwner::Footnote
                    | crate::FieldOwner::Endnote
                    | crate::FieldOwner::TableCell(_)
                    | crate::FieldOwner::FieldResult
            ) && !seen
        }) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field lacks its concrete owning-story event",
            ));
        }
        Ok(())
    }

    /// Write document header
    pub fn write_document_header(&mut self) -> io::Result<()> {
        let default_tab_width = match self.options.default_tab_width {
            DefaultTabWidthPolicy::PreserveDocument => None,
            DefaultTabWidthPolicy::Override(width) => Some(width),
        };
        self.write_document_header_with_origin(None, default_tab_width, None)
    }

    fn write_document_header_with_origin(
        &mut self,
        origin: Option<crate::DocumentOrigin>,
        default_tab_width: Option<u32>,
        default_fonts: Option<&crate::DocumentDefaultFonts>,
    ) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("rtf", Some(1))?;

        if self.options.use_ansi {
            self.write_control_word("ansi", None)?;
            self.write_control_word("ansicpg", Some(self.options.code_page as i32))?;
        }

        match origin {
            Some(crate::DocumentOrigin::PlainTextEmail) => {
                self.write_control_word("fromtext", None)?;
            },
            Some(crate::DocumentOrigin::HtmlEmail { version }) => {
                self.write_control_word(
                    "fromhtml",
                    version.map(crate::HtmlEmailVersion::rtf_value),
                )?;
            },
            None => {},
        }

        self.write_control_word(
            "deff",
            Some(i32::from(
                default_fonts
                    .and_then(|fonts| fonts.primary)
                    .unwrap_or(self.options.default_font),
            )),
        )?;
        if let Some(fonts) = default_fonts {
            if let Some(value) = fonts.associated {
                self.write_control_word("adeff", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_double_byte {
                self.write_control_word("stshfdbch", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_low_ansi {
                self.write_control_word("stshfloch", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_high_ansi {
                self.write_control_word("stshfhich", Some(i32::from(value)))?;
            }
            if let Some(value) = fonts.stylesheet_bidi {
                self.write_control_word("stshfbi", Some(i32::from(value)))?;
            }
        }
        if let Some(width) = default_tab_width {
            let width = i32::try_from(width).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    format!("RTF deftab width {width} exceeds {MAX_DEFAULT_TAB_WIDTH_TWIPS}"),
                )
            })?;
            self.write_control_word("deftab", Some(width))?;
        }

        Ok(())
    }

    pub fn write_document_auto_format_type(
        &mut self,
        document_type: Option<crate::DocumentAutoFormatType>,
    ) -> io::Result<()> {
        if let Some(document_type) = document_type {
            self.write_control_word("doctype", Some(document_type.rtf_value()))?;
        }
        Ok(())
    }

    pub fn write_language_defaults(
        &mut self,
        defaults: &crate::DocumentLanguageDefaults,
    ) -> io::Result<()> {
        defaults
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(language) = defaults.primary {
            self.write_control_word("deflang", Some(language.rtf_value()))?;
        }
        if let Some(language) = defaults.east_asian {
            self.write_control_word("deflangfe", Some(language.rtf_value()))?;
        }
        if let Some(language) = defaults.complex_script {
            self.write_control_word("adeflang", Some(language.rtf_value()))?;
        }
        Ok(())
    }

    pub fn write_document_direction(&mut self, doc: &RtfDocument<'_>) -> io::Result<()> {
        if let Some(direction) = doc.document_direction() {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrdoc",
                    TextDirection::RightToLeft => "rtldoc",
                },
                None,
            )?;
        }
        if doc.gutter_on_right() {
            self.write_control_word("rtlgutter", None)?;
        }
        Ok(())
    }

    /// Write explicit passive document-level hyphenation settings.
    pub fn write_document_hyphenation(
        &mut self,
        hyphenation: &crate::DocumentHyphenation,
    ) -> io::Result<()> {
        hyphenation
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(value) = hyphenation.hot_zone_twips {
            self.write_control_word("hyphhotz", Some(value as i32))?;
        }
        if let Some(value) = hyphenation.consecutive_line_limit {
            self.write_control_word("hyphconsec", Some(value as i32))?;
        }
        if let Some(value) = hyphenation.capitalized_words {
            self.write_control_word("hyphcaps", Some(i32::from(value)))?;
        }
        if let Some(value) = hyphenation.automatic {
            self.write_control_word("hyphauto", Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write inert names without opening, resolving, or invoking referenced files.
    pub fn write_document_external_references(
        &mut self,
        references: &crate::DocumentExternalReferences<'_>,
    ) -> io::Result<()> {
        references
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (control, value) in [
            ("nextfile", references.next_file.as_deref()),
            ("template", references.template.as_deref()),
        ] {
            let Some(value) = value else { continue };
            self.write_str("{\\*")?;
            self.write_control_word(control, None)?;
            self.write_str(" ")?;
            self.write_destination_text(value)?;
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write passive compatibility and output flags in stable specification order.
    pub fn write_document_output_settings(
        &mut self,
        settings: &crate::DocumentOutputSettings,
    ) -> io::Result<()> {
        if settings.word97_compatibility_marker {
            self.write_control_word("muser", None)?;
        }
        if settings.postscript_over_text {
            self.write_control_word("psover", None)?;
        }
        Ok(())
    }

    /// Write passive rendering flags in stable specification order.
    pub fn write_document_rendering_settings(
        &mut self,
        settings: &crate::DocumentRenderingSettings,
    ) -> io::Result<()> {
        if let Some(orientation) = settings.orientation {
            self.write_control_word(
                match orientation {
                    crate::DocumentRenderingOrientation::Horizontal => "horzdoc",
                    crate::DocumentRenderingOrientation::Vertical => "vertdoc",
                },
                None,
            )?;
        }
        if let Some(justification_mode) = settings.justification_mode {
            self.write_control_word(
                match justification_mode {
                    crate::DocumentJustificationMode::Compress => "jcompress",
                    crate::DocumentJustificationMode::Expand => "jexpand",
                },
                None,
            )?;
        }
        if settings.line_based_on_grid {
            self.write_control_word("lnongrid", None)?;
        }
        Ok(())
    }

    /// Write passive printing, cleanup, and event properties in stable order.
    pub fn write_document_processing_settings(
        &mut self,
        settings: &crate::DocumentProcessingSettings,
    ) -> io::Result<()> {
        if settings.fractional_character_widths_for_printing {
            self.write_control_word("fracwidth", None)?;
        }
        if let Some(cleanup) = settings.abstract_numbering_cleanup {
            self.write_control_word(
                "ilfomacatclnup",
                Some(match cleanup {
                    crate::AbstractNumberingCleanupStatus::Reviewed => 0,
                    crate::AbstractNumberingCleanupStatus::Incomplete => 1,
                }),
            )?;
        }
        if let Some(event_mask) = settings.event_mask {
            self.write_control_word("grfdocevents", Some(i32::from(event_mask.bits())))?;
        }
        Ok(())
    }

    /// Write passive document-level drawing-grid properties in specification order.
    pub fn write_document_drawing_grid(
        &mut self,
        grid: &crate::DocumentDrawingGrid,
    ) -> io::Result<()> {
        if let Some(value) = grid.horizontal_spacing {
            self.write_control_word("dghspace", Some(i32::from(value.get())))?;
        }
        if let Some(value) = grid.vertical_spacing {
            self.write_control_word("dgvspace", Some(i32::from(value.get())))?;
        }
        if let Some(value) = grid.horizontal_origin_twips {
            self.write_control_word("dghorigin", Some(i32::from(value)))?;
        }
        if let Some(value) = grid.vertical_origin_twips {
            self.write_control_word("dgvorigin", Some(i32::from(value)))?;
        }
        if let Some(value) = grid.horizontal_line_interval {
            self.write_control_word("dghshow", Some(i32::from(value.get())))?;
        }
        if let Some(value) = grid.vertical_line_interval {
            self.write_control_word("dgvshow", Some(i32::from(value.get())))?;
        }
        if grid.snap_to_grid {
            self.write_control_word("dgsnap", None)?;
        }
        if grid.follows_margins {
            self.write_control_word("dgmargin", None)?;
        }
        Ok(())
    }

    /// Write passive print-layout settings in stable specification order.
    pub fn write_document_print_layout_settings(
        &mut self,
        settings: &crate::DocumentPrintLayoutSettings,
    ) -> io::Result<()> {
        settings
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if settings.facing_pages {
            self.write_control_word("facingp", None)?;
        }
        if settings.mirror_margins {
            self.write_control_word("margmirror", None)?;
        }
        if let Some(value) = settings.document_gutter_twips {
            self.write_control_word("gutter", Some(value as i32))?;
        }
        if settings.parallel_gutter {
            self.write_control_word("gutterprl", None)?;
        }
        if settings.two_logical_pages_per_physical_page {
            self.write_control_word("twoonone", None)?;
        }
        Ok(())
    }

    /// Write passive theme languages in stable specification order.
    pub fn write_document_theme_languages(
        &mut self,
        languages: &crate::DocumentThemeLanguages,
    ) -> io::Result<()> {
        if let Some(language) = languages.primary {
            self.write_control_word("themelang", Some(language.rtf_value()))?;
        }
        if let Some(language) = languages.east_asian {
            self.write_control_word("themelangfe", Some(language.rtf_value()))?;
        }
        if let Some(language) = languages.complex_script {
            self.write_control_word("themelangcs", Some(language.rtf_value()))?;
        }
        Ok(())
    }

    /// Write passive web-save and custom-XML policies in specification order.
    pub fn write_document_xml_policies(
        &mut self,
        policies: &crate::DocumentXmlPolicies,
    ) -> io::Result<()> {
        for (name, value) in [
            ("relyonvml", policies.rely_on_vml),
            ("validatexml", policies.validate_custom_xml),
            ("showplaceholdtext", policies.show_placeholder_text),
            ("ignoremixedcontent", policies.ignore_mixed_content),
            ("saveinvalidxml", policies.save_invalid_xml),
            ("showxmlerrors", policies.show_xml_errors),
        ] {
            if let Some(value) = value {
                self.write_control_word(name, Some(i32::from(value)))?;
            }
        }
        Ok(())
    }

    /// Write passive embedding policies in stable specification order.
    pub fn write_document_embedding_policies(
        &mut self,
        policies: &crate::DocumentEmbeddingPolicies,
    ) -> io::Result<()> {
        if let Some(value) = policies.do_not_embed_system_fonts {
            self.write_control_word("donotembedsysfont", Some(i32::from(value)))?;
        }
        if let Some(value) = policies.do_not_embed_linguistic_data {
            self.write_control_word("donotembedlingdata", Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write passive revision policies in stable specification order.
    pub fn write_document_revision_policies(
        &mut self,
        policies: &crate::DocumentRevisionPolicies,
    ) -> io::Result<()> {
        if let Some(value) = policies.track_moves {
            self.write_control_word("trackmoves", Some(i32::from(value)))?;
        }
        if let Some(value) = policies.track_formatting {
            self.write_control_word("trackformatting", Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write passive style policies in stable specification order.
    pub fn write_document_style_policies(
        &mut self,
        policies: &crate::DocumentStylePolicies,
    ) -> io::Result<()> {
        if policies.update_styles_from_template {
            self.write_control_word("linkstyles", None)?;
        }
        if policies.lock_theme {
            self.write_control_word("stylelocktheme", None)?;
        }
        if policies.lock_quick_format_set {
            self.write_control_word("stylelockqfset", None)?;
        }
        if policies.use_normal_style_for_lists {
            self.write_control_word("usenormstyforlist", None)?;
        }
        Ok(())
    }

    /// Write passive legacy style restrictions in stable specification order.
    pub fn write_document_style_restrictions(
        &mut self,
        restrictions: &crate::DocumentStyleRestrictions,
    ) -> io::Result<()> {
        if restrictions.restrictions_present {
            self.write_control_word("stylelock", None)?;
        }
        if restrictions.enforced {
            self.write_control_word("stylelockenforced", None)?;
        }
        if restrictions.backward_compatibility {
            self.write_control_word("stylelockbackcomp", None)?;
        }
        if restrictions.allow_auto_format_override {
            self.write_control_word("autofmtoverride", None)?;
        }
        Ok(())
    }

    /// Write passive booklet-printing metadata in stable specification order.
    pub fn write_document_booklet_printing(
        &mut self,
        printing: &crate::DocumentBookletPrinting,
    ) -> io::Result<()> {
        if printing.book_fold {
            self.write_control_word("bookfold", None)?;
        }
        if printing.reverse_book_fold {
            self.write_control_word("bookfoldrev", None)?;
        }
        if let Some(value) = printing.sheets_per_booklet {
            if value > i32::MAX as u32 || value % 4 != 0 {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "booklet sheets must be an RTF signed nonnegative multiple of four",
                ));
            }
            self.write_control_word("bookfoldsheets", Some(value as i32))?;
        }
        Ok(())
    }

    /// Write passive privacy-removal requests in stable specification order.
    pub fn write_document_privacy_policies(
        &mut self,
        policies: &crate::DocumentPrivacyPolicies,
    ) -> io::Result<()> {
        if policies.remove_personal_information {
            self.write_control_word("rempersonalinfo", None)?;
        }
        if policies.remove_date_time_information {
            self.write_control_word("remdttm", None)?;
        }
        Ok(())
    }

    /// Write passive Word 2003 compatibility flags in specification order.
    pub fn write_document_word_2003_compatibility(
        &mut self,
        compatibility: &crate::DocumentWord2003Compatibility,
    ) -> io::Result<()> {
        if compatibility.preserve_autofit_table_width_around_shapes {
            self.write_control_word("noafcnsttbl", None)?;
        }
        if compatibility.use_hanging_indent_as_numbering_tab {
            self.write_control_word("noindnmbrts", None)?;
        }
        if compatibility.use_legacy_kinsoku_characters {
            self.write_control_word("felnbrelev", None)?;
        }
        if compatibility.use_legacy_floating_object_indentation {
            self.write_control_word("indrlsweleven", None)?;
        }
        if compatibility.allow_contextual_spacing_in_tables {
            self.write_control_word("nocxsptable", None)?;
        }
        if compatibility.ignore_cell_vertical_alignment_with_floating_objects {
            self.write_control_word("notcvasp", None)?;
        }
        if compatibility.ignore_text_box_vertical_alignment {
            self.write_control_word("notvatxbx", None)?;
        }
        if compatibility.split_page_break_paragraph {
            self.write_control_word("spltpgpar", None)?;
        }
        if compatibility.use_fixed_width_hangul {
            self.write_control_word("hwelev", None)?;
        }
        if compatibility.use_legacy_autofit_width_expansion {
            self.write_control_word("afelev", None)?;
        }
        if compatibility.use_cached_column_balancing {
            self.write_control_word("cachedcolbal", None)?;
        }
        if compatibility.underline_numbering_suffix {
            self.write_control_word("utinl", None)?;
        }
        if compatibility.do_not_split_rows_around_floating_tables {
            self.write_control_word("notbrkcnstfrctbl", None)?;
        }
        if compatibility.use_ansi_kerning_pairs {
            self.write_control_word("krnprsnet", None)?;
        }
        Ok(())
    }

    /// Write passive document compatibility policy in specification order.
    pub fn write_document_compatibility_policy(
        &mut self,
        policy: &crate::DocumentCompatibilityPolicy,
    ) -> io::Result<()> {
        if policy.reset_options_to_defaults {
            self.write_control_word("nocompatoptions", None)?;
        }
        if let Some(feature_throttle) = policy.feature_throttle {
            self.write_control_word("nofeaturethrottle", Some(feature_throttle.rtf_value()))?;
        }
        if policy.force_upgrade {
            self.write_control_word("forceupgrade", None)?;
        }
        Ok(())
    }

    /// Write passive Asian grid compatibility flags in specification order.
    pub fn write_document_asian_grid_compatibility(
        &mut self,
        compatibility: &crate::DocumentAsianGridCompatibility,
    ) -> io::Result<()> {
        if compatibility.apply_thai_line_breaking_rules {
            self.write_control_word("ApplyBrkRules", None)?;
        }
        if compatibility.snap_text_to_grid_inside_table {
            self.write_control_word("snaptogridincell", None)?;
        }
        if compatibility.allow_hanging_punctuation {
            self.write_control_word("wrppunct", None)?;
        }
        if compatibility.use_asian_line_breaking_rules {
            self.write_control_word("asianbrkrule", None)?;
        }
        if compatibility.compress_punctuation_at_line_start {
            self.write_control_word("toplinepunct", None)?;
        }
        Ok(())
    }

    /// Write passive legacy automatic-layout flags in specification order.
    pub fn write_document_legacy_layout_compatibility(
        &mut self,
        compatibility: &crate::DocumentLegacyLayoutCompatibility,
    ) -> io::Result<()> {
        if compatibility.do_not_use_word_97_shape_layout {
            self.write_control_word("splytwnine", None)?;
        }
        if compatibility.use_legacy_footnote_layout {
            self.write_control_word("ftnlytwnine", None)?;
        }
        if compatibility.use_html_paragraph_auto_spacing {
            self.write_control_word("htmautsp", None)?;
        }
        if compatibility.preserve_last_tab_alignment {
            self.write_control_word("useltbaln", None)?;
        }
        if compatibility.use_word_95_auto_spacing {
            self.write_control_word("oldas", None)?;
        }
        Ok(())
    }

    /// Write passive table-layout compatibility flags in specification order.
    pub fn write_document_table_layout_compatibility(
        &mut self,
        compatibility: &crate::DocumentTableLayoutCompatibility,
    ) -> io::Result<()> {
        if compatibility.combine_borders_like_word_5 {
            self.write_control_word("otblrul", None)?;
        }
        if compatibility.do_not_align_rows_independently {
            self.write_control_word("alntblind", None)?;
        }
        if compatibility.do_not_use_raw_table_width {
            self.write_control_word("lytcalctblwd", None)?;
        }
        if compatibility.keep_rows_together {
            self.write_control_word("lyttblrtgr", None)?;
        }
        if compatibility.do_not_adjust_line_height {
            self.write_control_word("nolnhtadjtbl", None)?;
        }
        if compatibility.do_not_break_wrapped_tables_across_pages {
            self.write_control_word("nobrkwrptbl", None)?;
        }
        if compatibility.prevent_autofit_growth_into_margins {
            self.write_control_word("nogrowautofit", None)?;
        }
        if compatibility.use_word_2003_table_style_rules {
            self.write_control_word("newtblstyruls", None)?;
        }
        Ok(())
    }

    /// Write passive East Asian compatibility flags in specification order.
    pub fn write_document_east_asian_compatibility(
        &mut self,
        compatibility: &crate::DocumentEastAsianCompatibility,
    ) -> io::Result<()> {
        if compatibility.do_not_balance_sbcs_dbcs {
            self.write_control_word("dntblnsbdb", None)?;
        }
        if compatibility.expand_spacing_at_shift_return {
            self.write_control_word("expshrtn", None)?;
        }
        if compatibility.do_not_add_space_for_underline {
            self.write_control_word("nospaceforul", None)?;
        }
        if compatibility.do_not_underline_trailing_spaces {
            self.write_control_word("noultrlspc", None)?;
        }
        if compatibility.do_not_translate_backslash_to_yen {
            self.write_control_word("noxlattoyen", None)?;
        }
        if compatibility.use_legacy_line_breaking_rules {
            self.write_control_word("lnbrkrule", None)?;
        }
        Ok(())
    }

    /// Write passive line-spacing compatibility flags in specification order.
    pub fn write_document_line_spacing_compatibility(
        &mut self,
        compatibility: &crate::DocumentLineSpacingCompatibility,
    ) -> io::Result<()> {
        if compatibility.suppress_extra_spacing_for_raised_lowered_text {
            self.write_control_word("noextrasprl", None)?;
        }
        if compatibility.suppress_extra_spacing_at_top_of_page {
            self.write_control_word("sprstsp", None)?;
        }
        if compatibility.suppress_space_before_after_hard_break {
            self.write_control_word("sprsspbf", None)?;
        }
        if compatibility.suppress_wordperfect_extra_line_spacing {
            self.write_control_word("sprslnsp", None)?;
        }
        if compatibility.suppress_extra_spacing_at_bottom_of_page {
            self.write_control_word("sprsbsp", None)?;
        }
        Ok(())
    }

    /// Write passive file and template flags in stable specification order.
    pub fn write_document_file_settings(
        &mut self,
        settings: &crate::DocumentFileSettings,
    ) -> io::Result<()> {
        if settings.automatic_backup {
            self.write_control_word("makebackup", None)?;
        }
        if settings.default_save_format_rtf {
            self.write_control_word("defformat", None)?;
        }
        if settings.template_or_stationery {
            self.write_control_word("doctemp", None)?;
        }
        Ok(())
    }

    /// Write passive view metadata in stable `viewkind`, `viewscale`, `viewzk` order.
    pub fn write_document_view(&mut self, view: &crate::DocumentView) -> io::Result<()> {
        view.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(kind) = view.kind {
            self.write_control_word("viewkind", Some(kind.rtf_value()))?;
        }
        if let Some(scale) = view.scale_percent {
            self.write_control_word("viewscale", Some(i32::from(scale)))?;
        }
        if let Some(kind) = view.zoom_kind {
            self.write_control_word("viewzk", Some(kind.rtf_value()))?;
        }
        if let Some(value) = view.background_shapes {
            self.write_control_word("viewbksp", Some(i32::from(value)))?;
        }
        if view.hide_page_boundaries {
            self.write_control_word("viewnobound", None)?;
        }
        Ok(())
    }

    /// Write passive review-display flags in stable specification order.
    pub fn write_review_display(
        &mut self,
        display: &crate::DocumentReviewDisplay,
    ) -> io::Result<()> {
        if display.hide_markup {
            self.write_control_word("donotshowmarkup", None)?;
        }
        if display.hide_comments {
            self.write_control_word("donotshowcomments", None)?;
        }
        if display.hide_insertions_and_deletions {
            self.write_control_word("donotshowinsdel", None)?;
        }
        Ok(())
    }

    /// Write an inert document-window caption as the canonical starred destination.
    pub fn write_window_caption(
        &mut self,
        caption: Option<&crate::DocumentWindowCaption<'_>>,
    ) -> io::Result<()> {
        let Some(caption) = caption else {
            return Ok(());
        };
        caption
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\windowcaption ")?;
        self.write_destination_text(caption.text.as_ref())?;
        self.write_str("}")
    }

    /// Write the inert custom kinsoku character sets and their language.
    pub fn write_kinsoku(&mut self, kinsoku: &crate::DocumentKinsoku<'_>) -> io::Result<()> {
        kinsoku
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(following) = &kinsoku.following {
            self.write_str("{\\*\\fchars ")?;
            self.write_destination_text(following.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(leading) = &kinsoku.leading {
            self.write_str("{\\*\\lchars ")?;
            self.write_destination_text(leading.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(language) = kinsoku.language {
            self.write_control_word("ksulang", Some(language as i32))?;
        }
        Ok(())
    }

    /// Write an inert custom XSL transform location as its required starred destination.
    pub fn write_xsl_transform(
        &mut self,
        transform: Option<&crate::DocumentXslTransform<'_>>,
    ) -> io::Result<()> {
        let Some(transform) = transform else {
            return Ok(());
        };
        transform
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\xform ")?;
        self.write_destination_text(transform.location.as_ref())?;
        self.write_str("}")
    }

    /// Write passive requested transform usage without applying the transform.
    pub fn write_xsl_transform_usage(
        &mut self,
        usage: crate::DocumentXslTransformUsage,
    ) -> io::Result<()> {
        if usage.is_requested() {
            self.write_control_word("usexform", None)?;
        }
        Ok(())
    }

    /// Write passive style-list filter suggestions as exactly four hexadecimal digits.
    pub fn write_style_list_filter(
        &mut self,
        filter: Option<crate::DocumentStyleListFilter>,
    ) -> io::Result<()> {
        let Some(filter) = filter else {
            return Ok(());
        };
        filter
            .validate_for_write()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\wgrffmtfilter ")?;
        self.write_str(&format!("{:04X}", filter.bits()))?;
        self.write_str("}")
    }

    /// Write a passive style-list sorting suggestion with an explicit value.
    pub fn write_style_sort_method(
        &mut self,
        method: Option<crate::DocumentStyleSortMethod>,
    ) -> io::Result<()> {
        if let Some(method) = method {
            self.write_control_word("stylesortmethod", Some(method.rtf_value()))?;
        }
        Ok(())
    }

    /// Write opaque reservation metadata without authenticating or decrypting it.
    pub fn write_document_write_reservations(
        &mut self,
        reservations: &crate::DocumentWriteReservations<'_>,
    ) -> io::Result<()> {
        reservations
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(hash) = &reservations.hash {
            self.write_str("{\\*\\writereservhash ")?;
            for byte in hash.data.iter() {
                self.write_str(&format!("{byte:02X}"))?;
            }
            self.write_str("}")?;
        }
        if let Some(legacy) = &reservations.legacy {
            self.write_str("{\\*\\writereservation ")?;
            self.write_destination_text(legacy.data.as_ref())?;
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write passive save preferences in stable specification order.
    pub fn write_document_save_preferences(
        &mut self,
        preferences: &crate::DocumentSavePreferences,
    ) -> io::Result<()> {
        if preferences.read_only == crate::DocumentReadOnlyRecommendation::Recommended {
            self.write_control_word("readonlyrecommended", None)?;
        }
        if preferences.thumbnail == crate::DocumentThumbnailPreference::RequiredIfSupported {
            self.write_control_word("saveprevpict", None)?;
        }
        Ok(())
    }

    /// Write font table
    fn write_font_table(&mut self) -> io::Result<()> {
        self.font_table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if self.font_table.fonts().is_empty() {
            return Ok(());
        }

        self.write_str("{")?;
        self.write_control_word("fonttbl", None)?;

        // Clone fonts to avoid borrowing issues
        let fonts: Vec<_> = self.font_table.fonts().to_vec();
        for (idx, font) in fonts.iter().enumerate() {
            if !self.font_table.is_defined(idx as FontRef) {
                continue;
            }
            font.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            self.write_str("{")?;
            self.write_control_word("f", Some(idx as i32))?;

            // Write font family
            match font.family {
                FontFamily::Roman => self.write_control_word("froman", None)?,
                FontFamily::Swiss => self.write_control_word("fswiss", None)?,
                FontFamily::Modern => self.write_control_word("fmodern", None)?,
                FontFamily::Script => self.write_control_word("fscript", None)?,
                FontFamily::Decor => self.write_control_word("fdecor", None)?,
                FontFamily::Tech => self.write_control_word("ftech", None)?,
                FontFamily::Nil => self.write_control_word("fnil", None)?,
            }

            // Write charset
            if font.charset != 0 {
                self.write_control_word("fcharset", Some(font.charset as i32))?;
            }
            if let Some(theme) = font.theme {
                self.write_control_word(theme.control_word(), None)?;
            }
            self.write_control_word(
                "fprq",
                Some(match font.pitch {
                    crate::FontPitch::Default => 0,
                    crate::FontPitch::Fixed => 1,
                    crate::FontPitch::Variable => 2,
                }),
            )?;
            if let Some(code_page) = font.code_page {
                self.write_control_word("cpg", Some(i32::from(code_page)))?;
            }
            if let Some(panose) = font.panose {
                self.write_str("{\\*")?;
                self.write_control_word("panose", None)?;
                self.write_str(" ")?;
                for byte in panose {
                    write!(self.writer, "{byte:02x}")?;
                }
                self.write_str("}")?;
            }
            if let Some(name) = font.non_tagged_name.as_deref() {
                self.write_str("{\\*")?;
                self.write_control_word("fname", None)?;
                self.write_str(" ")?;
                self.write_text(name)?;
                self.write_str("}")?;
            }
            if let Some(embedded) = &font.embedded {
                self.write_str("{\\*")?;
                self.write_control_word("fontemb", None)?;
                self.write_control_word(
                    match embedded.format {
                        crate::EmbeddedFontFormat::Nil => "ftnil",
                        crate::EmbeddedFontFormat::TrueType => "fttruetype",
                    },
                    None,
                )?;
                if let Some(name) = embedded.file_name.as_deref() {
                    self.write_str("{\\*")?;
                    self.write_control_word("fontfile", None)?;
                    if let Some(code_page) = embedded.file_code_page {
                        self.write_control_word("cpg", Some(i32::from(code_page)))?;
                    }
                    self.write_str(" ")?;
                    self.write_text(name)?;
                    self.write_str("}")?;
                }
                if let Some(data) = &embedded.data {
                    self.write_str(" ")?;
                    for byte in data {
                        write!(self.writer, "{byte:02x}")?;
                    }
                }
                self.write_str("}")?;
            }

            // Write font name
            self.write_str(" ")?;
            self.write_text(font.name.as_ref())?;
            if let Some(name) = font.alternate_name.as_deref() {
                self.write_str("{\\*")?;
                self.write_control_word("falt", None)?;
                self.write_str(" ")?;
                self.write_text(name)?;
                self.write_str("}")?;
            }
            self.write_str(";")?;
            self.write_str("}")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write the optional external-file metadata table.
    pub fn write_file_table(&mut self, table: Option<&crate::FileTable<'_>>) -> io::Result<()> {
        let Some(table) = table else {
            return Ok(());
        };
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        self.write_str("{\\*")?;
        self.write_control_word("filetbl", None)?;
        for entry in table.entries() {
            self.write_str("{")?;
            self.write_control_word("file", None)?;
            self.write_control_word("fid", Some(entry.id as i32))?;
            if let Some(level) = entry.relative_path_level {
                self.write_control_word("frelative", Some(i32::from(level)))?;
            }
            if let Some(os) = entry.operating_system {
                self.write_control_word("fosnum", Some(i32::from(os)))?;
            }
            if entry.valid_on.mac {
                self.write_control_word("fvalidmac", None)?;
            }
            if entry.valid_on.dos {
                self.write_control_word("fvaliddos", None)?;
            }
            if entry.valid_on.ntfs {
                self.write_control_word("fvalidntfs", None)?;
            }
            if entry.valid_on.hpfs {
                self.write_control_word("fvalidhpfs", None)?;
            }
            match entry.location {
                crate::FileLocation::Local => {},
                crate::FileLocation::Network => self.write_control_word("fnetwork", None)?,
                crate::FileLocation::NonFileSystem => {
                    self.write_control_word("fnonfilesys", None)?;
                },
            }
            self.write_str(" ")?;
            self.write_text(entry.name.as_ref())?;
            self.write_str(";}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    /// Write color table
    fn write_color_table(&mut self) -> io::Result<()> {
        if self.color_table.colors().is_empty() {
            return Ok(());
        }

        self.write_str("{")?;
        self.write_control_word("colortbl", None)?;

        // Clone colors to avoid borrowing issues
        let colors: Vec<_> = self.color_table.colors().to_vec();
        for color in &colors {
            self.write_control_word("red", Some(color.red as i32))?;
            self.write_control_word("green", Some(color.green as i32))?;
            self.write_control_word("blue", Some(color.blue as i32))?;
            self.write_str(";")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write the list-definition table.
    pub fn write_picture(&mut self, picture: &Picture<'_>) -> io::Result<()> {
        picture
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\pict")?;
        match picture.image_type {
            ImageType::Emf => self.write_control_word("emfblip", None)?,
            ImageType::Wmf => self.write_control_word("wmetafile", Some(8))?,
            ImageType::Png => self.write_control_word("pngblip", None)?,
            ImageType::Jpeg => self.write_control_word("jpegblip", None)?,
            ImageType::Dib if picture.bitmap.windows_bitmap => {
                self.write_control_word("wbitmap", Some(0))?
            },
            ImageType::Dib => self.write_control_word("dibitmap", Some(0))?,
            ImageType::Pict => self.write_control_word("macpict", None)?,
            ImageType::Unknown => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "cannot write a picture with unknown image type",
                ));
            },
        }
        for (control, value) in [
            ("picw", picture.width),
            ("pich", picture.height),
            ("picwgoal", picture.goal_width),
            ("pichgoal", picture.goal_height),
            ("picscalex", picture.scale_x),
            ("picscaley", picture.scale_y),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, Some(value))?;
            }
        }
        if picture.scaled {
            self.write_control_word("picscaled", None)?;
        }
        for (control, value) in [
            ("piccropl", picture.crop.left),
            ("piccropr", picture.crop.right),
            ("piccropt", picture.crop.top),
            ("piccropb", picture.crop.bottom),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, Some(value))?;
            }
        }
        if picture.bitmap.bitmap_source {
            self.write_control_word("picbmp", None)?;
        }
        for (control, value) in [
            ("picbpp", picture.bitmap.bits_per_pixel.map(i32::from)),
            (
                "wbmbitspixel",
                picture.bitmap.windows_bits_per_pixel.map(i32::from),
            ),
            ("wbmplanes", picture.bitmap.planes.map(i32::from)),
            (
                "wbmwidthbytes",
                picture
                    .bitmap
                    .width_bytes
                    .and_then(|value| i32::try_from(value).ok()),
            ),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, Some(value))?;
            }
        }
        if let Some(identity) = &picture.identity {
            if let Some(tag) = identity.tag {
                self.write_control_word("bliptag", Some(tag))?;
            }
            if let Some(upi) = identity.units_per_inch {
                self.write_control_word("blipupi", Some(i32::from(upi)))?;
            }
            if let Some(uid) = &identity.uid {
                self.write_str("{\\*\\blipuid ")?;
                for byte in uid.iter() {
                    write!(self.writer, "{byte:02x}")?;
                }
                self.write_str("}")?;
            }
        }
        if let Some(properties) = &picture.shape_properties {
            self.write_picture_shape_properties(properties)?;
        }
        self.write_str(" ")?;
        for byte in picture.data.iter() {
            write!(self.writer, "{byte:02x}")?;
        }
        self.write_str("}")
    }

    /// Write the list-definition table.
    pub fn write_list_table(&mut self, table: &ListTable<'_>) -> io::Result<()> {
        self.write_list_table_with_pictures(table, &[])
    }

    fn write_list_table_with_pictures(
        &mut self,
        table: &ListTable<'_>,
        pictures: &[Picture<'_>],
    ) -> io::Result<()> {
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if table.lists().is_empty() && table.picture_bullet_count == 0 {
            return Ok(());
        }
        if table.lists().len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF list table exceeds the supported list count",
            ));
        }
        self.write_str("{\\*\\listtable")?;
        if table.picture_bullet_count != 0 {
            self.write_str("{\\*\\listpicture")?;
            for slot in 0..table.picture_bullet_count as usize {
                let Some(index) = table
                    .picture_bullet_picture_indices()
                    .get(slot)
                    .copied()
                    .flatten()
                else {
                    continue;
                };
                let picture = pictures.get(index).ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF list-picture index is outside the document picture store",
                    )
                })?;
                self.write_str("{\\*\\shppict")?;
                self.write_picture(picture)?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        for list in table.lists() {
            self.write_list_definition(list)?;
        }
        self.write_str("}")?;
        Ok(())
    }

    fn write_list_definition(&mut self, list: &List<'_>) -> io::Result<()> {
        if list.levels.len() > 9 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF lists cannot contain more than nine levels",
            ));
        }
        if list.simple && list.levels.len() > 1 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "a simple RTF list cannot contain more than one level",
            ));
        }
        if list.simple && list.hybrid {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "an RTF list cannot be both simple and hybrid",
            ));
        }
        self.write_str("{")?;
        self.write_control_word("list", None)?;
        self.write_control_word("listtemplateid", Some(list.template_id))?;
        if list.simple {
            self.write_control_word("listsimple", None)?;
        }
        if list.hybrid {
            self.write_control_word("listhybrid", None)?;
        }
        for level in &list.levels {
            self.write_list_level(level)?;
        }
        if !list.name.is_empty() {
            self.write_str("{")?;
            self.write_control_word("listname", None)?;
            self.write_str(" ")?;
            self.write_text(list.name.as_ref())?;
            self.write_str(";}")?;
        }
        if !list.style_name.is_empty() {
            self.write_str("{\\*")?;
            self.write_control_word("liststylename", None)?;
            self.write_str(" ")?;
            self.write_text(list.style_name.as_ref())?;
            self.write_str(";}")?;
        }
        if let Some(priority) = list.style_priority {
            self.write_control_word("spriority", Some(priority))?;
        }
        self.write_control_word("listid", Some(list.id))?;
        self.write_str("}")?;
        Ok(())
    }

    fn write_list_level(&mut self, level: &ListLevel<'_>) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("listlevel", None)?;
        self.write_control_word(
            "levelnfc",
            Some(Self::list_level_type_value(level.level_type)),
        )?;
        self.write_control_word(
            "leveljc",
            Some(match level.justification {
                ListJustification::Left => 0,
                ListJustification::Center => 1,
                ListJustification::Right => 2,
            }),
        )?;
        self.write_control_word(
            "levelfollow",
            Some(match level.follow {
                ListFollow::Tab => 0,
                ListFollow::Space => 1,
                ListFollow::Nothing => 2,
            }),
        )?;
        self.write_control_word("levelstartat", Some(level.start_at))?;
        self.write_control_word("levelspace", Some(level.space))?;
        self.write_control_word("levelindent", Some(level.indent))?;
        if let Some(left_indent) = level.left_indent {
            self.write_control_word("li", Some(left_indent))?;
        }
        if let Some(first_line_indent) = level.first_line_indent {
            self.write_control_word("fi", Some(first_line_indent))?;
        }
        for tab in &level.tabs {
            self.write_control_word("tx", Some(*tab))?;
        }
        if level.tentative {
            self.write_control_word("lvltentative", None)?;
        }
        if level.legal_format {
            self.write_control_word("levellegal", None)?;
        }
        if level.no_restart {
            self.write_control_word("levelnorestart", None)?;
        }
        if level.legacy {
            self.write_control_word("levelold", None)?;
        }
        if level.include_previous {
            self.write_control_word("levelprev", None)?;
        }
        if level.include_previous_space {
            self.write_control_word("levelprevspace", None)?;
        }
        if let Some(template_id) = level.template_id {
            self.write_control_word("leveltemplateid", Some(template_id))?;
        }
        if let Some(picture_index) = level.picture_index {
            self.write_control_word(
                "levelpicture",
                Some(i32::try_from(picture_index).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "invalid list picture index")
                })?),
            )?;
        }
        self.write_list_level_text(level.number_text.as_ref(), level.number_positions.as_ref())?;
        if level.font_ref != 0 {
            self.write_control_word("f", Some(i32::from(level.font_ref)))?;
        }
        self.write_str("}")?;
        Ok(())
    }

    fn list_level_type_value(level_type: ListLevelType) -> i32 {
        match level_type {
            ListLevelType::Decimal => 0,
            ListLevelType::UpperRoman => 1,
            ListLevelType::LowerRoman => 2,
            ListLevelType::UpperLetter => 3,
            ListLevelType::LowerLetter => 4,
            ListLevelType::Ordinal => 5,
            ListLevelType::CardinalText => 6,
            ListLevelType::OrdinalText => 7,
            ListLevelType::Bullet => 23,
            ListLevelType::None => 255,
            ListLevelType::Other(value) => value,
        }
    }

    fn write_list_level_text(&mut self, text: &str, positions: &str) -> io::Result<()> {
        let count = u8::try_from(text.chars().count()).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF list level text cannot exceed 255 characters",
            )
        })?;
        self.write_str("{")?;
        self.write_control_word("leveltext", None)?;
        self.write_hex_byte(count)?;
        for ch in text.chars() {
            if u32::from(ch) <= u8::MAX.into() && (ch.is_control() || !ch.is_ascii()) {
                self.write_hex_byte(ch as u8)?;
            } else {
                let mut buffer = [0; 4];
                self.write_text(ch.encode_utf8(&mut buffer))?;
            }
        }
        self.write_str(";}")?;

        self.write_str("{")?;
        self.write_control_word("levelnumbers", None)?;
        if positions.is_empty() {
            for (index, ch) in text.chars().enumerate() {
                if u32::from(ch) <= 8 {
                    let position = u8::try_from(index + 1).map_err(|_| {
                        io::Error::new(io::ErrorKind::InvalidInput, "invalid RTF list placeholder")
                    })?;
                    self.write_hex_byte(position)?;
                }
            }
        } else {
            for byte in positions.bytes() {
                self.write_hex_byte(byte)?;
            }
        }
        self.write_str(";}")?;
        Ok(())
    }

    fn write_hex_byte(&mut self, value: u8) -> io::Result<()> {
        write!(self.writer, "\\'{value:02x}")
    }

    /// Write the list-override table.
    pub fn write_list_override_table(&mut self, table: &ListOverrideTable) -> io::Result<()> {
        if table.overrides().is_empty() {
            return Ok(());
        }
        if table.overrides().len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF list override table exceeds the supported entry count",
            ));
        }
        self.write_str("{\\*\\listoverridetable")?;
        for entry in table.overrides() {
            self.write_str("{")?;
            self.write_control_word("listoverride", None)?;
            self.write_control_word("listid", Some(entry.list_id))?;
            self.write_control_word(
                "listoverridecount",
                Some(i32::from(entry.level_count_override.unwrap_or_else(|| {
                    u8::try_from(entry.levels.len())
                        .unwrap_or_else(|_| u8::from(entry.start_at_override.is_some()))
                }))),
            )?;
            if entry.levels.is_empty() {
                if let Some(start_at) = entry.start_at_override {
                    self.write_str("{")?;
                    self.write_control_word("lfolevel", None)?;
                    self.write_control_word("listoverridestartat", None)?;
                    self.write_control_word("levelstartat", Some(start_at))?;
                    self.write_str("}")?;
                }
            }
            for level in &entry.levels {
                self.write_str("{")?;
                self.write_control_word("lfolevel", None)?;
                if level.format_override {
                    self.write_control_word("listoverrideformat", None)?;
                }
                if let Some(start_at) = level.start_at {
                    self.write_control_word("listoverridestartat", None)?;
                    self.write_control_word("levelstartat", Some(start_at))?;
                }
                self.write_str("}")?;
            }
            self.write_control_word("ls", Some(entry.index))?;
            self.write_str("}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    /// Write ordered inert legacy section-numbering defaults.
    pub fn write_legacy_section_numbering(
        &mut self,
        numbering: &crate::LegacySectionNumbering<'_>,
    ) -> io::Result<()> {
        numbering
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for level in numbering.levels() {
            self.write_str("{\\*")?;
            self.write_control_word("pnseclvl", Some(i32::from(level.level)))?;
            self.write_control_word(
                match level.format {
                    crate::LegacyNumberingFormat::Decimal => "pndec",
                    crate::LegacyNumberingFormat::UpperRoman => "pnucrm",
                    crate::LegacyNumberingFormat::LowerRoman => "pnlcrm",
                    crate::LegacyNumberingFormat::UpperLetter => "pnucltr",
                    crate::LegacyNumberingFormat::LowerLetter => "pnlcltr",
                },
                None,
            )?;
            if let Some(alignment) = level.alignment {
                self.write_control_word(
                    match alignment {
                        crate::LegacyNumberingAlignment::Left => "pnql",
                        crate::LegacyNumberingAlignment::Center => "pnqc",
                        crate::LegacyNumberingAlignment::Right => "pnqr",
                    },
                    None,
                )?;
            }
            if let Some(start_at) = level.start_at {
                self.write_control_word("pnstart", Some(start_at))?;
            }
            if let Some(indent) = level.indent {
                self.write_control_word("pnindent", Some(indent))?;
            }
            if let Some(space) = level.space {
                self.write_control_word("pnsp", Some(space))?;
            }
            if level.hanging {
                self.write_control_word("pnhang", None)?;
            }
            if level.previous {
                self.write_control_word("pnprev", None)?;
            }
            if let Some(font_ref) = level.font_ref {
                self.write_control_word("pnf", Some(i32::from(font_ref)))?;
            }
            if !level.text_before.is_empty() {
                self.write_str("{")?;
                self.write_control_word("pntxtb", None)?;
                self.write_str(" ")?;
                self.write_text(level.text_before.as_ref())?;
                self.write_str("}")?;
            }
            if !level.text_after.is_empty() {
                self.write_str("{")?;
                self.write_control_word("pntxta", None)?;
                self.write_str(" ")?;
                self.write_text(level.text_after.as_ref())?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write the inert paragraph-group property table.
    pub fn write_paragraph_group_table(
        &mut self,
        table: Option<&crate::ParagraphGroupPropertyTable>,
    ) -> io::Result<()> {
        let Some(table) = table else {
            return Ok(());
        };
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\pgptbl")?;
        for entry in table.entries() {
            self.write_str("{")?;
            self.write_control_word("pgp", None)?;
            self.write_control_word(
                "ipgp",
                Some(i32::try_from(entry.parent_id).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "invalid pgp parent ID")
                })?),
            )?;
            self.write_control_word("itap", Some(i32::from(entry.table_nesting_level)))?;
            self.write_control_word("li", Some(entry.left_indent))?;
            self.write_control_word("ri", Some(entry.right_indent))?;
            self.write_control_word("sb", Some(entry.space_before))?;
            self.write_control_word("sa", Some(entry.space_after))?;
            self.write_borders(&entry.borders)?;
            self.write_str("}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    /// Write explicit document-level footnote and endnote configuration.
    pub fn write_note_options(&mut self, options: &crate::NoteOptions) -> io::Result<()> {
        options
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        if let Some(value) = options.present_kinds {
            self.write_control_word(
                "fet",
                Some(match value {
                    crate::PresentNoteKinds::FootnotesOnly => 0,
                    crate::PresentNoteKinds::EndnotesOnly => 1,
                    crate::PresentNoteKinds::FootnotesAndEndnotes => 2,
                }),
            )?;
        }
        if let Some(value) = options.footnote_placement {
            self.write_control_word(
                match value {
                    crate::NotePlacement::EndOfSection => "endnotes",
                    crate::NotePlacement::EndOfDocument => "enddoc",
                    crate::NotePlacement::BeneathText => "ftntj",
                    crate::NotePlacement::BottomOfPage => "ftnbj",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_start {
            self.write_control_word("ftnstart", Some(value))?;
        }
        if let Some(value) = options.footnote_restart {
            self.write_control_word(
                match value {
                    crate::FootnoteRestart::Continuous => "ftnrstcont",
                    crate::FootnoteRestart::EachSection => "ftnrestart",
                    crate::FootnoteRestart::EachPage => "ftnrstpg",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_numbering {
            self.write_control_word(Self::note_numbering_control(value, false), None)?;
        }
        if let Some(value) = options.endnote_placement {
            self.write_control_word(
                match value {
                    crate::NotePlacement::EndOfSection => "aendnotes",
                    crate::NotePlacement::EndOfDocument => "aenddoc",
                    crate::NotePlacement::BeneathText => "aftntj",
                    crate::NotePlacement::BottomOfPage => "aftnbj",
                },
                None,
            )?;
        }
        if let Some(value) = options.endnote_start {
            self.write_control_word("aftnstart", Some(value))?;
        }
        if let Some(value) = options.endnote_restart {
            self.write_control_word(
                match value {
                    crate::EndnoteRestart::Continuous => "aftnrstcont",
                    crate::EndnoteRestart::EachSection => "aftnrestart",
                },
                None,
            )?;
        }
        if let Some(value) = options.endnote_numbering {
            self.write_control_word(Self::note_numbering_control(value, true), None)?;
        }
        Ok(())
    }

    fn note_numbering_control(style: crate::NoteNumberingStyle, endnote: bool) -> &'static str {
        match (endnote, style) {
            (false, crate::NoteNumberingStyle::Arabic) => "ftnnar",
            (false, crate::NoteNumberingStyle::LowercaseLetter) => "ftnnalc",
            (false, crate::NoteNumberingStyle::UppercaseLetter) => "ftnnauc",
            (false, crate::NoteNumberingStyle::LowercaseRoman) => "ftnnrlc",
            (false, crate::NoteNumberingStyle::UppercaseRoman) => "ftnnruc",
            (false, crate::NoteNumberingStyle::Chicago) => "ftnnchi",
            (false, crate::NoteNumberingStyle::KoreanChosung) => "ftnnchosung",
            (false, crate::NoteNumberingStyle::Circle) => "ftnncnum",
            (false, crate::NoteNumberingStyle::KanjiDigitless) => "ftnndbnum",
            (false, crate::NoteNumberingStyle::KanjiWithDigit) => "ftnndbnumd",
            (false, crate::NoteNumberingStyle::KanjiThree) => "ftnndbnumt",
            (false, crate::NoteNumberingStyle::KanjiFour) => "ftnndbnumk",
            (false, crate::NoteNumberingStyle::DoubleByte) => "ftnndbar",
            (false, crate::NoteNumberingStyle::KoreanGanada) => "ftnnganada",
            (false, crate::NoteNumberingStyle::ChineseOne) => "ftnngbnum",
            (false, crate::NoteNumberingStyle::ChineseTwo) => "ftnngbnumd",
            (false, crate::NoteNumberingStyle::ChineseThree) => "ftnngbnuml",
            (false, crate::NoteNumberingStyle::ChineseFour) => "ftnngbnumk",
            (false, crate::NoteNumberingStyle::ZodiacOne) => "ftnnzodiac",
            (false, crate::NoteNumberingStyle::ZodiacTwo) => "ftnnzodiacd",
            (false, crate::NoteNumberingStyle::ZodiacThree) => "ftnnzodiacl",
            (true, crate::NoteNumberingStyle::Arabic) => "aftnnar",
            (true, crate::NoteNumberingStyle::LowercaseLetter) => "aftnnalc",
            (true, crate::NoteNumberingStyle::UppercaseLetter) => "aftnnauc",
            (true, crate::NoteNumberingStyle::LowercaseRoman) => "aftnnrlc",
            (true, crate::NoteNumberingStyle::UppercaseRoman) => "aftnnruc",
            (true, crate::NoteNumberingStyle::Chicago) => "aftnnchi",
            (true, crate::NoteNumberingStyle::KoreanChosung) => "aftnnchosung",
            (true, crate::NoteNumberingStyle::Circle) => "aftnncnum",
            (true, crate::NoteNumberingStyle::KanjiDigitless) => "aftnndbnum",
            (true, crate::NoteNumberingStyle::KanjiWithDigit) => "aftnndbnumd",
            (true, crate::NoteNumberingStyle::KanjiThree) => "aftnndbnumt",
            (true, crate::NoteNumberingStyle::KanjiFour) => "aftnndbnumk",
            (true, crate::NoteNumberingStyle::DoubleByte) => "aftnndbar",
            (true, crate::NoteNumberingStyle::KoreanGanada) => "aftnnganada",
            (true, crate::NoteNumberingStyle::ChineseOne) => "aftnngbnum",
            (true, crate::NoteNumberingStyle::ChineseTwo) => "aftnngbnumd",
            (true, crate::NoteNumberingStyle::ChineseThree) => "aftnngbnuml",
            (true, crate::NoteNumberingStyle::ChineseFour) => "aftnngbnumk",
            (true, crate::NoteNumberingStyle::ZodiacOne) => "aftnnzodiac",
            (true, crate::NoteNumberingStyle::ZodiacTwo) => "aftnnzodiacd",
            (true, crate::NoteNumberingStyle::ZodiacThree) => "aftnnzodiacl",
        }
    }

    /// Write ordered semantic note-separator destinations.
    pub fn write_note_separators(
        &mut self,
        table: &crate::NoteSeparatorTable<'_>,
    ) -> io::Result<()> {
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for separator in table.entries() {
            self.write_str("{\\*")?;
            self.write_control_word(
                match separator.kind {
                    crate::NoteSeparatorKind::FootnoteSeparator => "ftnsep",
                    crate::NoteSeparatorKind::FootnoteContinuationSeparator => "ftnsepc",
                    crate::NoteSeparatorKind::FootnoteContinuationNotice => "ftncn",
                    crate::NoteSeparatorKind::EndnoteSeparator => "aftnsep",
                    crate::NoteSeparatorKind::EndnoteContinuationSeparator => "aftnsepc",
                    crate::NoteSeparatorKind::EndnoteContinuationNotice => "aftncn",
                },
                None,
            )?;
            self.write_str(" ")?;
            for element in &separator.elements {
                match element {
                    crate::NoteSeparatorElement::Text(text) => self.write_text(text.as_ref())?,
                    crate::NoteSeparatorElement::SeparatorMark => {
                        self.write_control_word("chftnsep", None)?
                    },
                    crate::NoteSeparatorElement::ContinuationSeparatorMark => {
                        self.write_control_word("chftnsepc", None)?
                    },
                    crate::NoteSeparatorElement::ParagraphBreak => {
                        self.write_control_word("par", None)?
                    },
                    crate::NoteSeparatorElement::LineBreak => {
                        self.write_control_word("line", None)?
                    },
                    crate::NoteSeparatorElement::Drawing(crate::StoryDrawing::Shape(index)) => {
                        self.write_root_shape(&separator.shapes[*index])?
                    },
                    crate::NoteSeparatorElement::Drawing(crate::StoryDrawing::ShapeGroup(
                        index,
                    )) => self.write_shape_group(&separator.shape_groups[*index], true)?,
                }
            }
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write the revision-author table referenced by tracked-change runs.
    pub fn write_revision_table(
        &mut self,
        authors: &[crate::RevisionAuthor<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        if authors.is_empty() && revisions.is_empty() {
            return Ok(());
        }
        if authors.len() > crate::annotation::MAX_REVISION_AUTHORS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF revision-author table exceeds the safety limit",
            ));
        }
        let author_bytes = authors.iter().try_fold(0usize, |total, author| {
            author
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            total.checked_add(author.name.len()).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF aggregate revision-author size overflow",
                )
            })
        })?;
        if author_bytes > crate::annotation::MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF aggregate revision-author text exceeds the safety limit",
            ));
        }
        for revision in revisions {
            revision
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let index = usize::try_from(revision.id).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision author indices cannot be negative",
                )
            })?;
            let author = authors.get(index).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision author index is outside revtbl",
                )
            })?;
            if author.name != revision.author {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision author does not match its revtbl entry",
                ));
            }
        }

        self.write_str("{\\*\\revtbl")?;
        for author in authors {
            self.write_str("{")?;
            self.write_text(author.name.as_ref())?;
            self.write_str(";}")?;
        }
        self.write_str("}")?;
        Ok(())
    }

    pub fn write_generator(
        &mut self,
        generator: Option<&crate::DocumentGenerator<'_>>,
    ) -> io::Result<()> {
        let Some(generator) = generator else {
            return Ok(());
        };
        generator
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\generator ")?;
        self.write_destination_text(generator.value.as_ref())?;
        self.write_str(";}")
    }

    pub fn write_revision_save_metadata(
        &mut self,
        metadata: Option<&crate::RevisionSaveMetadata>,
    ) -> io::Result<()> {
        let Some(metadata) = metadata else {
            return Ok(());
        };
        metadata
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\rsidtbl ")?;
        for id in metadata.ids() {
            self.write_control_word("rsid", Some(*id as i32))?;
        }
        self.write_str("}")?;
        if let Some(root) = metadata.root() {
            self.write_control_word("rsidroot", Some(root as i32))?;
        }
        Ok(())
    }

    pub fn write_xml_namespace_table(
        &mut self,
        namespaces: Option<&[crate::XmlNamespace<'_>]>,
    ) -> io::Result<()> {
        let Some(namespaces) = namespaces else {
            return Ok(());
        };
        if namespaces.len() > crate::xml_namespace::MAX_XML_NAMESPACES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF XML namespace count exceeds the safety limit",
            ));
        }
        let mut total = 0usize;
        self.write_str("{\\*\\xmlnstbl ")?;
        for (index, namespace) in namespaces.iter().enumerate() {
            namespace
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if namespaces[..index]
                .iter()
                .any(|existing| existing.id == namespace.id)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF XML namespace IDs must be unique",
                ));
            }
            total = total
                .checked_add(namespace.namespace.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF XML namespace aggregate size overflow",
                    )
                })?;
            if total > crate::xml_namespace::MAX_XML_NAMESPACE_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF XML namespace aggregate text exceeds the safety limit",
                ));
            }
            self.write_str("{")?;
            self.write_control_word("xmlns", Some(namespace.id as i32))?;
            self.write_str(" ")?;
            self.write_destination_text(namespace.namespace.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    /// Write inert range-protection usernames without resolving any identity.
    pub fn write_protection_user_table(
        &mut self,
        table: Option<&crate::ProtectionUserTable<'_>>,
    ) -> io::Result<()> {
        let Some(table) = table else {
            return Ok(());
        };
        table
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\protusertbl")?;
        for user in table.users() {
            self.write_str("{")?;
            self.write_destination_text(user.name.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    pub fn write_theme(&mut self, theme: Option<&crate::DocumentTheme<'_>>) -> io::Result<()> {
        let Some(theme) = theme else {
            return Ok(());
        };
        theme
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_hex_destination("themedata", theme.data.as_ref())?;
        if let Some(mapping) = theme.color_scheme_mapping.as_deref() {
            self.write_hex_destination("colorschememapping", mapping)?;
        }
        Ok(())
    }

    fn write_hex_destination(&mut self, control: &str, data: &[u8]) -> io::Result<()> {
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        for byte in data {
            write!(self.writer, "{byte:02x}")?;
        }
        self.write_str("}")
    }

    pub fn write_latent_styles(
        &mut self,
        styles: Option<&crate::LatentStyles<'_>>,
    ) -> io::Result<()> {
        let Some(styles) = styles else {
            return Ok(());
        };
        styles
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\latentstyles")?;
        self.write_control_word("lsdstimax", Some(styles.max_style_index as i32))?;
        self.write_optional_bool("lsdlockeddef", styles.locked_default)?;
        self.write_optional_bool("lsdsemihiddendef", styles.semi_hidden_default)?;
        self.write_optional_bool("lsdunhideuseddef", styles.unhide_when_used_default)?;
        self.write_optional_bool("lsdqformatdef", styles.quick_format_default)?;
        if let Some(priority) = styles.priority_default {
            self.write_control_word("lsdprioritydef", Some(i32::from(priority)))?;
        }
        if !styles.exceptions.is_empty() {
            self.write_str("{\\lsdlockedexcept ")?;
            for exception in &styles.exceptions {
                self.write_optional_bool("lsdlocked", exception.locked)?;
                self.write_optional_bool("lsdsemihidden", exception.semi_hidden)?;
                self.write_optional_bool("lsdunhideused", exception.unhide_when_used)?;
                self.write_optional_bool("lsdqformat", exception.quick_format)?;
                if let Some(priority) = exception.priority {
                    self.write_control_word("lsdpriority", Some(i32::from(priority)))?;
                }
                self.write_destination_text(exception.name.as_ref())?;
                self.write_str(";")?;
            }
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    pub fn write_data_store(
        &mut self,
        data_store: Option<&crate::DocumentDataStore<'_>>,
    ) -> io::Result<()> {
        let Some(data_store) = data_store else {
            return Ok(());
        };
        data_store
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_hex_destination("datastore", data_store.data.as_ref())
    }

    /// Write inert RTF 1.9.1 mail-merge metadata without evaluating it.
    pub fn write_mail_merge(&mut self, merge: Option<&crate::MailMerge<'_>>) -> io::Result<()> {
        let Some(merge) = merge else {
            return Ok(());
        };
        merge
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\mailmerge")?;
        self.write_mail_merge_text("mmconnectstr", merge.connect_string.as_deref())?;
        self.write_mail_merge_text("mmconnectstrdata", merge.connect_string_data.as_deref())?;
        self.write_mail_merge_text("mmquery", merge.query.as_deref())?;
        self.write_mail_merge_text("mmdatasource", merge.data_source.as_deref())?;
        self.write_mail_merge_text("mmheadersource", merge.header_source.as_deref())?;
        if merge.link_to_query {
            self.write_control_word("mmlinktoquery", None)?;
        }
        if let Some(object) = &merge.data_source_object {
            self.write_str("{\\*\\mmodso")?;
            if let Some(value) = object.active_record {
                self.write_control_word("mmodsoactive", Some(value as i32))?;
            }
            if let Some(value) = object.column_delimiter {
                self.write_control_word("mmodsocoldelim", Some(value))?;
            }
            if let Some(value) = object.column_count {
                self.write_control_word("mmodsocolumn", Some(value as i32))?;
            }
            self.write_optional_bool("mmodsodynaddr", object.dynamic_address)?;
            self.write_optional_bool("mmodsofhdr", object.first_row_header)?;
            if let Some(value) = object.hash {
                self.write_control_word("mmodsohash", Some(value))?;
            }
            if let Some(value) = object.id {
                self.write_control_word("mmodsolid", Some(value))?;
            }
            if let Some(value) = object.source_type {
                self.write_control_word("mmodsosrc", Some(value.rtf_value()))?;
            }
            self.write_mail_merge_text("mmodsofilter", object.filter.as_deref())?;
            self.write_mail_merge_text("mmodsoname", object.name.as_deref())?;
            self.write_mail_merge_text("mmodsosort", object.sort.as_deref())?;
            self.write_mail_merge_text("mmodsotable", object.table.as_deref())?;
            self.write_mail_merge_text("mmodsoudl", object.udl.as_deref())?;
            self.write_mail_merge_text("mmodsoudldata", object.udl_data.as_deref())?;
            self.write_mail_merge_text("mmodsouniquetag", object.unique_tag.as_deref())?;
            for mapping in &object.field_mappings {
                self.write_str("{\\*\\mmodsofldmpdata")?;
                self.write_control_word(
                    "mmodsofmcolumn",
                    Some(mapping.column.rtf_value().map_err(|error| {
                        io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                    })?),
                )?;
                self.write_mail_merge_text("mmodsoname", Some(mapping.name.as_ref()))?;
                self.write_mail_merge_text("mmodsomappedname", mapping.mapped_name.as_deref())?;
                self.write_str("}")?;
            }
            for value in &object.recipient_data {
                self.write_mail_merge_text("mmodsorecipdata", Some(value.as_ref()))?;
            }
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    fn write_mail_merge_text(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value else {
            return Ok(());
        };
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    pub fn write_math_properties(
        &mut self,
        properties: Option<&crate::DocumentMathProperties>,
    ) -> io::Result<()> {
        let Some(properties) = properties else {
            return Ok(());
        };
        properties
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\mmathPr")?;
        if let Some(value) = properties.binary_operator_break {
            self.write_control_word("mbrkBin", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.binary_subtraction_break {
            self.write_control_word("mbrkBinSub", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.default_justification {
            self.write_control_word("mdefJc", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.display_defaults {
            self.write_control_word("mdispDef", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.inter_equation_spacing {
            self.write_control_word("minterSp", Some(value))?;
        }
        if let Some(value) = properties.integral_limit_placement {
            self.write_control_word("mintLim", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.intra_equation_spacing {
            self.write_control_word("mintraSp", Some(value))?;
        }
        if let Some(value) = properties.left_margin {
            self.write_control_word("mlMargin", Some(value))?;
        }
        if let Some(value) = properties.math_font {
            self.write_control_word("mmathFont", Some(value as i32))?;
        }
        if let Some(value) = properties.nary_limit_placement {
            self.write_control_word("mnaryLim", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.post_spacing {
            self.write_control_word("mpostSp", Some(value))?;
        }
        if let Some(value) = properties.pre_spacing {
            self.write_control_word("mpreSp", Some(value))?;
        }
        if let Some(value) = properties.right_margin {
            self.write_control_word("mrMargin", Some(value))?;
        }
        if let Some(value) = properties.small_fractions {
            self.write_control_word("msmallFrac", Some(value.rtf_value()))?;
        }
        if let Some(value) = properties.wrap_indent {
            self.write_control_word("mwrapIndent", Some(value))?;
        }
        if let Some(value) = properties.wrap_right {
            self.write_control_word("mwrapRight", Some(value.rtf_value()))?;
        }
        self.write_str("}")
    }

    fn write_optional_bool(&mut self, control: &str, value: Option<bool>) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    /// Write an RTF stylesheet destination.
    pub fn write_stylesheet(&mut self, stylesheet: &StyleSheet<'_>) -> io::Result<()> {
        if stylesheet.styles().is_empty() {
            return Ok(());
        }
        stylesheet
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        self.write_str("{")?;
        self.write_control_word("stylesheet", None)?;
        for style in stylesheet.styles() {
            self.write_style_definition(style)?;
        }
        self.write_str("}")?;
        Ok(())
    }

    fn write_default_formatting_destinations(
        &mut self,
        defaults: &crate::DocumentDefaultFormatting,
    ) -> io::Result<()> {
        defaults
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for kind in defaults.destination_order() {
            self.write_str("{\\*")?;
            match kind {
                crate::DefaultFormattingDestination::Character => {
                    let value = defaults.character().ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "missing defchp value")
                    })?;
                    self.write_control_word("defchp", None)?;
                    self.write_formatting(&value.formatting)?;
                    for (control, font) in [
                        ("loch", value.low_ansi_font),
                        ("hich", value.high_ansi_font),
                        ("dbch", value.double_byte_font),
                    ] {
                        if let Some(font) = font {
                            self.write_control_word(control, None)?;
                            self.write_control_word("af", Some(i32::from(font)))?;
                        }
                    }
                },
                crate::DefaultFormattingDestination::Paragraph => {
                    let value = defaults.paragraph().ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "missing defpap value")
                    })?;
                    self.write_control_word("defpap", None)?;
                    self.write_paragraph_properties(&value.paragraph)?;
                    if let Some(level) = value.table_nesting_level {
                        self.write_control_word("itap", Some(i32::from(level)))?;
                    }
                },
            }
            self.write_str("}")?;
        }
        Ok(())
    }

    fn write_style_definition(&mut self, style: &Style<'_>) -> io::Result<()> {
        if style.name.contains(';') {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF style names cannot contain a semicolon",
            ));
        }
        self.write_str("{")?;
        let control = match style.style_type {
            StyleType::Paragraph => "s",
            StyleType::Character => "cs",
            StyleType::Section => "ds",
            StyleType::Table => "ts",
        };
        if style.style_type != StyleType::Paragraph {
            self.write_str("\\*")?;
        }
        self.write_control_word(control, Some(i32::from(style.id)))?;
        if style.style_type == StyleType::Table {
            if style.table_conditional.row_defaults_marker {
                self.write_control_word("tsrowd", None)?;
            }
        } else if !style.table_conditional.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table-style conditional metadata requires a table style",
            ));
        }
        self.write_formatting(&style.formatting)?;
        if let Some(paragraph) = &style.paragraph {
            self.write_paragraph_properties(paragraph)?;
        }
        if let Some(value) = style.based_on {
            self.write_control_word("sbasedon", Some(i32::from(value)))?;
        }
        if let Some(value) = style.next_style {
            self.write_control_word("snext", Some(i32::from(value)))?;
        }
        if let Some(value) = style.linked_style {
            self.write_control_word("slink", Some(i32::from(value)))?;
        }
        if style.additive {
            self.write_control_word("additive", None)?;
        }
        if style.auto_update {
            self.write_control_word("sautoupd", None)?;
        }
        if style.hidden {
            self.write_control_word("shidden", None)?;
        }
        if style.locked {
            self.write_control_word("slocked", None)?;
        }
        if style.semi_hidden {
            self.write_control_word("ssemihidden", None)?;
        }
        if style.unhide_when_used {
            self.write_control_word("sunhideused", None)?;
        }
        if style.quick_format {
            self.write_control_word("sqformat", None)?;
        }
        if let Some(value) = style.priority {
            self.write_control_word("spriority", Some(value))?;
        }
        if let Some(value) = style.revision_id {
            self.write_control_word("styrsid", Some(value))?;
        }
        if style.personal {
            self.write_control_word("spersonal", None)?;
        }
        if style.compose {
            self.write_control_word("scompose", None)?;
        }
        if style.reply {
            self.write_control_word("sreply", None)?;
        }
        if style.style_type == StyleType::Table {
            let conditional = &style.table_conditional;
            for (flag, word) in [
                (conditional.first_row, "tscfirstrow"),
                (conditional.last_row, "tsclastrow"),
                (conditional.first_column, "tscfirstcol"),
                (conditional.last_column, "tsclastcol"),
                (conditional.band_horizontal_odd, "tscbandhorzodd"),
                (conditional.band_horizontal_even, "tscbandhorzeven"),
                (conditional.band_vertical_odd, "tscbandvertodd"),
                (conditional.band_vertical_even, "tscbandverteven"),
            ] {
                if flag {
                    self.write_control_word(word, None)?;
                }
            }
            if let Some(size) = conditional.horizontal_band_size {
                self.write_control_word("tscbandsh", Some(i32::from(size)))?;
            }
            if let Some(size) = conditional.vertical_band_size {
                self.write_control_word("tscbandsv", Some(i32::from(size)))?;
            }
        }
        self.write_str(" ")?;
        self.write_text(style.name.as_ref())?;
        self.write_str(";}")?;
        Ok(())
    }

    /// Write the unique inert document-background shape destination.
    fn write_document_background(&mut self, shape: Option<&crate::Shape<'_>>) -> io::Result<()> {
        let Some(shape) = shape else {
            return Ok(());
        };
        let right = shape
            .geometry
            .x
            .checked_add(shape.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF background shape right edge overflows",
                )
            })?;
        let bottom = shape
            .geometry
            .y
            .checked_add(shape.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF background shape bottom edge overflows",
                )
            })?;
        if shape.properties.len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF background shape property count exceeds the safety limit",
            ));
        }
        self.write_str("{\\*")?;
        self.write_control_word("background", None)?;
        self.write_str("{")?;
        self.write_control_word("shp", None)?;
        self.write_str("{\\*")?;
        self.write_control_word("shpinst", None)?;
        self.write_control_word("shpleft", Some(shape.geometry.x))?;
        self.write_control_word("shptop", Some(shape.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(shape.geometry.z_order))?;
        self.write_shape_info(&shape.info)?;
        if shape.behind_doc
            && !shape.info.iter().any(|info| matches!(info, crate::ShapeGroupInfo::BelowText(_)))
        {
            self.write_control_word("shpfblwtxt", Some(1))?;
        }
        if shape.locked
            && !shape.info.iter().any(|info| matches!(info, crate::ShapeGroupInfo::LockAnchor))
        {
            self.write_control_word("shplockanchor", None)?;
        }
        let shape_type = match shape.shape_type {
            crate::ShapeType::Rectangle => Some(1),
            crate::ShapeType::RoundRectangle => Some(2),
            crate::ShapeType::Ellipse => Some(3),
            crate::ShapeType::Arc => Some(19),
            crate::ShapeType::Line => Some(20),
            crate::ShapeType::PictureFrame => Some(75),
            crate::ShapeType::TextBox => Some(202),
            crate::ShapeType::Group => Some(0),
            crate::ShapeType::Custom(value) => Some(value),
            crate::ShapeType::Polygon | crate::ShapeType::Unknown => None,
        };
        if let Some(value) = shape_type {
            self.write_shape_scalar_property("shapeType", &value.to_string())?;
        }
        for property in &shape.properties {
            if property.name == "shapeType" || property.name == "fBackground" {
                continue;
            }
            self.write_shape_property(property)?;
        }
        self.write_shape_scalar_property("fBackground", "1")?;
        if shape.text_destination_present
            || !shape.text.is_empty()
            || !shape.text_shapes.is_empty()
            || !shape.text_shape_groups.is_empty()
            || !shape.text_story_events.is_empty()
        {
            self.write_shape_text(shape)?;
        }
        self.write_str("}")?;
        if let Some(result) = &shape.result {
            self.write_shape_result(result)?;
        }
        self.write_str("}}")
    }

    fn write_shape_text(&mut self, shape: &crate::Shape<'_>) -> io::Result<()> {
        self.write_str("{\\shptxt ")?;
        if let Some(background_color) = shape
            .text_formatting
            .and_then(|formatting| formatting.background_color)
        {
            self.write_control_word("cb", Some(i32::from(background_color)))?;
        }
        self.write_field_story(
            shape.text.as_ref(),
            &shape.text_shapes,
            &shape.text_shape_groups,
            &shape.text_drawing_order,
            &shape.text_story_events,
            &[],
            crate::FieldOwner::Other,
            DrawingStoryTextMode::ShapeText,
            0,
        )?;
        self.write_str("}")
    }

    #[allow(clippy::too_many_arguments)]
    fn write_field_story(
        &mut self,
        text: &str,
        shapes: &[crate::Shape<'_>],
        shape_groups: &[crate::ShapeGroup<'_>],
        drawing_order: &[crate::StoryDrawing],
        story_events: &[crate::StoryEvent],
        fields: &[crate::Field<'_>],
        owner: crate::FieldOwner,
        mode: DrawingStoryTextMode,
        depth: usize,
    ) -> io::Result<()> {
        crate::field::validate_story_events(
            text,
            shapes,
            shape_groups,
            drawing_order,
            story_events,
            "generic-field story",
        )
        .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        let mut start = 0usize;
        for event in story_events {
            let offset = match *event {
                crate::StoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                    shapes[index].position
                },
                crate::StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                    shape_groups[index].position
                },
                crate::StoryEvent::Field(field) => field.position,
                crate::StoryEvent::PageBreak(page_break) => page_break.position,
            };
            self.write_drawing_story_fragment(&text[start..offset], mode)?;
            match *event {
                crate::StoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                    self.write_root_shape(&shapes[index])?
                },
                crate::StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                    self.write_shape_group(&shape_groups[index], true)?
                },
                crate::StoryEvent::Field(reference) => {
                    let field = fields
                        .get(reference.field_index)
                        .filter(|field| {
                            field.owner == owner
                                && field.position == reference.position
                                && field.range_end == reference.position
                        })
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF story has an invalid generic-field owner or reference",
                            )
                        })?;
                    self.write_field_with_fields(field, fields, depth + 1)?;
                },
                crate::StoryEvent::PageBreak(_) => self.write_str("\\page ")?,
            }
            start = offset;
        }
        self.write_drawing_story_fragment(&text[start..], mode)
    }

    fn write_drawing_story_fragment(
        &mut self,
        value: &str,
        mode: DrawingStoryTextMode,
    ) -> io::Result<()> {
        if matches!(mode, DrawingStoryTextMode::Destination) {
            return self.write_destination_text(value);
        }
        if matches!(mode, DrawingStoryTextMode::Note) {
            return self.write_text(value);
        }
        let mut start = 0usize;
        for (index, character) in value.char_indices() {
            if character != '\n' && character != '\t' {
                continue;
            }
            self.write_destination_text(&value[start..index])?;
            self.write_str(if character == '\n' {
                "\\par "
            } else {
                "\\tab "
            })?;
            start = index + character.len_utf8();
        }
        self.write_destination_text(&value[start..])?;
        Ok(())
    }

    fn write_shape_result(&mut self, result: &crate::ShapeResult<'_>) -> io::Result<()> {
        result
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\shprslt")?;
        self.write_legacy_drawing(&result.drawing)?;
        self.write_str("}")
    }

    fn write_shape_group(&mut self, group: &crate::ShapeGroup<'_>, root: bool) -> io::Result<()> {
        if root {
            group
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        } else if group.result.is_some() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF nested shape group cannot contain shprslt",
            ));
        }
        let right = group
            .geometry
            .x
            .checked_add(group.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape group right edge overflows",
                )
            })?;
        let bottom = group
            .geometry
            .y
            .checked_add(group.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape group bottom edge overflows",
                )
            })?;
        self.write_str("{\\shpgrp{\\*\\shpinst")?;
        self.write_control_word("shpleft", Some(group.geometry.x))?;
        self.write_control_word("shptop", Some(group.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(group.geometry.z_order))?;
        for info in &group.info {
            match *info {
                crate::ShapeGroupInfo::ShapeId(value) => {
                    self.write_control_word("shplid", Some(value))?
                },
                crate::ShapeGroupInfo::InHeader(value) => {
                    self.write_control_word("shpfhdr", Some(i32::from(value)))?
                },
                crate::ShapeGroupInfo::HorizontalPage => {
                    self.write_control_word("shpbxpage", None)?
                },
                crate::ShapeGroupInfo::HorizontalMargin => {
                    self.write_control_word("shpbxmargin", None)?
                },
                crate::ShapeGroupInfo::HorizontalColumn => {
                    self.write_control_word("shpbxcolumn", None)?
                },
                crate::ShapeGroupInfo::IgnoreHorizontal => {
                    self.write_control_word("shpbxignore", None)?
                },
                crate::ShapeGroupInfo::VerticalPage => {
                    self.write_control_word("shpbypage", None)?
                },
                crate::ShapeGroupInfo::VerticalMargin => {
                    self.write_control_word("shpbymargin", None)?
                },
                crate::ShapeGroupInfo::VerticalParagraph => {
                    self.write_control_word("shpbypara", None)?
                },
                crate::ShapeGroupInfo::IgnoreVertical => {
                    self.write_control_word("shpbyignore", None)?
                },
                crate::ShapeGroupInfo::Wrap(value) => {
                    self.write_control_word("shpwr", Some(value))?
                },
                crate::ShapeGroupInfo::WrapSide(value) => {
                    self.write_control_word("shpwrk", Some(value))?
                },
                crate::ShapeGroupInfo::BelowText(value) => {
                    self.write_control_word("shpfblwtxt", Some(i32::from(value)))?
                },
                crate::ShapeGroupInfo::LockAnchor => {
                    self.write_control_word("shplockanchor", None)?
                },
            }
        }
        for property in &group.properties {
            self.write_shape_property(property)?;
        }
        for child in &group.child_order {
            match *child {
                crate::ShapeGroupChild::Shape(index) => {
                    self.write_grouped_shape(&group.shapes[index])?
                },
                crate::ShapeGroupChild::Group(index) => {
                    self.write_shape_group(&group.groups[index], false)?
                },
            }
        }
        self.write_str("}")?;
        if let Some(result) = &group.result {
            self.write_shape_result(result)?;
        }
        self.write_str("}")
    }

    fn write_grouped_shape(&mut self, shape: &crate::Shape<'_>) -> io::Result<()> {
        if shape.result.is_some() || !shape.instruction_present {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF grouped shape cannot contain shprslt",
            ));
        }
        let right = shape
            .geometry
            .x
            .checked_add(shape.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF grouped shape right edge overflows",
                )
            })?;
        let bottom = shape
            .geometry
            .y
            .checked_add(shape.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF grouped shape bottom edge overflows",
                )
            })?;
        self.write_str("{\\shp{\\*\\shpinst")?;
        self.write_control_word("shpleft", Some(shape.geometry.x))?;
        self.write_control_word("shptop", Some(shape.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(shape.geometry.z_order))?;
        self.write_shape_info(&shape.info)?;
        if shape.behind_doc
            && !shape.info.iter().any(|info| matches!(info, crate::ShapeGroupInfo::BelowText(_)))
        {
            self.write_control_word("shpfblwtxt", Some(1))?;
        }
        if shape.locked
            && !shape.info.iter().any(|info| matches!(info, crate::ShapeGroupInfo::LockAnchor))
        {
            self.write_control_word("shplockanchor", None)?;
        }
        if !shape
            .properties
            .iter()
            .any(|property| property.name == "shapeType")
        {
            let shape_type = match shape.shape_type {
                crate::ShapeType::Rectangle => Some(1),
                crate::ShapeType::RoundRectangle => Some(2),
                crate::ShapeType::Ellipse => Some(3),
                crate::ShapeType::Arc => Some(19),
                crate::ShapeType::Line => Some(20),
                crate::ShapeType::PictureFrame => Some(75),
                crate::ShapeType::TextBox => Some(202),
                crate::ShapeType::Group => Some(0),
                crate::ShapeType::Custom(value) => Some(value),
                crate::ShapeType::Polygon | crate::ShapeType::Unknown => None,
            };
            if let Some(value) = shape_type {
                self.write_shape_scalar_property("shapeType", &value.to_string())?;
            }
        }
        for property in &shape.properties {
            self.write_shape_property(property)?;
        }
        if shape.text_destination_present
            || !shape.text.is_empty()
            || !shape.text_shapes.is_empty()
            || !shape.text_shape_groups.is_empty()
            || !shape.text_story_events.is_empty()
        {
            self.write_shape_text(shape)?;
        }
        self.write_str("}}")
    }

    fn write_shape_scalar_property(&mut self, name: &str, value: &str) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("sp", None)?;
        self.write_str("{")?;
        self.write_control_word("sn", None)?;
        self.write_str(" ")?;
        self.write_text(name)?;
        self.write_str("}{")?;
        self.write_control_word("sv", None)?;
        self.write_str(" ")?;
        self.write_text(value)?;
        self.write_str("}}")
    }

    fn write_shape_property(&mut self, property: &crate::ShapeProperty<'_>) -> io::Result<()> {
        property
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        self.write_str("{\\sp{\\sn ")?;
        self.write_destination_text(property.name.as_ref())?;
        self.write_str("}{\\sv")?;
        if let Some(value) = &property.binary_value {
            self.write_str("{\\*\\svb ")?;
            for byte in value.iter() {
                write!(self.writer, "{byte:02x}")?;
            }
            self.write_str("}")?;
        } else {
            self.write_str(" ")?;
            self.write_destination_text(property.value.as_ref())?;
        }
        self.write_str("}")?;
        if let Some(theme) = property.theme_value {
            self.write_str("{\\*\\hsv")?;
            self.write_control_word(
                match theme.color {
                    crate::ShapeThemeColor::Accent1 => "caccentone",
                    crate::ShapeThemeColor::Accent2 => "caccenttwo",
                    crate::ShapeThemeColor::Accent3 => "caccentthree",
                    crate::ShapeThemeColor::Accent4 => "caccentfour",
                    crate::ShapeThemeColor::Accent5 => "caccentfive",
                    crate::ShapeThemeColor::Accent6 => "caccentsix",
                    crate::ShapeThemeColor::Background1 => "cbackgroundone",
                    crate::ShapeThemeColor::Background2 => "cbackgroundtwo",
                    crate::ShapeThemeColor::Text1 => "ctextone",
                    crate::ShapeThemeColor::Text2 => "ctexttwo",
                },
                None,
            )?;
            self.write_control_word("ctint", Some(i32::from(theme.tint)))?;
            self.write_control_word("cshade", Some(i32::from(theme.shade)))?;
            self.write_str("}")?;
        }
        if let Some(hyperlink) = &property.hyperlink {
            hyperlink
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            self.write_str("{\\hl")?;
            if let Some(location) = &hyperlink.location {
                self.write_str("{\\hlloc ")?;
                self.write_destination_text(location.as_ref())?;
                self.write_str("}")?;
            }
            if let Some(source) = &hyperlink.source {
                self.write_str("{\\hlsrc ")?;
                self.write_destination_text(source.as_ref())?;
                self.write_str("}")?;
            }
            if let Some(friendly_name) = &hyperlink.friendly_name {
                self.write_str("{\\hlfr ")?;
                self.write_destination_text(friendly_name.as_ref())?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    /// Write the standard RTF document-information destination.
    pub fn write_document_info(&mut self, info: &DocumentInfo<'_>) -> io::Result<()> {
        let has_info = info.title.is_some()
            || info.subject.is_some()
            || info.author.is_some()
            || info.manager.is_some()
            || info.company.is_some()
            || info.operator.is_some()
            || info.category.is_some()
            || info.keywords.is_some()
            || info.comment.is_some()
            || info.document_comment.is_some()
            || info.hyperlink_base.is_some()
            || info.version.is_some()
            || info.revision.is_some()
            || info.creation_time.is_some()
            || info.revision_time.is_some()
            || info.print_time.is_some()
            || info.backup_time.is_some()
            || info.creation_timestamp.is_some()
            || info.revision_timestamp.is_some()
            || info.print_timestamp.is_some()
            || info.backup_timestamp.is_some()
            || info.editing_time.is_some()
            || info.pages.is_some()
            || info.words.is_some()
            || info.characters.is_some()
            || info.characters_with_spaces.is_some()
            || info.id.is_some()
            || info.protection.password_hash.is_some();
        info.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        if has_info {
            self.write_str("{")?;
            self.write_control_word("info", None)?;
            self.write_info_text("title", info.title.as_deref())?;
            self.write_info_text("subject", info.subject.as_deref())?;
            self.write_info_text("author", info.author.as_deref())?;
            self.write_info_text("manager", info.manager.as_deref())?;
            self.write_info_text("company", info.company.as_deref())?;
            self.write_info_text("operator", info.operator.as_deref())?;
            self.write_info_text("category", info.category.as_deref())?;
            self.write_info_text("keywords", info.keywords.as_deref())?;
            self.write_info_text("comment", info.comment.as_deref())?;
            self.write_info_text("doccomm", info.document_comment.as_deref())?;
            self.write_info_text("hlinkbase", info.hyperlink_base.as_deref())?;
            self.write_info_time(
                "creatim",
                info.creation_timestamp,
                info.creation_time.as_deref(),
            )?;
            self.write_info_time(
                "revtim",
                info.revision_timestamp,
                info.revision_time.as_deref(),
            )?;
            self.write_info_time("printim", info.print_timestamp, info.print_time.as_deref())?;
            self.write_info_time("buptim", info.backup_timestamp, info.backup_time.as_deref())?;
            self.write_optional_u32("version", info.version)?;
            self.write_optional_u32("vern", info.revision)?;
            self.write_optional_u32("edmins", info.editing_time)?;
            self.write_optional_u32("nofpages", info.pages)?;
            self.write_optional_u32("nofwords", info.words)?;
            self.write_optional_u32("nofchars", info.characters)?;
            self.write_optional_u32("nofcharsws", info.characters_with_spaces)?;
            self.write_optional_u32("id", info.id)?;
            if let Some(hash) = info.protection.password_hash.as_deref() {
                self.write_str("{\\*")?;
                self.write_control_word("password", None)?;
                self.write_str(" ")?;
                self.write_str(hash)?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        self.write_protection_controls(&info.protection)
    }

    fn write_protection_controls(
        &mut self,
        protection: &crate::DocumentProtection<'_>,
    ) -> io::Result<()> {
        for (control, value) in [
            ("formprot", protection.forms),
            ("annotprot", protection.annotations),
            ("revprot", protection.revisions),
            ("readprot", protection.read_only),
            ("allprot", protection.all),
        ] {
            if let Some(value) = value {
                self.write_control_word(control, (!value).then_some(0))?;
            }
        }
        if let Some(value) = protection.enforced {
            self.write_control_word("enforceprot", Some(i32::from(value)))?;
        }
        if let Some(level) = protection.level {
            self.write_control_word("protlevel", Some(level.rtf_value()))?;
        }
        Ok(())
    }

    fn write_info_text(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value else { return Ok(()) };
        self.write_str("{")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    fn write_info_time(
        &mut self,
        control: &str,
        typed: Option<RtfTimestamp>,
        legacy: Option<&str>,
    ) -> io::Result<()> {
        let timestamp = match (typed, legacy) {
            (Some(value), _) => value,
            (None, Some(value)) => RtfTimestamp::from_legacy(value)
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?,
            (None, None) => return Ok(()),
        };
        self.write_str("{")?;
        self.write_control_word(control, None)?;
        for (name, value) in [
            ("yr", timestamp.year),
            ("mo", timestamp.month),
            ("dy", timestamp.day),
            ("hr", timestamp.hour),
            ("min", timestamp.minute),
            ("sec", timestamp.second),
        ] {
            if let Some(value) = value {
                self.write_control_word(name, Some(value))?;
            }
        }
        self.write_str("}")
    }

    fn write_optional_u32(&mut self, control: &str, value: Option<u32>) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(value as i32))?;
        }
        Ok(())
    }

    fn write_optional_i32(&mut self, control: &str, value: Option<i32>) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(value))?;
        }
        Ok(())
    }

    /// Write the canonical starred RTF user-properties destination.
    pub fn write_user_properties(
        &mut self,
        properties: &[crate::UserProperty<'_>],
    ) -> io::Result<()> {
        if properties.is_empty() {
            return Ok(());
        }
        if properties.len() > crate::user_property::MAX_USER_PROPERTIES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF user-property count limit exceeded",
            ));
        }
        let mut names = std::collections::HashSet::with_capacity(properties.len());
        let mut aggregate = 0usize;
        for property in properties {
            property
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if !names.insert(property.name.as_ref()) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "duplicate RTF user-property name",
                ));
            }
            aggregate = aggregate
                .checked_add(property.text_bytes().ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "user-property size overflow")
                })?)
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "user-property size overflow")
                })?;
            if aggregate > crate::user_property::MAX_USER_PROPERTY_TEXT_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF user-property aggregate text limit exceeded",
                ));
            }
        }

        self.write_str("{\\*")?;
        self.write_control_word("userprops", None)?;
        for property in properties {
            self.write_str("{")?;
            self.write_control_word("propname", None)?;
            self.write_str(" ")?;
            self.write_destination_text(property.name.as_ref())?;
            self.write_str("}")?;
            self.write_control_word("proptype", Some(property.value.type_code()))?;
            self.write_str("{")?;
            self.write_control_word("staticval", None)?;
            self.write_str(" ")?;
            self.write_destination_text(property.value.lexical())?;
            self.write_str("}")?;
            if let Some(link) = &property.link_value {
                self.write_str("{")?;
                self.write_control_word("linkval", None)?;
                self.write_str(" ")?;
                self.write_destination_text(link.as_ref())?;
                self.write_str("}")?;
            }
        }
        self.write_str("}")
    }

    /// Write ordered standard RTF document-variable destinations.
    pub fn write_document_variables(
        &mut self,
        variables: &[crate::DocumentVariable<'_>],
    ) -> io::Result<()> {
        if variables.len() > crate::document_variable::MAX_DOCUMENT_VARIABLES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF document-variable count limit exceeded",
            ));
        }
        let mut aggregate = 0usize;
        for variable in variables {
            variable
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            aggregate = aggregate
                .checked_add(variable.name.len())
                .and_then(|size| size.checked_add(variable.value.len()))
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "document-variable size overflow",
                    )
                })?;
            if aggregate > crate::document_variable::MAX_DOCUMENT_VARIABLE_TEXT_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF document-variable aggregate text limit exceeded",
                ));
            }
            self.write_str("{\\*")?;
            self.write_control_word("docvar", None)?;
            self.write_str(" {")?;
            self.write_destination_text(variable.name.as_ref())?;
            self.write_str("}{")?;
            self.write_destination_text(variable.value.as_ref())?;
            self.write_str("}}")?;
        }
        Ok(())
    }

    fn write_destination_text(&mut self, text: &str) -> io::Result<()> {
        for character in text.chars() {
            match character {
                '\\' => self.write_str("\\\\")?,
                '{' => self.write_str("\\{")?,
                '}' => self.write_str("\\}")?,
                character if character.is_ascii_control() => {
                    write!(self.writer, "\\'{:02x}", character as u8)?;
                },
                character if character.is_ascii() => write!(self.writer, "{character}")?,
                character => {
                    for unit in character.encode_utf16(&mut [0; 2]).iter().copied() {
                        self.write_control_word("u", Some(i32::from(unit as i16)))?;
                        self.write_str("?")?;
                    }
                },
            }
        }
        Ok(())
    }

    /// Write a bookmark start destination.
    pub fn write_bookmark_start(&mut self, bookmark: &Bookmark<'_>) -> io::Result<()> {
        if bookmark.name.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark name cannot be empty",
            ));
        }
        self.write_str("{\\*")?;
        self.write_control_word("bkmkstart", None)?;
        self.write_optional_i32("bkmkcolf", bookmark.first_column)?;
        self.write_optional_i32("bkmkcoll", bookmark.last_column)?;
        if bookmark.is_public {
            self.write_control_word("bkmkpub", None)?;
        }
        self.write_str(" ")?;
        self.write_text(bookmark.name.as_ref())?;
        self.write_str("}")
    }

    /// Write a bookmark end destination.
    pub fn write_bookmark_end(&mut self, name: &str) -> io::Result<()> {
        if name.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark name cannot be empty",
            ));
        }
        self.write_str("{\\*")?;
        self.write_control_word("bkmkend", None)?;
        self.write_str(" ")?;
        self.write_text(name)?;
        self.write_str("}")
    }

    /// Write a custom XML tag open destination and its inert attributes.
    pub fn write_custom_xml_open(&mut self, tag: &crate::CustomXmlTag<'_>) -> io::Result<()> {
        tag.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word("xmlopen", None)?;
        if let Some(namespace) = tag.namespace {
            self.write_control_word("xmlns", Some(namespace as i32))?;
        }
        self.write_str(" ")?;
        self.write_destination_text(tag.name.as_ref())?;
        self.write_str("}")?;
        for attribute in &tag.attributes {
            self.write_str("{\\*")?;
            self.write_control_word("xmlattrname", None)?;
            self.write_str(" ")?;
            self.write_destination_text(attribute.name.as_ref())?;
            self.write_str("}{\\*")?;
            self.write_control_word("xmlattrvalue", None)?;
            self.write_str(" ")?;
            self.write_destination_text(attribute.value.as_ref())?;
            self.write_str("}")?;
        }
        Ok(())
    }

    /// Write a custom XML tag close destination.
    pub fn write_custom_xml_close(&mut self, tag: &crate::CustomXmlTag<'_>) -> io::Result<()> {
        if tag.name.is_empty() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF custom XML tag name cannot be empty",
            ));
        }
        self.write_str("{")?;
        self.write_control_word("xmlclose", None)?;
        self.write_str(" ")?;
        self.write_destination_text(tag.name.as_ref())?;
        self.write_str("}")
    }

    /// Write a protection-exception range marker destination.
    pub fn write_protection_range_marker(
        &mut self,
        control: &str,
        range: &crate::ProtectionRange<'_>,
    ) -> io::Result<()> {
        range
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(range.id.as_ref())?;
        self.write_str("}")
    }

    fn math_structure_control(kind: crate::MathStructureKind) -> &'static str {
        use crate::MathStructureKind as K;
        match kind {
            K::Accent => "macc",
            K::Bar => "mbar",
            K::BorderBox => "mborderBox",
            K::Box => "mbox",
            K::Delimiter => "md",
            K::EquationArray => "meqArr",
            K::Fraction => "mf",
            K::Function => "mfunc",
            K::GroupChar => "mgroupChr",
            K::LimitLower => "mlimlow",
            K::LimitUpper => "mlimupp",
            K::Matrix => "mm",
            K::Nary => "mnary",
            K::Phantom => "mphant",
            K::Radical => "mrad",
            K::ScriptPre => "msPre",
            K::ScriptSub => "msSub",
            K::ScriptSubSup => "msSubSup",
            K::ScriptSup => "msSup",
        }
    }

    fn math_structure_properties_control(kind: crate::MathStructureKind) -> &'static str {
        use crate::MathStructureKind as K;
        match kind {
            K::Accent => "maccPr",
            K::Bar => "mbarPr",
            K::BorderBox => "mborderBoxPr",
            K::Box => "mboxPr",
            K::Delimiter => "mdPr",
            K::EquationArray => "meqArrPr",
            K::Fraction => "mfPr",
            K::Function => "mfuncPr",
            K::GroupChar => "mgroupChrPr",
            K::LimitLower => "mlimlowPr",
            K::LimitUpper => "mlimuppPr",
            K::Matrix => "mmPr",
            K::Nary => "mnaryPr",
            K::Phantom => "mphantPr",
            K::Radical => "mradPr",
            K::ScriptPre => "msPrePr",
            K::ScriptSub => "msSubPr",
            K::ScriptSubSup => "msSubSupPr",
            K::ScriptSup => "msSupPr",
        }
    }

    fn math_element_control(role: crate::MathElementRole) -> &'static str {
        use crate::MathElementRole as R;
        match role {
            R::Element => "me",
            R::Numerator => "mnum",
            R::Denominator => "mden",
            R::Degree => "mdeg",
            R::Subscript => "msub",
            R::Superscript => "msup",
            R::Limit => "mlim",
            R::FunctionName => "mfName",
        }
    }

    fn math_property_control(name: crate::MathPropertyName) -> &'static str {
        use crate::MathPropertyName as N;
        match name {
            N::Type => "mtype",
            N::Grow => "mgrow",
            N::Char => "mchr",
            N::BeginChar => "mbegChr",
            N::EndChar => "mendChr",
            N::SeparatorChar => "msepChr",
            N::Position => "mpos",
            N::VerticalJustify => "mvertJc",
            N::BaseJustify => "mbaseJc",
            N::Justify => "mjc",
            N::Align => "maln",
            N::AlignScript => "malnScr",
            N::DegreeHide => "mdegHide",
            N::Differential => "mdiff",
            N::DifferentialStyle => "mdiffSty",
            N::HideBottom => "mhideBot",
            N::HideLeft => "mhideLeft",
            N::HideRight => "mhideRight",
            N::HideTop => "mhideTop",
            N::LimitLocation => "mlimLoc",
            N::PlaceholderHide => "mplcHide",
            N::SubscriptHide => "msubHide",
            N::SuperscriptHide => "msupHide",
            N::StrikeBottomLeftToTopRight => "mstrikeBLTR",
            N::StrikeHorizontal => "mstrikeH",
            N::StrikeTopLeftToBottomRight => "mstrikeTLBR",
            N::StrikeVertical => "mstrikeV",
            N::Style => "msty",
            N::Script => "mscr",
            N::Transparent => "mtransp",
            N::Show => "mshow",
            N::Shape => "mshp",
            N::ZeroAscent => "mzeroAsc",
            N::ZeroDescent => "mzeroDesc",
            N::ZeroWidth => "mzeroWid",
            N::OperatorEmulator => "mopEmu",
            N::NoBreak => "mnoBreak",
            N::NormalText => "mnor",
            N::Literal => "mlit",
            N::MatrixColumnGap => "mcGp",
            N::MatrixColumnGapRule => "mcGpRule",
            N::MatrixColumnSpacing => "mcSp",
            N::MatrixCellCount => "mcount",
            N::MatrixCellJustify => "mmcJc",
            N::RowSpacing => "mrSp",
            N::RowSpacingRule => "mrSpRule",
            N::Break => "mbrk",
            N::ArgumentSize => "margSz",
        }
    }

    /// Write an inert math zone destination.
    pub fn write_math_zone(&mut self, zone: &crate::MathZone<'_>) -> io::Result<()> {
        zone.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word(
            match zone.kind {
                crate::MathZoneKind::Inline => "mmath",
                crate::MathZoneKind::Display => "mmathPara",
            },
            None,
        )?;
        if let Some(properties) = &zone.paragraph_properties {
            self.write_math_properties_group("mmathParaPr", properties)?;
        }
        for object in &zone.content {
            self.write_math_object(object)?;
        }
        self.write_str("}")
    }

    fn write_math_object(&mut self, object: &crate::MathObject<'_>) -> io::Result<()> {
        match object {
            crate::MathObject::Structure(structure) => self.write_math_structure(structure),
            crate::MathObject::Run(run) => self.write_math_run(run),
        }
    }

    fn write_math_structure(&mut self, structure: &crate::MathStructure<'_>) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word(Self::math_structure_control(structure.kind), None)?;
        if let Some(properties) = &structure.properties {
            self.write_math_properties_group(
                Self::math_structure_properties_control(structure.kind),
                properties,
            )?;
        }
        for child in &structure.children {
            match child {
                crate::MathStructureChild::Element(element) => {
                    self.write_math_element(element)?
                },
                crate::MathStructureChild::MatrixRow(row) => {
                    self.write_str("{")?;
                    self.write_control_word("mmr", None)?;
                    for cell in &row.cells {
                        self.write_math_element(cell)?;
                    }
                    self.write_str("}")?;
                },
            }
        }
        self.write_str("}")
    }

    fn write_math_element(&mut self, element: &crate::MathElement<'_>) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word(Self::math_element_control(element.role), None)?;
        if let Some(properties) = &element.argument_properties {
            self.write_math_properties_group("margPr", properties)?;
        }
        self.write_str(" ")?;
        for object in &element.content {
            self.write_math_object(object)?;
        }
        self.write_str("}")
    }

    fn write_math_run(&mut self, run: &crate::MathRun<'_>) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("mr", None)?;
        if let Some(properties) = &run.properties {
            self.write_math_properties_group("mrPr", properties)?;
        }
        if run.normal_text {
            self.write_control_word("mnor", None)?;
        }
        self.write_str(" ")?;
        self.write_destination_text(run.text.as_ref())?;
        self.write_str("}")
    }

    fn write_math_properties_group(
        &mut self,
        destination: &str,
        properties: &crate::MathProperties<'_>,
    ) -> io::Result<()> {
        properties
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word(destination, None)?;
        for property in &properties.properties {
            self.write_str("{")?;
            self.write_control_word(Self::math_property_control(property.name), None)?;
            if !property.value.is_empty() {
                self.write_str(" ")?;
                self.write_destination_text(property.value.as_ref())?;
            }
            self.write_str("}")?;
        }
        if !properties.matrix_columns.is_empty() {
            self.write_str("{")?;
            self.write_control_word("mmcs", None)?;
            for column in &properties.matrix_columns {
                self.write_str("{")?;
                self.write_control_word("mmc", None)?;
                if let Some(column_properties) = &column.properties {
                    self.write_math_properties_group("mmcPr", column_properties)?;
                }
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        if let Some(control) = &properties.control {
            self.write_math_properties_group("mctrlPr", control)?;
        }
        self.write_str("}")
    }

    /// Write an annotation range-start destination.
    pub fn write_annotation_start(&mut self, annotation: &Annotation<'_>) -> io::Result<()> {
        if !annotation.has_reference {
            return Ok(());
        }
        self.write_str("{\\*")?;
        self.write_control_word("atrfstart", None)?;
        self.write_str(" ")?;
        write!(self.writer, "{}", annotation.id)?;
        self.write_str("}")
    }

    /// Write an annotation range end, author metadata, and inert comment body.
    pub fn write_annotation_end(&mut self, annotation: &Annotation<'_>) -> io::Result<()> {
        if annotation.annotation_type != AnnotationType::Comment {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "only comment annotations use the RTF annotation destination",
            ));
        }
        annotation
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        if annotation.has_reference {
            self.write_str("{\\*")?;
            self.write_control_word("atrfend", None)?;
            self.write_str(" ")?;
            write!(self.writer, "{}", annotation.id)?;
            self.write_str("}")?;
        }
        self.write_annotation_value("atnid", Some(annotation.initials.as_ref()))?;
        self.write_annotation_value("atnauthor", Some(annotation.author.as_ref()))?;
        self.write_control_word("chatn", None)?;
        self.write_str("{\\*")?;
        self.write_control_word("annotation", None)?;
        self.write_str(" ")?;
        let reference = annotation.has_reference.then(|| annotation.id.to_string());
        self.write_annotation_value("atnref", reference.as_deref())?;
        self.write_annotation_value("atndate", annotation.date.as_deref())?;
        self.write_annotation_value("atnparent", annotation.parent_id.as_deref())?;
        self.write_annotation_value("atnicn", annotation.icon.as_deref())?;
        self.write_annotation_value("atntime", annotation.time.as_deref())?;
        self.write_field_story(
            annotation.text.as_ref(),
            &annotation.shapes,
            &annotation.shape_groups,
            &annotation.drawing_order,
            &annotation.story_events,
            &[],
            crate::FieldOwner::Other,
            DrawingStoryTextMode::Destination,
            0,
        )?;
        self.write_str("}")
    }

    fn write_annotation_value(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value else {
            return Ok(());
        };
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    #[allow(clippy::too_many_arguments)]
    fn write_blocks_with_markup(
        &mut self,
        blocks: &[StyleBlock<'_>],
        bookmarks: &BookmarkTable<'_>,
        custom_xml_tags: &[crate::CustomXmlTag<'_>],
        math_zones: &[crate::MathZone<'_>],
        protection_ranges: &[crate::ProtectionRange<'_>],
        editable_regions: &[crate::EditableRegion<'_>],
        annotations: &[Annotation<'_>],
        notes: &[Note<'_>],
        revisions: &[Revision<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        generated_list_markers: &[crate::GeneratedListMarker<'_>],
        shapes: &[crate::Shape<'_>],
        shape_groups: &[crate::ShapeGroup<'_>],
        drawing_order: &[crate::StoryDrawing],
        picture_compatibility_records: &[crate::PictureCompatibilityRecord],
        pictures: &[crate::Picture<'_>],
        objects: &[crate::EmbeddedObject<'_>],
        legacy_text_boxes: &[crate::LegacyTextBox<'_>],
        legacy_drawings: &[crate::LegacyDrawing<'_>],
        form_fields: &[crate::FormField<'_>],
        fields: &[crate::Field<'_>],
        sections: &[Section<'_>],
        body_story_events: &[crate::BodyStoryEvent],
    ) -> io::Result<()> {
        if bookmarks.bookmarks().is_empty()
            && custom_xml_tags.is_empty()
            && math_zones.is_empty()
            && protection_ranges.is_empty()
            && editable_regions.is_empty()
            && annotations.is_empty()
            && notes.is_empty()
            && revisions.is_empty()
            && navigation_entries.is_empty()
            && generated_list_markers.is_empty()
            && shapes.iter().all(|shape| shape.is_background)
            && shape_groups.is_empty()
            && picture_compatibility_records.is_empty()
            && objects.is_empty()
            && legacy_text_boxes.is_empty()
            && legacy_drawings.is_empty()
            && form_fields.is_empty()
            && fields
                .iter()
                .all(|field| !matches!(field.owner, crate::FieldOwner::Body))
            && body_story_events.is_empty()
        {
            for block in blocks {
                self.write_style_block(block)?;
            }
            return Ok(());
        }

        let body: String = blocks.iter().map(|block| block.text.as_ref()).collect();
        let event_count = bookmarks
            .bookmarks()
            .len()
            .saturating_add(annotations.len())
            .saturating_add(notes.len())
            .saturating_add(revisions.len())
            .saturating_mul(2);
        let event_count = event_count.saturating_add(navigation_entries.len());
        let event_count = event_count.saturating_add(custom_xml_tags.len().saturating_mul(2));
        let event_count = event_count.saturating_add(math_zones.len());
        let event_count = event_count.saturating_add(protection_ranges.len().saturating_mul(2));
        let event_count = event_count.saturating_add(editable_regions.len().saturating_mul(2));
        let event_count = event_count.saturating_add(shapes.len());
        let event_count = event_count.saturating_add(shape_groups.len());
        let event_count = event_count.saturating_add(picture_compatibility_records.len());
        let event_count = event_count.saturating_add(objects.len());
        let event_count = event_count.saturating_add(legacy_drawings.len());
        let event_count = event_count.saturating_add(form_fields.len().saturating_mul(2));
        let event_count = event_count.saturating_add(body_story_events.len());
        let mut events = Vec::with_capacity(event_count);
        if notes.len() > crate::section::MAX_NOTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF note count exceeds the safety limit",
            ));
        }
        let mut previous_note_position = None;
        let mut note_text_bytes = 0usize;
        for note in notes {
            note.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(note.position..note.position).is_none()
                || previous_note_position.is_some_and(|position| position > note.position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF notes are outside or out of main-story order",
                ));
            }
            note_text_bytes = note_text_bytes
                .checked_add(note.text_bytes().ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF note text size overflow")
                })?)
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF note text size overflow")
                })?;
            if note_text_bytes > crate::section::MAX_NOTE_TEXT_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF note aggregate text exceeds the safety limit",
                ));
            }
            events.push(BodyEvent {
                offset: note.position,
                order: 1,
                kind: BodyEventKind::Note(note),
            });
            previous_note_position = Some(note.position);
        }
        let expected_drawings = shapes
            .iter()
            .filter(|shape| !shape.is_background)
            .count()
            .saturating_add(shape_groups.len());
        if drawing_order.len() != expected_drawings {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF body drawing order is incomplete",
            ));
        }
        if fields.len() > crate::field::MAX_GENERIC_FIELDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field count exceeds the safety limit",
            ));
        }
        if objects.len() > crate::object::MAX_EMBEDDED_OBJECTS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF embedded object count exceeds the safety limit",
            ));
        }
        let mut previous_object_position = None;
        for object in objects {
            object
                .validate(&body, pictures.len())
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if previous_object_position.is_some_and(|position| position > object.position) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF embedded objects are not ordered by body position",
                ));
            }
            for picture_index in &object.result_picture_indices {
                pictures[*picture_index].validate().map_err(|error| {
                    io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                })?;
            }
            events.push(BodyEvent {
                offset: object.position,
                order: 1,
                kind: BodyEventKind::Object(object, pictures),
            });
            previous_object_position = Some(object.position);
        }
        if picture_compatibility_records.len() > crate::MAX_PICTURE_COMPATIBILITY_RECORDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF picture-compatibility record count exceeds the safety limit",
            ));
        }
        let mut previous_picture_record = None;
        for record in picture_compatibility_records {
            record
                .validate(&body, pictures.len())
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if previous_picture_record.is_some_and(
                |previous: &crate::PictureCompatibilityRecord| {
                    previous.position > record.position
                        || (previous.position == record.position && previous.kind == record.kind)
                },
            ) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF picture-compatibility records are duplicated or out of body order",
                ));
            }
            let picture = &pictures[record.picture_index];
            picture
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            events.push(BodyEvent {
                offset: record.position,
                order: 1,
                kind: BodyEventKind::PictureCompatibility(record, picture),
            });
            previous_picture_record = Some(record);
        }
        if form_fields.len() > crate::form_field::MAX_FORM_FIELDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF form-field count exceeds the safety limit",
            ));
        }
        let mut form_field_bytes = 0usize;
        let mut form_field_ranges: Vec<&crate::FormField<'_>> = form_fields.iter().collect();
        form_field_ranges.sort_by_key(|field| (field.position, field.range_end));
        let mut previous_form_end = 0usize;
        for field in form_field_ranges {
            field
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let result = body.get(field.position..field.range_end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field range is outside body text or splits a character",
                )
            })?;
            if result != field.result_text {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field result does not match its visible body range",
                ));
            }
            if field.position != field.range_end && field.position < previous_form_end {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field result ranges cannot overlap",
                ));
            }
            if field.position != field.range_end {
                previous_form_end = field.range_end;
            }
            form_field_bytes = form_field_bytes
                .checked_add(field.text_bytes().ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF form-field aggregate size overflow",
                    )
                })?)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF form-field aggregate size overflow",
                    )
                })?;
            if form_field_bytes > crate::form_field::MAX_FORM_FIELD_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field aggregate text exceeds the safety limit",
                ));
            }
            let empty = field.position == field.range_end;
            events.push(BodyEvent {
                offset: field.position,
                order: 1,
                kind: BodyEventKind::FormFieldStart(field),
            });
            events.push(BodyEvent {
                offset: field.range_end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::FormFieldEnd,
            });
        }
        if generated_list_markers.len() > crate::generated_list_marker::MAX_GENERATED_LIST_MARKERS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generated list-marker count exceeds the safety limit",
            ));
        }
        let mut generated_marker_bytes = 0usize;
        let mut previous_generated_marker: Option<&crate::GeneratedListMarker<'_>> = None;
        for marker in generated_list_markers {
            marker
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(marker.position..marker.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF generated list-marker position is not a UTF-8 body boundary",
                ));
            }
            if previous_generated_marker.is_some_and(|previous| {
                previous.position > marker.position
                    || (previous.position == marker.position && previous.kind == marker.kind)
            }) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF generated list markers are duplicated or out of body order",
                ));
            }
            generated_marker_bytes = generated_marker_bytes
                .checked_add(marker.text.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF generated list-marker text size overflow",
                    )
                })?;
            if generated_marker_bytes
                > crate::generated_list_marker::MAX_GENERATED_LIST_MARKER_TOTAL_BYTES
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF generated list-marker text exceeds the aggregate safety limit",
                ));
            }
            events.push(BodyEvent {
                offset: marker.position,
                order: 1,
                kind: BodyEventKind::GeneratedListMarker(marker),
            });
            previous_generated_marker = Some(marker);
        }

        if legacy_text_boxes.len() > crate::legacy_text_box::MAX_LEGACY_TEXT_BOXES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF legacy text-box count exceeds the safety limit",
            ));
        }
        let mut legacy_text_box_bytes = 0usize;
        let mut previous_legacy_text_box_position = None;
        for text_box in legacy_text_boxes {
            text_box
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(text_box.position..text_box.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy text-box position is not a UTF-8 body boundary",
                ));
            }
            if previous_legacy_text_box_position
                .is_some_and(|position| position > text_box.position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy text boxes are not ordered by body position",
                ));
            }
            legacy_text_box_bytes = legacy_text_box_bytes
                .checked_add(text_box.text.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF legacy text-box text size overflow",
                    )
                })?;
            if legacy_text_box_bytes > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy text-box text exceeds the aggregate safety limit",
                ));
            }
            events.push(BodyEvent {
                offset: text_box.position,
                order: 1,
                kind: BodyEventKind::LegacyTextBox(text_box),
            });
            previous_legacy_text_box_position = Some(text_box.position);
        }

        if legacy_drawings.len() > crate::MAX_LEGACY_DRAWINGS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF legacy drawing count exceeds the safety limit",
            ));
        }
        let mut previous_legacy_drawing_position = None;
        for drawing in legacy_drawings {
            drawing
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(drawing.position..drawing.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy drawing position is not a UTF-8 body boundary",
                ));
            }
            if previous_legacy_drawing_position.is_some_and(|position| position > drawing.position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy drawings are not ordered by body position",
                ));
            }
            events.push(BodyEvent {
                offset: drawing.position,
                order: 1,
                kind: BodyEventKind::LegacyDrawing(drawing),
            });
            previous_legacy_drawing_position = Some(drawing.position);
        }

        if navigation_entries.len() > crate::navigation_entry::MAX_NAVIGATION_ENTRIES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF navigation-entry count limit exceeded",
            ));
        }
        let mut navigation_text_bytes = 0usize;
        let body_navigation: Vec<bool> = (0..navigation_entries.len())
            .map(|index| body_story_events.iter().any(|event| {
                matches!(event, crate::BodyStoryEvent::NavigationEntry(value) if *value == index)
            }))
            .collect();
        let body_revisions: Vec<bool> = (0..revisions.len())
            .map(|index| body_story_events.iter().any(|event| {
                matches!(
                    event,
                    crate::BodyStoryEvent::RevisionStart(value)
                        | crate::BodyStoryEvent::RevisionEnd(value)
                        | crate::BodyStoryEvent::RevisionDeletion(value)
                        if *value == index
                )
            }))
            .collect();
        for (entry_index, entry) in navigation_entries.iter().enumerate() {
            entry
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if body_navigation[entry_index]
                && body.get(entry.position()..entry.position()).is_none()
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF navigation-entry position is outside body text or splits a character",
                ));
            }
            navigation_text_bytes = navigation_text_bytes
                .checked_add(entry.text_bytes().ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "navigation-entry size overflow",
                    )
                })?)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "navigation-entry size overflow",
                    )
                })?;
            if navigation_text_bytes
                > crate::navigation_entry::MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF navigation-entry aggregate text limit exceeded",
                ));
            }
            if body_navigation[entry_index] {
                events.push(BodyEvent {
                    offset: entry.position(),
                    order: 1,
                    kind: BodyEventKind::NavigationEntry(entry),
                });
            }
        }
        for bookmark in bookmarks.bookmarks() {
            let end = bookmark
                .position
                .checked_add(bookmark.content.len())
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF bookmark range overflow")
                })?;
            let content = body.get(bookmark.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF bookmark range is outside body text or splits a character",
                )
            })?;
            if content != bookmark.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF bookmark content does not match its body range",
                ));
            }
            let empty = bookmark.content.is_empty();
            events.push(BodyEvent {
                offset: bookmark.position,
                order: 1,
                kind: BodyEventKind::BookmarkStart(bookmark),
            });
            events.push(BodyEvent {
                offset: end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::BookmarkEnd(bookmark),
            });
        }
        if custom_xml_tags.len() > crate::custom_xml::MAX_CUSTOM_XML_TAGS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF custom XML tag count exceeds the safety limit",
            ));
        }
        for tag in custom_xml_tags {
            tag.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let end = tag.position.checked_add(tag.content.len()).ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "RTF custom XML tag range overflow")
            })?;
            let content = body.get(tag.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF custom XML tag range is outside body text or splits a character",
                )
            })?;
            if content != tag.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF custom XML tag content does not match its body range",
                ));
            }
        }
        {
            let mut xml_stack: Vec<usize> = Vec::new();
            for event in body_story_events {
                match *event {
                    crate::BodyStoryEvent::CustomXmlOpen(index) => {
                        if index >= custom_xml_tags.len() {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF custom XML story event references a missing tag",
                            ));
                        }
                        xml_stack.push(index);
                    },
                    crate::BodyStoryEvent::CustomXmlClose(index) => {
                        let expected = xml_stack.pop();
                        if expected != Some(index) {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF custom XML tags are not properly nested",
                            ));
                        }
                    },
                    _ => {},
                }
            }
            if !xml_stack.is_empty() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF custom XML tags are not properly nested",
                ));
            }
        }
        if math_zones.len() > crate::math::MAX_MATH_ZONES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF math zone count exceeds the safety limit",
            ));
        }
        for zone in math_zones {
            zone.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(zone.position..zone.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF math zone anchor is outside body text or splits a character",
                ));
            }
        }
        if protection_ranges.len() > crate::protection_range::MAX_PROTECTION_RANGES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF protection-range count exceeds the safety limit",
            ));
        }
        for range in protection_ranges {
            range
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let end = range.position.checked_add(range.content.len()).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF protection-range range overflow",
                )
            })?;
            let content = body.get(range.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF protection range is outside body text or splits a character",
                )
            })?;
            if content != range.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF protection-range content does not match its body range",
                ));
            }
        }
        if editable_regions.len() > crate::editable_region::MAX_EDITABLE_REGIONS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF editable-region count exceeds the safety limit",
            ));
        }
        for region in editable_regions {
            region
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let end = region.position.checked_add(region.content.len()).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF editable-region range overflow",
                )
            })?;
            let content = body.get(region.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF editable region is outside body text or splits a character",
                )
            })?;
            if content != region.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF editable-region content does not match its body range",
                ));
            }
        }
        {
            let mut region_stack: Vec<usize> = Vec::new();
            for event in body_story_events {
                match *event {
                    crate::BodyStoryEvent::EditableRegionStart(index) => {
                        if index >= editable_regions.len() {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF editable-region story event references a missing region",
                            ));
                        }
                        region_stack.push(index);
                    },
                    crate::BodyStoryEvent::EditableRegionEnd(index) => {
                        let expected = region_stack.pop();
                        if expected != Some(index) {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF editable regions are not properly nested",
                            ));
                        }
                    },
                    _ => {},
                }
            }
            if !region_stack.is_empty() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF editable regions are not properly nested",
                ));
            }
        }
        for annotation in annotations {
            annotation
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if annotation.range_end < annotation.position
                || body
                    .get(annotation.position..annotation.range_end)
                    .is_none()
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF annotation range is outside body text or splits a character",
                ));
            }
            let empty = annotation.position == annotation.range_end;
            events.push(BodyEvent {
                offset: annotation.position,
                order: 1,
                kind: BodyEventKind::AnnotationStart(annotation),
            });
            events.push(BodyEvent {
                offset: annotation.range_end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::AnnotationEnd(annotation),
            });
        }
        let mut revision_ranges: Vec<(usize, &Revision<'_>)> = revisions
            .iter()
            .enumerate()
            .filter(|(index, revision)| body_revisions[*index] && revision.revision_type == RevisionType::Insertion)
            .collect();
        revision_ranges.sort_by_key(|(_, revision)| (revision.position, revision.range_end));
        let mut previous_end = 0usize;
        for (_, revision) in revision_ranges {
            if revision.range_end <= revision.position
                || revision.position < previous_end
                || body.get(revision.position..revision.range_end).is_none()
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision ranges overlap, leave the body, or split a character",
                ));
            }
            let content = &body[revision.position..revision.range_end];
            if content != revision.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision content does not match its body range",
                ));
            }
            previous_end = revision.range_end;
            events.push(BodyEvent {
                offset: revision.position,
                order: 1,
                kind: BodyEventKind::RevisionStart(revision),
            });
            events.push(BodyEvent {
                offset: revision.range_end,
                order: 0,
                kind: BodyEventKind::RevisionEnd,
            });
        }
        for revision in revisions
            .iter()
            .enumerate()
            .filter(|(index, revision)| body_revisions[*index] && revision.revision_type == RevisionType::Deletion)
            .map(|(_, revision)| revision)
        {
            revision
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(..revision.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF deletion position is outside body text or splits a character",
                ));
            }
            events.push(BodyEvent {
                offset: revision.position,
                order: 0,
                kind: BodyEventKind::RevisionDeletion(revision),
            });
        }
        events.clear();
        let bookmark_items = bookmarks.bookmarks();
        let mut saw_shapes = vec![false; shapes.len()];
        let mut saw_groups = vec![false; shape_groups.len()];
        let mut saw_fields = vec![false; fields.len()];
        let mut saw_bookmark_starts = vec![false; bookmark_items.len()];
        let mut saw_bookmark_ends = vec![false; bookmark_items.len()];
        let mut saw_custom_xml_opens = vec![false; custom_xml_tags.len()];
        let mut saw_custom_xml_closes = vec![false; custom_xml_tags.len()];
        let mut saw_math_zones = vec![false; math_zones.len()];
        let mut saw_protection_starts = vec![false; protection_ranges.len()];
        let mut saw_protection_ends = vec![false; protection_ranges.len()];
        let mut saw_editable_starts = vec![false; editable_regions.len()];
        let mut saw_editable_ends = vec![false; editable_regions.len()];
        let mut saw_annotation_starts = vec![false; annotations.len()];
        let mut saw_annotation_ends = vec![false; annotations.len()];
        let mut saw_notes = vec![false; notes.len()];
        let mut saw_objects = vec![false; objects.len()];
        let mut saw_picture_records = vec![false; picture_compatibility_records.len()];
        let mut saw_form_starts = vec![false; form_fields.len()];
        let mut saw_form_ends = vec![false; form_fields.len()];
        let mut saw_revision_starts = vec![false; revisions.len()];
        let mut saw_revision_ends = vec![false; revisions.len()];
        let mut saw_revision_deletions = vec![false; revisions.len()];
        let mut saw_generated_markers = vec![false; generated_list_markers.len()];
        let mut saw_legacy_text_boxes = vec![false; legacy_text_boxes.len()];
        let mut saw_legacy_drawings = vec![false; legacy_drawings.len()];
        let mut saw_navigation_entries = vec![false; navigation_entries.len()];
        let mut ordered_drawings = Vec::with_capacity(expected_drawings);
        let mut previous_story_position = None;
        let first_section_is_boundary_scoped = body_story_events.iter().any(|event| {
            matches!(
                event,
                crate::BodyStoryEvent::SectionBreak(section_break)
                    if section_break.next_section == Some(0)
            )
        });
        let mut next_section_index = if first_section_is_boundary_scoped {
            0
        } else {
            usize::from(!sections.is_empty())
        };
        for story_event in body_story_events {
            let (position, kind) = match *story_event {
                crate::BodyStoryEvent::Drawing(crate::StoryDrawing::Shape(index))
                    if index < shapes.len()
                        && !shapes[index].is_background
                        && !saw_shapes[index] =>
                {
                    saw_shapes[index] = true;
                    ordered_drawings.push(crate::StoryDrawing::Shape(index));
                    (shapes[index].position, BodyEventKind::Shape(&shapes[index]))
                },
                crate::BodyStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index))
                    if index < shape_groups.len() && !saw_groups[index] =>
                {
                    saw_groups[index] = true;
                    ordered_drawings.push(crate::StoryDrawing::ShapeGroup(index));
                    (
                        shape_groups[index].position,
                        BodyEventKind::ShapeGroup(&shape_groups[index]),
                    )
                },
                crate::BodyStoryEvent::Field(index)
                    if index < fields.len()
                        && !saw_fields[index]
                        && matches!(fields[index].owner, crate::FieldOwner::Body) =>
                {
                    saw_fields[index] = true;
                    (
                        fields[index].position,
                        BodyEventKind::GenericField(&fields[index]),
                    )
                },
                crate::BodyStoryEvent::PageBreak(page_break) => {
                    (page_break.position, BodyEventKind::PageBreak)
                },
                crate::BodyStoryEvent::SoftBreak(soft_break) => {
                    (soft_break.position, BodyEventKind::SoftBreak(soft_break))
                },
                crate::BodyStoryEvent::ColumnBreak(column_break) => {
                    (column_break.position, BodyEventKind::ColumnBreak)
                },
                crate::BodyStoryEvent::SectionBreak(section_break) => {
                    let section = match section_break.next_section {
                        None => None,
                        Some(index)
                            if index == next_section_index && index < sections.len() =>
                        {
                            next_section_index += 1;
                            Some(&sections[index])
                        },
                        Some(_) => {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF section boundary has an invalid or out-of-order section reference",
                            ));
                        },
                    };
                    (section_break.position, BodyEventKind::SectionBreak(section))
                },
                crate::BodyStoryEvent::BookmarkStart(index)
                    if index < bookmark_items.len() && !saw_bookmark_starts[index] =>
                {
                    saw_bookmark_starts[index] = true;
                    (
                        bookmark_items[index].position,
                        BodyEventKind::BookmarkStart(&bookmark_items[index]),
                    )
                },
                crate::BodyStoryEvent::BookmarkEnd(index)
                    if index < bookmark_items.len() && !saw_bookmark_ends[index] =>
                {
                    saw_bookmark_ends[index] = true;
                    let bookmark = &bookmark_items[index];
                    (
                        bookmark
                            .position
                            .checked_add(bookmark.content.len())
                            .ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF bookmark range overflow",
                                )
                            })?,
                        BodyEventKind::BookmarkEnd(bookmark),
                    )
                },
                crate::BodyStoryEvent::CustomXmlOpen(index)
                    if index < custom_xml_tags.len() && !saw_custom_xml_opens[index] =>
                {
                    saw_custom_xml_opens[index] = true;
                    (
                        custom_xml_tags[index].position,
                        BodyEventKind::CustomXmlOpen(&custom_xml_tags[index]),
                    )
                },
                crate::BodyStoryEvent::CustomXmlClose(index)
                    if index < custom_xml_tags.len() && !saw_custom_xml_closes[index] =>
                {
                    saw_custom_xml_closes[index] = true;
                    let tag = &custom_xml_tags[index];
                    (
                        tag.position.checked_add(tag.content.len()).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF custom XML tag range overflow",
                            )
                        })?,
                        BodyEventKind::CustomXmlClose(tag),
                    )
                },
                crate::BodyStoryEvent::MathZone(index)
                    if index < math_zones.len() && !saw_math_zones[index] =>
                {
                    saw_math_zones[index] = true;
                    (
                        math_zones[index].position,
                        BodyEventKind::MathZone(&math_zones[index]),
                    )
                },
                crate::BodyStoryEvent::ProtectionRangeStart(index)
                    if index < protection_ranges.len() && !saw_protection_starts[index] =>
                {
                    saw_protection_starts[index] = true;
                    (
                        protection_ranges[index].position,
                        BodyEventKind::ProtectionRangeStart(&protection_ranges[index]),
                    )
                },
                crate::BodyStoryEvent::ProtectionRangeEnd(index)
                    if index < protection_ranges.len() && !saw_protection_ends[index] =>
                {
                    saw_protection_ends[index] = true;
                    let range = &protection_ranges[index];
                    (
                        range.position.checked_add(range.content.len()).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF protection-range range overflow",
                            )
                        })?,
                        BodyEventKind::ProtectionRangeEnd(range),
                    )
                },
                crate::BodyStoryEvent::EditableRegionStart(index)
                    if index < editable_regions.len() && !saw_editable_starts[index] =>
                {
                    saw_editable_starts[index] = true;
                    (
                        editable_regions[index].position,
                        BodyEventKind::EditableRegionStart(&editable_regions[index]),
                    )
                },
                crate::BodyStoryEvent::EditableRegionEnd(index)
                    if index < editable_regions.len() && !saw_editable_ends[index] =>
                {
                    saw_editable_ends[index] = true;
                    let region = &editable_regions[index];
                    (
                        region.position.checked_add(region.content.len()).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF editable-region range overflow",
                            )
                        })?,
                        BodyEventKind::EditableRegionEnd(region),
                    )
                },
                crate::BodyStoryEvent::AnnotationStart(index)
                    if index < annotations.len() && !saw_annotation_starts[index] =>
                {
                    saw_annotation_starts[index] = true;
                    (
                        annotations[index].position,
                        BodyEventKind::AnnotationStart(&annotations[index]),
                    )
                },
                crate::BodyStoryEvent::AnnotationEnd(index)
                    if index < annotations.len() && !saw_annotation_ends[index] =>
                {
                    saw_annotation_ends[index] = true;
                    (
                        annotations[index].range_end,
                        BodyEventKind::AnnotationEnd(&annotations[index]),
                    )
                },
                crate::BodyStoryEvent::Note(index) if index < notes.len() && !saw_notes[index] => {
                    saw_notes[index] = true;
                    (notes[index].position, BodyEventKind::Note(&notes[index]))
                },
                crate::BodyStoryEvent::Object(index)
                    if index < objects.len() && !saw_objects[index] =>
                {
                    saw_objects[index] = true;
                    (
                        objects[index].position,
                        BodyEventKind::Object(&objects[index], pictures),
                    )
                },
                crate::BodyStoryEvent::PictureCompatibility(index)
                    if index < picture_compatibility_records.len()
                        && !saw_picture_records[index] =>
                {
                    saw_picture_records[index] = true;
                    let record = &picture_compatibility_records[index];
                    (
                        record.position,
                        BodyEventKind::PictureCompatibility(
                            record,
                            &pictures[record.picture_index],
                        ),
                    )
                },
                crate::BodyStoryEvent::FormFieldStart(index)
                    if index < form_fields.len() && !saw_form_starts[index] =>
                {
                    saw_form_starts[index] = true;
                    (
                        form_fields[index].position,
                        BodyEventKind::FormFieldStart(&form_fields[index]),
                    )
                },
                crate::BodyStoryEvent::FormFieldEnd(index)
                    if index < form_fields.len() && !saw_form_ends[index] =>
                {
                    saw_form_ends[index] = true;
                    (form_fields[index].range_end, BodyEventKind::FormFieldEnd)
                },
                crate::BodyStoryEvent::RevisionStart(index)
                    if index < revisions.len()
                        && !saw_revision_starts[index]
                        && revisions[index].revision_type == RevisionType::Insertion =>
                {
                    saw_revision_starts[index] = true;
                    (
                        revisions[index].position,
                        BodyEventKind::RevisionStart(&revisions[index]),
                    )
                },
                crate::BodyStoryEvent::RevisionEnd(index)
                    if index < revisions.len()
                        && !saw_revision_ends[index]
                        && revisions[index].revision_type == RevisionType::Insertion =>
                {
                    saw_revision_ends[index] = true;
                    (revisions[index].range_end, BodyEventKind::RevisionEnd)
                },
                crate::BodyStoryEvent::RevisionDeletion(index)
                    if index < revisions.len()
                        && !saw_revision_deletions[index]
                        && revisions[index].revision_type == RevisionType::Deletion =>
                {
                    saw_revision_deletions[index] = true;
                    (
                        revisions[index].position,
                        BodyEventKind::RevisionDeletion(&revisions[index]),
                    )
                },
                crate::BodyStoryEvent::GeneratedListMarker(index)
                    if index < generated_list_markers.len() && !saw_generated_markers[index] =>
                {
                    saw_generated_markers[index] = true;
                    (
                        generated_list_markers[index].position,
                        BodyEventKind::GeneratedListMarker(&generated_list_markers[index]),
                    )
                },
                crate::BodyStoryEvent::LegacyTextBox(index)
                    if index < legacy_text_boxes.len() && !saw_legacy_text_boxes[index] =>
                {
                    saw_legacy_text_boxes[index] = true;
                    (
                        legacy_text_boxes[index].position,
                        BodyEventKind::LegacyTextBox(&legacy_text_boxes[index]),
                    )
                },
                crate::BodyStoryEvent::LegacyDrawing(index)
                    if index < legacy_drawings.len() && !saw_legacy_drawings[index] =>
                {
                    saw_legacy_drawings[index] = true;
                    (
                        legacy_drawings[index].position,
                        BodyEventKind::LegacyDrawing(&legacy_drawings[index]),
                    )
                },
                crate::BodyStoryEvent::NavigationEntry(index)
                    if index < navigation_entries.len() && !saw_navigation_entries[index] =>
                {
                    saw_navigation_entries[index] = true;
                    (
                        navigation_entries[index].position(),
                        BodyEventKind::NavigationEntry(&navigation_entries[index]),
                    )
                },
                _ => {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF body story order has an invalid or duplicate reference",
                    ));
                },
            };
            if body.get(position..position).is_none()
                || previous_story_position.is_some_and(|previous| previous > position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF body story events are outside or out of story order",
                ));
            }
            events.push(BodyEvent {
                offset: position,
                order: 1,
                kind,
            });
            previous_story_position = Some(position);
        }
        let complete = ordered_drawings == drawing_order
            && saw_shapes
                .iter()
                .enumerate()
                .all(|(index, seen)| shapes[index].is_background || *seen)
            && saw_groups.iter().all(|seen| *seen)
            && saw_fields.iter().enumerate().all(|(index, seen)| {
                !matches!(fields[index].owner, crate::FieldOwner::Body) || *seen
            })
            && saw_bookmark_starts.iter().all(|seen| *seen)
            && saw_bookmark_ends.iter().all(|seen| *seen)
            && saw_custom_xml_opens.iter().all(|seen| *seen)
            && saw_custom_xml_closes.iter().all(|seen| *seen)
            && saw_math_zones.iter().all(|seen| *seen)
            && saw_protection_starts.iter().all(|seen| *seen)
            && saw_protection_ends.iter().all(|seen| *seen)
            && saw_editable_starts.iter().all(|seen| *seen)
            && saw_editable_ends.iter().all(|seen| *seen)
            && saw_annotation_starts.iter().all(|seen| *seen)
            && saw_annotation_ends.iter().all(|seen| *seen)
            && saw_notes.iter().all(|seen| *seen)
            && saw_objects.iter().all(|seen| *seen)
            && saw_picture_records.iter().all(|seen| *seen)
            && saw_form_starts.iter().all(|seen| *seen)
            && saw_form_ends.iter().all(|seen| *seen)
            && revisions
                .iter()
                .enumerate()
                .all(|(index, revision)| match revision.revision_type {
                    _ if !body_revisions[index] => true,
                    RevisionType::Insertion => {
                        saw_revision_starts[index]
                            && saw_revision_ends[index]
                            && !saw_revision_deletions[index]
                    },
                    RevisionType::Deletion => {
                        saw_revision_deletions[index]
                            && !saw_revision_starts[index]
                            && !saw_revision_ends[index]
                    },
                    _ => false,
                })
            && saw_generated_markers.iter().all(|seen| *seen)
            && saw_legacy_text_boxes.iter().all(|seen| *seen)
            && saw_legacy_drawings.iter().all(|seen| *seen)
            && next_section_index == sections.len()
            && saw_navigation_entries
                .iter()
                .enumerate()
                .all(|(index, seen)| !body_navigation[index] || *seen);
        if !complete {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF body story order is incomplete or changes drawing order",
            ));
        }
        events.sort_by_key(|event| (event.offset, event.order));

        let mut event_index = 0usize;
        let mut body_offset = 0usize;
        for block in blocks {
            let block_end = body_offset + block.text.len();
            let mut local_offset = 0usize;
            while event_index < events.len() && events[event_index].offset <= block_end {
                let event_offset = events[event_index].offset;
                if event_offset < body_offset {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF bookmark events are not ordered",
                    ));
                }
                let local_end = event_offset - body_offset;
                if local_end > local_offset {
                    self.write_style_block_fragment(block, local_offset, local_end)?;
                    local_offset = local_end;
                }
                while event_index < events.len() && events[event_index].offset == event_offset {
                    self.write_body_event(events[event_index], fields)?;
                    event_index += 1;
                }
            }
            if local_offset < block.text.len() {
                self.write_style_block_fragment(block, local_offset, block.text.len())?;
            }
            body_offset = block_end;
        }
        while event_index < events.len() && events[event_index].offset == body_offset {
            self.write_body_event(events[event_index], fields)?;
            event_index += 1;
        }
        if event_index != events.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark range extends beyond body text",
            ));
        }
        Ok(())
    }

    fn write_body_event(
        &mut self,
        event: BodyEvent<'_, '_>,
        fields: &[crate::Field<'_>],
    ) -> io::Result<()> {
        match event.kind {
            BodyEventKind::Shape(shape) => self.write_root_shape(shape),
            BodyEventKind::ShapeGroup(group) => self.write_shape_group(group, true),
            BodyEventKind::Object(object, pictures) => self.write_object(object, pictures),
            BodyEventKind::PictureCompatibility(record, picture) => {
                self.write_picture_compatibility(record.kind, picture)
            },
            BodyEventKind::GeneratedListMarker(marker) => self.write_generated_list_marker(marker),
            BodyEventKind::LegacyTextBox(text_box) => self.write_legacy_text_box(text_box),
            BodyEventKind::LegacyDrawing(drawing) => self.write_legacy_drawing(drawing),
            BodyEventKind::NavigationEntry(entry) => self.write_navigation_entry(entry),
            BodyEventKind::BookmarkStart(bookmark) => self.write_bookmark_start(bookmark),
            BodyEventKind::BookmarkEnd(bookmark) => self.write_bookmark_end(bookmark.name.as_ref()),
            BodyEventKind::CustomXmlOpen(tag) => self.write_custom_xml_open(tag),
            BodyEventKind::CustomXmlClose(tag) => self.write_custom_xml_close(tag),
            BodyEventKind::MathZone(zone) => self.write_math_zone(zone),
            BodyEventKind::ProtectionRangeStart(range) => {
                self.write_protection_range_marker("protstart", range)
            },
            BodyEventKind::ProtectionRangeEnd(range) => {
                self.write_protection_range_marker("protend", range)
            },
            BodyEventKind::EditableRegionStart(region) => {
                region
                    .validate()
                    .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
                self.write_str("\\ebcstart ")
            },
            BodyEventKind::EditableRegionEnd(region) => {
                region
                    .validate()
                    .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
                self.write_str("\\ebcend ")
            },
            BodyEventKind::AnnotationStart(annotation) => self.write_annotation_start(annotation),
            BodyEventKind::AnnotationEnd(annotation) => self.write_annotation_end(annotation),
            BodyEventKind::Note(note) => self.write_note_with_fields(note, fields),
            BodyEventKind::RevisionStart(revision) => self.write_revision_start(revision),
            BodyEventKind::RevisionEnd => self.write_str("}"),
            BodyEventKind::RevisionDeletion(revision) => self.write_revision(revision),
            BodyEventKind::FormFieldStart(field) => self.write_form_field_start(field),
            BodyEventKind::FormFieldEnd => self.write_str("}}"),
            BodyEventKind::GenericField(field) => self.write_field_with_fields(field, fields, 0),
            BodyEventKind::PageBreak => self.write_str("\\page "),
            BodyEventKind::SoftBreak(soft_break) => match soft_break.kind {
                crate::SoftBreakKind::Page => self.write_str("\\softpage "),
                crate::SoftBreakKind::Column => self.write_str("\\softcol "),
                crate::SoftBreakKind::Line => self.write_str("\\softline "),
                crate::SoftBreakKind::LineHeight(height) => {
                    self.write_control_word("softlheight", Some(height))?;
                    self.write_str(" ")
                },
            },
            BodyEventKind::ColumnBreak => self.write_str("\\column "),
            BodyEventKind::SectionBreak(section) => {
                self.write_control_word("sect", None)?;
                if let Some(section) = section {
                    self.write_section_with_fields(section, fields)?;
                }
                Ok(())
            },
        }
    }

    fn write_root_shape(&mut self, shape: &crate::Shape<'_>) -> io::Result<()> {
        shape
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if !shape.instruction_present {
            self.write_str("{\\shp")?;
            self.write_shape_result(
                shape
                    .result
                    .as_ref()
                    .expect("validated fallback-only shape result"),
            )?;
            return self.write_str("}");
        }
        let right = shape
            .geometry
            .x
            .checked_add(shape.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape right edge overflows",
                )
            })?;
        let bottom = shape
            .geometry
            .y
            .checked_add(shape.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape bottom edge overflows",
                )
            })?;
        self.write_str("{\\shp{\\*\\shpinst")?;
        self.write_control_word("shpleft", Some(shape.geometry.x))?;
        self.write_control_word("shptop", Some(shape.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(shape.geometry.z_order))?;
        self.write_shape_info(&shape.info)?;
        if !shape
            .properties
            .iter()
            .any(|property| property.name == "shapeType")
        {
            let shape_type = match shape.shape_type {
                crate::ShapeType::Rectangle => Some(1),
                crate::ShapeType::RoundRectangle => Some(2),
                crate::ShapeType::Ellipse => Some(3),
                crate::ShapeType::Arc => Some(19),
                crate::ShapeType::Line => Some(20),
                crate::ShapeType::PictureFrame => Some(75),
                crate::ShapeType::TextBox => Some(202),
                crate::ShapeType::Group => Some(0),
                crate::ShapeType::Custom(value) => Some(value),
                crate::ShapeType::Polygon | crate::ShapeType::Unknown => None,
            };
            if let Some(value) = shape_type {
                self.write_shape_scalar_property("shapeType", &value.to_string())?;
            }
        }
        for property in &shape.properties {
            self.write_shape_property(property)?;
        }
        if shape.text_destination_present
            || !shape.text.is_empty()
            || !shape.text_shapes.is_empty()
            || !shape.text_shape_groups.is_empty()
            || !shape.text_story_events.is_empty()
        {
            self.write_shape_text(shape)?;
        }
        self.write_str("}")?;
        if let Some(result) = &shape.result {
            self.write_shape_result(result)?;
        }
        self.write_str("}")
    }

    fn write_shape_info(&mut self, info: &[crate::ShapeGroupInfo]) -> io::Result<()> {
        for info in info {
            match *info {
                crate::ShapeGroupInfo::ShapeId(value) => {
                    self.write_control_word("shplid", Some(value))?
                },
                crate::ShapeGroupInfo::InHeader(value) => {
                    self.write_control_word("shpfhdr", Some(i32::from(value)))?
                },
                crate::ShapeGroupInfo::HorizontalPage => {
                    self.write_control_word("shpbxpage", None)?
                },
                crate::ShapeGroupInfo::HorizontalMargin => {
                    self.write_control_word("shpbxmargin", None)?
                },
                crate::ShapeGroupInfo::HorizontalColumn => {
                    self.write_control_word("shpbxcolumn", None)?
                },
                crate::ShapeGroupInfo::IgnoreHorizontal => {
                    self.write_control_word("shpbxignore", None)?
                },
                crate::ShapeGroupInfo::VerticalPage => {
                    self.write_control_word("shpbypage", None)?
                },
                crate::ShapeGroupInfo::VerticalMargin => {
                    self.write_control_word("shpbymargin", None)?
                },
                crate::ShapeGroupInfo::VerticalParagraph => {
                    self.write_control_word("shpbypara", None)?
                },
                crate::ShapeGroupInfo::IgnoreVertical => {
                    self.write_control_word("shpbyignore", None)?
                },
                crate::ShapeGroupInfo::Wrap(value) => {
                    self.write_control_word("shpwr", Some(value))?
                },
                crate::ShapeGroupInfo::WrapSide(value) => {
                    self.write_control_word("shpwrk", Some(value))?
                },
                crate::ShapeGroupInfo::BelowText(value) => {
                    self.write_control_word("shpfblwtxt", Some(i32::from(value)))?
                },
                crate::ShapeGroupInfo::LockAnchor => {
                    self.write_control_word("shplockanchor", None)?
                },
            }
        }
        Ok(())
    }

    fn write_object(
        &mut self,
        object: &crate::EmbeddedObject<'_>,
        pictures: &[crate::Picture<'_>],
    ) -> io::Result<()> {
        self.write_str("{\\object")?;
        self.write_str(match object.kind {
            crate::ObjectKind::Embedded => "\\objemb",
            crate::ObjectKind::Link => "\\objlink",
            crate::ObjectKind::AutoLink => "\\objautlink",
            crate::ObjectKind::Html => "\\objhtml",
            crate::ObjectKind::Subscriber => "\\objsub",
            crate::ObjectKind::Publisher => "\\objpub",
            crate::ObjectKind::InstallableCommand => "\\objicemb",
            crate::ObjectKind::OleControl => "\\objocx",
            crate::ObjectKind::Unknown => "",
        })?;
        if object.link_self {
            self.write_str("\\linkself")?;
        }
        if object.locked {
            self.write_str("\\objlock")?;
        }
        if object.update_requested {
            self.write_str("\\objupdate")?;
        }
        if !object.class_name.is_empty() {
            self.write_str("{\\*\\objclass ")?;
            self.write_destination_text(object.class_name.as_ref())?;
            self.write_str("}")?;
        }
        if !object.name.is_empty() {
            self.write_str("{\\*\\objname ")?;
            self.write_destination_text(object.name.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(alias) = &object.alias {
            self.write_str("{\\*\\objalias ")?;
            self.write_destination_text(alias.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(section) = &object.section {
            self.write_str("{\\*\\objsect ")?;
            self.write_destination_text(section.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(time) = object.time {
            self.write_str("{\\*\\objtime ")?;
            for (name, value) in [
                ("yr", time.year),
                ("mo", time.month),
                ("dy", time.day),
                ("hr", time.hour),
                ("min", time.minute),
                ("sec", time.second),
            ] {
                if let Some(value) = value {
                    self.write_control_word(name, Some(value))?;
                }
            }
            self.write_str("}")?;
        }
        if object.set_size {
            self.write_str("\\objsetsize")?;
        }
        self.write_optional_object_value("objalign", object.alignment)?;
        self.write_optional_object_value("objtransy", object.translation_y)?;
        if object.height != 0 {
            self.write_control_word("objh", Some(object.height))?;
        }
        if object.width != 0 {
            self.write_control_word("objw", Some(object.width))?;
        }
        self.write_optional_object_value("objcropt", object.crop_top)?;
        self.write_optional_object_value("objcropb", object.crop_bottom)?;
        self.write_optional_object_value("objcropl", object.crop_left)?;
        self.write_optional_object_value("objcropr", object.crop_right)?;
        self.write_optional_object_value("objscalex", object.scale_x)?;
        self.write_optional_object_value("objscaley", object.scale_y)?;
        if object.merge_result {
            self.write_str("\\rsltmerge")?;
        }
        if let Some(kind) = object.result_kind {
            self.write_str(match kind {
                crate::ObjectResultKind::Rtf => "\\rsltrtf",
                crate::ObjectResultKind::Text => "\\rslttxt",
                crate::ObjectResultKind::Picture => "\\rsltpict",
                crate::ObjectResultKind::Bitmap => "\\rsltbmp",
                crate::ObjectResultKind::Html => "\\rslthtml",
            })?;
        }
        if !object.class_id.is_empty() {
            self.write_str("{\\*\\oleclsid ")?;
            self.write_destination_text(object.class_id.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("{\\*\\objdata ")?;
        for byte in &object.data {
            write!(self.writer, "{byte:02x}")?;
        }
        self.write_str("}")?;
        if !object.result_text.is_empty() || !object.result_picture_indices.is_empty() {
            self.write_str("{\\result ")?;
            self.write_destination_text(object.result_text.as_ref())?;
            for index in &object.result_picture_indices {
                self.write_picture(&pictures[*index])?;
            }
            self.write_str("\\par}")?;
        }
        self.write_str("}")
    }

    fn write_optional_object_value(&mut self, control: &str, value: Option<i32>) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(value))?;
        }
        Ok(())
    }

    fn write_picture_compatibility(
        &mut self,
        kind: crate::PictureCompatibilityKind,
        picture: &crate::Picture<'_>,
    ) -> io::Result<()> {
        self.write_str(match kind {
            crate::PictureCompatibilityKind::ShapePicture => "{\\*\\shppict",
            crate::PictureCompatibilityKind::NonShapePicture => "{\\nonshppict",
        })?;
        self.write_picture(picture)?;
        self.write_str("}")
    }

    fn write_picture_shape_properties(
        &mut self,
        properties: &crate::PictureShapeProperties<'_>,
    ) -> io::Result<()> {
        properties
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\picprop")?;
        if let Some(shape_id) = properties.shape_id {
            self.write_control_word("shplid", Some(shape_id))?;
        }
        for property in &properties.properties {
            self.write_shape_property(property)?;
        }
        self.write_str("}")
    }

    /// Write one inert legacy drawing text box.
    pub fn write_legacy_text_box(&mut self, text_box: &crate::LegacyTextBox<'_>) -> io::Result<()> {
        text_box
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        self.write_str("{\\*\\do")?;
        if let Some(anchor) = text_box.horizontal_anchor {
            self.write_control_word(
                match anchor {
                    crate::LegacyHorizontalAnchor::Page => "dobxpage",
                    crate::LegacyHorizontalAnchor::Margin => "dobxmargin",
                    crate::LegacyHorizontalAnchor::Column => "dobxcolumn",
                },
                None,
            )?;
        }
        if let Some(anchor) = text_box.vertical_anchor {
            self.write_control_word(
                match anchor {
                    crate::LegacyVerticalAnchor::Page => "dobypage",
                    crate::LegacyVerticalAnchor::Margin => "dobymargin",
                    crate::LegacyVerticalAnchor::Paragraph => "dobypara",
                },
                None,
            )?;
        }
        if let Some(value) = text_box.z_order {
            self.write_control_word("dodhgt", Some(value))?;
        }
        self.write_control_word("dptxbx", None)?;
        if let Some(value) = text_box.margin {
            self.write_control_word("dptxbxmar", Some(value))?;
        }
        self.write_control_word(
            match text_box.direction {
                crate::LegacyTextDirection::LeftToRightTopToBottom => "dptxlrtb",
                crate::LegacyTextDirection::LeftToRightTopToBottomVertical => "dptxlrtbv",
                crate::LegacyTextDirection::TopToBottomRightToLeft => "dptxtbrl",
                crate::LegacyTextDirection::TopToBottomRightToLeftVertical => "dptxtbrlv",
                crate::LegacyTextDirection::BottomToTopLeftToRight => "dptxbtlr",
            },
            None,
        )?;
        if let Some(value) = text_box.x {
            self.write_control_word("dpx", Some(value))?;
        }
        if let Some(value) = text_box.y {
            self.write_control_word("dpy", Some(value))?;
        }
        if let Some(value) = text_box.width {
            self.write_control_word("dpxsize", Some(value))?;
        }
        if let Some(value) = text_box.height {
            self.write_control_word("dpysize", Some(value))?;
        }
        self.write_str("{\\dptxbxtext ")?;
        self.write_field_story(
            text_box.text.as_ref(),
            &text_box.shapes,
            &text_box.shape_groups,
            &text_box.drawing_order,
            &text_box.story_events,
            &[],
            crate::FieldOwner::Other,
            DrawingStoryTextMode::ShapeText,
            0,
        )?;
        self.write_str("}}")
    }

    /// Write one inert Word 6/95 drawing destination canonically.
    pub fn write_legacy_drawing(&mut self, drawing: &crate::LegacyDrawing<'_>) -> io::Result<()> {
        drawing
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\do")?;
        self.write_control_word(
            match drawing.horizontal_anchor {
                crate::LegacyHorizontalAnchor::Page => "dobxpage",
                crate::LegacyHorizontalAnchor::Margin => "dobxmargin",
                crate::LegacyHorizontalAnchor::Column => "dobxcolumn",
            },
            None,
        )?;
        self.write_control_word(
            match drawing.vertical_anchor {
                crate::LegacyVerticalAnchor::Page => "dobypage",
                crate::LegacyVerticalAnchor::Margin => "dobymargin",
                crate::LegacyVerticalAnchor::Paragraph => "dobypara",
            },
            None,
        )?;
        self.write_control_word("dodhgt", Some(drawing.z_order))?;
        if drawing.locked {
            self.write_control_word("dolock", None)?;
        }
        self.write_legacy_drawing_primitive(&drawing.primitive)?;
        self.write_str("}")
    }

    fn write_legacy_geometry(&mut self, geometry: crate::LegacyDrawingGeometry) -> io::Result<()> {
        self.write_control_word("dpx", Some(geometry.x))?;
        self.write_control_word("dpy", Some(geometry.y))?;
        self.write_control_word("dpxsize", Some(geometry.width))?;
        self.write_control_word("dpysize", Some(geometry.height))
    }

    fn write_legacy_point(&mut self, point: crate::LegacyDrawingPoint) -> io::Result<()> {
        self.write_control_word("dpptx", Some(point.x))?;
        self.write_control_word("dppty", Some(point.y))
    }

    fn write_legacy_drawing_primitive(
        &mut self,
        primitive: &crate::LegacyDrawingPrimitive<'_>,
    ) -> io::Result<()> {
        match primitive {
            crate::LegacyDrawingPrimitive::Group {
                geometry,
                children,
                end_geometry,
            } => {
                self.write_control_word("dpgroup", None)?;
                let count = i32::try_from(children.len()).map_err(|_| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "legacy drawing child count overflow",
                    )
                })?;
                self.write_control_word("dpcount", Some(count))?;
                self.write_legacy_geometry(*geometry)?;
                for child in children {
                    self.write_legacy_drawing_primitive(child)?;
                }
                self.write_control_word("dpendgroup", None)?;
                self.write_legacy_geometry(*end_geometry)
            },
            crate::LegacyDrawingPrimitive::Callout(callout) => {
                self.write_control_word("dpcallout", None)?;
                self.write_control_word(
                    match callout.callout_type {
                        crate::LegacyCalloutType::RightAngle => "dpcotright",
                        crate::LegacyCalloutType::Single => "dpcotsingle",
                        crate::LegacyCalloutType::Double => "dpcotdouble",
                        crate::LegacyCalloutType::Triple => "dpcottriple",
                    },
                    None,
                )?;
                if let Some(angle) = callout.angle {
                    self.write_control_word("dpcoa", Some(i32::from(angle)))?;
                }
                if callout.accent {
                    self.write_control_word("dpcoaccent", None)?;
                }
                if callout.smart_attach {
                    self.write_control_word("dpcosmarta", None)?;
                }
                if callout.best_fit {
                    self.write_control_word("dpcobestfit", None)?;
                }
                if callout.minus_x {
                    self.write_control_word("dpcominusx", None)?;
                }
                if callout.minus_y {
                    self.write_control_word("dpcominusy", None)?;
                }
                if callout.border {
                    self.write_control_word("dpcoborder", None)?;
                }
                if let Some(attachment) = callout.attachment {
                    self.write_control_word(
                        match attachment {
                            crate::LegacyCalloutAttachment::Top => "dpcodtop",
                            crate::LegacyCalloutAttachment::Center => "dpcodcenter",
                            crate::LegacyCalloutAttachment::Bottom => "dpcodbottom",
                            crate::LegacyCalloutAttachment::Absolute => "dpcodabs",
                        },
                        None,
                    )?;
                }
                if let Some(value) = callout.descent {
                    self.write_control_word("dpcodescent", Some(value))?;
                }
                self.write_control_word("dpcooffset", Some(callout.offset))?;
                self.write_control_word("dpcolength", Some(callout.length))?;
                self.write_legacy_geometry(callout.geometry)?;
                self.write_legacy_drawing_primitive(&callout.polyline)?;
                self.write_legacy_drawing_primitive(&callout.text_box)?;
                self.write_legacy_properties(callout.properties)
            },
            crate::LegacyDrawingPrimitive::Line {
                start,
                end,
                geometry,
                properties,
            } => {
                self.write_control_word("dpline", None)?;
                self.write_legacy_point(*start)?;
                self.write_legacy_point(*end)?;
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            crate::LegacyDrawingPrimitive::Rectangle {
                rounded,
                geometry,
                properties,
            } => {
                self.write_control_word("dprect", None)?;
                if *rounded {
                    self.write_control_word("dproundr", None)?;
                }
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            crate::LegacyDrawingPrimitive::TextBox {
                text_box,
                properties,
            } => {
                self.write_control_word("dptxbx", None)?;
                if let Some(value) = text_box.margin {
                    self.write_control_word("dptxbxmar", Some(value))?;
                }
                if text_box.direction != crate::LegacyTextDirection::LeftToRightTopToBottom {
                    self.write_control_word(
                        match text_box.direction {
                            crate::LegacyTextDirection::LeftToRightTopToBottom => "dptxlrtb",
                            crate::LegacyTextDirection::LeftToRightTopToBottomVertical => {
                                "dptxlrtbv"
                            },
                            crate::LegacyTextDirection::TopToBottomRightToLeft => "dptxtbrl",
                            crate::LegacyTextDirection::TopToBottomRightToLeftVertical => {
                                "dptxtbrlv"
                            },
                            crate::LegacyTextDirection::BottomToTopLeftToRight => "dptxbtlr",
                        },
                        None,
                    )?;
                }
                self.write_legacy_text_box_text(text_box)?;
                self.write_legacy_geometry(crate::LegacyDrawingGeometry {
                    x: text_box.x.unwrap_or(0),
                    y: text_box.y.unwrap_or(0),
                    width: text_box.width.unwrap_or(0),
                    height: text_box.height.unwrap_or(0),
                })?;
                self.write_legacy_properties(*properties)
            },
            crate::LegacyDrawingPrimitive::Ellipse {
                geometry,
                properties,
            } => {
                self.write_control_word("dpellipse", None)?;
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            crate::LegacyDrawingPrimitive::Polyline {
                closed,
                points,
                geometry,
                properties,
            } => {
                self.write_control_word("dppolyline", None)?;
                if *closed {
                    self.write_control_word("dppolygon", None)?;
                }
                let count = i32::try_from(points.len()).map_err(|_| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "legacy drawing point count overflow",
                    )
                })?;
                self.write_control_word("dppolycount", Some(count))?;
                for point in points {
                    self.write_legacy_point(*point)?;
                }
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            crate::LegacyDrawingPrimitive::Arc {
                flip_x,
                flip_y,
                geometry,
                properties,
            } => {
                self.write_control_word("dparc", None)?;
                if *flip_x {
                    self.write_control_word("dparcflipx", None)?;
                }
                if *flip_y {
                    self.write_control_word("dparcflipy", None)?;
                }
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
        }
    }

    fn write_legacy_text_box_text(
        &mut self,
        text_box: &crate::LegacyTextBox<'_>,
    ) -> io::Result<()> {
        self.write_str("{\\dptxbxtext ")?;
        self.write_field_story(
            text_box.text.as_ref(),
            &text_box.shapes,
            &text_box.shape_groups,
            &text_box.drawing_order,
            &text_box.story_events,
            &[],
            crate::FieldOwner::Other,
            DrawingStoryTextMode::ShapeText,
            0,
        )?;
        self.write_str("}")
    }

    fn write_legacy_color(
        &mut self,
        gray: &str,
        red_control: &str,
        green_control: &str,
        blue_control: &str,
        palette_control: &str,
        color: crate::LegacyDrawingColor,
    ) -> io::Result<()> {
        match color {
            crate::LegacyDrawingColor::Gray(value) => {
                self.write_control_word(gray, Some(i32::from(value)))
            },
            crate::LegacyDrawingColor::Rgb {
                red,
                green,
                blue,
                palette,
            } => {
                self.write_control_word(red_control, Some(i32::from(red)))?;
                self.write_control_word(green_control, Some(i32::from(green)))?;
                self.write_control_word(blue_control, Some(i32::from(blue)))?;
                if palette {
                    self.write_control_word(palette_control, None)?;
                }
                Ok(())
            },
        }
    }

    fn write_legacy_arrow(
        &mut self,
        prefix: &str,
        arrow: crate::LegacyDrawingArrow,
    ) -> io::Result<()> {
        self.write_control_word(
            &format!(
                "{prefix}{}",
                match arrow.fill {
                    crate::LegacyDrawingArrowFill::Solid => "sol",
                    crate::LegacyDrawingArrowFill::Hollow => "hol",
                }
            ),
            None,
        )?;
        self.write_control_word(&format!("{prefix}l"), Some(arrow.length as i32))?;
        self.write_control_word(&format!("{prefix}w"), Some(arrow.width as i32))
    }

    fn write_legacy_properties(
        &mut self,
        properties: crate::LegacyDrawingProperties,
    ) -> io::Result<()> {
        if let Some(line) = properties.line {
            self.write_control_word(
                match line.style {
                    crate::LegacyDrawingLineStyle::Solid => "dplinesolid",
                    crate::LegacyDrawingLineStyle::Hollow => "dplinehollow",
                    crate::LegacyDrawingLineStyle::Dashed => "dplinedash",
                    crate::LegacyDrawingLineStyle::Dotted => "dplinedot",
                    crate::LegacyDrawingLineStyle::DashDot => "dplinedado",
                    crate::LegacyDrawingLineStyle::DashDotDot => "dplinedadodo",
                },
                None,
            )?;
            self.write_legacy_color(
                "dplinegray",
                "dplinecor",
                "dplinecog",
                "dplinecob",
                "dplinepal",
                line.color,
            )?;
            self.write_control_word("dplinew", Some(line.width))?;
        }
        if let Some(fill) = properties.fill {
            self.write_legacy_color(
                "dpfillfggray",
                "dpfillfgcr",
                "dpfillfgcg",
                "dpfillfgcb",
                "dpfillfgpal",
                fill.foreground,
            )?;
            self.write_legacy_color(
                "dpfillbggray",
                "dpfillbgcr",
                "dpfillbgcg",
                "dpfillbgcb",
                "dpfillbgpal",
                fill.background,
            )?;
            self.write_control_word("dpfillpat", Some(fill.pattern as i32))?;
        }
        if let Some(arrow) = properties.start_arrow {
            self.write_legacy_arrow("dpastart", arrow)?;
        }
        if let Some(arrow) = properties.end_arrow {
            self.write_legacy_arrow("dpaend", arrow)?;
        }
        if let Some(shadow) = properties.shadow {
            self.write_control_word("dpshadow", None)?;
            self.write_control_word("dpshadx", Some(shadow.x_offset))?;
            self.write_control_word("dpshady", Some(shadow.y_offset))?;
        }
        Ok(())
    }

    /// Write one inert generated list-marker destination.
    pub fn write_generated_list_marker(
        &mut self,
        marker: &crate::GeneratedListMarker<'_>,
    ) -> io::Result<()> {
        marker
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word(
            match marker.kind {
                crate::GeneratedListMarkerKind::Modern => "listtext",
                crate::GeneratedListMarkerKind::Legacy => "pntext",
            },
            None,
        )?;
        self.write_str(" ")?;
        let mut segments = marker.text.split('\t').peekable();
        while let Some(segment) = segments.next() {
            self.write_destination_text(segment)?;
            if segments.peek().is_some() {
                self.write_control_word("tab", None)?;
            }
        }
        self.write_str("}")
    }

    fn write_form_field_start(&mut self, field: &crate::FormField<'_>) -> io::Result<()> {
        field
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\field{\\*\\fldinst ")?;
        self.write_str(match field.field_type {
            crate::FormFieldType::Text => "FORMTEXT",
            crate::FormFieldType::CheckBox => "FORMCHECKBOX",
            crate::FormFieldType::DropDown => "FORMDROPDOWN",
        })?;
        if !field.data.is_empty() {
            self.write_str("{\\*\\datafield ")?;
            for byte in field.data.iter() {
                write!(self.writer, "{byte:02x}")?;
            }
            self.write_str("}")?;
        }
        self.write_str("{\\*\\formfield{")?;
        self.write_control_word("fftype", Some(field.field_type.to_rtf()))?;
        if let Some(value) = field.text_type {
            self.write_control_word("fftypetxt", Some(value.to_rtf()))?;
        }
        if let Some(value) = field.max_length {
            self.write_control_word("ffmaxlen", Some(i32::from(value)))?;
        }
        if let Some(value) = field.half_point_size {
            self.write_control_word("ffhps", Some(value))?;
        }
        if field.protected {
            self.write_control_word("ffprot", None)?;
        }
        if field.calculate_on_exit {
            self.write_control_word("ffrecalc", None)?;
        }
        if field.size_automatically {
            self.write_control_word("ffsize", None)?;
        }
        if field.own_help {
            self.write_control_word("ffownhelp", None)?;
        }
        if field.own_status {
            self.write_control_word("ffownstat", None)?;
        }
        if field.has_list_box {
            self.write_control_word("ffhaslistbox", None)?;
        }
        if let Some(value) = field.default_result {
            self.write_control_word("ffdefres", Some(value))?;
        }
        if let Some(value) = field.result {
            self.write_control_word("ffres", Some(value))?;
        }
        self.write_form_field_value("ffname", field.name.as_deref())?;
        self.write_form_field_value("ffformat", field.format.as_deref())?;
        self.write_form_field_value("ffdeftext", field.default_text.as_deref())?;
        self.write_form_field_value("ffhelptext", field.help_text.as_deref())?;
        self.write_form_field_value("ffstattext", field.status_text.as_deref())?;
        self.write_form_field_value("ffentrymcr", field.entry_macro.as_deref())?;
        self.write_form_field_value("ffexitmcr", field.exit_macro.as_deref())?;
        for entry in &field.list_entries {
            self.write_form_field_value("ffl", Some(entry.as_ref()))?;
        }
        self.write_str("}}}")?; // formfield and fldinst
        self.write_str("{\\fldrslt ")
    }

    fn write_form_field_value(&mut self, control: &str, value: Option<&str>) -> io::Result<()> {
        let Some(value) = value else {
            return Ok(());
        };
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    /// Write an inert source mark. Marks are canonicalized as hidden; any
    /// originally visible entry text remains in the ordinary body stream.
    pub fn write_navigation_entry(&mut self, entry: &crate::NavigationEntry<'_>) -> io::Result<()> {
        entry
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        self.write_str("{")?;
        match entry {
            crate::NavigationEntry::Index(entry) => {
                self.write_control_word("xe", None)?;
                self.write_control_word("v", None)?;
                if let Some(index_id) = entry.index_id {
                    self.write_control_word("xef", Some(i32::from(index_id)))?;
                }
                if entry.bold_page_number {
                    self.write_control_word("bxe", None)?;
                }
                if entry.italic_page_number {
                    self.write_control_word("ixe", None)?;
                }
                self.write_str(" ")?;
                self.write_destination_text(entry.text.as_ref())?;
                match &entry.page_reference {
                    crate::IndexPageReference::CurrentPage => {},
                    crate::IndexPageReference::ReplacementText(value) => {
                        self.write_str("{")?;
                        self.write_control_word("txe", None)?;
                        self.write_str(" ")?;
                        self.write_destination_text(value.as_ref())?;
                        self.write_str("}")?;
                    },
                    crate::IndexPageReference::BookmarkRange(value) => {
                        self.write_str("{")?;
                        self.write_control_word("rxe", None)?;
                        self.write_str(" ")?;
                        self.write_destination_text(value.as_ref())?;
                        self.write_str("}")?;
                    },
                }
                if let Some(yomi) = &entry.yomi {
                    self.write_str("{")?;
                    self.write_control_word("yxe", None)?;
                    self.write_str("{\\*")?;
                    self.write_control_word("pxe", None)?;
                    self.write_str(" ")?;
                    self.write_destination_text(yomi.as_ref())?;
                    self.write_str("}}")?;
                }
            },
            crate::NavigationEntry::TableOfContents(entry) => {
                self.write_control_word(
                    if entry.suppress_page_number {
                        "tcn"
                    } else {
                        "tc"
                    },
                    None,
                )?;
                self.write_control_word("v", None)?;
                self.write_control_word("tcf", Some(i32::from(entry.table_id)))?;
                self.write_control_word("tcl", Some(i32::from(entry.level)))?;
                self.write_str(" ")?;
                self.write_destination_text(entry.text.as_ref())?;
            },
        }
        self.write_str("}")
    }

    fn write_style_block_fragment(
        &mut self,
        block: &StyleBlock<'_>,
        start: usize,
        end: usize,
    ) -> io::Result<()> {
        let text = block.text.get(start..end).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark boundary splits a UTF-8 character",
            )
        })?;
        let fragment = StyleBlock::new(
            std::borrow::Cow::Borrowed(text),
            block.formatting,
            block.paragraph,
        );
        self.write_style_block(&fragment)
    }

    /// Write a style block
    fn write_style_block(&mut self, block: &StyleBlock) -> io::Result<()> {
        self.write_str("{")?;

        // Write character formatting
        self.write_formatting(&block.formatting)?;

        // Write paragraph properties
        self.write_paragraph_properties(&block.paragraph)?;

        // Delimit the final control word from body text that starts with a letter.
        self.write_str(" ")?;

        // Write text content
        self.write_text(block.text.as_ref())?;

        self.write_str("}")?;
        Ok(())
    }

    /// Write character formatting
    fn write_formatting(&mut self, fmt: &Formatting) -> io::Result<()> {
        if let Some(character_style) = fmt.character_style {
            self.write_control_word("cs", Some(i32::from(character_style)))?;
        }
        if let Some(insert_rsid) = fmt.insert_rsid {
            self.write_control_word("insrsid", Some(insert_rsid as i32))?;
        }
        if let Some(delete_rsid) = fmt.delete_rsid {
            self.write_control_word("delrsid", Some(delete_rsid as i32))?;
        }
        if let Some(char_style_rsid) = fmt.char_style_rsid {
            self.write_control_word("charrsid", Some(char_style_rsid as i32))?;
        }
        if let Some(direction) = fmt.direction {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrch",
                    TextDirection::RightToLeft => "rtlch",
                },
                None,
            )?;
        }

        if let Some(complex_script) = fmt.complex_script {
            self.write_control_word("fcs", Some(i32::from(complex_script)))?;
        }
        if let Some(character_type) = fmt.character_type {
            self.write_control_word(
                match character_type {
                    crate::CharacterType::LowAnsi => "loch",
                    crate::CharacterType::HighAnsi => "hich",
                    crate::CharacterType::DoubleByte => "dbch",
                },
                None,
            )?;
        }
        if let Some(character_grid) = fmt.character_grid {
            self.write_control_word(
                "cgrid",
                match character_grid {
                    crate::CharacterGrid::Parameterless => None,
                    crate::CharacterGrid::Value(value) => Some(i32::from(value)),
                },
            )?;
        }
        if fmt.animated_text != crate::AnimatedTextEffect::None {
            self.write_control_word("animtext", Some(fmt.animated_text.rtf_value()))?;
        }
        if let Some(value) = fmt.fit_text.rtf_value() {
            self.write_control_word("fittext", Some(value))?;
        }
        if fmt.emphasis_mark != crate::EmphasisMark::None {
            self.write_control_word(fmt.emphasis_mark.control_word(), None)?;
        }

        if let Some(language) = fmt.language {
            self.write_control_word("lang", Some(language.rtf_value()))?;
        }
        if let Some(language) = fmt.east_asian_language {
            self.write_control_word("langfe", Some(language.rtf_value()))?;
        }
        if let Some(language) = fmt.language_no_proof {
            self.write_control_word("langnp", Some(language.rtf_value()))?;
        }
        if let Some(language) = fmt.east_asian_language_no_proof {
            self.write_control_word("langfenp", Some(language.rtf_value()))?;
        }
        if fmt.no_proof {
            self.write_control_word("noproof", None)?;
        }

        fmt.associated
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(bold) = fmt.associated.bold {
            self.write_control_word("ab", Some(i32::from(bold)))?;
        }
        if let Some(all_caps) = fmt.associated.all_caps {
            self.write_control_word("acaps", Some(i32::from(all_caps)))?;
        }
        if let Some(color_ref) = fmt.associated.color_ref {
            self.write_control_word("acf", Some(i32::from(color_ref)))?;
        }
        if let Some(crate::AssociatedCharacterBaseline::LoweredHalfPoints(value)) =
            fmt.associated.baseline
        {
            self.write_control_word("adn", Some(i32::from(value)))?;
        }
        if let Some(expansion) = fmt.associated.expansion_quarter_points {
            self.write_control_word("aexpnd", Some(i32::from(expansion)))?;
        }
        if let Some(font_ref) = fmt.associated.font_ref {
            self.write_control_word("af", Some(i32::from(font_ref)))?;
        }
        if let Some(font_size) = fmt.associated.font_size {
            self.write_control_word("afs", Some(i32::from(font_size.get())))?;
        }
        if let Some(italic) = fmt.associated.italic {
            self.write_control_word("ai", Some(i32::from(italic)))?;
        }
        if let Some(language) = fmt.associated.language {
            self.write_control_word("alang", Some(language.rtf_value()))?;
        }
        if let Some(outline) = fmt.associated.outline {
            self.write_control_word("aoutl", Some(i32::from(outline)))?;
        }
        if let Some(small_caps) = fmt.associated.small_caps {
            self.write_control_word("ascaps", Some(i32::from(small_caps)))?;
        }
        if let Some(shadow) = fmt.associated.shadow {
            self.write_control_word("ashad", Some(i32::from(shadow)))?;
        }
        if let Some(strike) = fmt.associated.strike {
            self.write_control_word("astrike", Some(i32::from(strike)))?;
        }
        if let Some(underline) = fmt.associated.underline {
            self.write_control_word(
                match underline {
                    crate::AssociatedUnderlineStyle::None => "aulnone",
                    crate::AssociatedUnderlineStyle::Single => "aul",
                    crate::AssociatedUnderlineStyle::Dotted => "auld",
                    crate::AssociatedUnderlineStyle::Double => "auldb",
                    crate::AssociatedUnderlineStyle::Words => "aulw",
                },
                None,
            )?;
        }
        if let Some(crate::AssociatedCharacterBaseline::RaisedHalfPoints(value)) =
            fmt.associated.baseline
        {
            self.write_control_word("aup", Some(i32::from(value)))?;
        }

        // Font
        if fmt.font_ref != 0 {
            self.write_control_word("f", Some(fmt.font_ref as i32))?;
        }

        // Font size
        self.write_control_word("fs", Some(fmt.font_size.get() as i32))?;

        // Color
        if fmt.color_ref != 0 {
            self.write_control_word("cf", Some(fmt.color_ref as i32))?;
        }

        // Exact character background color, independent of highlighting.
        if let Some(background_color) = fmt.background_color {
            self.write_control_word("cb", Some(i32::from(background_color)))?;
        }

        // Highlight
        if let Some(highlight) = fmt.highlight_color {
            self.write_control_word("highlight", Some(highlight as i32))?;
        }

        if let Some(border) = fmt.character_border {
            border
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            self.write_control_word("chbrdr", None)?;
            self.write_control_word(border.style.control_word(), None)?;
            self.write_control_word("brdrw", Some(i32::from(border.width)))?;
            self.write_control_word("brdrcf", Some(i32::from(border.color_ref)))?;
            self.write_control_word("brsp", Some(i32::from(border.space)))?;
            if border.shadow {
                self.write_control_word("brdrsh", None)?;
            }
            if border.frame {
                self.write_control_word("brdrframe", None)?;
            }
        }

        if let Some(shading) = fmt.character_shading {
            shading
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            self.write_control_word("chshdng", Some(i32::from(shading.amount)))?;
            self.write_control_word("chcfpat", Some(i32::from(shading.foreground_color)))?;
            self.write_control_word("chcbpat", Some(i32::from(shading.background_color)))?;
        }

        // Bold
        if fmt.bold {
            self.write_control_word("b", None)?;
        }

        // Italic
        if fmt.italic {
            self.write_control_word("i", None)?;
        }

        // Underline
        match fmt.underline {
            UnderlineStyle::None => {},
            UnderlineStyle::Single => self.write_control_word("ul", None)?,
            UnderlineStyle::Double => self.write_control_word("uldb", None)?,
            UnderlineStyle::Dotted => self.write_control_word("uld", None)?,
            UnderlineStyle::Dashed => self.write_control_word("uldash", None)?,
            UnderlineStyle::DashDot => self.write_control_word("uldashd", None)?,
            UnderlineStyle::DashDotDot => self.write_control_word("uldashdd", None)?,
            UnderlineStyle::Words => self.write_control_word("ulw", None)?,
            UnderlineStyle::Thick => self.write_control_word("ulth", None)?,
            UnderlineStyle::Wave => self.write_control_word("ulwave", None)?,
            UnderlineStyle::Hairline => self.write_control_word("ulhair", None)?,
            UnderlineStyle::ThickDotted => self.write_control_word("ulthd", None)?,
            UnderlineStyle::ThickDashed => self.write_control_word("ulthdash", None)?,
            UnderlineStyle::ThickDashDot => self.write_control_word("ulthdashd", None)?,
            UnderlineStyle::ThickDashDotDot => {
                self.write_control_word("ulthdashdd", None)?
            },
            UnderlineStyle::ThickLongDash => self.write_control_word("ulthldash", None)?,
            UnderlineStyle::LongDash => self.write_control_word("ulldash", None)?,
            UnderlineStyle::HeavyWave => self.write_control_word("ulhwave", None)?,
            UnderlineStyle::DoubleWave => self.write_control_word("ululdbwave", None)?,
        }
        if let Some(underline_color) = fmt.underline_color {
            self.write_control_word("ulc", Some(i32::from(underline_color)))?;
        }

        // Strike
        if fmt.strike {
            self.write_control_word("strike", None)?;
        }

        // Double strike
        if fmt.double_strike {
            self.write_control_word("striked", None)?;
        }

        fmt.character_positioning
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        match fmt.character_positioning.baseline {
            CharacterBaseline::Normal if fmt.superscript => {
                self.write_control_word("super", None)?
            },
            CharacterBaseline::Normal if fmt.subscript => self.write_control_word("sub", None)?,
            CharacterBaseline::Normal => {},
            CharacterBaseline::Superscript => self.write_control_word("super", None)?,
            CharacterBaseline::Subscript => self.write_control_word("sub", None)?,
            CharacterBaseline::RaisedHalfPoints(value) => {
                self.write_control_word("up", Some(i32::from(value)))?
            },
            CharacterBaseline::LoweredHalfPoints(value) => {
                self.write_control_word("dn", Some(i32::from(value)))?
            },
        }

        // Small caps
        if fmt.smallcaps {
            self.write_control_word("scaps", None)?;
        }

        // All caps
        if fmt.all_caps {
            self.write_control_word("caps", None)?;
        }

        // Hidden
        if fmt.hidden {
            self.write_control_word("v", None)?;
        }

        // Outline
        if fmt.outline {
            self.write_control_word("outl", None)?;
        }

        // Shadow
        if fmt.shadow {
            self.write_control_word("shad", None)?;
        }

        // Emboss
        if fmt.emboss {
            self.write_control_word("embo", None)?;
        }

        // Imprint
        if fmt.imprint {
            self.write_control_word("impr", None)?;
        }

        match fmt.character_positioning.expansion {
            CharacterExpansion::None if fmt.char_spacing != 0 => {
                self.write_control_word("expnd", Some(fmt.char_spacing))?
            },
            CharacterExpansion::None => {},
            CharacterExpansion::QuarterPoints(value) => {
                self.write_control_word("expnd", Some(i32::from(value)))?
            },
            CharacterExpansion::Twips(value) => {
                self.write_control_word("expndtw", Some(i32::from(value)))?
            },
        }
        let scale = if fmt.character_positioning.horizontal_scale_percent != 100 {
            i32::from(fmt.character_positioning.horizontal_scale_percent)
        } else {
            fmt.char_scale
        };
        if scale != 100 {
            self.write_control_word("charscalex", Some(scale))?;
        }
        let kerning = if fmt.character_positioning.kerning_half_points != 0 {
            i32::from(fmt.character_positioning.kerning_half_points)
        } else {
            fmt.kerning
        };
        if kerning != 0 {
            self.write_control_word("kerning", Some(kerning))?;
        }

        Ok(())
    }

    fn write_legacy_paragraph_numbering(&mut self, index: Option<u32>) -> io::Result<()> {
        let Some(index) = index else {
            return Ok(());
        };
        let record = self
            .legacy_paragraph_numbering
            .get(index as usize)
            .cloned()
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF paragraph references a missing legacy pn record",
                )
            })?;
        record
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;
        self.write_control_word("pn", None)?;
        match record.level {
            crate::LegacyParagraphNumberingLevel::Explicit(value) => {
                self.write_control_word("pnlvl", Some(i32::from(value)))?
            },
            crate::LegacyParagraphNumberingLevel::Bullet => {
                self.write_control_word("pnlvlblt", None)?
            },
            crate::LegacyParagraphNumberingLevel::Body => {
                self.write_control_word("pnlvlbody", None)?
            },
            crate::LegacyParagraphNumberingLevel::Continue => {
                self.write_control_word("pnlvlcont", None)?
            },
        }
        if let Some(format) = record.format {
            self.write_control_word(
                match format {
                    crate::LegacyParagraphNumberingFormat::Aiueo => "pnaiu",
                    crate::LegacyParagraphNumberingFormat::AiueoDbChar => "pnaiud",
                    crate::LegacyParagraphNumberingFormat::AiueoExtended => "pnaiueo",
                    crate::LegacyParagraphNumberingFormat::AiueoExtendedDbChar => "pnaiueod",
                    crate::LegacyParagraphNumberingFormat::Chosung => "pnchosung",
                    crate::LegacyParagraphNumberingFormat::CardinalText => "pncard",
                    crate::LegacyParagraphNumberingFormat::Decimal => "pndec",
                    crate::LegacyParagraphNumberingFormat::DecimalWithPeriod => "pndecd",
                    crate::LegacyParagraphNumberingFormat::UpperRoman => "pnucrm",
                    crate::LegacyParagraphNumberingFormat::LowerRoman => "pnlcrm",
                    crate::LegacyParagraphNumberingFormat::UpperLetter => "pnucltr",
                    crate::LegacyParagraphNumberingFormat::LowerLetter => "pnlcltr",
                    crate::LegacyParagraphNumberingFormat::Ordinal => "pnord",
                    crate::LegacyParagraphNumberingFormat::OrdinalText => "pnordt",
                    crate::LegacyParagraphNumberingFormat::ChineseCounting => "pncnum",
                    crate::LegacyParagraphNumberingFormat::ChineseCountingDbChar => "pndbnum",
                    crate::LegacyParagraphNumberingFormat::ChineseCountingKorean => "pndbnumd",
                    crate::LegacyParagraphNumberingFormat::ChineseCountingLegal => "pndbnumk",
                    crate::LegacyParagraphNumberingFormat::ChineseCountingThousand => "pndbnuml",
                    crate::LegacyParagraphNumberingFormat::ChineseCountingTraditional => "pndbnumt",
                    crate::LegacyParagraphNumberingFormat::Ganada => "pnganada",
                    crate::LegacyParagraphNumberingFormat::GbCounting => "pngbnum",
                    crate::LegacyParagraphNumberingFormat::GbCountingDbChar => "pngbnumd",
                    crate::LegacyParagraphNumberingFormat::GbCountingKorean => "pngbnumk",
                    crate::LegacyParagraphNumberingFormat::GbCountingLegal => "pngbnuml",
                    crate::LegacyParagraphNumberingFormat::GbLip => "pngblip",
                    crate::LegacyParagraphNumberingFormat::Iroha => "pniroha",
                    crate::LegacyParagraphNumberingFormat::IrohaDbChar => "pnirohad",
                    crate::LegacyParagraphNumberingFormat::Zodiac => "pnzodiac",
                    crate::LegacyParagraphNumberingFormat::ZodiacDbChar => "pnzodiacd",
                    crate::LegacyParagraphNumberingFormat::ZodiacLegal => "pnzodiacl",
                },
                None,
            )?;
        }
        if let Some(value) = record.alignment {
            self.write_control_word(
                match value {
                    crate::LegacyParagraphNumberingAlignment::Left => "pnql",
                    crate::LegacyParagraphNumberingAlignment::Center => "pnqc",
                    crate::LegacyParagraphNumberingAlignment::Right => "pnqr",
                },
                None,
            )?;
        }
        for (enabled, name) in [
            (record.across, "pnacross"),
            (record.number_once, "pnnumonce"),
            (record.previous, "pnprev"),
            (record.restart, "pnrestart"),
            (record.hanging, "pnhang"),
        ] {
            if enabled {
                self.write_control_word(name, None)?;
            }
        }
        if let Some(value) = record.bidi {
            self.write_control_word(
                match value {
                    crate::LegacyParagraphNumberingBidi::A => "pnbidia",
                    crate::LegacyParagraphNumberingBidi::B => "pnbidib",
                },
                None,
            )?;
        }
        if let Some(value) = record.start_at {
            self.write_control_word("pnstart", Some(value))?;
        }
        if let Some(value) = record.indent {
            self.write_control_word("pnindent", Some(value))?;
        }
        if let Some(value) = record.space {
            self.write_control_word("pnsp", Some(value))?;
        }
        if let Some(value) = record.font_ref {
            self.write_control_word("pnf", Some(i32::from(value)))?;
        }
        if let Some(value) = record.font_size {
            self.write_control_word("pnfs", Some(i32::from(value)))?;
        }
        if let Some(value) = record.color_ref {
            self.write_control_word("pncf", Some(i32::from(value)))?;
        }
        for (value, name) in [
            (record.bold, "pnb"),
            (record.italic, "pni"),
            (record.caps, "pncaps"),
            (record.small_caps, "pnscaps"),
            (record.strike, "pnstrike"),
        ] {
            if let Some(value) = value {
                self.write_control_word(name, (!value).then_some(0))?;
            }
        }
        if let Some(value) = record.underline {
            self.write_control_word(
                match value {
                    crate::LegacyParagraphNumberingUnderline::None => "pnulnone",
                    crate::LegacyParagraphNumberingUnderline::Single => "pnul",
                    crate::LegacyParagraphNumberingUnderline::Dotted => "pnuld",
                    crate::LegacyParagraphNumberingUnderline::Dashed => "pnuldash",
                    crate::LegacyParagraphNumberingUnderline::DashDot => "pnuldashd",
                    crate::LegacyParagraphNumberingUnderline::DashDotDot => "pnuldashdd",
                    crate::LegacyParagraphNumberingUnderline::Double => "pnuldb",
                    crate::LegacyParagraphNumberingUnderline::Hairline => "pnulhair",
                    crate::LegacyParagraphNumberingUnderline::Thick => "pnulth",
                    crate::LegacyParagraphNumberingUnderline::Words => "pnulw",
                    crate::LegacyParagraphNumberingUnderline::Wave => "pnulwave",
                },
                None,
            )?;
        }
        let revision = &record.revision;
        if let Some(value) = revision.author {
            self.write_control_word("pnrauth", Some(i32::from(value)))?;
        }
        if let Some(value) = revision.date {
            self.write_control_word("pnrdate", Some(value))?;
        }
        if let Some(value) = revision.number_format {
            self.write_control_word("pnrnfc", Some(value))?;
        }
        if revision.no_tracking {
            self.write_control_word("pnrnot", None)?;
        }
        if let Some(value) = revision.paragraph_number {
            self.write_control_word("pnrpnbr", Some(value))?;
        }
        if let Some(value) = revision.rgb {
            self.write_control_word("pnrrgb", Some(value as i32))?;
        }
        if let Some(value) = revision.start {
            self.write_control_word("pnrstart", Some(value))?;
        }
        if let Some(value) = revision.stop {
            self.write_control_word("pnrstop", Some(value))?;
        }
        if let Some(value) = revision.text_start {
            self.write_control_word("pnrxst", Some(value))?;
        }
        if let Some(value) = record.text_before {
            self.write_str("{")?;
            self.write_control_word("pntxtb", None)?;
            self.write_str(" ")?;
            self.write_text(value.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(value) = record.text_after {
            self.write_str("{")?;
            self.write_control_word("pntxta", None)?;
            self.write_str(" ")?;
            self.write_text(value.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("}")
    }

    /// Write author/date metadata for a structural revision marker, after
    /// validating the author index.
    fn write_revision_metadata(
        &mut self,
        author_control: &'static str,
        date_control: &'static str,
        metadata: crate::RevisionMetadata,
    ) -> io::Result<()> {
        metadata
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(author) = metadata.author {
            self.write_control_word(author_control, Some(author))?;
        }
        if let Some(date) = metadata.date {
            self.write_control_word(date_control, Some(date))?;
        }
        Ok(())
    }

    /// Write paragraph properties
    fn write_paragraph_properties(&mut self, para: &Paragraph) -> io::Result<()> {
        if let Some(paragraph_style) = para.paragraph_style {
            self.write_control_word("s", Some(i32::from(paragraph_style)))?;
        }
        if let Some(paragraph_rsid) = para.paragraph_rsid {
            self.write_control_word("pararsid", Some(paragraph_rsid as i32))?;
        }
        if let Some(outline_level) = para.outline_level {
            self.write_control_word("outlinelevel", Some(i32::from(outline_level)))?;
        }
        self.write_revision_metadata("prauth", "prdate", para.revision)?;
        self.write_legacy_paragraph_numbering(para.legacy_numbering)?;
        if let Some(direction) = para.direction {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrpar",
                    TextDirection::RightToLeft => "rtlpar",
                },
                None,
            )?;
        }

        // Alignment
        match para.alignment {
            Alignment::Left => self.write_control_word("ql", None)?,
            Alignment::Right => self.write_control_word("qr", None)?,
            Alignment::Center => self.write_control_word("qc", None)?,
            Alignment::Justify => self.write_control_word("qj", None)?,
        }

        // Spacing
        if para.spacing.before != 0 {
            self.write_control_word("sb", Some(para.spacing.before))?;
        }
        if para.spacing.after != 0 {
            self.write_control_word("sa", Some(para.spacing.after))?;
        }
        if let Some(value) = para.spacing_policy.list_before {
            self.write_control_word(
                "lisb",
                Some(i32::try_from(value).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF lisb exceeds i32")
                })?),
            )?;
        }
        if let Some(value) = para.spacing_policy.list_after {
            self.write_control_word(
                "lisa",
                Some(i32::try_from(value).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF lisa exceeds i32")
                })?),
            )?;
        }
        if para.spacing_policy.automatic_before {
            self.write_control_word("sbauto", Some(1))?;
        }
        if para.spacing_policy.automatic_after {
            self.write_control_word("saauto", Some(1))?;
        }
        if !para.spacing_policy.snap_to_line_grid {
            self.write_control_word("nosnaplinegrid", None)?;
        }
        if para.spacing_policy.contextual_spacing {
            self.write_control_word("contextualspace", None)?;
        }
        if para.spacing.line != 0 {
            self.write_control_word("sl", Some(para.spacing.line))?;
            if para.spacing.line_multiple {
                self.write_control_word("slmult", Some(1))?;
            }
        }

        // Indentation
        if para.indentation.left != 0 {
            self.write_control_word("li", Some(para.indentation.left))?;
        }
        if para.indentation.right != 0 {
            self.write_control_word("ri", Some(para.indentation.right))?;
        }
        if para.indentation.first_line != 0 {
            self.write_control_word("fi", Some(para.indentation.first_line))?;
        }
        let logical = para.logical_indentation;
        if let Some(v) = logical.start {
            self.write_control_word("lin", Some(v))?;
        }
        if let Some(v) = logical.end {
            self.write_control_word("rin", Some(v))?;
        }
        if let Some(v) = logical.first_line_character_units {
            self.write_control_word("cufi", Some(v))?;
        }
        if let Some(v) = logical.left_character_units {
            self.write_control_word("culi", Some(v))?;
        }
        if let Some(v) = logical.right_character_units {
            self.write_control_word("curi", Some(v))?;
        }
        if logical.mirrored {
            self.write_control_word("indmirror", None)?;
        }

        // Borders (if any)
        self.write_borders(&para.borders)?;

        // Shading (if any)
        self.write_shading(&para.shading)?;

        // Custom tab stops, retained in declaration order.
        for tab in &para.tab_stops {
            self.write_tab_stop(tab)?;
        }

        if let Some(drop_cap) = para.drop_cap {
            drop_cap.validate().map_err(|error| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    format!("invalid RTF paragraph drop cap: {error}"),
                )
            })?;
            self.write_control_word("dropcapli", Some(i32::from(drop_cap.line_count())))?;
            self.write_control_word("dropcapt", Some(drop_cap.kind().as_rtf_value()))?;
        }

        // Keep together
        if para.keep_together {
            self.write_control_word("keep", None)?;
        }

        // Keep with next
        if para.keep_next {
            self.write_control_word("keepn", None)?;
        }

        // Side-by-side
        if para.side_by_side {
            self.write_control_word("sbys", None)?;
        }

        // Page break before
        if para.page_break_before {
            self.write_control_word("pagebb", None)?;
        }

        // Widow control
        if para.widow_control {
            self.write_control_word("widctlpar", None)?;
        }
        if para.no_line_numbering {
            self.write_control_word("noline", None)?;
        }
        if para.no_auto_tab_indent {
            self.write_control_word("notabind", None)?;
        }

        let breaking = para.line_breaking;
        if breaking.automatic_hyphenation {
            self.write_control_word("hyphpar", None)?;
        }
        match breaking.wrapping {
            crate::ParagraphWrapping::Default => {},
            crate::ParagraphWrapping::NoCharacterWrap => {
                self.write_control_word("nocwrap", None)?
            },
            crate::ParagraphWrapping::NoWordWrap => self.write_control_word("nowwrap", None)?,
            crate::ParagraphWrapping::NoOverflow => self.write_control_word("nooverflow", None)?,
        }
        if breaking.auto_space_alphabetic {
            self.write_control_word("aspalpha", None)?;
        }
        if breaking.auto_space_numbers {
            self.write_control_word("aspnum", None)?;
        }
        match breaking.font_alignment {
            crate::ParagraphFontAlignment::Auto => {},
            crate::ParagraphFontAlignment::Hanging => self.write_control_word("fahang", None)?,
            crate::ParagraphFontAlignment::Center => self.write_control_word("facenter", None)?,
            crate::ParagraphFontAlignment::Roman => self.write_control_word("faroman", None)?,
            crate::ParagraphFontAlignment::Variable => self.write_control_word("favar", None)?,
            crate::ParagraphFontAlignment::Fixed => self.write_control_word("fafixed", None)?,
        }
        if breaking.adjust_right_indent {
            self.write_control_word("adjustright", None)?;
        }

        if let Some(list_override) = para.list_override {
            self.write_control_word("ls", Some(list_override))?;
        }
        if let Some(list_level) = para.list_level {
            if list_level > 8 {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF paragraph list levels must be between zero and eight",
                ));
            }
            self.write_control_word("ilvl", Some(i32::from(list_level)))?;
        }

        Ok(())
    }

    /// Write borders
    fn write_borders(&mut self, borders: &Borders) -> io::Result<()> {
        if !borders.has_any_border() {
            return Ok(());
        }

        // Top border
        if borders.top.is_visible() {
            self.write_border("brdrt", &borders.top)?;
        }

        // Bottom border
        if borders.bottom.is_visible() {
            self.write_border("brdrb", &borders.bottom)?;
        }

        // Left border
        if borders.left.is_visible() {
            self.write_border("brdrl", &borders.left)?;
        }

        // Right border
        if borders.right.is_visible() {
            self.write_border("brdrr", &borders.right)?;
        }

        // Bar border
        if borders.bar.is_visible() {
            self.write_border("brdrbar", &borders.bar)?;
        }

        // Between border
        if borders.between.is_visible() {
            self.write_border("brdrbtw", &borders.between)?;
        }

        Ok(())
    }

    /// Write a single border
    fn write_border(&mut self, control: &str, border: &Border) -> io::Result<()> {
        self.write_control_word(control, None)?;

        // Border style
        let style_word = match border.style {
            BorderStyle::None => return Ok(()),
            BorderStyle::Single => "brdrs",
            BorderStyle::Thick => "brdrth",
            BorderStyle::Dotted => "brdrdot",
            BorderStyle::Dashed => "brdrdash",
            BorderStyle::DashSmallGap => "brdrdashsm",
            BorderStyle::DotDash => "brdrdashd",
            BorderStyle::DotDotDash => "brdrdashdd",
            BorderStyle::Double => "brdrdb",
            BorderStyle::Triple => "brdrtriple",
            BorderStyle::ThickThinSmall => "brdrtnthsg",
            BorderStyle::ThinThickSmall => "brdrtnthmg",
            BorderStyle::ThinThickThinSmall => "brdrtnthtnsg",
            BorderStyle::ThickThinMedium => "brdrtnthmg",
            BorderStyle::ThinThickMedium => "brdrthtnmg",
            BorderStyle::ThinThickThinMedium => "brdrtnthtnmg",
            BorderStyle::ThickThinLarge => "brdrtnthlg",
            BorderStyle::ThinThickLarge => "brdrththlg",
            BorderStyle::ThinThickThinLarge => "brdrtnthtnlg",
            BorderStyle::Wavy => "brdrwavy",
            BorderStyle::WavyDouble => "brdrwavydb",
            BorderStyle::Striped => "brdrdashdotstr",
            BorderStyle::Embossed => "brdremboss",
            BorderStyle::Engraved => "brdrengrave",
            BorderStyle::Outset => "brdroutset",
            BorderStyle::Inset => "brdrinset",
            BorderStyle::Hairline => "brdrhair",
        };
        self.write_control_word(style_word, None)?;

        // Border width
        self.write_control_word("brdrw", Some(border.width))?;

        // Border color
        if border.color_ref != 0 {
            self.write_control_word("brdrcf", Some(border.color_ref as i32))?;
        }

        // Border space
        if border.space != 0 {
            self.write_control_word("brsp", Some(border.space))?;
        }

        // Border shadow
        if border.shadow {
            self.write_control_word("brdrsh", None)?;
        }

        // Border frame
        if border.frame {
            self.write_control_word("brdrframe", None)?;
        }

        Ok(())
    }

    /// Write shading
    fn write_shading(&mut self, shading: &Shading) -> io::Result<()> {
        shading
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if !shading.is_present() {
            return Ok(());
        }

        let pattern_value = match (shading.amount, shading.pattern) {
            (Some(amount), _) => Some(i32::from(amount)),
            (None, None) => None,
            (None, Some(ShadingPattern::Clear)) => Some(0),
            (None, Some(ShadingPattern::Solid)) => Some(10_000),
            (None, Some(ShadingPattern::Percent5)) => Some(500),
            (None, Some(ShadingPattern::Percent10)) => Some(1000),
            (None, Some(ShadingPattern::Percent12)) => Some(1250),
            (None, Some(ShadingPattern::Percent15)) => Some(1500),
            (None, Some(ShadingPattern::Percent20)) => Some(2000),
            (None, Some(ShadingPattern::Percent25)) => Some(2500),
            (None, Some(ShadingPattern::Percent30)) => Some(3000),
            (None, Some(ShadingPattern::Percent35)) => Some(3500),
            (None, Some(ShadingPattern::Percent40)) => Some(4000),
            (None, Some(ShadingPattern::Percent45)) => Some(4500),
            (None, Some(ShadingPattern::Percent50)) => Some(5000),
            (None, Some(ShadingPattern::Percent55)) => Some(5500),
            (None, Some(ShadingPattern::Percent60)) => Some(6000),
            (None, Some(ShadingPattern::Percent62)) => Some(6250),
            (None, Some(ShadingPattern::Percent65)) => Some(6500),
            (None, Some(ShadingPattern::Percent70)) => Some(7000),
            (None, Some(ShadingPattern::Percent75)) => Some(7500),
            (None, Some(ShadingPattern::Percent80)) => Some(8000),
            (None, Some(ShadingPattern::Percent85)) => Some(8500),
            (None, Some(ShadingPattern::Percent87)) => Some(8750),
            (None, Some(ShadingPattern::Percent90)) => Some(9000),
            (None, Some(ShadingPattern::Percent95)) => Some(9500),
            (None, Some(_)) => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "explicit paragraph shading patterns are not representable by the numeric shading family",
                ));
            },
        };

        if let Some(pattern_value) = pattern_value {
            self.write_control_word("shading", Some(pattern_value))?;
        }

        // Foreground color
        if let Some(color) = shading.foreground_color {
            self.write_control_word("cfpat", Some(i32::from(color)))?;
        }

        // Background color
        if let Some(color) = shading.background_color {
            self.write_control_word("cbpat", Some(i32::from(color)))?;
        }

        Ok(())
    }

    /// Write tab stop
    ///
    fn write_tab_stop(&mut self, tab: &TabStop) -> io::Result<()> {
        // The left kind is implicit. A bar tab uses `tbN` as its terminator.
        match tab.alignment {
            TabAlignment::Left | TabAlignment::Bar => {},
            TabAlignment::Right => self.write_control_word("tqr", None)?,
            TabAlignment::Center => self.write_control_word("tqc", None)?,
            TabAlignment::Decimal => self.write_control_word("tqdec", None)?,
        }

        // Tab leader
        match tab.leader {
            TabLeader::None => {},
            TabLeader::Dot => self.write_control_word("tldot", None)?,
            TabLeader::MiddleDot => self.write_control_word("tlmdot", None)?,
            TabLeader::Hyphen => self.write_control_word("tlhyph", None)?,
            TabLeader::Underscore => self.write_control_word("tlul", None)?,
            TabLeader::ThickLine => self.write_control_word("tlth", None)?,
            TabLeader::Equal => self.write_control_word("tleq", None)?,
        }

        self.write_control_word(
            if tab.alignment == TabAlignment::Bar {
                "tb"
            } else {
                "tx"
            },
            Some(tab.position),
        )?;

        Ok(())
    }

    /// Write a table
    fn write_table(
        &mut self,
        table: &Table,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        if let Some(first) = table.rows().first()
            && table
                .rows()
                .iter()
                .skip(1)
                .any(|row| row.positioning() != first.positioning())
        {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF positioned-table properties must be identical for all rows in one logical table",
            ));
        }
        for row in table.rows() {
            self.write_table_row(row, table.direction(), fields, navigation_entries, revisions)?;
        }
        Ok(())
    }

    fn validate_table_tree(table: &Table, depth: usize, count: &mut usize) -> io::Result<()> {
        if depth > crate::MAX_TABLE_NESTING_DEPTH {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF table nesting exceeds 32 levels",
            ));
        }
        *count = count.checked_add(1).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF logical table count overflow",
            )
        })?;
        if *count > 4096 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF document exceeds 4096 logical tables",
            ));
        }
        if table.row_count() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF logical table exceeds 65536 rows",
            ));
        }
        table
            .validate_merges()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        if let Some(first) = table.rows().first()
            && table
                .rows()
                .iter()
                .skip(1)
                .any(|row| row.positioning() != first.positioning())
        {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF positioned-table properties must be identical for all rows in one logical table",
            ));
        }
        for row in table.rows() {
            row.geometry()
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if row.cell_count() > crate::MAX_TABLE_CELLS_PER_ROW {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF table row exceeds 4096 cells",
                ));
            }
            for cell in row.cells() {
                cell.validate_drawings().map_err(|error| {
                    io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                })?;
                if let Some(width) = cell.preferred_width() {
                    width.validate().map_err(|error| {
                        io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                    })?;
                }
                let mut previous = 0;
                for nested in cell.nested_tables() {
                    if nested.text_offset < previous
                        || nested.text_offset > cell.text().len()
                        || !cell.text().is_char_boundary(nested.text_offset)
                    {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "invalid nested-table text insertion offset",
                        ));
                    }
                    previous = nested.text_offset;
                    Self::validate_table_tree(&nested.table, depth + 1, count)?;
                }
            }
        }
        Ok(())
    }

    /// Write a table row
    fn write_table_row(
        &mut self,
        row: &Row,
        table_direction: Option<TextDirection>,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        // Row defaults
        self.write_control_word("trowd", None)?;
        if let Some(table_style) = row.table_style() {
            self.write_control_word("ts", Some(i32::from(table_style)))?;
        }
        if let Some(table_rsid) = row.table_rsid() {
            self.write_control_word("tblrsid", Some(table_rsid as i32))?;
        }
        self.write_revision_metadata("trauth", "trdate", row.revision())?;

        if let Some(direction) = table_direction {
            self.write_control_word(
                "taprtl",
                (direction == TextDirection::LeftToRight).then_some(0),
            )?;
        }
        self.write_table_row_banding(row)?;
        self.write_table_row_layout(row)?;
        self.write_table_row_geometry(row.geometry())?;
        self.write_table_row_borders(row.borders())?;
        self.write_table_shading("tr", row.shading())?;
        self.write_table_distances("trpadd", "trpaddf", row.padding())?;
        self.write_table_distances("trspd", "trspdf", row.spacing())?;
        self.write_table_row_cell_defaults(row.cell_defaults())?;
        self.write_floating_table_position(row.positioning())?;

        // Cell boundaries
        let cell_width = 2880; // Default cell width (2 inches)
        for (i, cell) in row.cells().iter().enumerate() {
            self.write_table_preferred_width("clftsWidth", "clwWidth", cell.preferred_width())?;
            self.write_table_cell_merge(cell)?;
            self.write_table_cell_revision(cell)?;
            self.write_table_cell_layout(cell)?;
            self.write_table_cell_borders(cell.borders())?;
            self.write_table_shading("cl", cell.shading())?;
            self.write_table_distances("clpad", "clpadf", cell.padding())?;
            self.write_table_distances("clspd", "clspdf", cell.spacing())?;
            let boundary = cell
                .right_boundary()
                .unwrap_or(cell_width * ((i + 1) as i32));
            self.write_control_word("cellx", Some(boundary))?;
        }

        // Write cells
        for cell in row.cells() {
            self.write_str("{")?;
            self.write_control_word("intbl", None)?;
            self.write_str(" ")?;
            self.write_cell_content(cell, 1, fields, navigation_entries, revisions)?;
            self.write_control_word("cell", None)?;
            self.write_str("}")?;
        }

        // Row end
        self.write_control_word("row", None)?;
        self.write_str("\n")?;

        Ok(())
    }

    fn write_cell_content(
        &mut self,
        cell: &crate::Cell<'_>,
        depth: usize,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        let mut offset = 0;
        for event in cell.story_events() {
            let position = match *event {
                crate::CellStoryEvent::NestedTable(index) => {
                    cell.nested_tables()[index].text_offset
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                    cell.shapes()[index].position
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                    cell.shape_groups()[index].position
                },
                crate::CellStoryEvent::Field(field) => field.position,
                crate::CellStoryEvent::PageBreak(page_break) => page_break.position,
                crate::CellStoryEvent::ColumnBreak(column_break) => column_break.position,
                crate::CellStoryEvent::NavigationEntry(reference)
                | crate::CellStoryEvent::RevisionStart(reference)
                | crate::CellStoryEvent::RevisionEnd(reference)
                | crate::CellStoryEvent::RevisionDeletion(reference) => reference.position,
            };
            self.write_text(&cell.text()[offset..position])?;
            match *event {
                crate::CellStoryEvent::NestedTable(index) => {
                    self.write_nested_table(&cell.nested_tables()[index].table, depth + 1, fields, navigation_entries, revisions)?
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => {
                    self.write_root_shape(&cell.shapes()[index])?
                },
                crate::CellStoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => {
                    self.write_shape_group(&cell.shape_groups()[index], true)?
                },
                crate::CellStoryEvent::Field(reference) => {
                    let field=fields.get(reference.field_index).filter(|field|field.owner==crate::FieldOwner::TableCell(depth as u8)&&field.position==reference.position&&field.range_end==reference.position).ok_or_else(||io::Error::new(io::ErrorKind::InvalidInput,"RTF table-cell story has an invalid generic-field owner or reference"))?;
                    self.write_field_with_fields(field, fields, 0)?;
                },
                crate::CellStoryEvent::PageBreak(_) => self.write_str("\\page ")?,
                crate::CellStoryEvent::ColumnBreak(_) => self.write_str("\\column ")?,
                crate::CellStoryEvent::NavigationEntry(reference) => {
                    let entry = navigation_entries.get(reference.index).filter(|entry| {
                        entry.position() == reference.position
                    }).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell navigation reference is invalid"))?;
                    self.write_navigation_entry(entry)?;
                },
                crate::CellStoryEvent::RevisionStart(reference) => {
                    let revision = revisions.get(reference.index).filter(|revision| {
                        revision.revision_type == RevisionType::Insertion
                            && revision.position == reference.position
                            && cell.text().get(revision.position..revision.range_end)
                                == Some(revision.content.as_ref())
                    }).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell insertion revision reference is invalid"))?;
                    self.write_revision_start(revision)?;
                },
                crate::CellStoryEvent::RevisionEnd(reference) => {
                    revisions.get(reference.index).filter(|revision| {
                        revision.revision_type == RevisionType::Insertion
                            && revision.range_end == reference.position
                    }).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell revision end reference is invalid"))?;
                    self.write_str("}")?;
                },
                crate::CellStoryEvent::RevisionDeletion(reference) => {
                    let revision = revisions.get(reference.index).filter(|revision| {
                        revision.revision_type == RevisionType::Deletion
                            && revision.position == reference.position
                    }).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF table-cell deletion revision reference is invalid"))?;
                    self.write_revision(revision)?;
                },
            }
            offset = position;
        }
        self.write_text(&cell.text()[offset..])
    }

    fn write_nested_table(
        &mut self,
        table: &Table,
        depth: usize,
        fields: &[crate::Field<'_>],
        navigation_entries: &[crate::NavigationEntry<'_>],
        revisions: &[Revision<'_>],
    ) -> io::Result<()> {
        for row in table.rows() {
            for cell in row.cells() {
                self.write_str("{")?;
                self.write_control_word("intbl", None)?;
                self.write_control_word("itap", Some(depth as i32))?;
                self.write_str(" ")?;
                self.write_cell_content(cell, depth, fields, navigation_entries, revisions)?;
                self.write_control_word("nestcell", None)?;
                self.write_str("}")?;
            }
            self.write_str("{\\*")?;
            self.write_control_word("nesttableprops", None)?;
            self.write_control_word("itap", Some(depth as i32))?;
            self.write_control_word("trowd", None)?;
            if let Some(table_style) = row.table_style() {
                self.write_control_word("ts", Some(i32::from(table_style)))?;
            }
            if let Some(table_rsid) = row.table_rsid() {
                self.write_control_word("tblrsid", Some(table_rsid as i32))?;
            }
            self.write_revision_metadata("trauth", "trdate", row.revision())?;
            if let Some(direction) = table.direction() {
                self.write_control_word(
                    "taprtl",
                    (direction == TextDirection::LeftToRight).then_some(0),
                )?;
            }
            self.write_table_row_banding(row)?;
            self.write_table_row_layout(row)?;
            self.write_table_row_geometry(row.geometry())?;
            self.write_table_row_borders(row.borders())?;
            self.write_table_shading("tr", row.shading())?;
            self.write_table_distances("trpadd", "trpaddf", row.padding())?;
            self.write_table_distances("trspd", "trspdf", row.spacing())?;
            self.write_table_row_cell_defaults(row.cell_defaults())?;
            self.write_floating_table_position(row.positioning())?;
            for (index, cell) in row.cells().iter().enumerate() {
                self.write_table_preferred_width("clftsWidth", "clwWidth", cell.preferred_width())?;
                self.write_table_cell_merge(cell)?;
                self.write_table_cell_revision(cell)?;
                self.write_table_cell_layout(cell)?;
                self.write_table_cell_borders(cell.borders())?;
                self.write_table_shading("cl", cell.shading())?;
                self.write_table_distances("clpad", "clpadf", cell.padding())?;
                self.write_table_distances("clspd", "clspdf", cell.spacing())?;
                self.write_control_word(
                    "cellx",
                    Some(cell.right_boundary().unwrap_or(2880 * ((index + 1) as i32))),
                )?;
            }
            self.write_control_word("nestrow", None)?;
            self.write_str("}")?;
            self.write_str("{")?;
            self.write_control_word("nonesttables", None)?;
            self.write_control_word("par", None)?;
            self.write_str("}")?;
        }
        Ok(())
    }

    fn write_table_preferred_width(
        &mut self,
        unit_control: &str,
        value_control: &str,
        width: Option<crate::TablePreferredWidth>,
    ) -> io::Result<()> {
        let Some(width) = width else { return Ok(()) };
        width
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(
            unit_control,
            Some(match width.unit() {
                crate::TablePreferredWidthUnit::Null => 0,
                crate::TablePreferredWidthUnit::Auto => 1,
                crate::TablePreferredWidthUnit::Percent => 2,
                crate::TablePreferredWidthUnit::Twips => 3,
            }),
        )?;
        if let Some(value) = width.value() {
            self.write_control_word(value_control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    fn write_table_invisible_width(
        &mut self,
        unit_control: &str,
        value_control: &str,
        width: Option<crate::TablePreferredWidth>,
    ) -> io::Result<()> {
        let Some(width) = width else { return Ok(()) };
        width
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(
            unit_control,
            Some(match width.unit() {
                crate::TablePreferredWidthUnit::Null => 0,
                crate::TablePreferredWidthUnit::Auto => 1,
                crate::TablePreferredWidthUnit::Percent => 2,
                crate::TablePreferredWidthUnit::Twips => 3,
            }),
        )?;
        if let Some(value) = width.value().filter(|value| *value != 0) {
            self.write_control_word(value_control, Some(i32::from(value)))?;
        }
        Ok(())
    }

    fn write_table_row_geometry(&mut self, geometry: crate::TableRowGeometry) -> io::Result<()> {
        geometry
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(value) = geometry.half_gap_twips() {
            self.write_control_word("trgaph", Some(i32::from(value)))?;
        }
        if let Some(value) = geometry.left_edge_twips() {
            self.write_control_word("trleft", Some(value))?;
        }
        match geometry.height() {
            crate::TableRowHeight::Automatic => {},
            crate::TableRowHeight::Minimum(value) => {
                self.write_control_word("trrh", Some(i32::from(value)))?
            },
            crate::TableRowHeight::Exact(value) => {
                self.write_control_word("trrh", Some(-i32::from(value)))?
            },
        }
        self.write_table_preferred_width("trftsWidth", "trwWidth", geometry.preferred_width())?;
        self.write_table_invisible_width(
            "trftsWidthB",
            "trwWidthB",
            geometry.leading_invisible_width(),
        )?;
        self.write_table_invisible_width(
            "trftsWidthA",
            "trwWidthA",
            geometry.trailing_invisible_width(),
        )?;
        if geometry.auto_fit() {
            self.write_control_word("trautofit", Some(1))?;
        }
        if let Some(indent) = geometry.indent() {
            self.write_control_word("tblind", Some(indent.value()))?;
            self.write_control_word(
                "tblindtype",
                Some(match indent.unit() {
                    crate::TableIndentUnit::Auto => 0,
                    crate::TableIndentUnit::Twips => 1,
                    crate::TableIndentUnit::Nil => 2,
                    crate::TableIndentUnit::Percent => 3,
                }),
            )?;
        }
        Ok(())
    }

    fn write_table_row_layout(&mut self, row: &Row<'_>) -> io::Result<()> {
        if let Some(alignment) = row.layout().alignment {
            self.write_control_word(
                match alignment {
                    crate::TableRowAlignment::Left => "trql",
                    crate::TableRowAlignment::Center => "trqc",
                    crate::TableRowAlignment::Right => "trqr",
                },
                None,
            )?;
        }
        if let Some(direction) = row.direction() {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrrow",
                    TextDirection::RightToLeft => "rtlrow",
                },
                None,
            )?;
        }
        if row.layout().header {
            self.write_control_word("trhdr", None)?;
        }
        if row.layout().keep_together {
            self.write_control_word("trkeep", None)?;
        }
        if row.layout().keep_with_following {
            self.write_control_word("trkeepfollow", None)?;
        }
        Ok(())
    }

    fn write_table_row_banding(&mut self, row: &Row<'_>) -> io::Result<()> {
        let banding = row.banding();
        if let Some(value) = banding.row_index {
            self.write_control_word("irow", Some(i32::from(value)))?;
        }
        if let Some(value) = banding.band_index {
            self.write_control_word(
                "irowband",
                Some(match value {
                    crate::TableRowBandIndex::Header => -1,
                    crate::TableRowBandIndex::Row(value) => i32::from(value),
                }),
            )?;
        }
        let flags = row.autoformat_flags();
        for (flag, word) in [
            (crate::TableAutoformatFlag::Border, "tbllkborder"),
            (crate::TableAutoformatFlag::Shading, "tbllkshading"),
            (crate::TableAutoformatFlag::Font, "tbllkfont"),
            (crate::TableAutoformatFlag::Color, "tbllkcolor"),
            (crate::TableAutoformatFlag::BestFit, "tbllkbestfit"),
            (crate::TableAutoformatFlag::HeaderRows, "tbllkhdrrows"),
            (crate::TableAutoformatFlag::LastRow, "tbllklastrow"),
            (crate::TableAutoformatFlag::HeaderColumns, "tbllkhdrcols"),
            (crate::TableAutoformatFlag::LastColumn, "tbllklastcol"),
            (crate::TableAutoformatFlag::NoRowBanding, "tbllknorowband"),
            (
                crate::TableAutoformatFlag::NoColumnBanding,
                "tbllknocolband",
            ),
        ] {
            if flags.contains(flag) {
                self.write_control_word(word, None)?;
            }
        }
        if banding.last_row {
            self.write_control_word("lastrow", None)?;
        }
        Ok(())
    }

    fn write_table_cell_layout(&mut self, cell: &crate::Cell<'_>) -> io::Result<()> {
        let layout = cell.layout();
        if let Some(alignment) = layout.vertical_alignment {
            self.write_control_word(
                match alignment {
                    crate::TableCellVerticalAlignment::Top => "clvertalt",
                    crate::TableCellVerticalAlignment::Center => "clvertalc",
                    crate::TableCellVerticalAlignment::Bottom => "clvertalb",
                },
                None,
            )?;
        }
        if let Some(flow) = layout.text_flow {
            self.write_control_word(
                match flow {
                    crate::TableCellTextFlow::LeftToRightTopToBottom => "cltxlrtb",
                    crate::TableCellTextFlow::RightToLeftTopToBottom => "cltxtbrl",
                    crate::TableCellTextFlow::LeftToRightBottomToTop => "cltxbtlr",
                    crate::TableCellTextFlow::LeftToRightTopToBottomVertical => "cltxlrtbv",
                    crate::TableCellTextFlow::TopToBottomRightToLeftVertical => "cltxtbrlv",
                },
                None,
            )?;
        }
        if layout.fit_text {
            self.write_control_word("clFitText", None)?;
        }
        if layout.no_wrap {
            self.write_control_word("clNoWrap", None)?;
        }
        if layout.hide_mark {
            self.write_control_word("clhidemark", None)?;
        }
        Ok(())
    }

    fn write_table_cell_merge(&mut self, cell: &crate::Cell<'_>) -> io::Result<()> {
        let merge = cell.merge();
        if let Some(role) = merge.horizontal {
            self.write_control_word(
                match role {
                    crate::TableCellMergeRole::First => "clmgf",
                    crate::TableCellMergeRole::Continuation => "clmrg",
                },
                None,
            )?;
        }
        if let Some(role) = merge.vertical {
            self.write_control_word(
                match role {
                    crate::TableCellMergeRole::First => "clvmgf",
                    crate::TableCellMergeRole::Continuation => "clvmrg",
                },
                None,
            )?;
        }
        Ok(())
    }

    fn write_table_cell_revision(&mut self, cell: &crate::Cell<'_>) -> io::Result<()> {
        if let Some(revision) = cell.revision() {
            self.write_control_word(revision.kind.control_word(), None)?;
            self.write_revision_metadata(
                revision.kind.author_control_word(),
                revision.kind.date_control_word(),
                revision.metadata,
            )?;
        }
        Ok(())
    }

    fn write_table_border(&mut self, selector: &str, border: &Border) -> io::Result<()> {
        border
            .validate_table()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word(selector, None)?;
        self.write_control_word(border.style.control_word(), None)?;
        if border.style == BorderStyle::None {
            return Ok(());
        }
        self.write_control_word("brdrw", Some(border.width))?;
        self.write_control_word("brdrcf", Some(i32::from(border.color_ref)))?;
        self.write_control_word("brsp", Some(border.space))?;
        if border.shadow {
            self.write_control_word("brdrsh", None)?;
        }
        if border.frame {
            self.write_control_word("brdrframe", None)?;
        }
        Ok(())
    }
    fn write_table_row_borders(&mut self, borders: &crate::TableRowBorders) -> io::Result<()> {
        borders
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (selector, border) in [
            ("trbrdrt", borders.top),
            ("trbrdrl", borders.left),
            ("trbrdrb", borders.bottom),
            ("trbrdrr", borders.right),
            ("trbrdrh", borders.horizontal),
            ("trbrdrv", borders.vertical),
        ] {
            if let Some(border) = border {
                self.write_table_border(selector, &border)?;
            }
        }
        Ok(())
    }
    fn write_table_cell_borders(&mut self, borders: &crate::TableCellBorders) -> io::Result<()> {
        borders
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (selector, border) in [
            ("clbrdrt", borders.top),
            ("clbrdrl", borders.left),
            ("clbrdrb", borders.bottom),
            ("clbrdrr", borders.right),
            ("cldglu", borders.upper_left_to_lower_right),
            ("cldgll", borders.upper_right_to_lower_left),
        ] {
            if let Some(border) = border {
                self.write_table_border(selector, &border)?;
            }
        }
        Ok(())
    }
    fn write_table_row_cell_defaults(
        &mut self,
        defaults: &crate::TableRowCellDefaults,
    ) -> io::Result<()> {
        defaults
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        let borders = &defaults.borders;
        for (selector, border) in [
            ("tsbrdrt", borders.top),
            ("tsbrdrl", borders.left),
            ("tsbrdrb", borders.bottom),
            ("tsbrdrr", borders.right),
            ("tsbrdrh", borders.horizontal_inside),
            ("tsbrdrv", borders.vertical_inside),
            ("tsbrdrdgl", borders.diagonal_upper_left_to_lower_right),
            ("tsbrdrdg", borders.diagonal_upper_right_to_lower_left),
        ] {
            if let Some(border) = border {
                self.write_table_border(selector, &border)?;
            }
        }
        self.write_table_distances("tscellpadd", "tscellpaddf", &defaults.padding)?;
        self.write_table_distances("tscellspc", "tscellspcf", &defaults.spacing)?;
        self.write_table_preferred_width(
            "tscellwidthfts",
            "tscellwidth",
            defaults.preferred_cell_width,
        )?;
        Ok(())
    }

    fn write_table_shading(
        &mut self,
        prefix: &str,
        shading: crate::TableShading,
    ) -> io::Result<()> {
        shading
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(index) = shading.pattern_index {
            if prefix != "tr" {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "trpat is only valid for row shading",
                ));
            }
            self.write_control_word("trpat", Some(i32::from(index)))?;
        }
        if let Some(pattern) = shading.pattern {
            let suffix = match pattern {
                crate::ShadingPattern::Horizontal => "bghoriz",
                crate::ShadingPattern::Vertical => "bgvert",
                crate::ShadingPattern::ForwardDiagonal => "bgfdiag",
                crate::ShadingPattern::BackwardDiagonal => "bgbdiag",
                crate::ShadingPattern::Cross => "bgcross",
                crate::ShadingPattern::DiagonalCross => "bgdcross",
                crate::ShadingPattern::DarkHorizontal => "bgdkhor",
                crate::ShadingPattern::DarkVertical => "bgdkvert",
                crate::ShadingPattern::DarkForwardDiagonal => "bgdkfdiag",
                crate::ShadingPattern::DarkBackwardDiagonal => "bgdkbdiag",
                crate::ShadingPattern::DarkCross => "bgdkcross",
                crate::ShadingPattern::DarkDiagonalCross => "bgdkdcross",
                _ => {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "invalid explicit RTF table shading pattern",
                    ));
                },
            };
            self.write_control_word(&format!("{prefix}{suffix}"), None)?;
        }
        if let Some(color) = shading.foreground_color {
            self.write_control_word(&format!("{prefix}cfpat"), Some(i32::from(color)))?;
        }
        if let Some(color) = shading.background_color {
            self.write_control_word(&format!("{prefix}cbpat"), Some(i32::from(color)))?;
        }
        if let Some(amount) = shading.amount {
            self.write_control_word(&format!("{prefix}shdng"), Some(i32::from(amount)))?;
        }
        Ok(())
    }

    fn write_table_distances(
        &mut self,
        value_prefix: &str,
        unit_prefix: &str,
        distances: &TableEdgeDistances,
    ) -> io::Result<()> {
        distances
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        for (suffix, side) in [
            ("l", distances.left),
            ("r", distances.right),
            ("t", distances.top),
            ("b", distances.bottom),
        ] {
            if let Some(value) = side.value {
                self.write_control_word(
                    &format!("{value_prefix}{suffix}"),
                    Some(i32::from(value)),
                )?;
            }
            if let Some(unit) = side.unit {
                self.write_control_word(
                    &format!("{unit_prefix}{suffix}"),
                    Some(match unit {
                        TableDistanceUnit::Null => 0,
                        TableDistanceUnit::Twips => 3,
                    }),
                )?;
            }
        }
        Ok(())
    }

    fn write_floating_table_position(
        &mut self,
        position: &crate::FloatingTablePosition,
    ) -> io::Result<()> {
        position
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(reference) = position.horizontal_reference {
            self.write_control_word(
                match reference {
                    crate::TableHorizontalReference::Column => "tphcol",
                    crate::TableHorizontalReference::Margin => "tphmrg",
                    crate::TableHorizontalReference::Page => "tphpg",
                },
                None,
            )?
        }
        if let Some(value) = position.horizontal_position {
            let (word, param) = match value {
                crate::TableHorizontalPosition::Offset(value) => ("tposx", Some(value)),
                crate::TableHorizontalPosition::NegativeOffset(value) => ("tposnegx", Some(value)),
                crate::TableHorizontalPosition::Center => ("tposxc", None),
                crate::TableHorizontalPosition::Inside => ("tposxi", None),
                crate::TableHorizontalPosition::Left => ("tposxl", None),
                crate::TableHorizontalPosition::Outside => ("tposxo", None),
                crate::TableHorizontalPosition::Right => ("tposxr", None),
            };
            self.write_control_word(word, param)?
        }
        if let Some(reference) = position.vertical_reference {
            self.write_control_word(
                match reference {
                    crate::TableVerticalReference::Margin => "tpvmrg",
                    crate::TableVerticalReference::Paragraph => "tpvpara",
                    crate::TableVerticalReference::Page => "tpvpg",
                },
                None,
            )?
        }
        if let Some(value) = position.vertical_position {
            let (word, param) = match value {
                crate::TableVerticalPosition::Offset(value) => ("tposy", Some(value)),
                crate::TableVerticalPosition::NegativeOffset(value) => ("tposnegy", Some(value)),
                crate::TableVerticalPosition::Bottom => ("tposyb", None),
                crate::TableVerticalPosition::Center => ("tposyc", None),
                crate::TableVerticalPosition::Inline => ("tposyil", None),
                crate::TableVerticalPosition::Inside => ("tposyin", None),
                crate::TableVerticalPosition::Outside => ("tposyout", None),
                crate::TableVerticalPosition::Top => ("tposyt", None),
            };
            self.write_control_word(word, param)?
        }
        for (word, value) in [
            ("tdfrmtxtLeft", position.wrap_distances.left),
            ("tdfrmtxtRight", position.wrap_distances.right),
            ("tdfrmtxtTop", position.wrap_distances.top),
            ("tdfrmtxtBottom", position.wrap_distances.bottom),
        ] {
            if let Some(value) = value {
                self.write_control_word(word, Some(i32::from(value)))?
            }
        }
        if position.no_overlap {
            self.write_control_word("tabsnoovrlp", None)?
        }
        Ok(())
    }

    /// Write a control word
    pub fn write_control_word(&mut self, word: &str, param: Option<i32>) -> io::Result<()> {
        self.write_str("\\")?;
        self.write_str(word)?;
        if let Some(p) = param {
            write!(self.writer, "{}", p)?;
        }
        Ok(())
    }

    /// Write plain text (with proper escaping)
    pub fn write_text(&mut self, text: &str) -> io::Result<()> {
        for ch in text.chars() {
            match ch {
                '\\' => self.write_str("\\\\")?,
                '{' => self.write_str("\\{")?,
                '}' => self.write_str("\\}")?,
                // The trailing space delimits the control word. Without it the
                // following character is absorbed into the word itself (`\partwo`)
                // or misread as its numeric parameter (`\par2`), silently
                // destroying the text that follows the break. RTF always consumes
                // a single delimiting space, so it never reappears as content.
                '\n' => self.write_str("\\par ")?,
                '\t' => self.write_str("\\tab ")?,
                // RTF special characters with dedicated control words keep their
                // source spelling instead of a generic \u escape. The trailing
                // space delimits the control word exactly like \par and \tab.
                '\u{2014}' => self.write_str("\\emdash ")?,
                '\u{2013}' => self.write_str("\\endash ")?,
                '\u{2003}' => self.write_str("\\emspace ")?,
                '\u{2002}' => self.write_str("\\enspace ")?,
                '\u{2005}' => self.write_str("\\qmspace ")?,
                '\u{2022}' => self.write_str("\\bullet ")?,
                '\u{200E}' => self.write_str("\\ltrmark ")?,
                '\u{200F}' => self.write_str("\\rtlmark ")?,
                '\u{200D}' => self.write_str("\\zwj ")?,
                '\u{200C}' => self.write_str("\\zwnj ")?,
                '\u{200B}' => self.write_str("\\zwbo ")?,
                '\u{FEFF}' => self.write_str("\\zwnbo ")?,
                // Readers discard raw carriage returns and other bare control
                // bytes as line-ending noise, so emit them as hex escapes to keep
                // them part of the document text.
                c if (c as u32) < ASCII_CONTROL_LIMIT => {
                    write!(self.writer, "\\'{:02x}", c as u32)?;
                },
                c if c.is_ascii() => {
                    write!(self.writer, "{}", c)?;
                },
                c => {
                    // Write Unicode character
                    let code = c as i32;
                    self.write_control_word("u", Some(code))?;
                    // Fallback character
                    self.write_str("?")?;
                },
            }
        }
        Ok(())
    }

    /// Write a string
    pub fn write_str(&mut self, s: &str) -> io::Result<()> {
        self.writer.write_all(s.as_bytes())
    }

    /// Flush the writer
    pub fn flush(&mut self) -> io::Result<()> {
        self.writer.flush()
    }

    /// Write a header or footer
    pub fn write_header_footer(&mut self, hf: &HeaderFooter) -> io::Result<()> {
        self.write_header_footer_with_fields(hf, &[])
    }

    fn write_header_footer_with_fields(
        &mut self,
        hf: &HeaderFooter,
        fields: &[crate::Field<'_>],
    ) -> io::Result<()> {
        self.write_str("{")?;

        // Write header/footer type control word
        match hf.header_type {
            HeaderFooterType::Header => self.write_control_word("header", None)?,
            HeaderFooterType::HeaderFirst => self.write_control_word("headerf", None)?,
            HeaderFooterType::HeaderLeft => self.write_control_word("headerl", None)?,
            HeaderFooterType::HeaderRight => self.write_control_word("headerr", None)?,
            HeaderFooterType::Footer => self.write_control_word("footer", None)?,
            HeaderFooterType::FooterFirst => self.write_control_word("footerf", None)?,
            HeaderFooterType::FooterLeft => self.write_control_word("footerl", None)?,
            HeaderFooterType::FooterRight => self.write_control_word("footerr", None)?,
        }

        enum EventKind<'b, 'a> {
            Shape(&'b crate::Shape<'a>),
            Group(&'b crate::ShapeGroup<'a>),
            Field(&'b crate::Field<'a>),
            PageBreak,
        }
        struct Event<'b, 'a> {
            offset: usize,
            kind: EventKind<'b, 'a>,
        }
        let story = hf.text();
        crate::field::validate_story_events(
            &story,
            &hf.shapes,
            &hf.shape_groups,
            &hf.drawing_order,
            &hf.story_events,
            "header/footer",
        )
        .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        let owner = match hf.header_type {
            HeaderFooterType::Header
            | HeaderFooterType::HeaderFirst
            | HeaderFooterType::HeaderLeft
            | HeaderFooterType::HeaderRight => crate::FieldOwner::Header,
            HeaderFooterType::Footer
            | HeaderFooterType::FooterFirst
            | HeaderFooterType::FooterLeft
            | HeaderFooterType::FooterRight => crate::FieldOwner::Footer,
        };
        let mut events = Vec::with_capacity(hf.story_events.len());
        for story_event in &hf.story_events {
            match *story_event {
                crate::StoryEvent::Drawing(crate::StoryDrawing::Shape(index)) => events.push(Event {
                    offset: hf.shapes[index].position,
                    kind: EventKind::Shape(&hf.shapes[index]),
                }),
                crate::StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(index)) => events.push(Event {
                    offset: hf.shape_groups[index].position,
                    kind: EventKind::Group(&hf.shape_groups[index]),
                }),
                crate::StoryEvent::Field(reference) => events.push(Event {
                    offset: reference.position,
                    kind: EventKind::Field(fields.get(reference.field_index).filter(|field| field.owner == owner && field.position == reference.position && field.range_end == reference.position).ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "RTF header/footer story has an invalid generic-field owner or reference"))?),
                }),
                crate::StoryEvent::PageBreak(page_break) => events.push(Event {
                    offset: page_break.position,
                    kind: EventKind::PageBreak,
                }),
            }
        }
        /* Exact event order is taken from the concrete header/footer story. */
        for event in &events {
            match event.kind {
                EventKind::Shape(_)
                | EventKind::Group(_)
                | EventKind::Field(_)
                | EventKind::PageBreak => {},
            }
        }
        /* Keep paragraph formatting while splitting its text around story events. */
        let mut next_event = 0usize;
        let mut story_offset = 0usize;

        for para in &hf.paragraphs {
            self.write_formatting(&para.formatting)?;
            self.write_paragraph_properties(&para.paragraph)?;
            self.write_str(" ")?;
            let text = para.text.as_ref();
            let end = story_offset.saturating_add(text.len());
            let mut local = 0usize;
            while let Some(event) = events.get(next_event).filter(|event| event.offset <= end) {
                let split = event.offset.checked_sub(story_offset).ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF header/footer events are out of story order",
                    )
                })?;
                self.write_text(&text[local..split])?;
                match event.kind {
                    EventKind::Shape(shape) => self.write_root_shape(shape)?,
                    EventKind::Group(group) => self.write_shape_group(group, true)?,
                    EventKind::Field(field) => self.write_field_with_fields(field, fields, 0)?,
                    EventKind::PageBreak => self.write_str("\\page ")?,
                }
                local = split;
                next_event += 1;
            }
            self.write_text(&text[local..])?;
            self.write_control_word("par", None)?;
            story_offset = end.saturating_add(1);
        }
        while let Some(event) = events.get(next_event) {
            if event.offset != story.len() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF header/footer event position is unreachable",
                ));
            }
            match event.kind {
                EventKind::Shape(shape) => self.write_root_shape(shape)?,
                EventKind::Group(group) => self.write_shape_group(group, true)?,
                EventKind::Field(field) => self.write_field_with_fields(field, fields, 0)?,
                EventKind::PageBreak => self.write_str("\\page ")?,
            }
            next_event += 1;
        }

        self.write_str("}")?;
        Ok(())
    }

    /*
            match *drawing {
                crate::StoryDrawing::Shape(index) => events.push(Event {
                    offset: hf.shapes[index].position,
                    drawing: Drawing::Shape(&hf.shapes[index]),
                }),
                crate::StoryDrawing::ShapeGroup(index) => events.push(Event {
                    offset: hf.shape_groups[index].position,
                    drawing: Drawing::Group(&hf.shape_groups[index]),
                }),
            }
        }
        let mut next_event = 0usize;
        let mut story_offset = 0usize;

        // Write paragraphs and merge story-owned drawings at UTF-8 boundaries.
        for para in &hf.paragraphs {
            self.write_formatting(&para.formatting)?;
            self.write_paragraph_properties(&para.paragraph)?;
            self.write_str(" ")?;
            let text = para.text.as_ref();
            let end = story_offset.saturating_add(text.len());
            let mut local = 0usize;
            while let Some(event) = events.get(next_event).filter(|event| event.offset <= end) {
                let split = event.offset.checked_sub(story_offset).ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF header/footer drawings are out of story order")
                })?;
                self.write_text(&text[local..split])?;
                match event.drawing {
                    Drawing::Shape(shape) => self.write_root_shape(shape)?,
                    Drawing::Group(group) => self.write_shape_group(group, true)?,
                }
                local = split;
                next_event += 1;
            }
            self.write_text(&text[local..])?;
            self.write_control_word("par", None)?;
            story_offset = end.saturating_add(1);
        }
        while let Some(event) = events.get(next_event) {
            if event.offset != story.len() {
                return Err(io::Error::new(io::ErrorKind::InvalidInput, "RTF header/footer drawing position is unreachable"));
            }
            match event.drawing {
                Drawing::Shape(shape) => self.write_root_shape(shape)?,
                Drawing::Group(group) => self.write_shape_group(group, true)?,
            }
            next_event += 1;
        }

        self.write_str("}")?;
        Ok(())
    } */

    /// Write a footnote or endnote
    pub fn write_note(&mut self, note: &Note) -> io::Result<()> {
        self.write_note_with_fields(note, &[])
    }

    fn write_note_with_fields(
        &mut self,
        note: &Note,
        fields: &[crate::Field<'_>],
    ) -> io::Result<()> {
        note.validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{")?;

        // Write note type control word
        if note.is_footnote {
            self.write_control_word("footnote", None)?;
        } else {
            self.write_control_word("endnote", None)?;
        }

        // Write reference number/marker
        if !note.reference.is_empty()
            && let Ok(num) = note.reference.parse::<i32>()
        {
            self.write_control_word("chftn", Some(num))?;
        }

        // Write note content
        self.write_str(" {")?;
        self.write_formatting(&note.formatting)?;
        self.write_field_story(
            note.content.as_ref(),
            &note.shapes,
            &note.shape_groups,
            &note.drawing_order,
            &note.story_events,
            fields,
            if note.is_footnote {
                crate::FieldOwner::Footnote
            } else {
                crate::FieldOwner::Endnote
            },
            DrawingStoryTextMode::Note,
            0,
        )?;
        self.write_str("}")?;

        self.write_str("}")?;
        Ok(())
    }

    /// Write a hyperlink field
    pub fn write_hyperlink(&mut self, url: &str, display_text: &str) -> io::Result<()> {
        let instruction = format!("HYPERLINK {}", crate::field::quoted_field_operand(url));
        self.write_hyperlink_instruction(&instruction, display_text)
    }

    /// Write an internal bookmark hyperlink without exposing raw field syntax.
    pub fn write_internal_hyperlink(
        &mut self,
        bookmark: &str,
        display_text: &str,
    ) -> io::Result<()> {
        let instruction = format!(
            "HYPERLINK \\l {}",
            crate::field::quoted_field_operand(bookmark)
        );
        self.write_hyperlink_instruction(&instruction, display_text)
    }

    fn write_hyperlink_instruction(
        &mut self,
        instruction: &str,
        display_text: &str,
    ) -> io::Result<()> {
        self.write_str("{\\field")?;
        // Field instruction
        self.write_str("{\\*\\fldinst{")?;
        self.write_text(instruction)?;
        self.write_str("}}")?;

        // Field result (display text)
        self.write_str("{\\fldrslt{")?;
        self.write_control_word("ul", None)?; // Underline hyperlinks by default
        self.write_control_word("cf", Some(1))?; // Blue color for hyperlinks
        self.write_text(display_text)?;
        self.write_str("}}}")?;

        Ok(())
    }

    /// Write a field (generic)
    pub fn write_field(&mut self, field: &Field) -> io::Result<()> {
        self.write_field_with_fields(field, &[], 0)
    }

    /// Write a caller-provided legacy `EQ` expression as an inert RTF field.
    ///
    /// The expression is escaped for the field instruction and emitted with
    /// the empty cached-result group conventionally used for `EQ`. It is never
    /// parsed, calculated, formatted, or rendered by this library.
    pub fn write_equation(&mut self, expression: &str) -> io::Result<()> {
        let field = Field::new_equation(expression)
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_field(&field)
    }

    fn write_field_with_fields(
        &mut self,
        field: &Field,
        fields: &[crate::Field<'_>],
        depth: usize,
    ) -> io::Result<()> {
        if depth > 64 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF nested generic fields exceed 64 levels",
            ));
        }
        field
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\field")?;
        if field.status.dirty {
            self.write_str("\\flddirty")?;
        }
        if field.status.edited {
            self.write_str("\\fldedit")?;
        }
        if field.status.locked {
            self.write_str("\\fldlock")?;
        }
        if field.status.private {
            self.write_str("\\fldpriv")?;
        }

        // Field instruction
        self.write_str("{\\*\\fldinst{")?;
        self.write_text(field.instruction.as_ref())?;
        self.write_str("}}")?;

        // Field result
        if field.field_type == FieldType::Equation
            && field.result.is_empty()
            && field.result_events.is_empty()
        {
            // RTF 1.9.1 examples write a null fldrslt group for EQ fields.
            self.write_str("{\\fldrslt}")?;
        } else if !field.result.is_empty() || !field.result_events.is_empty() {
            self.write_str("{\\fldrslt{")?;
            self.write_field_story(
                field.result.as_ref(),
                &field.shapes,
                &field.shape_groups,
                &field.drawing_order,
                &field.result_events,
                fields,
                crate::FieldOwner::FieldResult,
                DrawingStoryTextMode::ShapeText,
                depth,
            )?;
            self.write_str("}}")?;
        }

        self.write_str("}")?;
        Ok(())
    }

    /// Write a revision mark (track changes)
    pub fn write_revision(&mut self, revision: &Revision) -> io::Result<()> {
        revision
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_revision_start(revision)?;
        self.write_text(revision.content.as_ref())?;
        self.write_str("}")?;
        Ok(())
    }

    fn write_revision_start(&mut self, revision: &Revision<'_>) -> io::Result<()> {
        self.write_str("{")?;
        let (kind, author, date) = match revision.revision_type {
            RevisionType::Insertion => ("revised", "revauth", "revdttm"),
            RevisionType::Deletion => ("deleted", "revauthdel", "revdttmdel"),
            RevisionType::FormatChange | RevisionType::MovedFrom | RevisionType::MovedTo => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "this RTF revision kind has no lossless scoped-run representation",
                ));
            },
        };
        self.write_control_word(kind, None)?;
        self.write_control_word(author, Some(revision.id))?;
        if let Some(date_value) = revision.date.as_deref() {
            let packed = date_value.parse::<i32>().map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision dates must contain the packed signed DTTM value",
                )
            })?;
            self.write_control_word(date, Some(packed))?;
        }
        self.write_str(" ")?;
        Ok(())
    }

    /// Write a section with headers and footers
    pub fn write_section(&mut self, section: &Section) -> io::Result<()> {
        self.write_section_with_fields(section, &[])
    }

    fn write_section_with_fields(
        &mut self,
        section: &Section,
        fields: &[crate::Field<'_>],
    ) -> io::Result<()> {
        // Write section properties
        self.write_control_word("sectd", None)?;
        if let Some(section_style) = section.properties.section_style {
            self.write_control_word("ds", Some(i32::from(section_style)))?;
        }
        if let Some(section_rsid) = section.properties.section_rsid {
            self.write_control_word("sectrsid", Some(section_rsid as i32))?;
        }
        self.write_revision_metadata(
            "srauth",
            "srdate",
            section.properties.revision,
        )?;
        if section.properties.title_page {
            self.write_control_word("titlepg", None)?;
        }
        self.write_section_note_options(&section.properties.note_options)?;
        self.write_page_borders(&section.properties.page_borders)?;

        if let Some(direction) = section.properties.direction {
            self.write_control_word(
                match direction {
                    TextDirection::LeftToRight => "ltrsect",
                    TextDirection::RightToLeft => "rtlsect",
                },
                None,
            )?;
        }

        match section.properties.break_type {
            SectionBreakType::Continuous => self.write_control_word("sbknone", None)?,
            SectionBreakType::Column => self.write_control_word("sbkcol", None)?,
            SectionBreakType::Page => self.write_control_word("sbkpage", None)?,
            SectionBreakType::EvenPage => self.write_control_word("sbkeven", None)?,
            SectionBreakType::OddPage => self.write_control_word("sbkodd", None)?,
        }

        // Page size
        self.write_control_word("pgwsxn", Some(section.properties.page_width))?;
        self.write_control_word("pghsxn", Some(section.properties.page_height))?;

        // Margins
        self.write_control_word("marglsxn", Some(section.properties.margin_left))?;
        self.write_control_word("margrsxn", Some(section.properties.margin_right))?;
        self.write_control_word("margtsxn", Some(section.properties.margin_top))?;
        self.write_control_word("margbsxn", Some(section.properties.margin_bottom))?;
        self.write_control_word("guttersxn", Some(section.properties.margin_gutter))?;

        // Paper-source bins
        if let Some(first) = section.properties.paper_source.first {
            self.write_control_word("binfsxn", Some(i32::from(first)))?;
        }
        if let Some(other) = section.properties.paper_source.other {
            self.write_control_word("binsxn", Some(i32::from(other)))?;
        }

        // Header/footer distance
        self.write_control_word("headery", Some(section.properties.header_distance))?;
        self.write_control_word("footery", Some(section.properties.footer_distance))?;

        if section.properties.orientation == PageOrientation::Landscape {
            self.write_control_word("lndscpsxn", None)?;
        }
        if let Some(rendering) = section.properties.rendering {
            self.write_control_word(
                match rendering {
                    crate::SectionRendering::Horizontal => "horzsect",
                    crate::SectionRendering::Vertical => "vertsect",
                },
                None,
            )?;
        }
        section
            .properties
            .columns
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_control_word("cols", Some(i32::from(section.properties.columns.count)))?;
        if !section.properties.balance_columns {
            self.write_control_word("nocolbal", None)?;
        }
        if section.properties.columns.separator {
            self.write_control_word("linebetcol", None)?;
        }
        self.write_control_word("colsx", Some(section.properties.columns.default_spacing))?;
        for (index, column) in section.properties.columns.explicit.iter().enumerate() {
            self.write_control_word("colno", Some((index + 1) as i32))?;
            self.write_control_word("colw", Some(column.width))?;
            if let Some(space) = column.space_after {
                self.write_control_word("colsr", Some(space))?;
            }
        }
        self.write_control_word("pgnstarts", Some(section.properties.page_number_start))?;
        self.write_control_word(
            section.properties.page_number_format.control_word(),
            None,
        )?;
        if let Some(restart) = section.properties.page_number_restart {
            self.write_control_word(restart.control_word(), None)?;
        }
        if let Some(offset_x) = section.properties.page_number_offset_x {
            self.write_control_word("pgnx", Some(offset_x))?;
        }
        if let Some(offset_y) = section.properties.page_number_offset_y {
            self.write_control_word("pgny", Some(offset_y))?;
        }
        section
            .properties
            .page_number_heading
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(level) = section.properties.page_number_heading.level {
            self.write_control_word("pgnhn", Some(i32::from(level)))?;
        }
        if let Some(separator) = section.properties.page_number_heading.separator {
            self.write_control_word(separator.control_word(), None)?;
        }
        section
            .properties
            .document_grid
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(line_grid) = section.properties.document_grid.line_grid {
            self.write_control_word("sectlinegrid", Some(line_grid))?;
        }
        if let Some(grid_type) = section.properties.document_grid.grid_type {
            self.write_control_word(grid_type.control_word(), None)?;
        }
        self.write_control_word(
            match section.properties.vertical_alignment {
                VerticalAlignment::Top => "vertalt",
                VerticalAlignment::Center => "vertalc",
                VerticalAlignment::Justify => "vertalj",
                VerticalAlignment::Bottom => "vertalb",
            },
            None,
        )?;
        section
            .properties
            .line_numbering
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if let Some(increment) = section.properties.line_numbering.increment {
            self.write_control_word("linemod", Some(i32::from(increment)))?;
        }
        if let Some(distance) = section.properties.line_numbering.distance {
            self.write_control_word("linex", Some(distance))?;
        }
        if let Some(start) = section.properties.line_numbering.start {
            self.write_control_word("linestarts", Some(start as i32))?;
        }
        if let Some(restart) = section.properties.line_numbering.restart {
            self.write_control_word(
                match restart {
                    crate::SectionLineNumberRestart::Section => "linerestart",
                    crate::SectionLineNumberRestart::Page => "lineppage",
                    crate::SectionLineNumberRestart::Continuous => "linecont",
                },
                None,
            )?;
        }

        // Write all headers and footers for this section
        for hf in &section.headers_footers {
            self.write_header_footer_with_fields(hf, fields)?;
        }

        Ok(())
    }

    /// Write canonical section page-border options and edges.
    pub fn write_page_borders(&mut self, borders: &crate::PageBorders) -> io::Result<()> {
        borders
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if borders.is_empty() {
            return Ok(());
        }
        if borders.option_value() != 0 {
            self.write_control_word("pgbrdropt", Some(borders.option_value()))?;
        }
        if borders.surround_header {
            self.write_control_word("pgbrdrhead", None)?;
        }
        if borders.surround_footer {
            self.write_control_word("pgbrdrfoot", None)?;
        }
        if borders.snap_to_text_borders {
            self.write_control_word("pgbrdrsnap", None)?;
        }
        for (control, border) in [
            ("pgbrdrt", borders.top),
            ("pgbrdrl", borders.left),
            ("pgbrdrb", borders.bottom),
            ("pgbrdrr", borders.right),
        ] {
            let Some(border) = border else {
                continue;
            };
            self.write_control_word(control, None)?;
            if let Some(art) = border.art {
                self.write_control_word("brdrart", Some(i32::from(art)))?;
            } else {
                self.write_control_word(border.style.control_word(), None)?;
            }
            self.write_control_word("brdrw", Some(i32::from(border.width)))?;
            self.write_control_word("brdrcf", Some(i32::from(border.color_ref)))?;
            self.write_control_word("brsp", Some(i32::from(border.space)))?;
            if border.shadow {
                self.write_control_word("brdrsh", None)?;
            }
            if border.frame {
                self.write_control_word("brdrframe", None)?;
            }
        }
        Ok(())
    }

    /// Write explicit section-level footnote and endnote overrides.
    pub fn write_section_note_options(
        &mut self,
        options: &crate::SectionNoteOptions,
    ) -> io::Result<()> {
        options
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if options.endnote_here {
            self.write_control_word("endnhere", None)?;
        }
        if let Some(value) = options.footnote_placement {
            self.write_control_word(
                match value {
                    crate::SectionFootnotePlacement::BeneathText => "sftntj",
                    crate::SectionFootnotePlacement::BottomOfPage => "sftnbj",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_start {
            self.write_control_word("sftnstart", Some(value))?;
        }
        if let Some(value) = options.footnote_restart {
            self.write_control_word(
                match value {
                    crate::FootnoteRestart::Continuous => "sftnrstcont",
                    crate::FootnoteRestart::EachSection => "sftnrestart",
                    crate::FootnoteRestart::EachPage => "sftnrstpg",
                },
                None,
            )?;
        }
        if let Some(value) = options.footnote_numbering {
            self.write_control_word(Self::section_note_numbering_control(value, false), None)?;
        }
        if let Some(value) = options.endnote_start {
            self.write_control_word("saftnstart", Some(value))?;
        }
        if let Some(value) = options.endnote_restart {
            self.write_control_word(
                match value {
                    crate::EndnoteRestart::Continuous => "saftnrstcont",
                    crate::EndnoteRestart::EachSection => "saftnrestart",
                },
                None,
            )?;
        }
        if let Some(value) = options.endnote_numbering {
            self.write_control_word(Self::section_note_numbering_control(value, true), None)?;
        }
        Ok(())
    }

    fn section_note_numbering_control(
        style: crate::NoteNumberingStyle,
        endnote: bool,
    ) -> &'static str {
        match (endnote, style) {
            (false, crate::NoteNumberingStyle::Arabic) => "sftnnar",
            (false, crate::NoteNumberingStyle::LowercaseLetter) => "sftnnalc",
            (false, crate::NoteNumberingStyle::UppercaseLetter) => "sftnnauc",
            (false, crate::NoteNumberingStyle::LowercaseRoman) => "sftnnrlc",
            (false, crate::NoteNumberingStyle::UppercaseRoman) => "sftnnruc",
            (false, crate::NoteNumberingStyle::Chicago) => "sftnnchi",
            (false, crate::NoteNumberingStyle::KoreanChosung) => "sftnnchosung",
            (false, crate::NoteNumberingStyle::Circle) => "sftnncnum",
            (false, crate::NoteNumberingStyle::KanjiDigitless) => "sftnndbnum",
            (false, crate::NoteNumberingStyle::KanjiWithDigit) => "sftnndbnumd",
            (false, crate::NoteNumberingStyle::KanjiThree) => "sftnndbnumt",
            (false, crate::NoteNumberingStyle::KanjiFour) => "sftnndbnumk",
            (false, crate::NoteNumberingStyle::DoubleByte) => "sftnndbar",
            (false, crate::NoteNumberingStyle::KoreanGanada) => "sftnnganada",
            (false, crate::NoteNumberingStyle::ChineseOne) => "sftnngbnum",
            (false, crate::NoteNumberingStyle::ChineseTwo) => "sftnngbnumd",
            (false, crate::NoteNumberingStyle::ChineseThree) => "sftnngbnuml",
            (false, crate::NoteNumberingStyle::ChineseFour) => "sftnngbnumk",
            (false, crate::NoteNumberingStyle::ZodiacOne) => "sftnnzodiac",
            (false, crate::NoteNumberingStyle::ZodiacTwo) => "sftnnzodiacd",
            (false, crate::NoteNumberingStyle::ZodiacThree) => "sftnnzodiacl",
            (true, crate::NoteNumberingStyle::Arabic) => "saftnnar",
            (true, crate::NoteNumberingStyle::LowercaseLetter) => "saftnnalc",
            (true, crate::NoteNumberingStyle::UppercaseLetter) => "saftnnauc",
            (true, crate::NoteNumberingStyle::LowercaseRoman) => "saftnnrlc",
            (true, crate::NoteNumberingStyle::UppercaseRoman) => "saftnnruc",
            (true, crate::NoteNumberingStyle::Chicago) => "saftnnchi",
            (true, crate::NoteNumberingStyle::KoreanChosung) => "saftnnchosung",
            (true, crate::NoteNumberingStyle::Circle) => "saftnncnum",
            (true, crate::NoteNumberingStyle::KanjiDigitless) => "saftnndbnum",
            (true, crate::NoteNumberingStyle::KanjiWithDigit) => "saftnndbnumd",
            (true, crate::NoteNumberingStyle::KanjiThree) => "saftnndbnumt",
            (true, crate::NoteNumberingStyle::KanjiFour) => "saftnndbnumk",
            (true, crate::NoteNumberingStyle::DoubleByte) => "saftnndbar",
            (true, crate::NoteNumberingStyle::KoreanGanada) => "saftnnganada",
            (true, crate::NoteNumberingStyle::ChineseOne) => "saftnngbnum",
            (true, crate::NoteNumberingStyle::ChineseTwo) => "saftnngbnumd",
            (true, crate::NoteNumberingStyle::ChineseThree) => "saftnngbnuml",
            (true, crate::NoteNumberingStyle::ChineseFour) => "saftnngbnumk",
            (true, crate::NoteNumberingStyle::ZodiacOne) => "saftnnzodiac",
            (true, crate::NoteNumberingStyle::ZodiacTwo) => "saftnnzodiacd",
            (true, crate::NoteNumberingStyle::ZodiacThree) => "saftnnzodiacl",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_document() {
        let mut output = Vec::new();
        let mut writer = RtfWriter::new(&mut output);

        writer.write_document_header().unwrap();
        writer.write_text("Hello World").unwrap();
        writer.write_str("}").unwrap();

        let result = String::from_utf8(output).unwrap();
        assert!(result.contains("rtf1"));
        assert!(result.contains("Hello World"));
    }

    #[test]
    fn test_control_words() {
        let mut output = Vec::new();
        let mut writer = RtfWriter::new(&mut output);

        writer.write_control_word("test", Some(42)).unwrap();
        writer.write_control_word("flag", None).unwrap();

        let result = String::from_utf8(output).unwrap();
        assert_eq!(result, "\\test42\\flag");
    }

    #[test]
    fn equation_writer_uses_an_empty_cached_result_without_evaluation() {
        let mut output = Vec::new();
        let mut writer = RtfWriter::new(&mut output);
        writer.write_document_header().unwrap();
        writer.write_equation(r"\f(1,2)").unwrap();
        writer.write_str("}").unwrap();

        let serialized = String::from_utf8(output).unwrap();
        assert!(serialized.contains(r"EQ \\f(1,2)"));
        assert!(serialized.contains(r"{\fldrslt}"));

        let document = RtfDocument::parse(&serialized).unwrap();
        let equations = document.equations();
        assert_eq!(equations.len(), 1);
        assert_eq!(equations[0].expression(), r"\f(1,2)");
        assert_eq!(equations[0].cached_result(), None);
    }

    #[test]
    fn document_writer_round_trips_caller_authored_eq_fields() {
        let mut document = RtfDocument::parse(r"{\rtf1\ansi BeforeAfter}").unwrap();
        let mut equation = Field::new_equation(r"\f(1,2)").unwrap();
        equation.owner = FieldOwner::Body;
        equation.position = "Before".len();
        equation.range_end = equation.position;
        document.push_field(equation).unwrap();

        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();
        let serialized = String::from_utf8(output).unwrap();
        assert!(serialized.contains(r"{\fldrslt}"));

        let reparsed = RtfDocument::parse(&serialized).unwrap();
        assert_eq!(reparsed.text(), "BeforeAfter");
        assert_eq!(reparsed.equation_count(), 1);
        assert_eq!(reparsed.equations()[0].expression(), r"\f(1,2)");
    }

    #[test]
    fn document_info_writer_round_trips() {
        let mut info = DocumentInfo::new().with_title(std::borrow::Cow::Borrowed("Résumé 你"));
        info.author = Some(std::borrow::Cow::Borrowed("Ada"));
        info.creation_time = Some(std::borrow::Cow::Borrowed("2026-07-15T12:34:56"));
        info.pages = Some(3);
        info.characters_with_spaces = Some(42);

        let mut output = Vec::new();
        let mut writer = RtfWriter::new(&mut output);
        writer.write_document_header().unwrap();
        writer.write_document_info(&info).unwrap();
        writer.write_text("Body").unwrap();
        writer.write_str("}").unwrap();

        let rtf = String::from_utf8(output).unwrap();
        let parsed = RtfDocument::parse(&rtf).unwrap();
        assert_eq!(parsed.info().title.as_deref(), Some("Résumé 你"));
        assert_eq!(parsed.info().author.as_deref(), Some("Ada"));
        assert_eq!(
            parsed.info().creation_time.as_deref(),
            Some("2026-07-15T12:34:56")
        );
        assert_eq!(parsed.info().pages, Some(3));
        assert_eq!(parsed.info().characters_with_spaces, Some(42));
        assert_eq!(parsed.text(), "Body");
    }

    #[test]
    fn document_writer_round_trips_bookmark_ranges() {
        let source = r#"{\rtf1\ansi Start {\*\bkmkstart\bkmkcolf2\bkmkcoll4\bkmkpub Link}R\'e9sum\'e9 \u20320?{\*\bkmkend Link} end}"#;
        let document = RtfDocument::parse(source).unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        let bookmark = reparsed.bookmarks().get("Link").unwrap();
        assert_eq!(bookmark.content, "Résumé 你");
        assert_eq!(bookmark.first_column, Some(2));
        assert_eq!(bookmark.last_column, Some(4));
        assert!(bookmark.is_public);
    }

    #[test]
    fn document_writer_preserves_bookmark_in_empty_body() {
        let document =
            RtfDocument::parse(r#"{\rtf1{\*\bkmkstart Empty}{\*\bkmkend Empty}}"#).unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        let bookmark = reparsed.bookmarks().get("Empty").unwrap();
        assert_eq!(bookmark.position, 0);
        assert!(bookmark.content.is_empty());
        assert!(reparsed.text().is_empty());
    }

    #[test]
    fn document_writer_round_trips_annotations() {
        let source = r#"{\rtf1\ansi Before {\*\atrfstart 12}range{\*\atrfend 12}{\*\atnid AM}{\*\atnauthor Ada M}\chatn{\*\annotation{\*\atnref 12}{\*\atndate 12345}{\*\atnparent 4}{\*\atnicn 3}{\*\atntime 99}Review \u20320? now} after}"#;
        let document = RtfDocument::parse(source).unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.annotations().len(), 1);
        let annotation = &reparsed.annotations()[0];
        assert_eq!(annotation.id, 12);
        assert_eq!(annotation.author, "Ada M");
        assert_eq!(annotation.initials, "AM");
        assert_eq!(annotation.date.as_deref(), Some("12345"));
        assert_eq!(annotation.parent_id.as_deref(), Some("4"));
        assert_eq!(annotation.icon.as_deref(), Some("3"));
        assert_eq!(annotation.time.as_deref(), Some("99"));
        assert_eq!(annotation.text, "Review 你 now");
        assert_eq!(annotation.position, "Before ".len());
        assert_eq!(annotation.range_end, "Before range".len());
    }

    #[test]
    fn document_writer_round_trips_headers_and_footers() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi\sectd\sbkodd\pgwsxn11000\pghsxn15000\marglsxn910\margrsxn810\margtsxn710\margbsxn610\guttersxn130\headery310\footery410\lndscpsxn\cols3\colsx370\pgnstarts6\pgnlcltr\vertalb\linemod1\lineppage{\header Header \u20320? one\par Header two}{\footer Footer}Body}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), "Body");
        assert_eq!(reparsed.sections().len(), 1);
        let section = &reparsed.sections()[0];
        assert_eq!(section.properties, document.sections()[0].properties);
        assert_eq!(
            section.get_header(HeaderFooterType::Header).unwrap().text(),
            "Header 你 one\nHeader two"
        );
        assert_eq!(
            section.get_header(HeaderFooterType::Footer).unwrap().text(),
            "Footer"
        );
    }

    #[test]
    fn document_writer_round_trips_stylesheets() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi{\stylesheet{\s0\snext0 Normal;}{\s1\b\qc\sbasedon0\snext0\slink2\sautoupd\shidden\slocked\ssemihidden\sunhideused\sqformat\spriority9\styrsid42 Heading \u20320?;}{\*\cs2\i\additive\slink1 Emphasis;}{\*\ds3 Section;}{\*\ts4 Table;}}Body}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), "Body");
        assert_eq!(reparsed.stylesheet().styles().len(), 5);
        let heading = reparsed
            .stylesheet()
            .get_typed(StyleType::Paragraph, 1)
            .unwrap();
        assert_eq!(heading.name, "Heading 你");
        assert!(heading.formatting.bold);
        assert_eq!(heading.paragraph.unwrap().alignment, Alignment::Center);
        assert_eq!(heading.linked_style, Some(2));
        assert!(heading.auto_update);
        assert!(heading.hidden);
        assert!(heading.locked);
        assert!(heading.semi_hidden);
        assert!(heading.unhide_when_used);
        assert!(heading.quick_format);
        assert_eq!(heading.priority, Some(9));
        assert_eq!(heading.revision_id, Some(42));

        let character = reparsed
            .stylesheet()
            .get_typed(StyleType::Character, 2)
            .unwrap();
        assert!(character.additive);
        assert!(character.formatting.italic);
        assert!(
            reparsed
                .stylesheet()
                .get_typed(StyleType::Section, 3)
                .is_some()
        );
        assert!(
            reparsed
                .stylesheet()
                .get_typed(StyleType::Table, 4)
                .is_some()
        );
    }

    #[test]
    fn document_writer_round_trips_list_tables() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi{\*\listtable{\list\listtemplateid42\listhybrid{\listlevel\levelnfc0\leveljc2\levelfollow1\levelstartat3\levelspace120\levelindent360{\leveltext\'02\'00.;}{\levelnumbers\'01;}\f2}{\listlevel\levelnfc77\leveljc0\levelfollow2\levelstartat1{\leveltext\'01\u8226?;}{\levelnumbers;}}{\listname Outline;}\listid77}}{\*\listoverridetable{\listoverride\listid77\listoverridecount1{\lfolevel\listoverridestartat\levelstartat9}\ls4}}\pard\ls4\ilvl1 Body}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), "Body");
        let paragraph = reparsed.blocks().last().unwrap().paragraph;
        assert_eq!(paragraph.list_override, Some(4));
        assert_eq!(paragraph.list_level, Some(1));
        let list = reparsed.list_table().get(77).unwrap();
        assert_eq!(list.template_id, 42);
        assert!(list.hybrid);
        assert_eq!(list.name, "Outline");
        assert_eq!(list.levels.len(), 2);
        assert_eq!(list.levels[0].number_text, "\0.");
        assert_eq!(list.levels[0].follow, ListFollow::Space);
        assert_eq!(list.levels[1].level_type, ListLevelType::Other(77));
        assert_eq!(list.levels[1].number_text, "•");
        assert_eq!(list.levels[1].follow, ListFollow::Nothing);
        let list_override = reparsed.list_override_table().get(4).unwrap();
        assert_eq!(list_override.list_id, 77);
        assert_eq!(list_override.level_count_override, Some(1));
        assert_eq!(list_override.start_at_override, Some(9));
    }

    #[test]
    fn document_writer_round_trips_tracked_revision_ranges() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi{\*\revtbl{Unknown;}{Ada;}}Before {\deleted\revauthdel1\revdttmdel123 old}{\revised\revauth1\revdttm-456 new \u20320?} after}"#,
        )
        .unwrap();
        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();

        let reparsed = RtfDocument::from_bytes(&output).unwrap_or_else(|error| {
            panic!(
                "failed to parse revision writer output: {error}\n{}",
                String::from_utf8_lossy(&output)
            )
        });
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.revisions().len(), 2);
        for (actual, expected) in reparsed.revisions().iter().zip(document.revisions()) {
            assert_eq!(actual.revision_type, expected.revision_type);
            assert_eq!(actual.id, expected.id);
            assert_eq!(actual.author, expected.author);
            assert_eq!(actual.date, expected.date);
            assert_eq!(actual.content, expected.content);
            assert_eq!(actual.position, expected.position);
            assert_eq!(actual.range_end, expected.range_end);
        }
    }

    #[test]
    fn document_writer_round_trips_multiple_section_boundaries() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi\sectd\sbkpage\pgwsxn10000{\header First}One\sect\sectd\sbknone\pgwsxn12000{\header Second}Two\sect\sectd\sbkeven\pgwsxn14000{\header Third}Three}"#,
        )
        .unwrap();
        assert_eq!(document.text(), "OneTwoThree");
        assert_eq!(document.sections().len(), 3);
        assert_eq!(
            document.section_breaks().copied().collect::<Vec<_>>(),
            vec![
                crate::SectionBreak::new("One".len(), Some(1)),
                crate::SectionBreak::new("OneTwo".len(), Some(2)),
            ]
        );

        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();
        let serialized = String::from_utf8_lossy(&output);
        assert!(serialized.contains("\\sect\\sectd"));

        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.sections().len(), 3);
        assert_eq!(
            reparsed.section_breaks().copied().collect::<Vec<_>>(),
            document.section_breaks().copied().collect::<Vec<_>>(),
        );
        assert_eq!(
            reparsed.sections()[0].properties.break_type,
            crate::SectionBreakType::Page
        );
        assert_eq!(reparsed.sections()[0].properties.page_width, 10000);
        assert_eq!(
            reparsed.sections()[1].properties.break_type,
            crate::SectionBreakType::Continuous
        );
        assert_eq!(reparsed.sections()[1].properties.page_width, 12000);
        assert_eq!(
            reparsed.sections()[2].properties.break_type,
            crate::SectionBreakType::EvenPage
        );
        assert_eq!(reparsed.sections()[2].properties.page_width, 14000);
        for (section, expected_header) in reparsed
            .sections()
            .iter()
            .zip(["First", "Second", "Third"])
        {
            assert_eq!(
                section
                    .get_header(HeaderFooterType::Header)
                    .unwrap()
                    .text(),
                expected_header
            );
        }
    }

    #[test]
    fn document_writer_round_trips_boundary_to_first_explicit_section() {
        let document = RtfDocument::parse(
            r#"{\rtf1\ansi Before\sect\sectd\sbknone{\header Second}After}"#,
        )
        .unwrap();
        assert_eq!(document.text(), "BeforeAfter");
        assert_eq!(document.sections().len(), 1);
        assert_eq!(
            document.section_breaks().copied().collect::<Vec<_>>(),
            vec![crate::SectionBreak::new("Before".len(), Some(0))]
        );

        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();
        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(
            reparsed.section_breaks().copied().collect::<Vec<_>>(),
            document.section_breaks().copied().collect::<Vec<_>>(),
        );
        assert_eq!(reparsed.sections().len(), 1);
        assert_eq!(
            reparsed.sections()[0]
                .get_header(HeaderFooterType::Header)
                .unwrap()
                .text(),
            "Second"
        );
    }

    #[test]
    fn document_writer_preserves_inherited_section_boundary() {
        let document = RtfDocument::parse(r#"{\rtf1\ansi\sectd\sbknone One\sect Two}"#).unwrap();
        assert_eq!(document.text(), "OneTwo");
        assert_eq!(document.sections().len(), 1);
        assert_eq!(
            document.section_breaks().copied().collect::<Vec<_>>(),
            vec![crate::SectionBreak::new("One".len(), None)]
        );

        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();
        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.sections().len(), 1);
        assert_eq!(
            reparsed.section_breaks().copied().collect::<Vec<_>>(),
            document.section_breaks().copied().collect::<Vec<_>>(),
        );
    }

    #[test]
    fn document_writer_preserves_page_then_section_break_order() {
        let document =
            RtfDocument::parse(r#"{\rtf1\ansi\sectd One\page\sect\sectd\sbknone Two}"#)
                .unwrap();
        assert_eq!(document.text(), "OneTwo");
        assert_eq!(document.sections().len(), 2);
        assert_eq!(
            document.body_story_events(),
            [
                crate::BodyStoryEvent::PageBreak(crate::PageBreak::new("One".len())),
                crate::BodyStoryEvent::SectionBreak(crate::SectionBreak::new(
                    "One".len(),
                    Some(1),
                )),
            ]
        );

        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();
        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.body_story_events(), document.body_story_events());
    }

    #[test]
    fn document_writer_round_trips_libreoffice_multi_section_fixture() {
        let document = RtfDocument::parse(include_str!(
            "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf94043.rtf"
        ))
        .unwrap();
        assert!(document.sections().len() >= 3);
        assert!(document.section_breaks().count() >= 3);

        let mut output = Vec::new();
        RtfWriter::new(&mut output)
            .write_document(&document)
            .unwrap();
        let reparsed = RtfDocument::from_bytes(&output).unwrap();
        assert_eq!(reparsed.text(), document.text());
        assert_eq!(reparsed.sections().len(), document.sections().len());
        assert_eq!(
            reparsed.section_breaks().copied().collect::<Vec<_>>(),
            document.section_breaks().copied().collect::<Vec<_>>(),
        );
    }
}
