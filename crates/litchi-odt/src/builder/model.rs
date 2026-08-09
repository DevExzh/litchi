use crate::elements::table::Table;
use crate::elements::text::{Heading, Hyperlink, List, ListItem, Paragraph, Span};
use litchi_core::{Metadata, Result};
use std::ops::Range;

/// Builder for creating new ODT documents.
///
/// This builder allows you to create ODT documents programmatically by adding
/// paragraphs, tables, and other elements, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odt::Builder;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut builder = Builder::new();
/// builder.add_paragraph("Hello, World!")?;
/// builder.add_paragraph("This is a new document.")?;
/// builder.save("document.odt")?;
/// # Ok(())
/// # }
/// ```
/// Document element - can be paragraph, heading, table, or list
#[derive(Debug, Clone)]
pub(super) enum DocumentElement {
    Paragraph(Paragraph),
    Heading(Heading),
    Table(Table),
    List(List),
    Section(String),
}

#[derive(Debug, Clone)]
pub(super) enum AnnotationInsertion {
    Append {
        paragraph_index: usize,
        annotation: crate::ruby_family::Annotation,
    },
    Wrap {
        paragraph_index: usize,
        range: Range<usize>,
        annotation: crate::ruby_family::Annotation,
    },
}

pub struct Builder {
    pub(super) elements: Vec<DocumentElement>,
    pub(super) text_indexes: Vec<String>,
    pub(super) text_index_marks: Vec<(usize, crate::TextIndexMark)>,
    pub(super) reference_marks: Vec<(usize, crate::ReferenceMark)>,
    pub(super) bookmark_targets: Vec<(usize, crate::BookmarkTarget)>,
    pub(super) notes: Vec<(usize, crate::Note)>,
    pub(super) property_forms: Vec<crate::form::PropertyForm>,
    pub(super) control_forms: Vec<crate::form::ControlForm>,
    pub(super) ruby_annotations: Vec<AnnotationInsertion>,
    pub(super) ruby_styles: Vec<crate::ruby_family::Style>,
    pub(super) interactive_forms: Vec<crate::form::InteractiveForm>,
    pub(super) selection_forms: Vec<crate::form::SelectionForm>,
    pub(super) visual_forms: Vec<crate::form::VisualForm>,
    pub(super) generic_forms: Vec<crate::form::GenericForm>,
    pub(super) password_file_forms: Vec<crate::form::PasswordFileForm>,
    pub(super) image_frame_forms: Vec<crate::form::ImageFrameForm>,
    pub(super) value_range_forms: Vec<crate::form::ValueRangeForm>,
    pub(super) typed_value_forms: Vec<crate::form::TypedValueForm>,
    pub(super) grid_forms: Vec<crate::form::GridForm>,
    pub(super) connection_resource_forms: Vec<crate::form::ConnectionResourceForm>,
    pub(super) metadata: Metadata,
    pub(super) paragraph_tab_styles: Vec<crate::style::paragraph::tab_stop::Style>,
    pub(super) paragraph_drop_cap_styles: Vec<crate::style::paragraph::drop_cap::Style>,
    pub(super) list_level_label_alignments: Vec<crate::list_label_alignment::Style>,
    pub(super) paragraph_flow_styles: Vec<crate::style::paragraph::flow::Style>,
    pub(super) paragraph_margin_styles: Vec<crate::style::paragraph::margin::Style>,
    pub(super) paragraph_border_styles: Vec<crate::style::paragraph::border::Style>,
    pub(super) paragraph_alignment_styles: Vec<crate::style::paragraph::alignment::Style>,
    pub(super) paragraph_break_styles: Vec<crate::style::paragraph::breaks::Style>,
    pub(super) paragraph_writing_mode_styles: Vec<crate::style::paragraph::writing_mode::Style>,
    pub(super) table_row_property_styles: Vec<crate::style::table::row::Style>,
    pub(super) table_column_property_styles: Vec<crate::style::table::column::Style>,
    pub(super) table_cell_property_styles: Vec<crate::style::table::cell::Style>,
    pub(super) table_property_styles: Vec<crate::style::table::table::Style>,
    pub(super) section_property_styles: Vec<crate::SectionStyleProperties>,
    pub(super) section_names: std::collections::HashSet<String>,
    pub(super) section_xml_ids: std::collections::HashSet<String>,
    pub(super) page_layout_columns: Vec<(String, crate::style::columns::Columns)>,
    pub(super) page_layout_footnote_separators: Vec<(String, crate::footnote_separator::Separator)>,
    pub(super) page_layout_header_footer_properties: Vec<(
        String,
        crate::header_footer::properties::Region,
        crate::header_footer::properties::StyleProperties,
    )>,
    pub(super) notes_configurations: crate::notes_configuration::Configurations,
    pub(super) line_numbering_configuration: Option<crate::line_numbering::Configuration>,
    pub(super) page_sequence: Option<crate::Sequence>,
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

impl Builder {
    /// Create a new document builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    ///
    /// let builder = Builder::new();
    /// ```
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            text_indexes: Vec::new(),
            text_index_marks: Vec::new(),
            reference_marks: Vec::new(),
            bookmark_targets: Vec::new(),
            ruby_annotations: Vec::new(),
            notes: Vec::new(),
            ruby_styles: Vec::new(),
            property_forms: Vec::new(),
            control_forms: Vec::new(),
            interactive_forms: Vec::new(),
            selection_forms: Vec::new(),
            visual_forms: Vec::new(),
            generic_forms: Vec::new(),
            password_file_forms: Vec::new(),
            image_frame_forms: Vec::new(),
            value_range_forms: Vec::new(),
            typed_value_forms: Vec::new(),
            grid_forms: Vec::new(),
            connection_resource_forms: Vec::new(),
            metadata: Metadata::default(),
            paragraph_tab_styles: Vec::new(),
            paragraph_drop_cap_styles: Vec::new(),
            list_level_label_alignments: Vec::new(),
            paragraph_flow_styles: Vec::new(),
            paragraph_margin_styles: Vec::new(),
            paragraph_border_styles: Vec::new(),
            paragraph_alignment_styles: Vec::new(),
            paragraph_break_styles: Vec::new(),
            paragraph_writing_mode_styles: Vec::new(),
            table_row_property_styles: Vec::new(),
            table_column_property_styles: Vec::new(),
            table_cell_property_styles: Vec::new(),
            table_property_styles: Vec::new(),
            section_property_styles: Vec::new(),
            section_names: std::collections::HashSet::new(),
            section_xml_ids: std::collections::HashSet::new(),
            page_layout_columns: Vec::new(),
            page_layout_footnote_separators: Vec::new(),
            page_layout_header_footer_properties: Vec::new(),
            notes_configurations: crate::notes_configuration::Configurations::default(),
            line_numbering_configuration: None,
            page_sequence: None,
        }
    }

    /// Add a paragraph containing one validated, inert dynamic text field.
    pub fn add_dynamic_text_field(
        &mut self,
        field: &crate::elements::field::DynamicTextField,
    ) -> Result<&mut Self> {
        let mut paragraph = Paragraph::new();
        paragraph.add_dynamic_text_field(field)?;
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(self)
    }

    /// Return footnote and endnote configurations emitted to `styles.xml`.
    pub fn notes_configurations(&self) -> &crate::notes_configuration::Configurations {
        &self.notes_configurations
    }

    /// Replace the validated footnote and endnote configurations.
    pub fn set_notes_configurations(
        &mut self,
        configurations: crate::notes_configuration::Configurations,
    ) -> Result<&mut Self> {
        configurations.validate()?;
        self.notes_configurations = configurations;
        Ok(self)
    }

    /// Set one validated note-class configuration.
    pub fn set_notes_configuration(
        &mut self,
        configuration: crate::notes_configuration::Configuration,
    ) -> Result<&mut Self> {
        configuration.validate()?;
        match configuration.note_class {
            crate::notes_configuration::Class::Footnote => {
                self.notes_configurations.footnote = Some(configuration);
            },
            crate::notes_configuration::Class::Endnote => {
                self.notes_configurations.endnote = Some(configuration);
            },
        }
        Ok(self)
    }

    /// Return the optional document line-numbering configuration to emit.
    ///
    /// The configuration is serialized as style metadata only. Building a
    /// document never calculates page or line numbers.
    pub fn line_numbering_configuration(&self) -> Option<&crate::line_numbering::Configuration> {
        self.line_numbering_configuration.as_ref()
    }

    /// Set the validated document line-numbering configuration to emit.
    ///
    /// This stores presentation metadata only. It never performs pagination or
    /// generates line numbers.
    pub fn set_line_numbering_configuration(
        &mut self,
        configuration: crate::line_numbering::Configuration,
    ) -> Result<&mut Self> {
        configuration.validate()?;
        self.line_numbering_configuration = Some(configuration);
        Ok(self)
    }

    /// Omit document line-numbering configuration from generated styles.
    pub fn clear_line_numbering_configuration(&mut self) -> &mut Self {
        self.line_numbering_configuration = None;
        self
    }

    /// Add a named or default paragraph style carrying typed tab stops.
    pub fn add_paragraph_tab_style(
        &mut self,
        style: crate::style::paragraph::tab_stop::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_tab_styles.len() >= 4_096 {
            return Err(litchi_core::Error::InvalidFormat(
                "document builder exceeds 4096 paragraph tab styles".to_owned(),
            ));
        }
        if self.paragraph_tab_styles.iter().any(|existing| {
            existing.is_default_style == style.is_default_style && existing.name == style.name
        }) {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate paragraph tab style identity".to_owned(),
            ));
        }
        self.paragraph_tab_styles.push(style);
        Ok(self)
    }

    /// Add a typed paragraph drop-cap style definition.
    pub fn add_paragraph_drop_cap_style(
        &mut self,
        style: crate::style::paragraph::drop_cap::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_drop_cap_styles.len() >= 4_096 {
            return Err(litchi_core::Error::InvalidFormat(
                "too many paragraph drop-cap styles".to_string(),
            ));
        }
        if self
            .paragraph_drop_cap_styles
            .iter()
            .any(|old| old.name == style.name && old.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate paragraph drop-cap style identity".to_string(),
            ));
        }
        if let Some(tabs) = self
            .paragraph_tab_styles
            .iter()
            .find(|tabs| crate::style::paragraph::drop_cap::same_style_identity(&style, tabs))
            && tabs.parent_style_name != style.parent_style_name
        {
            return Err(litchi_core::Error::InvalidFormat(
                "paragraph style parent definitions conflict".to_string(),
            ));
        }
        self.paragraph_drop_cap_styles.push(style);
        Ok(self)
    }

    /// Customize one of the three generated `L1` numbered-list levels.
    pub fn set_numbered_list_level_label_alignment(
        &mut self,
        level: u16,
        alignment: crate::list_label_alignment::Alignment,
    ) -> Result<&mut Self> {
        if !(1..=3).contains(&level) {
            return Err(litchi_core::Error::InvalidFormat(
                "generated numbered-list level must be 1..=3".to_string(),
            ));
        }
        let item = crate::list_label_alignment::Style::new("L1", level, alignment)?;
        if let Some(old) = self
            .list_level_label_alignments
            .iter_mut()
            .find(|x| x.level == level)
        {
            *old = item;
        } else {
            self.list_level_label_alignments.push(item);
        }
        Ok(self)
    }
    pub fn add_paragraph_flow_style(
        &mut self,
        style: crate::style::paragraph::flow::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_flow_styles.len() >= 4096
            || self
                .paragraph_flow_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive paragraph flow style".to_string(),
            ));
        }
        if self
            .paragraph_tab_styles
            .iter()
            .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
            || self
                .paragraph_drop_cap_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "paragraph flow style conflicts with another typed paragraph style".to_string(),
            ));
        }
        self.paragraph_flow_styles.push(style);
        Ok(self)
    }

    /// Whether another typed paragraph style family already uses this identity.
    fn paragraph_style_identity_taken(&self, name: &Option<String>, is_default: bool) -> bool {
        let matches =
            |x_name: &Option<String>, x_default: bool| *x_name == *name && x_default == is_default;
        self.paragraph_tab_styles
            .iter()
            .any(|x| matches(&x.name, x.is_default_style))
            || self
                .paragraph_drop_cap_styles
                .iter()
                .any(|x| matches(&x.name, x.is_default_style))
            || self
                .paragraph_flow_styles
                .iter()
                .any(|x| matches(&x.name, x.is_default_style))
            || self
                .paragraph_margin_styles
                .iter()
                .any(|x| matches(&x.name, x.is_default_style))
            || self
                .paragraph_border_styles
                .iter()
                .any(|x| matches(&x.name, x.is_default_style))
            || self
                .paragraph_alignment_styles
                .iter()
                .any(|x| matches(&x.name, x.is_default_style))
            || self
                .paragraph_break_styles
                .iter()
                .any(|x| matches(&x.name, x.is_default_style))
            || self
                .paragraph_writing_mode_styles
                .iter()
                .any(|x| matches(&x.name, x.is_default_style))
    }

    /// Add a named or default paragraph style carrying typed margin properties.
    pub fn add_paragraph_margin_style(
        &mut self,
        style: crate::style::paragraph::margin::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_margin_styles.len() >= 4096
            || self
                .paragraph_margin_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive paragraph margin style".to_string(),
            ));
        }
        if self.paragraph_style_identity_taken(&style.name, style.is_default_style) {
            return Err(litchi_core::Error::InvalidFormat(
                "paragraph margin style conflicts with another typed paragraph style".to_string(),
            ));
        }
        self.paragraph_margin_styles.push(style);
        Ok(self)
    }

    /// Add a named or default paragraph style carrying typed border and
    /// background properties.
    pub fn add_paragraph_border_style(
        &mut self,
        style: crate::style::paragraph::border::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_border_styles.len() >= 4096
            || self
                .paragraph_border_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive paragraph border style".to_string(),
            ));
        }
        if self.paragraph_style_identity_taken(&style.name, style.is_default_style) {
            return Err(litchi_core::Error::InvalidFormat(
                "paragraph border style conflicts with another typed paragraph style".to_string(),
            ));
        }
        self.paragraph_border_styles.push(style);
        Ok(self)
    }

    /// Add a named or default paragraph style carrying typed alignment
    /// properties.
    pub fn add_paragraph_alignment_style(
        &mut self,
        style: crate::style::paragraph::alignment::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_alignment_styles.len() >= 4096
            || self
                .paragraph_alignment_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive paragraph alignment style".to_string(),
            ));
        }
        if self.paragraph_style_identity_taken(&style.name, style.is_default_style) {
            return Err(litchi_core::Error::InvalidFormat(
                "paragraph alignment style conflicts with another typed paragraph style"
                    .to_string(),
            ));
        }
        self.paragraph_alignment_styles.push(style);
        Ok(self)
    }

    /// Add a named or default paragraph style carrying typed break, page-number,
    /// and line-numbering properties.
    pub fn add_paragraph_break_style(
        &mut self,
        style: crate::style::paragraph::breaks::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_break_styles.len() >= 4096
            || self
                .paragraph_break_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive paragraph break style".to_string(),
            ));
        }
        if self.paragraph_style_identity_taken(&style.name, style.is_default_style) {
            return Err(litchi_core::Error::InvalidFormat(
                "paragraph break style conflicts with another typed paragraph style".to_string(),
            ));
        }
        self.paragraph_break_styles.push(style);
        Ok(self)
    }

    /// Add a named or default paragraph style carrying typed writing-mode and
    /// register properties.
    pub fn add_paragraph_writing_mode_style(
        &mut self,
        style: crate::style::paragraph::writing_mode::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.paragraph_writing_mode_styles.len() >= 4096
            || self
                .paragraph_writing_mode_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive paragraph writing-mode style".to_string(),
            ));
        }
        if self.paragraph_style_identity_taken(&style.name, style.is_default_style) {
            return Err(litchi_core::Error::InvalidFormat(
                "paragraph writing-mode style conflicts with another typed paragraph style"
                    .to_string(),
            ));
        }
        self.paragraph_writing_mode_styles.push(style);
        Ok(self)
    }

    /// Add a named or default table-row style carrying typed row properties.
    pub fn add_table_row_property_style(
        &mut self,
        style: crate::style::table::row::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.table_row_property_styles.len() >= 4096
            || self
                .table_row_property_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive table-row property style".to_string(),
            ));
        }
        self.table_row_property_styles.push(style);
        Ok(self)
    }

    /// Add a named or default table-column style carrying typed column properties.
    pub fn add_table_column_property_style(
        &mut self,
        style: crate::style::table::column::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.table_column_property_styles.len() >= 4096
            || self
                .table_column_property_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive table-column property style".to_string(),
            ));
        }
        self.table_column_property_styles.push(style);
        Ok(self)
    }

    /// Add a named or default table-cell style carrying typed cell properties.
    pub fn add_table_cell_property_style(
        &mut self,
        style: crate::style::table::cell::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.table_cell_property_styles.len() >= 4096
            || self
                .table_cell_property_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive table-cell property style".to_string(),
            ));
        }
        self.table_cell_property_styles.push(style);
        Ok(self)
    }

    /// Add a named or default table style carrying typed table properties.
    pub fn add_table_property_style(
        &mut self,
        style: crate::style::table::table::Style,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.table_property_styles.len() >= 4096
            || self
                .table_property_styles
                .iter()
                .any(|x| x.name == style.name && x.is_default_style == style.is_default_style)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive table property style".to_string(),
            ));
        }
        self.table_property_styles.push(style);
        Ok(self)
    }

    /// Add a named section style carrying typed residual section properties.
    pub fn add_section_property_style(
        &mut self,
        style: crate::SectionStyleProperties,
    ) -> Result<&mut Self> {
        style.validate()?;
        if self.section_property_styles.len() >= 4096
            || self
                .section_property_styles
                .iter()
                .any(|item| item.name == style.name)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive section property style".to_string(),
            ));
        }
        self.section_property_styles.push(style);
        Ok(self)
    }

    /// Add a named automatic page layout with typed multi-column properties.
    pub fn add_page_layout_columns(
        &mut self,
        page_layout_name: impl Into<String>,
        columns: crate::style::columns::Columns,
    ) -> Result<&mut Self> {
        let name = page_layout_name.into();
        columns.to_page_layout_fragment(&name)?;
        if self.page_layout_columns.len() >= 4_096 {
            return Err(litchi_core::Error::InvalidFormat(
                "document builder exceeds 4096 column page layouts".to_owned(),
            ));
        }
        if self
            .page_layout_columns
            .iter()
            .any(|(existing, _)| existing == &name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate page layout '{name}'"
            )));
        }
        self.page_layout_columns.push((name, columns));
        Ok(self)
    }

    /// Add a typed footnote separator to a named automatic page layout.
    pub fn add_page_layout_footnote_separator(
        &mut self,
        page_layout_name: impl Into<String>,
        separator: crate::footnote_separator::Separator,
    ) -> Result<&mut Self> {
        let name = page_layout_name.into();
        separator.to_page_layout_fragment(&name)?;
        if self
            .page_layout_footnote_separators
            .iter()
            .any(|(existing, _)| existing == &name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate footnote separator page layout '{name}'"
            )));
        }
        let shares_column_layout = self
            .page_layout_columns
            .iter()
            .any(|(existing, _)| existing == &name);
        if !shares_column_layout
            && self.page_layout_columns.len() + self.page_layout_footnote_separators.len() >= 4_096
        {
            return Err(litchi_core::Error::InvalidFormat(
                "document builder exceeds 4096 page layouts".to_owned(),
            ));
        }
        self.page_layout_footnote_separators.push((name, separator));
        Ok(self)
    }

    pub fn add_page_layout_header_footer_properties(
        &mut self,
        page_layout_name: impl Into<String>,
        region: crate::header_footer::properties::Region,
        properties: crate::header_footer::properties::StyleProperties,
    ) -> Result<&mut Self> {
        let name = page_layout_name.into();
        properties.validate()?;
        if self.page_layout_header_footer_properties.len() >= 8192
            || self
                .page_layout_header_footer_properties
                .iter()
                .any(|(n, r, _)| n == &name && *r == region)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "duplicate or excessive page-layout header/footer properties".to_string(),
            ));
        }
        self.page_layout_header_footer_properties
            .push((name, region, properties));
        Ok(self)
    }

    /// Set document metadata
    ///
    /// # Arguments
    ///
    /// * `metadata` - Document metadata (title, author, etc.)
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    /// use litchi_core::Metadata;
    ///
    /// let mut builder = Builder::new();
    /// let mut metadata = Metadata::default();
    /// metadata.title = Some("My Document".to_string());
    /// metadata.author = Some("John Doe".to_string());
    /// builder.set_metadata(metadata);
    /// ```
    pub fn set_metadata(&mut self, metadata: Metadata) {
        self.metadata = metadata;
    }

    /// Add a paragraph with text
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_paragraph(&mut self, text: &str) -> Result<&mut Self> {
        let mut para = Paragraph::new();
        para.set_text(text);
        self.elements.push(DocumentElement::Paragraph(para));
        Ok(self)
    }

    /// Add a paragraph containing one simple ODF hyperlink.
    ///
    /// Hyperlink targets are serialized as inert `text:a` `XLink` attributes;
    /// this library never follows or fetches them.
    pub fn add_hyperlink(
        &mut self,
        href: impl AsRef<str>,
        text: impl AsRef<str>,
    ) -> Result<&mut Self> {
        let hyperlink = Hyperlink::with_href(href, text)?;
        self.add_hyperlink_element(hyperlink)
    }

    /// Add a paragraph containing a fully configured ODF hyperlink.
    pub fn add_hyperlink_element(&mut self, hyperlink: Hyperlink) -> Result<&mut Self> {
        let mut paragraph = Paragraph::new();
        paragraph.add_hyperlink(hyperlink)?;
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(self)
    }

    /// Add a heading
    ///
    /// # Arguments
    ///
    /// * `text` - Heading text
    /// * `level` - Heading level (1-6)
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_heading("Chapter 1", 1)?;
    /// builder.add_heading("Section 1.1", 2)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_heading(&mut self, text: &str, level: u8) -> Result<&mut Self> {
        if !(1..=6).contains(&level) {
            return Err(litchi_core::Error::Other(
                "Heading level must be between 1 and 6".to_string(),
            ));
        }
        let mut heading = Heading::new(level);
        heading.set_text(text);
        self.elements.push(DocumentElement::Heading(heading));
        Ok(self)
    }

    /// Add a paragraph with rich text formatting
    ///
    /// # Arguments
    ///
    /// * `spans` - Vector of (text, `style_name`) tuples for formatted text
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_rich_paragraph(vec![
    ///     ("This is ", None),
    ///     ("bold", Some("Bold")),
    ///     (" and ", None),
    ///     ("italic", Some("Italic")),
    ///     (" text.", None),
    /// ])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_rich_paragraph(&mut self, spans: Vec<(&str, Option<&str>)>) -> Result<&mut Self> {
        let mut para = Paragraph::new();

        for (text, style) in spans {
            let mut span = Span::new();
            span.set_text(text);
            if let Some(style_name) = style {
                span.set_style_name(style_name);
            }
            para.add_span(span);
        }

        self.elements.push(DocumentElement::Paragraph(para));
        Ok(self)
    }

    /// Add a bulleted list
    ///
    /// # Arguments
    ///
    /// * `items` - Vector of list item texts
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_bulleted_list(vec!["Item 1", "Item 2", "Item 3"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_bulleted_list(&mut self, items: Vec<&str>) -> Result<&mut Self> {
        let mut list = List::new();

        for item_text in items {
            let mut item = ListItem::new();
            let mut para = Paragraph::new();
            para.set_text(item_text);
            item.add_paragraph(para);
            list.add_item(item);
        }

        self.elements.push(DocumentElement::List(list));
        Ok(self)
    }

    /// Add a numbered list
    ///
    /// # Arguments
    ///
    /// * `items` - Vector of list item texts
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_numbered_list(vec!["First", "Second", "Third"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_numbered_list(&mut self, items: Vec<&str>) -> Result<&mut Self> {
        let mut list = List::new();
        // Set the numbered list style
        list.set_style_name("L1");

        for item_text in items {
            let mut item = ListItem::new();
            let mut para = Paragraph::new();
            para.set_text(item_text);
            item.add_paragraph(para);
            list.add_item(item);
        }

        self.elements.push(DocumentElement::List(list));
        Ok(self)
    }

    /// Add a paragraph element
    ///
    /// # Arguments
    ///
    /// * `paragraph` - A `Paragraph` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    /// use litchi_odt::elements::text::Paragraph;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// let mut para = Paragraph::new();
    /// para.set_text("Styled paragraph");
    /// para.set_style_name("Heading1");
    /// builder.add_paragraph_element(para)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_paragraph_element(&mut self, paragraph: Paragraph) -> Result<&mut Self> {
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(self)
    }

    /// Add a heading element
    ///
    /// # Arguments
    ///
    /// * `heading` - A `Heading` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    /// use litchi_odt::elements::text::Heading;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// let mut heading = Heading::new(1);
    /// heading.set_text("Chapter Title");
    /// builder.add_heading_element(heading)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_heading_element(&mut self, heading: Heading) -> Result<&mut Self> {
        self.elements.push(DocumentElement::Heading(heading));
        Ok(self)
    }

    /// Add a list element
    ///
    /// # Arguments
    ///
    /// * `list` - A `List` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    /// use litchi_odt::elements::text::{List, ListItem};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// let mut list = List::new();
    /// let mut item = ListItem::new();
    /// item.set_text("First item");
    /// list.add_item(item);
    /// builder.add_list_element(list)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_list_element(&mut self, list: List) -> Result<&mut Self> {
        self.elements.push(DocumentElement::List(list));
        Ok(self)
    }

    /// Add a table to the document
    ///
    /// # Arguments
    ///
    /// * `table` - A `Table` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::Builder;
    /// use litchi_odt::elements::table::Table;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// let mut table = Table::new();
    /// table.set_name("Table1");
    /// builder.add_table(table)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_table(&mut self, table: Table) -> Result<&mut Self> {
        self.elements.push(DocumentElement::Table(table));
        Ok(self)
    }

    /// Add a validated typed section in mixed document order.
    pub fn add_section(&mut self, section: crate::Section) -> Result<&mut Self> {
        section.validate()?;
        if self.section_names.contains(&section.name) {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate section name '{}'",
                section.name
            )));
        }
        if let Some(id) = &section.xml_id
            && self.section_xml_ids.contains(id)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate section xml:id '{id}'"
            )));
        }
        let fragment = section.to_xml_fragment()?;
        self.section_names.insert(section.name.clone());
        if let Some(id) = &section.xml_id {
            self.section_xml_ids.insert(id.clone());
        }
        self.elements.push(DocumentElement::Section(fragment));
        Ok(self)
    }

    /// Add caller-authored, schema-validated generated-index markup.
    pub fn add_text_index(&mut self, index: &crate::TextIndex) -> Result<&mut Self> {
        self.text_indexes.push(index.to_xml_fragment()?);
        Ok(self)
    }

    /// Insert a validated point mark at a current paragraph end, or wrap it with a range mark.
    pub fn add_text_index_mark(
        &mut self,
        paragraph_index: usize,
        mark: &crate::TextIndexMark,
    ) -> Result<&mut Self> {
        mark.to_xml_fragments()?;
        let paragraph_count = self
            .elements
            .iter()
            .filter(|element| {
                matches!(
                    element,
                    DocumentElement::Paragraph(_) | DocumentElement::Heading(_)
                )
            })
            .count();
        if paragraph_index >= paragraph_count {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "paragraph index {paragraph_index} is out of bounds"
            )));
        }
        self.text_index_marks.push((paragraph_index, mark.clone()));
        Ok(self)
    }

    /// Insert a point reference at a current paragraph end, or wrap it with a range reference.
    pub fn add_reference_mark(
        &mut self,
        paragraph_index: usize,
        mark: &crate::ReferenceMark,
    ) -> Result<&mut Self> {
        mark.to_xml_fragments()?;
        let paragraph_count = self
            .elements
            .iter()
            .filter(|element| {
                matches!(
                    element,
                    DocumentElement::Paragraph(_) | DocumentElement::Heading(_)
                )
            })
            .count();
        if paragraph_index >= paragraph_count {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "paragraph index {paragraph_index} is out of bounds"
            )));
        }
        if self
            .reference_marks
            .iter()
            .any(|(_, existing)| existing.name() == mark.name())
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate reference-mark identity '{}'",
                mark.name()
            )));
        }
        self.reference_marks.push((paragraph_index, mark.clone()));
        Ok(self)
    }

    /// Insert a point bookmark at a current paragraph end, or wrap it with a range bookmark.
    pub fn add_bookmark_target(
        &mut self,
        paragraph_index: usize,
        target: &crate::BookmarkTarget,
    ) -> Result<&mut Self> {
        target.to_xml_fragments()?;
        let paragraph_count = self
            .elements
            .iter()
            .filter(|element| {
                matches!(
                    element,
                    DocumentElement::Paragraph(_) | DocumentElement::Heading(_)
                )
            })
            .count();
        if paragraph_index >= paragraph_count {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "paragraph index {paragraph_index} is out of bounds"
            )));
        }
        if self
            .bookmark_targets
            .iter()
            .any(|(_, existing)| existing.name() == target.name())
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate bookmark identity '{}'",
                target.name()
            )));
        }
        self.bookmark_targets
            .push((paragraph_index, target.clone()));
        Ok(self)
    }

    /// Append a typed ruby annotation to one `text:p` paragraph.
    ///
    /// The annotation is inserted at the end of the paragraph selected in
    /// document order, including paragraphs nested in lists and table cells.
    /// Its base may contain validated inline content, while its pronunciation
    /// is plain text as required by ODF `text:ruby`.
    pub fn add_ruby_annotation(
        &mut self,
        paragraph_index: usize,
        annotation: &crate::ruby_family::Annotation,
    ) -> Result<&mut Self> {
        annotation.validate()?;
        let body = self.generate_content_body()?;
        let xml = format!(
            r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</office:text>"#
        );
        crate::insert_ruby_annotation_xml(&xml, paragraph_index, annotation)?;
        self.ruby_annotations.push(AnnotationInsertion::Append {
            paragraph_index,
            annotation: annotation.clone(),
        });
        Ok(self)
    }

    /// Wrap one UTF-8 text range in a selected paragraph with ruby.
    ///
    /// The range follows `wrap_ruby_annotation_xml`: a plain base may span
    /// adjacent character data under one parent, while an XML base may span
    /// balanced legal inline elements. The builder emits queued ruby mutations
    /// in call order without splitting ancestors or evaluating active content.
    pub fn wrap_ruby_annotation(
        &mut self,
        paragraph_index: usize,
        range: Range<usize>,
        annotation: &crate::ruby_family::Annotation,
    ) -> Result<&mut Self> {
        annotation.validate()?;
        let body = self.generate_content_body()?;
        let xml = format!(
            r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</office:text>"#
        );
        crate::wrap_ruby_annotation_xml(&xml, paragraph_index, range.clone(), annotation)?;
        self.ruby_annotations.push(AnnotationInsertion::Wrap {
            paragraph_index,
            range,
            annotation: annotation.clone(),
        });
        Ok(self)
    }

    /// Append a validated footnote or endnote to one `text:p` paragraph.
    ///
    /// The note is inserted at the paragraph end selected in document order,
    /// including paragraphs nested in lists, table cells, and note bodies.
    /// Plain-text body newlines create separate ODF paragraphs; a note with an
    /// `NoteBodyContent` retains its validated structured body. No field,
    /// link, script, macro, event listener, or embedded payload is evaluated.
    pub fn add_note(&mut self, paragraph_index: usize, note: &crate::Note) -> Result<&mut Self> {
        note.validate()?;
        let body = self.generate_content_body()?;
        let xml = format!(
            r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</office:text>"#
        );
        crate::insert_note_xml(&xml, paragraph_index, note)?;
        self.notes.push((paragraph_index, note.clone()));
        Ok(self)
    }

    /// Add a named ODF ruby style definition to `styles.xml`.
    pub fn add_ruby_style(&mut self, style: crate::ruby_family::Style) -> Result<&mut Self> {
        style.validate()?;
        if self
            .ruby_styles
            .iter()
            .any(|existing| existing.name == style.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate ruby style name '{}'",
                style.name
            )));
        }
        self.ruby_styles.push(style);
        Ok(self)
    }

    /// Set or clear the explicit `text:page-sequence` master-page assignments.
    ///
    /// The sequence is written as the first child of `office:text`, matching
    /// the element order of ODF 1.3 §5.1 and §5.3. Master-page names are
    /// stored lexically and never resolved against `styles.xml`.
    pub fn set_page_sequence(&mut self, sequence: Option<crate::Sequence>) -> Result<&mut Self> {
        if let Some(sequence) = &sequence {
            sequence.to_xml_fragment()?;
        }
        self.page_sequence = sequence;
        Ok(self)
    }

    /// Add a minimal inert form containing typed custom properties.
    pub fn add_property_form(&mut self, form: &crate::form::PropertyForm) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .property_forms
            .iter()
            .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate property form '{}'",
                form.name
            )));
        }
        self.property_forms.push(form.clone());
        Ok(self)
    }

    /// Add a canonical form containing typed text and textarea controls.
    pub fn add_control_form(&mut self, form: &crate::form::ControlForm) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .control_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.control_forms.push(form.clone());
        Ok(self)
    }

    /// Add a canonical form containing typed button and checkbox controls.
    pub fn add_interactive_form(
        &mut self,
        form: &crate::form::InteractiveForm,
    ) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .interactive_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.interactive_forms.push(form.clone());
        Ok(self)
    }

    /// Add a canonical form containing typed listbox and combobox controls.
    pub fn add_selection_form(&mut self, form: &crate::form::SelectionForm) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .selection_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.selection_forms.push(form.clone());
        Ok(self)
    }

    /// Add a canonical form containing radio, frame, and image-button controls.
    pub fn add_visual_form(&mut self, form: &crate::form::VisualForm) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .visual_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.visual_forms.push(form.clone());
        Ok(self)
    }

    /// Add a canonical form containing fixed-text, hidden, and generic controls.
    pub fn add_generic_form(&mut self, form: &crate::form::GenericForm) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .generic_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .visual_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.generic_forms.push(form.clone());
        Ok(self)
    }

    /// Add a canonical form containing password and file controls.
    pub fn add_password_file_form(
        &mut self,
        form: &crate::form::PasswordFileForm,
    ) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .password_file_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .generic_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .visual_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.password_file_forms.push(form.clone());
        Ok(self)
    }

    /// Add a canonical form containing image-frame controls.
    pub fn add_image_frame_form(
        &mut self,
        form: &crate::form::ImageFrameForm,
    ) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .image_frame_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .password_file_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .generic_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .visual_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.image_frame_forms.push(form.clone());
        Ok(self)
    }
}

impl Builder {
    /// Add a validated value-range form to the document.
    pub fn add_value_range_form(
        &mut self,
        form: &crate::form::ValueRangeForm,
    ) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .value_range_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .image_frame_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .password_file_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .generic_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .visual_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.value_range_forms.push(form.clone());
        Ok(self)
    }
}

impl Builder {
    /// Add a validated form containing formatted-text, number, date, or time controls.
    pub fn add_typed_value_form(
        &mut self,
        form: &crate::form::TypedValueForm,
    ) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .typed_value_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .value_range_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .image_frame_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .password_file_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .generic_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .visual_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.typed_value_forms.push(form.clone());
        Ok(self)
    }
}

impl Builder {
    /// Adds a form whose final child is an inert `form:connection-resource`.
    pub fn add_connection_resource_form(
        &mut self,
        form: &crate::form::ConnectionResourceForm,
    ) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .connection_resource_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .grid_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .typed_value_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .value_range_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .image_frame_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .password_file_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .generic_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .visual_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.connection_resource_forms.push(form.clone());
        Ok(self)
    }

    pub fn add_grid_form(&mut self, form: &crate::form::GridForm) -> Result<&mut Self> {
        form.to_xml_fragment()?;
        if self
            .grid_forms
            .iter()
            .any(|existing| existing.name == form.name)
            || self
                .typed_value_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .value_range_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .image_frame_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .password_file_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .generic_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .visual_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .selection_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .interactive_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .control_forms
                .iter()
                .any(|existing| existing.name == form.name)
            || self
                .property_forms
                .iter()
                .any(|existing| existing.name == form.name)
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "duplicate form '{}'",
                form.name
            )));
        }
        self.grid_forms.push(form.clone());
        Ok(self)
    }
}
