//! OpenDocument Text document builder.
//!
//! This module provides a builder pattern for creating new ODT documents from scratch.

use crate::core::PackageWriter;
use crate::elements::table::Table;
use crate::elements::text::{Heading, Hyperlink, List, ListItem, Paragraph, Span};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::{ops::Range, path::Path};

/// Builder for creating new ODT documents.
///
/// This builder allows you to create ODT documents programmatically by adding
/// paragraphs, tables, and other elements, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::DocumentBuilder;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut builder = DocumentBuilder::new();
/// builder.add_paragraph("Hello, World!")?;
/// builder.add_paragraph("This is a new document.")?;
/// builder.save("document.odt")?;
/// # Ok(())
/// # }
/// ```
/// Document element - can be paragraph, heading, table, or list
#[derive(Debug, Clone)]
enum DocumentElement {
    Paragraph(Paragraph),
    Heading(Heading),
    Table(Table),
    List(List),
    Section(String),
}

#[derive(Debug, Clone)]
enum RubyAnnotationInsertion {
    Append {
        paragraph_index: usize,
        annotation: crate::RubyAnnotation,
    },
    Wrap {
        paragraph_index: usize,
        range: Range<usize>,
        annotation: crate::RubyAnnotation,
    },
}

pub struct DocumentBuilder {
    elements: Vec<DocumentElement>,
    text_indexes: Vec<String>,
    text_index_marks: Vec<(usize, crate::TextIndexMark)>,
    reference_marks: Vec<(usize, crate::ReferenceMark)>,
    bookmark_targets: Vec<(usize, crate::BookmarkTarget)>,
    ruby_annotations: Vec<RubyAnnotationInsertion>,
    notes: Vec<(usize, crate::Note)>,
    ruby_styles: Vec<crate::RubyStyle>,
    property_forms: Vec<crate::OdfPropertyForm>,
    control_forms: Vec<crate::OdfControlForm>,
    interactive_forms: Vec<crate::OdfInteractiveForm>,
    selection_forms: Vec<crate::OdfSelectionForm>,
    visual_forms: Vec<crate::OdfVisualForm>,
    generic_forms: Vec<crate::OdfGenericForm>,
    password_file_forms: Vec<crate::OdfPasswordFileForm>,
    image_frame_forms: Vec<crate::OdfImageFrameForm>,
    value_range_forms: Vec<crate::OdfValueRangeForm>,
    typed_value_forms: Vec<crate::OdfTypedValueForm>,
    grid_forms: Vec<crate::OdfGridForm>,
    connection_resource_forms: Vec<crate::OdfConnectionResourceForm>,
    metadata: Metadata,
    paragraph_tab_styles: Vec<crate::ParagraphStyleTabStops>,
    paragraph_drop_cap_styles: Vec<crate::ParagraphStyleDropCap>,
    list_level_label_alignments: Vec<crate::ListStyleLevelLabelAlignment>,
    paragraph_flow_styles: Vec<crate::ParagraphStyleFlow>,
    table_row_property_styles: Vec<crate::TableRowStyleProperties>,
    table_property_styles: Vec<crate::TableStyleProperties>,
    section_property_styles: Vec<crate::SectionStyleProperties>,
    section_names: std::collections::HashSet<String>,
    section_xml_ids: std::collections::HashSet<String>,
    page_layout_columns: Vec<(String, crate::StyleColumns)>,
    page_layout_footnote_separators: Vec<(String, crate::StyleFootnoteSeparator)>,
    page_layout_header_footer_properties: Vec<(
        String,
        crate::PageHeaderFooterRegion,
        crate::HeaderFooterStyleProperties,
    )>,
    notes_configurations: crate::OdfNotesConfigurations,
    line_numbering_configuration: Option<crate::OdfLineNumberingConfiguration>,
}

impl Default for DocumentBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl DocumentBuilder {
    /// Create a new document builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::DocumentBuilder;
    ///
    /// let builder = DocumentBuilder::new();
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
            table_row_property_styles: Vec::new(),
            table_property_styles: Vec::new(),
            section_property_styles: Vec::new(),
            section_names: std::collections::HashSet::new(),
            section_xml_ids: std::collections::HashSet::new(),
            page_layout_columns: Vec::new(),
            page_layout_footnote_separators: Vec::new(),
            page_layout_header_footer_properties: Vec::new(),
            notes_configurations: crate::OdfNotesConfigurations::default(),
            line_numbering_configuration: None,
        }
    }

    /// Add a paragraph containing one validated, inert dynamic text field.
    pub fn add_dynamic_text_field(
        &mut self,
        field: &crate::elements::field::OdfDynamicTextField,
    ) -> Result<&mut Self> {
        let mut paragraph = Paragraph::new();
        paragraph.add_dynamic_text_field(field)?;
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(self)
    }

    /// Return footnote and endnote configurations emitted to `styles.xml`.
    pub fn notes_configurations(&self) -> &crate::OdfNotesConfigurations {
        &self.notes_configurations
    }

    /// Replace the validated footnote and endnote configurations.
    pub fn set_notes_configurations(
        &mut self,
        configurations: crate::OdfNotesConfigurations,
    ) -> Result<&mut Self> {
        configurations.validate()?;
        self.notes_configurations = configurations;
        Ok(self)
    }

    /// Set one validated note-class configuration.
    pub fn set_notes_configuration(
        &mut self,
        configuration: crate::OdfNotesConfiguration,
    ) -> Result<&mut Self> {
        configuration.validate()?;
        match configuration.note_class {
            crate::OdfNoteClass::Footnote => {
                self.notes_configurations.footnote = Some(configuration)
            },
            crate::OdfNoteClass::Endnote => self.notes_configurations.endnote = Some(configuration),
        }
        Ok(self)
    }

    /// Return the optional document line-numbering configuration to emit.
    ///
    /// The configuration is serialized as style metadata only. Building a
    /// document never calculates page or line numbers.
    pub fn line_numbering_configuration(&self) -> Option<&crate::OdfLineNumberingConfiguration> {
        self.line_numbering_configuration.as_ref()
    }

    /// Set the validated document line-numbering configuration to emit.
    ///
    /// This stores presentation metadata only. It never performs pagination or
    /// generates line numbers.
    pub fn set_line_numbering_configuration(
        &mut self,
        configuration: crate::OdfLineNumberingConfiguration,
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
        style: crate::ParagraphStyleTabStops,
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
        style: crate::ParagraphStyleDropCap,
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
            .find(|tabs| crate::paragraph_drop_cap::same_style_identity(&style, tabs))
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
        alignment: crate::ListLevelLabelAlignment,
    ) -> Result<&mut Self> {
        if !(1..=3).contains(&level) {
            return Err(litchi_core::Error::InvalidFormat(
                "generated numbered-list level must be 1..=3".to_string(),
            ));
        }
        let item = crate::ListStyleLevelLabelAlignment::new("L1", level, alignment)?;
        if let Some(old) = self
            .list_level_label_alignments
            .iter_mut()
            .find(|x| x.level == level)
        {
            *old = item
        } else {
            self.list_level_label_alignments.push(item)
        }
        Ok(self)
    }
    pub fn add_paragraph_flow_style(
        &mut self,
        style: crate::ParagraphStyleFlow,
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

    /// Add a named or default table-row style carrying typed row properties.
    pub fn add_table_row_property_style(
        &mut self,
        style: crate::TableRowStyleProperties,
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

    /// Add a named or default table style carrying typed table properties.
    pub fn add_table_property_style(
        &mut self,
        style: crate::TableStyleProperties,
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
        columns: crate::StyleColumns,
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
        separator: crate::StyleFootnoteSeparator,
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
        region: crate::PageHeaderFooterRegion,
        properties: crate::HeaderFooterStyleProperties,
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
    /// use litchi_odf::DocumentBuilder;
    /// use litchi_core::Metadata;
    ///
    /// let mut builder = DocumentBuilder::new();
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
    /// use litchi_odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// Hyperlink targets are serialized as inert `text:a` XLink attributes;
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
    /// use litchi_odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// * `spans` - Vector of (text, style_name) tuples for formatted text
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// use litchi_odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// use litchi_odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// use litchi_odf::{DocumentBuilder, Paragraph};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// use litchi_odf::DocumentBuilder;
    /// use litchi_odf::elements::text::Heading;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// use litchi_odf::DocumentBuilder;
    /// use litchi_odf::elements::text::{List, ListItem};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
    /// use litchi_odf::{DocumentBuilder, Table};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
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
        annotation: &crate::RubyAnnotation,
    ) -> Result<&mut Self> {
        annotation.validate()?;
        let body = self.generate_content_body();
        let xml = format!(
            r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</office:text>"#
        );
        crate::insert_ruby_annotation_xml(&xml, paragraph_index, annotation)?;
        self.ruby_annotations.push(RubyAnnotationInsertion::Append {
            paragraph_index,
            annotation: annotation.clone(),
        });
        Ok(self)
    }

    /// Wrap one UTF-8 text-node range in a selected paragraph with ruby.
    ///
    /// The range follows `wrap_ruby_annotation_xml`: it must be non-empty,
    /// fit inside one text/CDATA/entity node, and equal the annotation's
    /// plain-text base. The builder emits queued ruby mutations in call order
    /// without splitting surrounding inline markup or evaluating any active
    /// document content.
    pub fn wrap_ruby_annotation(
        &mut self,
        paragraph_index: usize,
        range: Range<usize>,
        annotation: &crate::RubyAnnotation,
    ) -> Result<&mut Self> {
        annotation.validate()?;
        let body = self.generate_content_body();
        let xml = format!(
            r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</office:text>"#
        );
        crate::wrap_ruby_annotation_xml(&xml, paragraph_index, range.clone(), annotation)?;
        self.ruby_annotations.push(RubyAnnotationInsertion::Wrap {
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
    /// `OdfNoteBodyContent` retains its validated structured body. No field,
    /// link, script, macro, event listener, or embedded payload is evaluated.
    pub fn add_note(&mut self, paragraph_index: usize, note: &crate::Note) -> Result<&mut Self> {
        note.validate()?;
        let body = self.generate_content_body();
        let xml = format!(
            r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</office:text>"#
        );
        crate::insert_note_xml(&xml, paragraph_index, note)?;
        self.notes.push((paragraph_index, note.clone()));
        Ok(self)
    }

    /// Add a named ODF ruby style definition to `styles.xml`.
    pub fn add_ruby_style(&mut self, style: crate::RubyStyle) -> Result<&mut Self> {
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

    /// Add a minimal inert form containing typed custom properties.
    pub fn add_property_form(&mut self, form: &crate::OdfPropertyForm) -> Result<&mut Self> {
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
    pub fn add_control_form(&mut self, form: &crate::OdfControlForm) -> Result<&mut Self> {
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
    pub fn add_interactive_form(&mut self, form: &crate::OdfInteractiveForm) -> Result<&mut Self> {
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
    pub fn add_selection_form(&mut self, form: &crate::OdfSelectionForm) -> Result<&mut Self> {
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
    pub fn add_visual_form(&mut self, form: &crate::OdfVisualForm) -> Result<&mut Self> {
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
    pub fn add_generic_form(&mut self, form: &crate::OdfGenericForm) -> Result<&mut Self> {
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
        form: &crate::OdfPasswordFileForm,
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
    pub fn add_image_frame_form(&mut self, form: &crate::OdfImageFrameForm) -> Result<&mut Self> {
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

    /// Generate the content.xml body
    fn generate_content_body(&self) -> String {
        let mut estimated = 256usize;
        estimated += self.elements.len() * 96;
        estimated += self
            .elements
            .iter()
            .map(|e| match e {
                DocumentElement::Paragraph(p) => p.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Heading(h) => h.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Table(_) => 256,
                DocumentElement::List(_) => 256,
                DocumentElement::Section(xml) => xml.len(),
            })
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        if !self.property_forms.is_empty()
            || !self.control_forms.is_empty()
            || !self.interactive_forms.is_empty()
            || !self.selection_forms.is_empty()
            || !self.visual_forms.is_empty()
            || !self.generic_forms.is_empty()
            || !self.password_file_forms.is_empty()
            || !self.image_frame_forms.is_empty()
            || !self.value_range_forms.is_empty()
            || !self.typed_value_forms.is_empty()
            || !self.grid_forms.is_empty()
            || !self.connection_resource_forms.is_empty()
        {
            body.push_str("<office:forms>");
            for form in &self.property_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder property form"),
                );
            }
            for form in &self.control_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder control form"),
                );
            }
            for form in &self.interactive_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder interactive form"),
                );
            }
            for form in &self.selection_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder selection form"),
                );
            }
            for form in &self.visual_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder visual form"),
                );
            }
            for form in &self.generic_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder generic form"),
                );
            }
            for form in &self.password_file_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder password/file form"),
                );
            }
            for form in &self.image_frame_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder image-frame form"),
                );
            }
            for form in &self.value_range_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder value-range form"),
                );
            }
            for form in &self.typed_value_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder typed-value form"),
                );
            }
            for form in &self.grid_forms {
                body.push_str(&form.to_xml_fragment().expect("validated builder grid form"));
            }
            for form in &self.connection_resource_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder connection-resource form"),
                );
            }
            body.push_str("</office:forms>");
        }

        // Add all elements in order they were added
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Heading(heading) => {
                    let elem: crate::elements::element::Element = heading.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::List(list) => {
                    let elem: crate::elements::element::Element = list.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Section(section) => body.push_str(section),
            }
        }

        for index in &self.text_indexes {
            body.push_str(index);
        }

        if !self.text_index_marks.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, mark) in &self.text_index_marks {
                wrapped = crate::odt::insert_text_index_mark_xml(&wrapped, *paragraph_index, mark)
                    .expect("validated builder index mark");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        if !self.reference_marks.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, mark) in &self.reference_marks {
                wrapped = crate::odt::insert_reference_mark_xml(&wrapped, *paragraph_index, mark)
                    .expect("validated builder reference mark");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        if !self.bookmark_targets.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, target) in &self.bookmark_targets {
                wrapped = crate::insert_bookmark_xml(&wrapped, *paragraph_index, target)
                    .expect("validated builder bookmark target");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        if !self.ruby_annotations.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for insertion in &self.ruby_annotations {
                wrapped = match insertion {
                    RubyAnnotationInsertion::Append {
                        paragraph_index,
                        annotation,
                    } => crate::insert_ruby_annotation_xml(&wrapped, *paragraph_index, annotation)
                        .expect("validated builder ruby annotation"),
                    RubyAnnotationInsertion::Wrap {
                        paragraph_index,
                        range,
                        annotation,
                    } => crate::wrap_ruby_annotation_xml(
                        &wrapped,
                        *paragraph_index,
                        range.clone(),
                        annotation,
                    )
                    .expect("validated builder ruby annotation range"),
                };
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        if !self.notes.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, note) in &self.notes {
                wrapped = crate::insert_note_xml(&wrapped, *paragraph_index, note)
                    .expect("validated builder note");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        body
    }

    /// Generate the complete content.xml
    fn generate_content_xml(&self) -> String {
        let body = self.generate_content_body();

        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with metadata
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();

        let mut meta = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><meta:creation-date>{}</meta:creation-date><dc:date>{}</dc:date>"#,
            now, now
        );

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str(&format!("<dc:title>{}</dc:title>", escape_xml(title)));
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str(&format!("<dc:creator>{}</dc:creator>", escape_xml(author)));
        }

        if let Some(ref subject) = self.metadata.subject {
            meta.push_str(&format!("<dc:subject>{}</dc:subject>", escape_xml(subject)));
        }

        if let Some(ref description) = self.metadata.description {
            meta.push_str(&format!(
                "<dc:description>{}</dc:description>",
                escape_xml(description)
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            meta.push_str(&format!(
                "<meta:keyword>{}</meta:keyword>",
                escape_xml(keywords)
            ));
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Generate styles.xml with list styles
    fn generate_styles_xml(&self) -> String {
        let mut xml = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:font-face-decls/><office:styles><!-- Numbered list style --><text:list-style style:name="L1"><text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.27cm" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.905cm" fo:text-indent="-0.635cm" fo:margin-left="1.905cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2.54cm" fo:text-indent="-0.635cm" fo:margin-left="2.54cm"/></style:list-level-properties></text:list-level-style-number></text:list-style></office:styles><office:automatic-styles/><office:master-styles/></office:document-styles>"#.to_string();
        if !self.paragraph_tab_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_tab_styles
                .iter()
                .map(|style| {
                    self.paragraph_drop_cap_styles
                        .iter()
                        .find(|cap| crate::paragraph_drop_cap::same_style_identity(cap, style))
                        .map_or_else(
                            || {
                                style
                                    .to_xml_fragment()
                                    .expect("validated paragraph tab style")
                            },
                            |cap| {
                                crate::paragraph_drop_cap::merge_with_tab_style(style, cap)
                                    .expect("validated merged paragraph style")
                            },
                        )
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_drop_cap_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_drop_cap_styles
                .iter()
                .filter(|cap| {
                    !self
                        .paragraph_tab_styles
                        .iter()
                        .any(|tabs| crate::paragraph_drop_cap::same_style_identity(cap, tabs))
                })
                .map(|style| {
                    style
                        .to_xml_fragment()
                        .expect("validated paragraph drop-cap style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_flow_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_flow_styles
                .iter()
                .map(|x| x.to_xml_fragment().expect("validated paragraph flow style"))
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_row_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .table_row_property_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated table-row property style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .table_property_styles
                .iter()
                .map(|x| x.to_xml_fragment().expect("validated table property style"))
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.section_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .section_property_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated section property style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if let Some(configuration) = &self.line_numbering_configuration {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragment = configuration
                .to_xml()
                .expect("validated line-numbering configuration");
            xml.insert_str(insertion, &fragment);
        }
        if self.notes_configurations.footnote.is_some()
            || self.notes_configurations.endnote.is_some()
        {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .notes_configurations
                .to_xml_fragment()
                .expect("validated notes configurations");
            xml.insert_str(insertion, &fragments);
        }
        if !self.ruby_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .ruby_styles
                .iter()
                .map(|style| style.to_xml_fragment().expect("validated ruby style"))
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.page_layout_columns.is_empty()
            || !self.page_layout_footnote_separators.is_empty()
            || !self.page_layout_header_footer_properties.is_empty()
        {
            let mut fragments = self
                .page_layout_columns
                .iter()
                .map(|(name, columns)| {
                    let mut fragment = columns
                        .to_page_layout_fragment(name)
                        .expect("validated column page layout");
                    if let Some((_, separator)) = self
                        .page_layout_footnote_separators
                        .iter()
                        .find(|(separator_name, _)| separator_name == name)
                    {
                        let insertion = fragment
                            .find("</style:page-layout-properties>")
                            .expect("static column page layout fragment");
                        fragment.insert_str(
                            insertion,
                            &separator
                                .to_xml_fragment()
                                .expect("validated footnote separator"),
                        );
                    }
                    for (_, region, properties) in self
                        .page_layout_header_footer_properties
                        .iter()
                        .filter(|(property_name, _, _)| property_name == name)
                    {
                        let insertion = fragment
                            .rfind("</style:page-layout>")
                            .expect("page layout fragment");
                        fragment.insert_str(
                            insertion,
                            &properties
                                .to_region_fragment(*region)
                                .expect("validated header/footer properties"),
                        );
                    }
                    fragment
                })
                .collect::<String>();
            for (name, separator) in &self.page_layout_footnote_separators {
                if !self
                    .page_layout_columns
                    .iter()
                    .any(|(column_name, _)| column_name == name)
                {
                    let mut fragment = separator
                        .to_page_layout_fragment(name)
                        .expect("validated footnote separator page layout");
                    for (_, region, properties) in self
                        .page_layout_header_footer_properties
                        .iter()
                        .filter(|(property_name, _, _)| property_name == name)
                    {
                        let insertion = fragment
                            .rfind("</style:page-layout>")
                            .expect("page layout fragment");
                        fragment.insert_str(
                            insertion,
                            &properties
                                .to_region_fragment(*region)
                                .expect("validated header/footer properties"),
                        );
                    }
                    fragments.push_str(&fragment);
                }
            }
            for (index, (name, _, _)) in
                self.page_layout_header_footer_properties.iter().enumerate()
            {
                if self.page_layout_columns.iter().any(|(n, _)| n == name)
                    || self
                        .page_layout_footnote_separators
                        .iter()
                        .any(|(n, _)| n == name)
                    || self.page_layout_header_footer_properties[..index]
                        .iter()
                        .any(|(n, _, _)| n == name)
                {
                    continue;
                }
                let mut fragment = format!(
                    "<style:page-layout style:name=\"{}\">",
                    litchi_core::xml::escape_xml(name)
                );
                for (_, region, properties) in self
                    .page_layout_header_footer_properties
                    .iter()
                    .filter(|(n, _, _)| n == name)
                {
                    fragment.push_str(
                        &properties
                            .to_region_fragment(*region)
                            .expect("validated header/footer properties"),
                    );
                }
                fragment.push_str("</style:page-layout>");
                fragments.push_str(&fragment);
            }
            xml = xml.replacen(
                "<office:automatic-styles/>",
                &format!("<office:automatic-styles>{fragments}</office:automatic-styles>"),
                1,
            );
        }
        for alignment in &self.list_level_label_alignments {
            xml = crate::list_label_alignment::replace_list_level_label_alignment_xml(
                &xml, alignment,
            )
            .expect("validated generated list alignment");
        }
        xml
    }

    /// Build the document and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype("application/vnd.oasis.opendocument.text")?;

        // Add content.xml
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml with list styles
        let styles_xml = self.generate_styles_xml();
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }

    /// Build and save the document to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// builder.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn test_document_builder_new() {
        let builder = DocumentBuilder::new();
        assert!(builder.elements.is_empty());
    }

    #[test]
    fn test_document_builder_default() {
        let builder: DocumentBuilder = Default::default();
        assert!(builder.elements.is_empty());
    }

    #[test]
    fn test_add_paragraph() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Hello, World!").unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn hyperlink_authoring_round_trips_through_an_odt_package() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_hyperlink("https://example.test/a?x=1&y=2", "Example & link")
            .unwrap();
        let mut configured = Hyperlink::with_href("#bookmark", "Jump").unwrap();
        configured.set_target_frame_name("_self");
        configured.set_show(Some(crate::TextHyperlinkShow::Replace));
        configured.set_actuate(Some(crate::TextHyperlinkActuate::OnRequest));
        configured.set_title("Jump to bookmark");
        builder.add_hyperlink_element(configured).unwrap();

        let document = crate::odt::Document::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(
            document.hyperlinks().unwrap(),
            vec![
                (
                    "Example & link".to_string(),
                    "https://example.test/a?x=1&y=2".to_string(),
                ),
                ("Jump".to_string(), "#bookmark".to_string()),
            ]
        );
        let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains("xlink:type=\"simple\""));
        assert!(content.contains("office:target-frame-name=\"_self\""));
        assert!(content.contains("xlink:show=\"replace\""));
        assert!(content.contains("xlink:actuate=\"onRequest\""));
        assert!(content.contains("xlink:href=\"https://example.test/a?x=1&amp;y=2\""));

        let mut invalid = DocumentBuilder::new();
        assert!(invalid.add_hyperlink("", "missing target").is_err());
        assert!(invalid.elements.is_empty());
    }

    #[test]
    fn ruby_annotation_authoring_round_trips_through_an_odt_package() {
        let style = crate::RubyStyle::new(
            "RubyAbove",
            Some(crate::RubyProperties {
                position: Some(crate::RubyPosition::Above),
                alignment: Some(crate::RubyAlignment::Center),
            }),
        )
        .unwrap();
        let annotation = crate::RubyAnnotation::new(
            Some(style.name.clone()),
            crate::RubyBase::from_text("漢").unwrap(),
            "かん",
            None,
        )
        .unwrap();

        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Read ").unwrap();
        builder.add_ruby_style(style.clone()).unwrap();
        assert!(builder.add_ruby_style(style.clone()).is_err());
        builder.add_ruby_annotation(0, &annotation).unwrap();

        let document = crate::odt::Document::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(document.ruby_styles().unwrap().styles, vec![style]);
        assert_eq!(
            document.ruby_annotations().unwrap().annotations,
            vec![annotation.clone()]
        );
        let rubies = document.rubies().unwrap();
        let ruby = rubies.first().unwrap();
        assert_eq!(ruby.base(), "漢");
        assert_eq!(ruby.text(), "かん");

        let mut invalid = DocumentBuilder::new();
        assert!(invalid.add_ruby_annotation(0, &annotation).is_err());
        assert!(invalid.elements.is_empty());
    }

    #[test]
    fn ruby_range_annotation_authoring_round_trips_through_an_odt_package() {
        let annotation =
            crate::RubyAnnotation::new(None, crate::RubyBase::from_text("字").unwrap(), "じ", None)
                .unwrap();
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Read 漢字").unwrap();
        let start = "Read 漢".len();
        builder
            .wrap_ruby_annotation(0, start..start + "字".len(), &annotation)
            .unwrap();

        let document = crate::odt::Document::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(
            document.ruby_annotations().unwrap().annotations,
            vec![annotation]
        );
    }

    #[test]
    fn note_authoring_round_trips_through_an_odt_package() {
        let mut note = crate::Note::new(crate::NoteClass::Footnote, "1", "First\nSecond").unwrap();
        note.set_id(Some("note-1".to_string())).unwrap();
        note.set_label(Some("*".to_string())).unwrap();

        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Body text").unwrap();
        builder.add_note(0, &note).unwrap();

        let document = crate::odt::Document::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(document.notes().unwrap(), vec![note.clone()]);
        assert_eq!(document.footnotes().unwrap(), vec![note]);
        assert!(document.endnotes().unwrap().is_empty());

        let mut invalid = DocumentBuilder::new();
        invalid.add_paragraph("Only paragraph").unwrap();
        assert!(
            invalid
                .add_note(
                    1,
                    &crate::Note::new(crate::NoteClass::Endnote, "i", "No").unwrap()
                )
                .is_err()
        );
        assert!(invalid.notes.is_empty());
    }

    #[test]
    fn test_add_heading() {
        let mut builder = DocumentBuilder::new();
        builder.add_heading("Chapter 1", 1).unwrap();
        builder.add_heading("Section 1.1", 2).unwrap();
        assert_eq!(builder.elements.len(), 2);
    }

    #[test]
    fn test_add_heading_invalid_level() {
        let mut builder = DocumentBuilder::new();
        let result = builder.add_heading("Invalid", 0);
        assert!(result.is_err());

        let result = builder.add_heading("Invalid", 7);
        assert!(result.is_err());
    }

    #[test]
    fn test_add_rich_paragraph() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_rich_paragraph(vec![
                ("This is ", None),
                ("bold", Some("Bold")),
                (" text.", None),
            ])
            .unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_bulleted_list() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_bulleted_list(vec!["Item 1", "Item 2", "Item 3"])
            .unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_numbered_list() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_numbered_list(vec!["First", "Second", "Third"])
            .unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_paragraph_element() {
        let mut builder = DocumentBuilder::new();
        let mut para = Paragraph::new();
        para.set_text("Custom paragraph");
        builder.add_paragraph_element(para).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_heading_element() {
        let mut builder = DocumentBuilder::new();
        let mut heading = Heading::new(1);
        heading.set_text("Custom heading");
        builder.add_heading_element(heading).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_list_element() {
        let mut builder = DocumentBuilder::new();
        let mut list = List::new();
        let mut item = ListItem::new();
        item.set_text("Item");
        list.add_item(item);
        builder.add_list_element(list).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_table() {
        let mut builder = DocumentBuilder::new();
        let mut table = Table::new();
        table.set_name("Table1");
        builder.add_table(table).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_set_metadata() {
        let mut builder = DocumentBuilder::new();
        let metadata = Metadata {
            title: Some("Test Title".to_string()),
            author: Some("Test Author".to_string()),
            subject: Some("Test Subject".to_string()),
            description: Some("Test Description".to_string()),
            keywords: Some("test, keywords".to_string()),
            ..Metadata::default()
        };
        builder.set_metadata(metadata);

        assert_eq!(builder.metadata.title, Some("Test Title".to_string()));
        assert_eq!(builder.metadata.author, Some("Test Author".to_string()));
    }

    #[test]
    fn test_generate_content_body() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Paragraph 1").unwrap();
        builder.add_heading("Heading", 1).unwrap();

        let body = builder.generate_content_body();
        assert!(body.contains("Paragraph 1"));
        assert!(body.contains("Heading"));
    }

    #[test]
    fn test_generate_content_xml() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test").unwrap();

        let xml = builder.generate_content_xml();
        assert!(xml.starts_with(r#"<?xml version="1.0" encoding="UTF-8"?"#));
        assert!(xml.contains("office:document-content"));
        assert!(xml.contains("office:text"));
        assert!(xml.contains("Test"));
    }

    #[test]
    fn test_generate_meta_xml() {
        let mut builder = DocumentBuilder::new();
        builder.metadata.title = Some("My Title".to_string());
        builder.metadata.author = Some("My Author".to_string());
        builder.metadata.subject = Some("My Subject".to_string());
        builder.metadata.description = Some("My Description".to_string());
        builder.metadata.keywords = Some("my, keywords".to_string());

        let meta_xml = builder.generate_meta_xml();
        assert!(meta_xml.contains("office:document-meta"));
        assert!(meta_xml.contains("Litchi/"));
        assert!(meta_xml.contains("My Title"));
        assert!(meta_xml.contains("My Author"));
        assert!(meta_xml.contains("My Subject"));
        assert!(meta_xml.contains("My Description"));
        assert!(meta_xml.contains("my, keywords"));
    }

    #[test]
    fn test_generate_styles_xml() {
        let builder = DocumentBuilder::new();
        let styles_xml = builder.generate_styles_xml();
        assert!(styles_xml.contains("office:document-styles"));
        assert!(styles_xml.contains("L1")); // Numbered list style
    }

    #[test]
    fn line_numbering_configuration_round_trips_through_an_odt_package() {
        let configuration = crate::OdfLineNumberingConfiguration {
            number_lines: Some(true),
            number_format: Some(crate::OdfLineNumberFormat::UpperAlpha),
            letter_sync: Some(true),
            style_name: Some("LineNumbers".to_string()),
            increment: Some(5),
            number_position: Some(crate::OdfLineNumberPosition::Outer),
            offset: Some(crate::OdfNonNegativeLength::new("0.25in").unwrap()),
            count_empty_lines: Some(true),
            count_in_text_boxes: Some(false),
            restart_on_page: Some(true),
            separator: Some(crate::OdfLineNumberingSeparator {
                increment: Some(10),
                text: " / ".to_string(),
            }),
        };

        let mut builder = DocumentBuilder::new();
        assert!(builder.line_numbering_configuration().is_none());
        builder
            .set_line_numbering_configuration(configuration.clone())
            .unwrap();
        assert_eq!(builder.line_numbering_configuration(), Some(&configuration));
        builder.clear_line_numbering_configuration();
        assert!(builder.line_numbering_configuration().is_none());
        builder
            .set_line_numbering_configuration(configuration.clone())
            .unwrap();

        let bytes = builder.build().unwrap();
        let document = crate::odt::Document::from_bytes(bytes.clone()).unwrap();
        assert_eq!(
            document.line_numbering_configuration().unwrap(),
            Some(configuration.clone())
        );
        let package = crate::OpenDocumentPackage::from_bytes(bytes).unwrap();
        assert_eq!(
            package.line_numbering_configuration().unwrap(),
            Some(configuration)
        );
    }

    #[test]
    fn test_build() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test content").unwrap();

        let result = builder.build();
        assert!(result.is_ok());
        let bytes = result.unwrap();
        assert!(!bytes.is_empty());
        // Check it's a valid ZIP (starts with PK)
        assert_eq!(&bytes[0..2], b"PK");
    }

    #[test]
    fn test_save() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.odt");

        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test content").unwrap();

        let result = builder.save(&path);
        assert!(result.is_ok());
        assert!(path.exists());

        // Verify the file is a valid ZIP
        let bytes = std::fs::read(&path).unwrap();
        assert_eq!(&bytes[0..2], b"PK");
    }

    #[test]
    fn test_chained_builder_api() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_heading("Title", 1)
            .unwrap()
            .add_paragraph("Introduction")
            .unwrap()
            .add_bulleted_list(vec!["Point 1", "Point 2"])
            .unwrap()
            .add_numbered_list(vec!["Step 1", "Step 2"])
            .unwrap();

        assert_eq!(builder.elements.len(), 4);
    }

    #[test]
    fn test_document_element_clone() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test").unwrap();

        let cloned = builder.elements[0].clone();
        match (&builder.elements[0], &cloned) {
            (DocumentElement::Paragraph(_), DocumentElement::Paragraph(_)) => {},
            _ => panic!("Clone mismatch"),
        }
    }

    #[test]
    fn test_document_element_debug() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test").unwrap();

        let debug_str = format!("{:?}", builder.elements[0]);
        assert!(debug_str.contains("Paragraph"));
    }

    #[test]
    fn test_complete_document() {
        let mut builder = DocumentBuilder::new();

        // Set metadata
        let metadata = Metadata {
            title: Some("Complete Document".to_string()),
            author: Some("Test Author".to_string()),
            ..Metadata::default()
        };
        builder.set_metadata(metadata);

        // Add various elements
        builder.add_heading("Title", 1).unwrap();
        builder.add_paragraph("This is a paragraph.").unwrap();
        builder
            .add_rich_paragraph(vec![
                ("Normal ", None),
                ("styled", Some("Emphasis")),
                (" text", None),
            ])
            .unwrap();
        builder
            .add_bulleted_list(vec!["Bullet 1", "Bullet 2"])
            .unwrap();
        builder
            .add_numbered_list(vec!["Number 1", "Number 2"])
            .unwrap();

        // Build and verify
        let result = builder.build();
        assert!(result.is_ok());
    }

    #[test]
    fn test_empty_document_build() {
        let builder = DocumentBuilder::new();
        let result = builder.build();
        assert!(result.is_ok());
    }

    #[test]
    fn test_heading_levels() {
        let mut builder = DocumentBuilder::new();
        for level in 1..=6 {
            builder
                .add_heading(&format!("Level {}", level), level)
                .unwrap();
        }
        assert_eq!(builder.elements.len(), 6);
    }

    #[test]
    fn test_list_with_empty_items() {
        let mut builder = DocumentBuilder::new();
        builder.add_bulleted_list(vec![]).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }
}

impl DocumentBuilder {
    /// Add a validated value-range form to the document.
    pub fn add_value_range_form(&mut self, form: &crate::OdfValueRangeForm) -> Result<&mut Self> {
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

impl DocumentBuilder {
    /// Add a validated form containing formatted-text, number, date, or time controls.
    pub fn add_typed_value_form(&mut self, form: &crate::OdfTypedValueForm) -> Result<&mut Self> {
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

impl DocumentBuilder {
    /// Adds a form whose final child is an inert `form:connection-resource`.
    pub fn add_connection_resource_form(
        &mut self,
        form: &crate::OdfConnectionResourceForm,
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

    pub fn add_grid_form(&mut self, form: &crate::OdfGridForm) -> Result<&mut Self> {
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
