//! Semantic ODT document queries and the public read facade.

use crate::core::{Meta, Styles};
use crate::elements::style::StyleRegistry;
use crate::elements::table::Table as ElementTable;
use crate::elements::text::{Paragraph as ElementParagraph, TextElements};
use litchi_core::{Metadata, Result};

use crate::header_footer::{Master, read};
use crate::page_layout::{PageLayout, parse_page_layouts};
use crate::page_sequence::{Sequence, parse_page_sequence};

use super::codec::{parse_bookmark_names, parse_hyperlinks, parse_image_references};
use super::model::Document;

impl Document {
    crate::package::scripts::script_facade_methods!();
    crate::package::annotation::annotation_facade_methods!(Text);

    /// Extract all text content from the document.
    ///
    /// This method extracts plain text from all paragraphs, headings, and text elements
    /// in the document, preserving paragraph breaks. Formatting, styles, and non-text
    /// elements are omitted.
    ///
    /// # Performance
    ///
    /// This method parses the XML on each call. For repeated access, consider caching
    /// the result.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let text = doc.text()?;
    /// println!("Text content:\n{}", text);
    /// # Ok(())
    /// # }
    /// ```
    pub fn text(&self) -> Result<String> {
        TextElements::extract_text(self.content.xml_content())
    }

    /// Get the number of paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraph count: {}", count);
    /// # Ok(())
    /// # }
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        Ok(self.paragraphs()?.len())
    }

    /// Get all paragraphs in the document as structured elements.
    ///
    /// Returns a vector of `Paragraph` elements that can be used to access
    /// individual paragraph content, styles, and attributes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let paragraphs = doc.paragraphs()?;
    ///
    /// for para in paragraphs {
    ///     println!("Paragraph: {}", para.text()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<ElementParagraph>> {
        TextElements::parse_paragraphs(self.content.xml_content())
    }

    /// Get all tables in the document.
    ///
    /// Returns a vector of `Table` elements representing all tables found
    /// in the document body.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let tables = doc.tables()?;
    ///
    /// println!("Found {} tables", tables.len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn tables(&self) -> Result<Vec<ElementTable>> {
        use crate::elements::table::TableElements;
        TableElements::parse_tables_from_content(self.content.xml_content())
    }

    /// Get all document elements (paragraphs, headings, and tables) in document order.
    ///
    /// This method extracts both paragraphs (including headings) and tables, interleaved
    /// in the order they appear in the document. This provides a more efficient way to
    /// iterate through document content than calling `paragraphs()` and `tables()` separately,
    /// and preserves the exact document order.
    ///
    /// # Returns
    ///
    /// A vector of `OrderElement` containing all paragraphs, headings, tables, and
    /// lists in the order they appear in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    /// use litchi_odt::elements::parser::OrderElement;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let elements = doc.elements()?;
    ///
    /// for element in elements {
    ///     match element {
    ///         OrderElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text()?);
    ///         },
    ///         OrderElement::NumberedParagraph(para) => {
    ///             println!("Numbered paragraph: {}", para.text()?);
    ///         },
    ///         OrderElement::Heading(heading) => {
    ///             println!("Heading: {}", heading.text()?);
    ///         },
    ///         OrderElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count()?);
    ///         },
    ///         OrderElement::List(_) => {
    ///             println!("List element");
    ///         },
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn elements(&self) -> Result<Vec<crate::elements::parser::OrderElement>> {
        use crate::elements::parser::Parser;

        // Parse all elements in document order using the generic ODF parser
        Parser::parse_elements_in_order(self.content.xml_content())
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file, including title, author,
    /// creation date, modification date, word count, and other document properties.
    ///
    /// # Returns
    ///
    /// A `Metadata` struct containing all available metadata fields. Fields that
    /// are not present in the document will be `None`.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let metadata = doc.metadata()?;
    ///
    /// if let Some(title) = &metadata.title {
    ///     println!("Title: {}", title);
    /// }
    /// if let Some(author) = &metadata.author {
    ///     println!("Author: {}", author);
    /// }
    /// if let Some(word_count) = metadata.word_count {
    ///     println!("Words: {}", word_count);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            meta.try_extract_metadata()
        } else {
            Ok(Metadata::default())
        }
    }

    /// Get the complete format-specific OpenDocument metadata model.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.meta.as_ref().map(Meta::odf_metadata).transpose()
    }

    /// Get the style registry for this document.
    ///
    /// The style registry contains all styles defined in the document,
    /// including both automatic styles and named styles.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let styles = doc.styles();
    /// // Use the style registry to query styles...
    /// # Ok(())
    /// # }
    /// ```
    pub fn styles(&self) -> &StyleRegistry {
        &self.style_registry
    }

    /// Get typed named ruby style definitions from `styles.xml`.
    pub fn ruby_styles(&self) -> Result<crate::ruby_family::Styles> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::parse_ruby_styles(styles.xml_content()),
        )
    }

    /// Return font-face declarations stored in `content.xml`.
    ///
    /// This preserves stored metadata only. It does not fetch linked font
    /// resources, load fonts, or inspect embedded font data.
    pub fn content_font_face_declarations(&self) -> Result<Option<crate::font_face::Declarations>> {
        crate::font_face::parse_content_font_face_declarations(self.content.xml_content())
    }

    /// Return font-face declarations stored in `styles.xml`.
    ///
    /// This preserves stored metadata only. It does not fetch linked font
    /// resources, load fonts, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<crate::font_face::Declarations>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| crate::font_face::parse_styles_font_face_declarations(styles.xml_content()),
        )
    }

    /// Return named legacy and SVG drawing gradients from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites, load external data, or render gradients.
    pub fn drawing_gradients(&self) -> Result<crate::drawing::resources::gradient::Collection> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| {
                crate::drawing::resources::gradient::parse_drawing_gradients(styles.xml_content())
            },
        )
    }

    /// Return named drawing hatch resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<crate::drawing::resources::hatch::Collection> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing::resources::hatch::parse_drawing_hatches(styles.xml_content()),
        )
    }

    /// Return named drawing stroke-dash resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing::resources::stroke_dash::Collection> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| {
                crate::drawing::resources::stroke_dash::parse_drawing_stroke_dashes(
                    styles.xml_content(),
                )
            },
        )
    }

    /// Return named drawing fill-image definitions from `styles.xml`.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites, follow links, load linked resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing::resources::fill_image::Collection> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| {
                crate::drawing::resources::fill_image::parse_drawing_fill_images(
                    styles.xml_content(),
                )
            },
        )
    }

    /// Return named drawing marker definitions from `styles.xml`.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing::resources::marker::Collection> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing::resources::marker::parse_drawing_markers(styles.xml_content()),
        )
    }

    /// Return named drawing opacity definitions from `styles.xml`.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing::resources::opacity::Collection> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| {
                crate::drawing::resources::opacity::parse_drawing_opacities(styles.xml_content())
            },
        )
    }

    /// Parse master pages and their losslessly retained headers and footers.
    pub fn master_pages(&self) -> Result<Vec<Master>> {
        self.styles
            .as_ref()
            .map_or_else(|| Ok(Vec::new()), |styles| read(styles.xml_content()))
    }

    /// Parse automatic page layouts, their properties, and header/footer styles.
    pub fn page_layouts(&self) -> Result<Vec<PageLayout>> {
        self.styles.as_ref().map_or_else(
            || Ok(Vec::new()),
            |styles| parse_page_layouts(styles.xml_content()),
        )
    }

    /// The unnamed fallback page layout (`style:default-page-layout`), when
    /// the document declares one.
    pub fn default_page_layout(&self) -> Result<Option<PageLayout>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| crate::page_layout::parse_default_page_layout(styles.xml_content()),
        )
    }

    /// Return the stored footnote and endnote presentation configurations.
    ///
    /// These style declarations are retained as metadata only. This API does
    /// not renumber notes, resolve style references, or render continuation
    /// notices.
    pub fn notes_configurations(&self) -> Result<crate::notes_configuration::Configurations> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::notes_configuration::parse(styles.xml_content()),
        )
    }

    /// Return stored outline numbering styles.
    ///
    /// These are styles metadata only. This API does not apply styles to
    /// headings, generate labels, or update tables of contents.
    pub fn outline_styles(&self) -> Result<crate::outline_style::Styles> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::outline_style::parse_outline_styles(styles.xml_content()),
        )
    }

    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// This styles metadata remains inert: the API does not generate
    /// bibliography entries, resolve citations, or access external sources.
    pub fn bibliography_configuration(
        &self,
    ) -> Result<Option<crate::bibliography_configuration::Configuration>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| {
                crate::bibliography_configuration::parse_bibliography_configuration(
                    styles.xml_content(),
                )
            },
        )
    }

    /// Return the explicit master-page assignments from `text:page-sequence`, if present.
    ///
    /// This exposes stored page metadata only. It does not resolve master-page
    /// styles, calculate page breaks, or paginate the document.
    pub fn page_sequence(&self) -> Result<Option<Sequence>> {
        parse_page_sequence(self.content.xml_content())
    }

    /// Return the stored document line-numbering configuration.
    ///
    /// This is presentation metadata from styles.xml only. It is never used to
    /// paginate the document or generate line numbers.
    pub fn line_numbering_configuration(
        &self,
    ) -> Result<Option<crate::line_numbering::Configuration>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| crate::line_numbering::parse(styles.xml_content()),
        )
    }

    /// Get resolved style properties for a given style name.
    ///
    /// This method resolves style inheritance to provide the complete set of
    /// properties that apply to elements using the specified style.
    ///
    /// # Arguments
    ///
    /// * `style_name` - Name of the style to resolve
    ///
    /// # Returns
    ///
    /// A `StyleProperties` struct containing all resolved properties for the style.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let props = doc.get_style_properties("Heading1");
    ///
    /// if let Some(font_size) = &props.text.font_size {
    ///     println!("Heading 1 font size: {}", font_size);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_style_properties(
        &self,
        style_name: &str,
    ) -> crate::elements::style::StyleProperties<'_> {
        self.style_registry.get_resolved_properties(style_name)
    }

    /// Get all tracked changes in the document.
    ///
    /// Tracked changes include insertions, deletions, and format changes made
    /// by document collaborators.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let changes = doc.track_changes()?;
    ///
    /// for change in changes {
    ///     println!("Change by {:?}: {:?}", change.author, change.change_type);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn track_changes(&self) -> Result<Vec<crate::parser::TrackChange>> {
        crate::parser::Parser::parse_track_changes(self.content.xml_content())
    }

    /// Get tracked changes together with their inert container policy metadata.
    ///
    /// Protection-key material and digest identifiers are retained for round-trip and
    /// inspection only; this method never unlocks, accepts, rejects, or evaluates changes.
    pub fn tracked_changes(&self) -> Result<crate::parser::TrackedChanges> {
        crate::parser::Parser::parse_tracked_changes(self.content.xml_content())
    }

    /// Get all comments/annotations in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let comments = doc.comments()?;
    ///
    /// for comment in comments {
    ///     println!("Comment by {:?}: {}", comment.author, comment.content);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn comments(&self) -> Result<Vec<crate::parser::Comment>> {
        crate::parser::Parser::parse_comments(self.content.xml_content())
    }

    /// Get all sections in the document.
    ///
    /// Sections are document subdivisions that can have protected content,
    /// different formatting, or special layout properties.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let sections = doc.sections()?;
    ///
    /// for section in sections {
    ///     println!("Section '{}': protected={}", section.name, section.protected);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn sections(&self) -> Result<Vec<crate::parser::Section>> {
        crate::parser::Parser::parse_sections(self.content.xml_content())
    }

    /// Get all generated text indexes, including their inert source definitions and cached bodies.
    ///
    /// This covers tables of contents and illustration, table, object, user, alphabetical,
    /// and bibliography indexes. Stored source declarations are never evaluated, and external
    /// alphabetical auto-mark files are never fetched.
    pub fn text_indexes(&self) -> Result<Vec<crate::TextIndex>> {
        crate::index::parse_text_indexes(self.content.xml_content())
    }

    /// Get point and resolved range marks that contribute entries to generated text indexes.
    pub fn text_index_marks(&self) -> Result<Vec<crate::TextIndexMark>> {
        crate::index_mark::parse_text_index_marks(self.content.xml_content())
    }

    /// Get point and range targets used by `text:reference-ref` fields.
    pub fn reference_marks(&self) -> Result<Vec<crate::ReferenceMark>> {
        crate::reference_mark::parse_reference_marks(self.content.xml_content())
    }

    /// Get all semantic footnotes and endnotes in document order.
    pub fn notes(&self) -> Result<Vec<crate::Note>> {
        crate::note::parse_notes(self.content.xml_content())
    }

    pub fn footnotes(&self) -> Result<Vec<crate::Note>> {
        Ok(self
            .notes()?
            .into_iter()
            .filter(|note| note.class() == crate::NoteClass::Footnote)
            .collect())
    }

    pub fn endnotes(&self) -> Result<Vec<crate::Note>> {
        Ok(self
            .notes()?
            .into_iter()
            .filter(|note| note.class() == crate::NoteClass::Endnote)
            .collect())
    }

    /// Get structure-preserving ruby annotations in document order.
    pub fn ruby_annotations(&self) -> Result<crate::ruby_family::Annotations> {
        crate::parse_ruby_annotations(self.content.xml_content())
    }

    /// Get simplified ruby base/pronunciation pairs in document order.
    pub fn rubies(&self) -> Result<Vec<crate::ruby::Annotation>> {
        crate::ruby::parse_rubies(self.content.xml_content())
    }

    /// Get all bookmarks in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let bookmarks = doc.bookmarks()?;
    ///
    /// for bookmark in bookmarks {
    ///     if let Some(name) = bookmark.name() {
    ///         println!("Bookmark: {}", name);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn bookmarks(&self) -> Result<Vec<crate::elements::bookmark::Bookmark>> {
        use crate::elements::bookmark::BookmarkParser;
        BookmarkParser::parse_bookmarks(self.content.xml_content())
    }

    /// Get all bookmark ranges in the document.
    ///
    /// Bookmark ranges span multiple paragraphs or sections.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let ranges = doc.bookmark_ranges()?;
    ///
    /// for range in ranges {
    ///     if range.is_complete() {
    ///         println!("Complete bookmark range: {}", range.name);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn bookmark_ranges(&self) -> Result<Vec<crate::elements::bookmark::BookmarkRange>> {
        use crate::elements::bookmark::BookmarkParser;
        BookmarkParser::parse_bookmark_ranges(self.content.xml_content())
    }

    /// Get all fields in the document.
    ///
    /// Fields are dynamic content elements like page numbers, dates, and references.
    /// Returned values are the document's cached display text; expressions, database
    /// fields, and other dynamic sources are never evaluated or executed.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let fields = doc.fields()?;
    ///
    /// for field in fields {
    ///     println!("Field type: {}, value: {}", field.field_type(), field.value());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn fields(&self) -> Result<Vec<crate::elements::field::Field>> {
        use crate::elements::field::FieldParser;
        FieldParser::parse_fields(self.content.xml_content())
    }

    /// Get conditional, hidden, and placeholder text fields in document order.
    ///
    /// Conditions and formulas are returned as inert strings. This method never
    /// evaluates them and returns only the cached display text stored in the file.
    pub fn dynamic_text_fields(&self) -> Result<Vec<crate::elements::field::DynamicTextField>> {
        use crate::elements::field::FieldParser;
        FieldParser::parse_dynamic_text_fields(self.content.xml_content())
    }

    /// Get all tables with repeated cells and rows expanded.
    ///
    /// ODF files can store repeated cells/rows compactly. This method expands
    /// them into their full representation.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let expanded_tables = doc.tables_expanded()?;
    ///
    /// for table in expanded_tables {
    ///     println!("Expanded table has {} rows", table.row_count()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn tables_expanded(&self) -> Result<Vec<crate::elements::table::Table>> {
        use crate::elements::table_expansion::TableExpander;
        let tables = self.tables()?;
        TableExpander::expand_tables(tables)
    }

    /// Extract all hyperlinks from the document
    ///
    /// Returns a vector of tuples containing (link text, URL).
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let hyperlinks = doc.hyperlinks()?;
    /// for (text, url) in hyperlinks {
    ///     println!("{}: {}", text, url);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn hyperlinks(&self) -> Result<Vec<(String, String)>> {
        parse_hyperlinks(self.content.xml_content())
    }

    /// Extract all bookmark names from the document
    ///
    /// Returns a vector of bookmark names.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let bookmark_names = doc.bookmark_names()?;
    /// for bookmark in bookmark_names {
    ///     println!("Bookmark: {}", bookmark);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn bookmark_names(&self) -> Result<Vec<String>> {
        parse_bookmark_names(self.content.xml_content())
    }

    /// Extract all linked image references from the document
    ///
    /// Returns package-local and external `xlink:href` values in document order.
    /// Images stored inline as `office:binary-data` have no path and are omitted.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let images = doc.image_paths()?;
    /// for img_path in images {
    ///     println!("Image: {}", img_path);
    ///     // You can extract the image bytes with:
    ///     // let bytes = doc.get_file(&img_path)?;
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn image_paths(&self) -> Result<Vec<String>> {
        parse_image_references(self.content.xml_content())
    }

    /// Discover referenced, inline, missing, and inert linked images.
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        let package = self.package.package()?;
        crate::media::scan_package(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            &package,
        )
    }

    /// Inspect classic forms without executing bindings, events, or external resources.
    pub fn forms(&self) -> Result<crate::form::Forms> {
        let mut parts = vec![(self.content.xml_content(), crate::form::Part::Content)];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::form::Part::Styles));
        }
        crate::form::parse_form_parts(&parts)
    }

    pub fn rdf_graphs(&self) -> Result<Vec<crate::rdf::Graph>> {
        crate::rdf::graphs(&self.package)
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::variable_declaration::Declarations> {
        let mut parts = vec![(
            self.content.xml_content(),
            crate::variable_declaration::Part::Content,
        )];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::variable_declaration::Part::Styles));
        }
        crate::variable_declaration::parse_parts(&parts)
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::Object>> {
        let package = self.package.package()?;
        crate::embedded::scan_package(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            &package,
        )
    }

    /// Open one inert embedded chart as a standalone chart document.
    pub fn embedded_chart(&self, index: usize) -> Result<crate::odc::Document> {
        crate::package::charts::open_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )
    }
}
