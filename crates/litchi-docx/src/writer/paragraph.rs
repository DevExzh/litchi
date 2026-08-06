//! Paragraph types and implementation for DOCX documents.
use crate::error::{Error, Result};
use crate::namespace::normalize_xml_integer;
use crate::paragraph::extensions::Extensions;
use crate::{OfficeMath, OfficeMathParagraph};
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::{LineSpacing, ParagraphAlignment};
// Import other writer types
use super::bookmark::MutableBookmark;
use super::field::MutableField;
use super::hyperlink::MutableHyperlink;
use super::image::MutableInlineImage;
use super::ole_object::MutableOleObject;
use super::revision::{
    MutableRevision, ParagraphPropertyChange, RevisionKind, RevisionMetadata, RevisionTextMode,
};
use super::run::MutableRun;
use super::section::SectionProperties;
use super::smart_tag::MutableSmartTag;
use super::smartart::MutableSmartArt;
use super::textbox::MutableTextBox;
use super::vml_shape::MutableVmlShape;

/// Elements that can appear in a paragraph.
#[derive(Debug)]
pub(crate) enum ParagraphElement {
    Run(MutableRun),
    /// Display `<m:oMathPara>` content directly inside the paragraph.
    DisplayOfficeMath(OfficeMathParagraph),
    Hyperlink(MutableHyperlink),
    InlineImage(MutableInlineImage),
    /// Inline DrawingML text box (wordprocessing shape).
    TextBox(MutableTextBox),
    /// Embedded OLE/package object (`w:object` with `o:OLEObject`).
    OleObject(MutableOleObject),
    /// SmartArt (DrawingML diagram) anchor with `dgm:relIds`.
    SmartArt(MutableSmartArt),
    /// Legacy VML shape (`w:pict` with `v:rect`/`v:oval`/…).
    VmlShape(MutableVmlShape),
    /// Bookmark start marker
    BookmarkStart(MutableBookmark),
    /// Bookmark end marker (ID only)
    BookmarkEnd(u32),
    /// Field
    Field(MutableField),
    /// Run-level smart tag.
    SmartTag(MutableSmartTag),
    /// Typed tracked-change wrapper.
    Revision(MutableRevision),
}

impl ParagraphElement {
    pub(crate) fn write_placeholder(
        &self,
        xml: &mut String,
        hyperlink_index: &mut usize,
        image_index: &mut usize,
    ) -> Result<()> {
        self.write_placeholder_mode(xml, hyperlink_index, image_index, RevisionTextMode::Normal)
    }

    pub(crate) fn write_placeholder_mode(
        &self,
        xml: &mut String,
        hyperlink_index: &mut usize,
        image_index: &mut usize,
        mode: RevisionTextMode,
    ) -> Result<()> {
        match self {
            Self::Run(run) => run.to_xml_mode(xml, mode),
            Self::DisplayOfficeMath(math) => {
                xml.push_str(math.xml());
                Ok(())
            },
            Self::Hyperlink(hyperlink) => {
                let placeholder = format!("{{{{HYPERLINK_{}}}}}", *hyperlink_index);
                hyperlink.to_xml_mode(xml, Some(&placeholder), mode)?;
                *hyperlink_index += 1;
                Ok(())
            },
            Self::InlineImage(image) => {
                xml.push_str("<w:r>");
                let placeholder = format!("{{{{IMAGE_{}}}}}", *image_index);
                image.to_xml(xml, &placeholder)?;
                xml.push_str("</w:r>");
                *image_index += 1;
                Ok(())
            },
            // Text boxes carry no relationships; both modes serialize alike.
            Self::TextBox(text_box) => {
                xml.push_str("<w:r>");
                text_box.to_xml(xml)?;
                xml.push_str("</w:r>");
                Ok(())
            },
            // Without a relationship mapper the serializer emits
            // deterministic `{{OLE_*}}` placeholders.
            Self::OleObject(object) => {
                xml.push_str("<w:r>");
                object.to_xml(xml, None, None)?;
                xml.push_str("</w:r>");
                Ok(())
            },
            // Without a relationship mapper the serializer emits
            // deterministic `{{SMARTART_*}}` placeholders.
            Self::SmartArt(smartart) => {
                xml.push_str("<w:r>");
                smartart.to_xml(xml, None)?;
                xml.push_str("</w:r>");
                Ok(())
            },
            // VML shapes carry no relationships; both modes serialize alike.
            Self::VmlShape(shape) => {
                xml.push_str("<w:r>");
                shape.to_xml(xml)?;
                xml.push_str("</w:r>");
                Ok(())
            },
            Self::BookmarkStart(bookmark) => {
                xml.push_str(&bookmark.to_xml_start()?);
                Ok(())
            },
            Self::BookmarkEnd(id) => write!(xml, r#"<w:bookmarkEnd w:id="{}"/>"#, id)
                .map_err(|error| Error::Xml(error.to_string())),
            Self::Field(field) => {
                xml.push_str("<w:r>");
                xml.push_str(&field.to_xml_mode(mode)?);
                xml.push_str("</w:r>");
                Ok(())
            },
            Self::SmartTag(tag) => {
                tag.write_placeholder_mode(xml, hyperlink_index, image_index, mode)
            },
            Self::Revision(revision) => {
                revision.write_placeholder(xml, hyperlink_index, image_index)
            },
        }
    }

    pub(crate) fn write_with_rels(
        &self,
        xml: &mut String,
        rel_mapper: &crate::writer::relmap::RelationshipMapper,
        hyperlink_index: &mut usize,
        image_index: &mut usize,
    ) -> Result<()> {
        self.write_with_rels_mode(
            xml,
            rel_mapper,
            hyperlink_index,
            image_index,
            RevisionTextMode::Normal,
        )
    }

    pub(crate) fn write_with_rels_mode(
        &self,
        xml: &mut String,
        rel_mapper: &crate::writer::relmap::RelationshipMapper,
        hyperlink_index: &mut usize,
        image_index: &mut usize,
        mode: RevisionTextMode,
    ) -> Result<()> {
        match self {
            Self::Run(run) => run.to_xml_mode(xml, mode),
            Self::DisplayOfficeMath(math) => {
                xml.push_str(math.xml());
                Ok(())
            },
            Self::Hyperlink(hyperlink) => {
                if hyperlink.url.is_some() {
                    if let Some(rel_id) = rel_mapper.get_hyperlink_id(*hyperlink_index) {
                        hyperlink.to_xml_mode(xml, Some(rel_id), mode)?;
                    } else {
                        let placeholder = format!("{{{{HYPERLINK_{}}}}}", *hyperlink_index);
                        hyperlink.to_xml_mode(xml, Some(&placeholder), mode)?;
                    }
                    *hyperlink_index += 1;
                } else {
                    hyperlink.to_xml_mode(xml, None, mode)?;
                }
                Ok(())
            },
            Self::InlineImage(image) => {
                xml.push_str("<w:r>");
                if let Some(rel_id) = rel_mapper.get_image_id(*image_index) {
                    image.to_xml(xml, rel_id)?;
                } else {
                    let placeholder = format!("{{{{IMAGE_{}}}}}", *image_index);
                    image.to_xml(xml, &placeholder)?;
                }
                xml.push_str("</w:r>");
                *image_index += 1;
                Ok(())
            },
            // Text boxes carry no relationships; both modes serialize alike.
            Self::TextBox(text_box) => {
                xml.push_str("<w:r>");
                text_box.to_xml(xml)?;
                xml.push_str("</w:r>");
                Ok(())
            },
            // OLE object payload/preview relationships resolve by shape ID.
            Self::OleObject(object) => {
                xml.push_str("<w:r>");
                object.to_xml(
                    xml,
                    rel_mapper.get_ole_object_id(object.shape_id()),
                    rel_mapper.get_ole_preview_id(object.shape_id()),
                )?;
                xml.push_str("</w:r>");
                Ok(())
            },
            // SmartArt diagram relationships resolve by anchor key.
            Self::SmartArt(smartart) => {
                xml.push_str("<w:r>");
                smartart.to_xml(xml, rel_mapper.get_smart_art_ids(smartart.anchor_key()))?;
                xml.push_str("</w:r>");
                Ok(())
            },
            // VML shapes carry no relationships; both modes serialize alike.
            Self::VmlShape(shape) => {
                xml.push_str("<w:r>");
                shape.to_xml(xml)?;
                xml.push_str("</w:r>");
                Ok(())
            },
            Self::BookmarkStart(bookmark) => {
                xml.push_str(&bookmark.to_xml_start()?);
                Ok(())
            },
            Self::BookmarkEnd(id) => write!(xml, r#"<w:bookmarkEnd w:id="{}"/>"#, id)
                .map_err(|error| Error::Xml(error.to_string())),
            Self::Field(field) => {
                xml.push_str("<w:r>");
                xml.push_str(&field.to_xml_mode(mode)?);
                xml.push_str("</w:r>");
                Ok(())
            },
            Self::SmartTag(tag) => {
                tag.write_with_rels_mode(xml, rel_mapper, hyperlink_index, image_index, mode)
            },
            Self::Revision(revision) => {
                revision.write_with_rels(xml, rel_mapper, hyperlink_index, image_index)
            },
        }
    }

    pub(crate) fn collect_hyperlink_urls(&self, urls: &mut Vec<String>) {
        match self {
            Self::Hyperlink(link) => {
                if let Some(url) = &link.url {
                    urls.push(url.clone());
                }
            },
            Self::SmartTag(tag) => tag.collect_hyperlink_urls(urls),
            Self::Revision(revision) => revision.collect_hyperlink_urls(urls),
            _ => {},
        }
    }

    pub(crate) fn collect_images<'a>(
        &'a self,
        images: &mut Vec<(&'a [u8], crate::format::ImageFormat)>,
    ) {
        match self {
            Self::InlineImage(image) => images.push((image.data(), image.format())),
            Self::SmartTag(tag) => tag.collect_images(images),
            Self::Revision(revision) => revision.collect_images(images),
            _ => {},
        }
    }

    pub(crate) fn append_run_text(&self, text: &mut String) {
        match self {
            Self::Run(run) => text.push_str(&run.get_text()),
            Self::SmartTag(tag) => tag.append_run_text(text),
            Self::Revision(revision) => revision.append_run_text(text),
            _ => {},
        }
    }
}

/// A mutable paragraph in a document.
#[derive(Debug)]
pub struct MutableParagraph {
    /// Elements (runs and hyperlinks) in this paragraph
    pub(crate) elements: Vec<ParagraphElement>,
    /// Paragraph style ID
    pub(crate) style: Option<String>,
    /// Paragraph properties
    pub(crate) properties: ParagraphProperties,
    /// Word 2010 paragraph-level extension attributes.
    pub(crate) extension_values: Extensions,
    pub(crate) property_change: Option<ParagraphPropertyChange>,
}

impl MutableParagraph {
    pub(crate) fn new() -> Self {
        Self {
            elements: Vec::new(),
            style: None,
            properties: ParagraphProperties::default(),
            extension_values: Extensions::new(),
            property_change: None,
        }
    }

    /// Add a new run to the paragraph.
    pub fn add_run(&mut self) -> &mut MutableRun {
        self.elements.push(ParagraphElement::Run(MutableRun::new()));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::Run(r) => r,
            _ => unreachable!(),
        }
    }

    /// Add a run with text.
    pub fn add_run_with_text(&mut self, text: &str) -> &mut MutableRun {
        let run = self.add_run();
        run.set_text(text);
        run
    }

    /// Add an inline Office Math equation to this paragraph.
    ///
    /// The equation is emitted as `<w:r><m:oMath>…</m:oMath></w:r>`, allowing
    /// it to appear between ordinary text runs.
    pub fn add_inline_office_math(&mut self, equation: OfficeMath) -> &mut Self {
        self.add_run().set_office_math(equation);
        self
    }

    /// Parse and add an inline Office Math equation.
    pub fn add_inline_office_math_xml(&mut self, xml: impl Into<String>) -> Result<&mut Self> {
        let equation = OfficeMath::from_xml(xml)?;
        Ok(self.add_inline_office_math(equation))
    }

    /// Add a display Office Math equation to this paragraph.
    ///
    /// Display equations are enclosed in an `<m:oMathPara>` container as
    /// required by Word's display-math layout model.
    pub fn add_display_office_math(&mut self, equation: OfficeMath) -> &mut Self {
        self.add_office_math_paragraph(OfficeMathParagraph::from_equation(equation))
    }

    /// Parse and add a display Office Math equation.
    pub fn add_display_office_math_xml(&mut self, xml: impl Into<String>) -> Result<&mut Self> {
        let equation = OfficeMath::from_xml(xml)?;
        Ok(self.add_display_office_math(equation))
    }

    /// Add a fully specified display-math paragraph.
    pub fn add_office_math_paragraph(&mut self, paragraph: OfficeMathParagraph) -> &mut Self {
        self.elements
            .push(ParagraphElement::DisplayOfficeMath(paragraph));
        self
    }

    /// Parse and add a fully specified display-math paragraph.
    pub fn add_office_math_paragraph_xml(&mut self, xml: impl Into<String>) -> Result<&mut Self> {
        let paragraph = OfficeMathParagraph::from_xml(xml)?;
        Ok(self.add_office_math_paragraph(paragraph))
    }

    /// Add a hyperlink to the paragraph.
    ///
    /// # Arguments
    ///
    /// * `url` - The URL to link to
    /// * `text` - The display text for the hyperlink
    pub fn add_hyperlink(&mut self, url: &str, text: &str) -> &mut MutableHyperlink {
        self.elements
            .push(ParagraphElement::Hyperlink(MutableHyperlink::new(
                url.to_string(),  // URL first (matches MutableHyperlink::new signature)
                text.to_string(), // Text second
            )));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::Hyperlink(h) => h,
            _ => unreachable!(),
        }
    }

    /// Add an inline image to the paragraph.
    pub fn add_picture(
        &mut self,
        image_path: &str,
        width_emu: Option<i64>,
        height_emu: Option<i64>,
    ) -> Result<&mut MutableInlineImage> {
        use std::fs;
        let data = fs::read(image_path).map_err(Error::Io)?;
        let image = MutableInlineImage::from_bytes(data, width_emu, height_emu)?;
        self.elements.push(ParagraphElement::InlineImage(image));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::InlineImage(img) => Ok(img),
            _ => unreachable!(),
        }
    }

    /// Add an inline image from bytes to the paragraph.
    pub fn add_picture_from_bytes(
        &mut self,
        data: Vec<u8>,
        width_emu: Option<i64>,
        height_emu: Option<i64>,
    ) -> Result<&mut MutableInlineImage> {
        let image = MutableInlineImage::from_bytes(data, width_emu, height_emu)?;
        self.elements.push(ParagraphElement::InlineImage(image));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::InlineImage(img) => Ok(img),
            _ => unreachable!(),
        }
    }

    /// Add an inline text box to the paragraph.
    ///
    /// The text box is serialized as a DrawingML wordprocessing shape
    /// (`wps:wsp`) and reappears in the
    /// [`crate::Document::text_boxes`] inventory after save and reopen.
    pub fn add_text_box(&mut self, text_box: MutableTextBox) -> &mut MutableTextBox {
        self.elements.push(ParagraphElement::TextBox(text_box));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::TextBox(text_box) => text_box,
            _ => unreachable!(),
        }
    }

    /// Add an embedded OLE/package object to the paragraph.
    ///
    /// Crate-internal: public embedding goes through
    /// [`crate::writer::MutableDocument::add_ole_object`], which assigns
    /// and validates the shape identity first.
    pub(crate) fn add_ole_object(&mut self, object: MutableOleObject) -> &mut MutableOleObject {
        self.elements.push(ParagraphElement::OleObject(object));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::OleObject(object) => object,
            _ => unreachable!(),
        }
    }

    /// Add a SmartArt (DrawingML diagram) anchor to the paragraph.
    ///
    /// Crate-internal: public authoring goes through
    /// [`crate::writer::MutableDocument::add_smart_art`], which assigns
    /// the anchor key first.
    pub(crate) fn add_smart_art(&mut self, smartart: MutableSmartArt) -> &mut MutableSmartArt {
        self.elements.push(ParagraphElement::SmartArt(smartart));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::SmartArt(smartart) => smartart,
            _ => unreachable!(),
        }
    }

    /// Add a legacy VML shape to the paragraph.
    ///
    /// Crate-internal: public authoring goes through
    /// [`crate::writer::MutableDocument::add_vml_shape`], which assigns
    /// the shape identity first.
    pub(crate) fn add_vml_shape(&mut self, shape: MutableVmlShape) -> &mut MutableVmlShape {
        self.elements.push(ParagraphElement::VmlShape(shape));
        match self.elements.last_mut().unwrap() {
            ParagraphElement::VmlShape(shape) => shape,
            _ => unreachable!(),
        }
    }

    /// Add a bookmark start marker.
    ///
    /// Bookmarks mark named locations in the document. You must call `add_bookmark_end`
    /// with the same ID after adding content to close the bookmark.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique bookmark ID (must be unique within the document)
    /// * `name` - Bookmark name (for cross-referencing)
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// para.add_bookmark_start(1, "MyBookmark");
    /// para.add_run_with_text("Bookmarked text");
    /// para.add_bookmark_end(1);
    /// ```
    pub fn add_bookmark_start(&mut self, id: u32, name: &str) {
        let bookmark = MutableBookmark::new(id, name.to_string());
        self.elements
            .push(ParagraphElement::BookmarkStart(bookmark));
    }

    /// Add a bookmark end marker.
    ///
    /// This closes a bookmark started with `add_bookmark_start`.
    ///
    /// # Arguments
    ///
    /// * `id` - ID of the bookmark to close (must match a previous bookmark start)
    pub fn add_bookmark_end(&mut self, id: u32) {
        self.elements.push(ParagraphElement::BookmarkEnd(id));
    }

    /// Add a field to the paragraph.
    ///
    /// Fields are dynamic content like page numbers, dates, cross-references, etc.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// // Add a page number field
    /// para.add_field(MutableField::page());
    ///
    /// // Add a date field
    /// para.add_field(MutableField::date(Some("MMMM d, yyyy")));
    ///
    /// // Add a cross-reference field
    /// para.add_field(MutableField::reference("MyBookmark"));
    /// ```
    pub fn add_field(&mut self, field: MutableField) {
        self.elements.push(ParagraphElement::Field(field));
    }

    /// Add a run-level smart tag to this paragraph.
    pub fn add_smart_tag(&mut self, element: impl Into<String>) -> &mut MutableSmartTag {
        self.elements
            .push(ParagraphElement::SmartTag(MutableSmartTag::new(element)));
        match self.elements.last_mut() {
            Some(ParagraphElement::SmartTag(tag)) => tag,
            _ => unreachable!(),
        }
    }

    /// Add a typed tracked-change wrapper to this paragraph.
    pub fn add_revision(
        &mut self,
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> &mut MutableRevision {
        self.elements
            .push(ParagraphElement::Revision(MutableRevision::new(
                kind, metadata,
            )));
        match self.elements.last_mut() {
            Some(ParagraphElement::Revision(revision)) => revision,
            _ => unreachable!(),
        }
    }

    /// Record the paragraph properties that existed before this formatting revision.
    pub fn set_property_change(
        &mut self,
        metadata: RevisionMetadata,
        previous: &MutableParagraph,
    ) -> &mut Self {
        self.property_change = Some(ParagraphPropertyChange::snapshot(metadata, previous));
        self
    }

    /// Set the paragraph style.
    pub fn set_style(&mut self, style_id: &str) {
        self.style = Some(style_id.to_string());
    }

    /// Set the HTML division ID referenced by this paragraph.
    ///
    /// Division IDs are XML Schema integers and are kept as strings so values
    /// larger than the native integer types can be written without truncation.
    pub fn set_division_id(&mut self, id: impl Into<String>) -> Result<&mut Self> {
        self.properties.division_id = Some(normalize_xml_integer(
            id.into(),
            "Word paragraph division ID",
        )?);
        Ok(self)
    }

    /// Return the HTML division ID referenced by this paragraph, if set.
    pub fn division_id(&self) -> Option<&str> {
        self.properties.division_id.as_deref()
    }

    /// Remove the HTML division reference from this paragraph.
    pub fn clear_division_id(&mut self) -> &mut Self {
        self.properties.division_id = None;
        self
    }

    /// Set paragraph alignment.
    pub fn set_alignment(&mut self, alignment: ParagraphAlignment) {
        self.properties.alignment = Some(alignment);
    }

    /// Set spacing before this paragraph (in points).
    pub fn set_space_before(&mut self, points: f64) {
        self.properties.space_before = Some((points * 20.0) as u32);
    }

    /// Set spacing after this paragraph (in points).
    pub fn set_space_after(&mut self, points: f64) {
        self.properties.space_after = Some((points * 20.0) as u32);
    }

    /// Set line spacing for this paragraph.
    pub fn set_line_spacing(&mut self, spacing: LineSpacing) {
        self.properties.line_spacing = Some(spacing);
    }

    /// Set left indentation (in inches).
    pub fn set_indent_left(&mut self, inches: f64) {
        self.properties.indent_left = Some((inches * 1440.0) as i32);
    }

    /// Set right indentation (in inches).
    pub fn set_indent_right(&mut self, inches: f64) {
        self.properties.indent_right = Some((inches * 1440.0) as u32);
    }

    /// Set first line indentation (in inches).
    pub fn set_indent_first_line(&mut self, inches: f64) {
        self.properties.indent_first_line = Some((inches * 1440.0) as i32);
    }

    /// Set this paragraph as a list item.
    ///
    /// The num_id values correspond to the numbering definitions in numbering.xml:
    /// - numId 1: Bullet list (using Symbol font)
    /// - numId 9: Decimal list (1. 2. 3. ...)
    /// - Other formats use different IDs as needed
    pub fn set_list(&mut self, list_type: ListType, level: u32) {
        let num_id = match list_type {
            ListType::Bullet => 1,       // References abstractNumId 8 (bullet)
            ListType::Decimal => 9,      // References abstractNumId 0 (decimal)
            ListType::LowerLetter => 10, // References abstractNumId 9 (lower letter a, b, c)
            ListType::UpperLetter => 11, // References abstractNumId 10 (upper letter A, B, C)
            ListType::LowerRoman => 12,  // References abstractNumId 11 (lower roman i, ii, iii)
            ListType::UpperRoman => 13,  // References abstractNumId 12 (upper roman I, II, III)
        };

        self.properties.numbering = Some(NumberingProperties {
            num_id,
            ilvl: level,
        });
    }

    /// Get the number of elements (runs and hyperlinks).
    pub fn element_count(&self) -> usize {
        self.elements.len()
    }

    /// Clear all elements from the paragraph.
    pub fn clear(&mut self) {
        self.elements.clear();
    }

    /// Set the section break ending at this paragraph.
    pub fn set_section_break(&mut self, properties: SectionProperties) -> Result<()> {
        properties.validate()?;
        self.properties.section = Some(properties);
        Ok(())
    }

    /// Remove and return the section break ending at this paragraph.
    pub fn remove_section_break(&mut self) -> Option<SectionProperties> {
        self.properties.section.take()
    }

    /// Return this paragraph's section-break properties.
    pub fn section_break(&self) -> Option<&SectionProperties> {
        self.properties.section.as_ref()
    }

    /// Mutably access this paragraph's section-break properties.
    pub fn section_break_mut(&mut self) -> Option<&mut SectionProperties> {
        self.properties.section.as_mut()
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:p");
        crate::paragraph::extensions::append_paragraph_attributes(&self.extension_values, xml)?;
        xml.push('>');

        // Write paragraph properties
        if self.style.is_some()
            || self.properties.has_properties()
            || self.property_change.is_some()
        {
            xml.push_str("<w:pPr>");

            if let Some(ref style) = self.style {
                write!(xml, "<w:pStyle w:val=\"{}\"/>", escape_xml(style))
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(alignment) = self.properties.alignment {
                write!(xml, "<w:jc w:val=\"{}\"/>", alignment.as_str())
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            // Write numbering properties for lists
            if let Some(ref numbering) = self.properties.numbering {
                xml.push_str("<w:numPr>");
                write!(xml, "<w:ilvl w:val=\"{}\"/>", numbering.ilvl)
                    .map_err(|e| Error::Xml(e.to_string()))?;
                write!(xml, "<w:numId w:val=\"{}\"/>", numbering.num_id)
                    .map_err(|e| Error::Xml(e.to_string()))?;
                xml.push_str("</w:numPr>");
            }

            // Write spacing
            if self.properties.space_before.is_some()
                || self.properties.space_after.is_some()
                || self.properties.line_spacing.is_some()
            {
                xml.push_str("<w:spacing");
                if let Some(before) = self.properties.space_before {
                    write!(xml, " w:before=\"{}\"", before)
                        .map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(after) = self.properties.space_after {
                    write!(xml, " w:after=\"{}\"", after).map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(ref line_spacing) = self.properties.line_spacing {
                    match line_spacing {
                        LineSpacing::Single => {
                            write!(xml, " w:line=\"240\" w:lineRule=\"auto\"")
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::OneAndHalf => {
                            write!(xml, " w:line=\"360\" w:lineRule=\"auto\"")
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::Double => {
                            write!(xml, " w:line=\"480\" w:lineRule=\"auto\"")
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::Multiple(factor) => {
                            let value = (factor * 240.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"auto\"", value)
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::Exact(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"exact\"", value)
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::AtLeast(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"atLeast\"", value)
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                    }
                }
                xml.push_str("/>");
            }

            // Write indentation
            if self.properties.indent_left.is_some()
                || self.properties.indent_right.is_some()
                || self.properties.indent_first_line.is_some()
            {
                xml.push_str("<w:ind");
                if let Some(left) = self.properties.indent_left {
                    write!(xml, " w:left=\"{}\"", left).map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(right) = self.properties.indent_right {
                    write!(xml, " w:right=\"{}\"", right).map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(first_line) = self.properties.indent_first_line {
                    if first_line >= 0 {
                        write!(xml, " w:firstLine=\"{}\"", first_line)
                            .map_err(|e| Error::Xml(e.to_string()))?;
                    } else {
                        write!(xml, " w:hanging=\"{}\"", -first_line)
                            .map_err(|e| Error::Xml(e.to_string()))?;
                    }
                }
                xml.push_str("/>");
            }

            // Write tab stops
            if !self.properties.tab_stops.is_empty() {
                xml.push_str("<w:tabs>");
                for tab_stop in &self.properties.tab_stops {
                    xml.push_str("<w:tab");
                    write!(xml, " w:val=\"{}\"", escape_xml(&tab_stop.alignment))
                        .map_err(|e| Error::Xml(e.to_string()))?;
                    write!(xml, " w:pos=\"{}\"", tab_stop.position)
                        .map_err(|e| Error::Xml(e.to_string()))?;
                    if let Some(ref leader) = tab_stop.leader {
                        write!(xml, " w:leader=\"{}\"", escape_xml(leader))
                            .map_err(|e| Error::Xml(e.to_string()))?;
                    }
                    xml.push_str("/>");
                }
                xml.push_str("</w:tabs>");
            }

            if let Some(ref division_id) = self.properties.division_id {
                write!(xml, "<w:divId w:val=\"{}\"/>", escape_xml(division_id))
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(section) = &self.properties.section {
                section.write_xml(xml, None)?;
            }

            if let Some(change) = &self.property_change {
                change.write_xml(xml, None)?;
            }

            xml.push_str("</w:pPr>");
        }

        // Write elements (runs, hyperlinks, bookmarks, fields)
        // Use placeholders for relationship IDs that will be replaced after relationships are created
        let mut hyperlink_idx = 0;
        let mut image_idx = 0;
        for element in &self.elements {
            element.write_placeholder(xml, &mut hyperlink_idx, &mut image_idx)?;
        }

        xml.push_str("</w:p>");
        Ok(())
    }

    /// Generate XML with actual relationship IDs from the mapper.
    ///
    /// The hyperlink_counter and image_counter are used to track the global index
    /// across all paragraphs, and are updated as elements are processed.
    pub(crate) fn to_xml_with_rels(
        &self,
        xml: &mut String,
        rel_mapper: &crate::writer::relmap::RelationshipMapper,
        hyperlink_counter: &mut usize,
        image_counter: &mut usize,
    ) -> Result<()> {
        xml.push_str("<w:p");
        crate::paragraph::extensions::append_paragraph_attributes(&self.extension_values, xml)?;
        xml.push('>');

        // Write paragraph properties (same as to_xml)
        if self.style.is_some()
            || self.properties.has_properties()
            || self.property_change.is_some()
        {
            xml.push_str("<w:pPr>");

            if let Some(ref style) = self.style {
                write!(xml, "<w:pStyle w:val=\"{}\"/>", escape_xml(style))
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(alignment) = self.properties.alignment {
                write!(xml, "<w:jc w:val=\"{}\"/>", alignment.as_str())
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            // Write numbering properties for lists
            if let Some(ref numbering) = self.properties.numbering {
                xml.push_str("<w:numPr>");
                write!(xml, "<w:ilvl w:val=\"{}\"/>", numbering.ilvl)
                    .map_err(|e| Error::Xml(e.to_string()))?;
                write!(xml, "<w:numId w:val=\"{}\"/>", numbering.num_id)
                    .map_err(|e| Error::Xml(e.to_string()))?;
                xml.push_str("</w:numPr>");
            }

            // Write spacing
            if self.properties.space_before.is_some()
                || self.properties.space_after.is_some()
                || self.properties.line_spacing.is_some()
            {
                xml.push_str("<w:spacing");
                if let Some(before) = self.properties.space_before {
                    write!(xml, " w:before=\"{}\"", before)
                        .map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(after) = self.properties.space_after {
                    write!(xml, " w:after=\"{}\"", after).map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(ref line_spacing) = self.properties.line_spacing {
                    match line_spacing {
                        LineSpacing::Single => {
                            write!(xml, " w:line=\"240\" w:lineRule=\"auto\"")
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::OneAndHalf => {
                            write!(xml, " w:line=\"360\" w:lineRule=\"auto\"")
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::Double => {
                            write!(xml, " w:line=\"480\" w:lineRule=\"auto\"")
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::Multiple(factor) => {
                            let value = (factor * 240.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"auto\"", value)
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::Exact(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"exact\"", value)
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                        LineSpacing::AtLeast(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"atLeast\"", value)
                                .map_err(|e| Error::Xml(e.to_string()))?;
                        },
                    }
                }
                xml.push_str("/>");
            }

            // Write indentation
            if self.properties.indent_left.is_some()
                || self.properties.indent_right.is_some()
                || self.properties.indent_first_line.is_some()
            {
                xml.push_str("<w:ind");
                if let Some(left) = self.properties.indent_left {
                    write!(xml, " w:left=\"{}\"", left).map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(right) = self.properties.indent_right {
                    write!(xml, " w:right=\"{}\"", right).map_err(|e| Error::Xml(e.to_string()))?;
                }
                if let Some(first_line) = self.properties.indent_first_line {
                    if first_line >= 0 {
                        write!(xml, " w:firstLine=\"{}\"", first_line)
                            .map_err(|e| Error::Xml(e.to_string()))?;
                    } else {
                        write!(xml, " w:hanging=\"{}\"", -first_line)
                            .map_err(|e| Error::Xml(e.to_string()))?;
                    }
                }
                xml.push_str("/>");
            }

            // Write tab stops
            if !self.properties.tab_stops.is_empty() {
                xml.push_str("<w:tabs>");
                for tab_stop in &self.properties.tab_stops {
                    xml.push_str("<w:tab");
                    write!(xml, " w:val=\"{}\"", escape_xml(&tab_stop.alignment))
                        .map_err(|e| Error::Xml(e.to_string()))?;
                    write!(xml, " w:pos=\"{}\"", tab_stop.position)
                        .map_err(|e| Error::Xml(e.to_string()))?;
                    if let Some(ref leader) = tab_stop.leader {
                        write!(xml, " w:leader=\"{}\"", escape_xml(leader))
                            .map_err(|e| Error::Xml(e.to_string()))?;
                    }
                    xml.push_str("/>");
                }
                xml.push_str("</w:tabs>");
            }

            if let Some(ref division_id) = self.properties.division_id {
                write!(xml, "<w:divId w:val=\"{}\"/>", escape_xml(division_id))
                    .map_err(|e| Error::Xml(e.to_string()))?;
            }

            if let Some(section) = &self.properties.section {
                section.write_xml(xml, Some(rel_mapper))?;
            }

            if let Some(change) = &self.property_change {
                change.write_xml(xml, Some(rel_mapper))?;
            }

            xml.push_str("</w:pPr>");
        }

        // Write elements with actual relationship IDs
        // Use the passed-in counters to maintain global indices across all paragraphs
        for element in &self.elements {
            element.write_with_rels(xml, rel_mapper, hyperlink_counter, image_counter)?;
        }

        xml.push_str("</w:p>");
        Ok(())
    }
}

/// Tab stop definition for paragraphs.
#[derive(Debug, Clone)]
pub(crate) struct TabStop {
    pub(crate) position: u32,
    pub(crate) alignment: String,
    pub(crate) leader: Option<String>,
}

/// Paragraph properties.
#[derive(Debug, Default, Clone)]
pub(crate) struct ParagraphProperties {
    pub(crate) alignment: Option<ParagraphAlignment>,
    pub(crate) numbering: Option<NumberingProperties>,
    pub(crate) space_before: Option<u32>,
    pub(crate) space_after: Option<u32>,
    pub(crate) line_spacing: Option<LineSpacing>,
    pub(crate) indent_left: Option<i32>,
    pub(crate) indent_right: Option<u32>,
    pub(crate) indent_first_line: Option<i32>,
    pub(crate) tab_stops: Vec<TabStop>,
    pub(crate) division_id: Option<String>,
    pub(crate) section: Option<SectionProperties>,
}

impl ParagraphProperties {
    pub(crate) fn has_properties(&self) -> bool {
        self.alignment.is_some()
            || self.numbering.is_some()
            || self.space_before.is_some()
            || self.space_after.is_some()
            || self.line_spacing.is_some()
            || self.indent_left.is_some()
            || self.indent_right.is_some()
            || self.indent_first_line.is_some()
            || !self.tab_stops.is_empty()
            || self.division_id.is_some()
            || self.section.is_some()
    }

    pub(crate) fn write_values(
        &self,
        xml: &mut String,
        style: Option<&str>,
        rel_mapper: Option<&crate::writer::relmap::RelationshipMapper>,
    ) -> Result<()> {
        if let Some(style) = style {
            write!(xml, "<w:pStyle w:val=\"{}\"/>", escape_xml(style))?;
        }
        if let Some(alignment) = self.alignment {
            write!(xml, "<w:jc w:val=\"{}\"/>", alignment.as_str())?;
        }
        if let Some(n) = &self.numbering {
            write!(
                xml,
                "<w:numPr><w:ilvl w:val=\"{}\"/><w:numId w:val=\"{}\"/></w:numPr>",
                n.ilvl, n.num_id
            )?;
        }
        if self.space_before.is_some() || self.space_after.is_some() || self.line_spacing.is_some()
        {
            xml.push_str("<w:spacing");
            if let Some(v) = self.space_before {
                write!(xml, " w:before=\"{v}\"")?;
            }
            if let Some(v) = self.space_after {
                write!(xml, " w:after=\"{v}\"")?;
            }
            if let Some(v) = &self.line_spacing {
                let (line, rule) = match v {
                    LineSpacing::Single => (240, "auto"),
                    LineSpacing::OneAndHalf => (360, "auto"),
                    LineSpacing::Double => (480, "auto"),
                    LineSpacing::Multiple(f) => ((*f * 240.0) as u32, "auto"),
                    LineSpacing::Exact(f) => ((*f * 20.0) as u32, "exact"),
                    LineSpacing::AtLeast(f) => ((*f * 20.0) as u32, "atLeast"),
                };
                write!(xml, " w:line=\"{line}\" w:lineRule=\"{rule}\"")?;
            }
            xml.push_str("/>");
        }
        if self.indent_left.is_some()
            || self.indent_right.is_some()
            || self.indent_first_line.is_some()
        {
            xml.push_str("<w:ind");
            if let Some(v) = self.indent_left {
                write!(xml, " w:left=\"{v}\"")?;
            }
            if let Some(v) = self.indent_right {
                write!(xml, " w:right=\"{v}\"")?;
            }
            if let Some(v) = self.indent_first_line {
                if v >= 0 {
                    write!(xml, " w:firstLine=\"{v}\"")?;
                } else {
                    write!(xml, " w:hanging=\"{}\"", -v)?;
                }
            }
            xml.push_str("/>");
        }
        if !self.tab_stops.is_empty() {
            xml.push_str("<w:tabs>");
            for tab in &self.tab_stops {
                write!(
                    xml,
                    "<w:tab w:val=\"{}\" w:pos=\"{}\"",
                    escape_xml(&tab.alignment),
                    tab.position
                )?;
                if let Some(leader) = &tab.leader {
                    write!(xml, " w:leader=\"{}\"", escape_xml(leader))?;
                }
                xml.push_str("/>");
            }
            xml.push_str("</w:tabs>");
        }
        if let Some(v) = &self.division_id {
            write!(xml, "<w:divId w:val=\"{}\"/>", escape_xml(v))?;
        }
        if let Some(section) = &self.section {
            section.write_xml(xml, rel_mapper)?;
        }
        Ok(())
    }
}

/// Numbering properties for lists.
#[derive(Debug, Clone)]
pub(crate) struct NumberingProperties {
    pub(crate) num_id: u32,
    pub(crate) ilvl: u32,
}

/// List types for paragraphs.
#[derive(Debug, Clone, Copy)]
pub enum ListType {
    Bullet,
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
}

#[cfg(test)]
mod tests {
    use super::MutableParagraph;
    use crate::writer::relmap::RelationshipMapper;
    use crate::{OfficeMath, Paragraph};

    #[test]
    fn writes_paragraph_division_id_for_reader_round_trip() {
        let mut paragraph = MutableParagraph::new();
        paragraph
            .set_division_id("+123456789012345678901234567890")
            .unwrap();

        assert_eq!(
            paragraph.division_id(),
            Some("+123456789012345678901234567890")
        );

        let mut xml = String::new();
        paragraph.to_xml(&mut xml).unwrap();
        let mut relationship_xml = String::new();
        paragraph
            .to_xml_with_rels(
                &mut relationship_xml,
                &RelationshipMapper::new(),
                &mut 0,
                &mut 0,
            )
            .unwrap();
        assert_eq!(relationship_xml, xml);

        let parsed = Paragraph::new(xml.into_bytes());
        assert_eq!(
            parsed.division_id().unwrap().as_deref(),
            Some("+123456789012345678901234567890")
        );

        paragraph.clear_division_id();
        let mut cleared_xml = String::new();
        paragraph.to_xml(&mut cleared_xml).unwrap();
        assert!(!cleared_xml.contains("<w:divId"));
    }

    #[test]
    fn rejects_invalid_paragraph_division_id() {
        let mut paragraph = MutableParagraph::new();
        assert!(paragraph.set_division_id("12.5").is_err());
        assert_eq!(paragraph.division_id(), None);
    }

    #[test]
    fn writes_and_reopens_inline_and_display_office_math() {
        let inline = OfficeMath::text("x + y");
        let display = OfficeMath::from_xml("<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>").unwrap();
        let mut paragraph = MutableParagraph::new();
        paragraph.add_run_with_text("before ");
        paragraph.add_inline_office_math(inline.clone());
        paragraph.add_run_with_text(" after");
        paragraph.add_display_office_math(display.clone());

        let mut xml = String::new();
        paragraph.to_xml(&mut xml).unwrap();
        assert!(xml.contains("<w:r><m:oMath"));
        assert!(xml.contains("<m:oMathPara"));

        let parsed = Paragraph::new(xml.into_bytes());
        assert_eq!(parsed.inline_office_math().unwrap(), vec![inline]);
        assert_eq!(parsed.display_office_math().unwrap(), vec![display]);

        let count = paragraph.element_count();
        assert!(
            paragraph
                .add_inline_office_math_xml("<m:notMath/>")
                .is_err()
        );
        assert_eq!(paragraph.element_count(), count);
    }
}
