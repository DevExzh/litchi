/// Paragraph types and implementation for DOCX documents.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::{LineSpacing, ParagraphAlignment};
// Import other writer types
use super::hyperlink::MutableHyperlink;
use super::image::MutableInlineImage;
use super::run::MutableRun;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Elements that can appear in a paragraph.
#[derive(Debug)]
pub(crate) enum ParagraphElement {
    Run(MutableRun),
    Hyperlink(MutableHyperlink),
    InlineImage(MutableInlineImage),
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
}

impl MutableParagraph {
    pub(crate) fn new() -> Self {
        Self {
            elements: Vec::new(),
            style: None,
            properties: ParagraphProperties::default(),
        }
    }

    /// Create a MutableParagraph from existing XML.
    ///
    /// For now, this stores the raw XML to preserve formatting when re-writing.
    /// Full parsing could be added later for more sophisticated editing.
    pub(crate) fn from_xml(_xml: &str) -> Result<Self> {
        // For initial implementation, create an empty paragraph
        // Future enhancement: parse XML to extract runs, hyperlinks, and properties
        Ok(Self::new())
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

    /// Add a hyperlink to the paragraph.
    pub fn add_hyperlink(&mut self, text: &str, url: &str) -> &mut MutableHyperlink {
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
        let data = fs::read(image_path).map_err(OoxmlError::IoError)?;
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

    /// Set the paragraph style.
    pub fn set_style(&mut self, style_id: &str) {
        self.style = Some(style_id.to_string());
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
        self.properties.indent_left = Some((inches * 1440.0) as u32);
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
            ListType::LowerLetter => 10, // TODO: Add to numbering.xml
            ListType::UpperLetter => 11, // TODO: Add to numbering.xml
            ListType::LowerRoman => 12,  // TODO: Add to numbering.xml
            ListType::UpperRoman => 13,  // TODO: Add to numbering.xml
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

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:p>");

        // Write paragraph properties
        if self.style.is_some() || self.properties.has_properties() {
            xml.push_str("<w:pPr>");

            if let Some(ref style) = self.style {
                write!(xml, "<w:pStyle w:val=\"{}\"/>", escape_xml(style))
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(alignment) = self.properties.alignment {
                write!(xml, "<w:jc w:val=\"{}\"/>", alignment.as_str())
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            // Write numbering properties for lists
            if let Some(ref numbering) = self.properties.numbering {
                xml.push_str("<w:numPr>");
                write!(xml, "<w:ilvl w:val=\"{}\"/>", numbering.ilvl)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, "<w:numId w:val=\"{}\"/>", numbering.num_id)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
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
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(after) = self.properties.space_after {
                    write!(xml, " w:after=\"{}\"", after)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(ref line_spacing) = self.properties.line_spacing {
                    match line_spacing {
                        LineSpacing::Single => {
                            write!(xml, " w:line=\"240\" w:lineRule=\"auto\"")
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::OneAndHalf => {
                            write!(xml, " w:line=\"360\" w:lineRule=\"auto\"")
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::Double => {
                            write!(xml, " w:line=\"480\" w:lineRule=\"auto\"")
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::Multiple(factor) => {
                            let value = (factor * 240.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"auto\"", value)
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::Exact(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"exact\"", value)
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::AtLeast(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"atLeast\"", value)
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
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
                    write!(xml, " w:left=\"{}\"", left)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(right) = self.properties.indent_right {
                    write!(xml, " w:right=\"{}\"", right)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(first_line) = self.properties.indent_first_line {
                    if first_line >= 0 {
                        write!(xml, " w:firstLine=\"{}\"", first_line)
                            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    } else {
                        write!(xml, " w:hanging=\"{}\"", -first_line)
                            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    }
                }
                xml.push_str("/>");
            }

            xml.push_str("</w:pPr>");
        }

        // Write elements (runs and hyperlinks)
        // Use placeholders for relationship IDs that will be replaced after relationships are created
        let mut hyperlink_idx = 0;
        let mut image_idx = 0;
        for element in self.elements.iter() {
            match element {
                ParagraphElement::Run(run) => run.to_xml(xml)?,
                ParagraphElement::Hyperlink(hyperlink) => {
                    let placeholder = format!("{{{{HYPERLINK_{}}}}}", hyperlink_idx);
                    hyperlink.to_xml(xml, &placeholder)?;
                    hyperlink_idx += 1;
                },
                ParagraphElement::InlineImage(image) => {
                    xml.push_str("<w:r>");
                    let placeholder = format!("{{{{IMAGE_{}}}}}", image_idx);
                    image.to_xml(xml, &placeholder)?;
                    xml.push_str("</w:r>");
                    image_idx += 1;
                },
            }
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
        rel_mapper: &crate::ooxml::docx::writer::relmap::RelationshipMapper,
        hyperlink_counter: &mut usize,
        image_counter: &mut usize,
    ) -> Result<()> {
        xml.push_str("<w:p>");

        // Write paragraph properties (same as to_xml)
        if self.style.is_some() || self.properties.has_properties() {
            xml.push_str("<w:pPr>");

            if let Some(ref style) = self.style {
                write!(xml, "<w:pStyle w:val=\"{}\"/>", escape_xml(style))
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(alignment) = self.properties.alignment {
                write!(xml, "<w:jc w:val=\"{}\"/>", alignment.as_str())
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            // Write numbering properties for lists
            if let Some(ref numbering) = self.properties.numbering {
                xml.push_str("<w:numPr>");
                write!(xml, "<w:ilvl w:val=\"{}\"/>", numbering.ilvl)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, "<w:numId w:val=\"{}\"/>", numbering.num_id)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
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
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(after) = self.properties.space_after {
                    write!(xml, " w:after=\"{}\"", after)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(ref line_spacing) = self.properties.line_spacing {
                    match line_spacing {
                        LineSpacing::Single => {
                            write!(xml, " w:line=\"240\" w:lineRule=\"auto\"")
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::OneAndHalf => {
                            write!(xml, " w:line=\"360\" w:lineRule=\"auto\"")
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::Double => {
                            write!(xml, " w:line=\"480\" w:lineRule=\"auto\"")
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::Multiple(factor) => {
                            let value = (factor * 240.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"auto\"", value)
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::Exact(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"exact\"", value)
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                        },
                        LineSpacing::AtLeast(points) => {
                            let value = (points * 20.0) as u32;
                            write!(xml, " w:line=\"{}\" w:lineRule=\"atLeast\"", value)
                                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
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
                    write!(xml, " w:left=\"{}\"", left)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(right) = self.properties.indent_right {
                    write!(xml, " w:right=\"{}\"", right)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }
                if let Some(first_line) = self.properties.indent_first_line {
                    if first_line >= 0 {
                        write!(xml, " w:firstLine=\"{}\"", first_line)
                            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    } else {
                        write!(xml, " w:hanging=\"{}\"", -first_line)
                            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    }
                }
                xml.push_str("/>");
            }

            xml.push_str("</w:pPr>");
        }

        // Write elements with actual relationship IDs
        // Use the passed-in counters to maintain global indices across all paragraphs
        for element in self.elements.iter() {
            match element {
                ParagraphElement::Run(run) => run.to_xml(xml)?,
                ParagraphElement::Hyperlink(hyperlink) => {
                    if let Some(rel_id) = rel_mapper.get_hyperlink_id(*hyperlink_counter) {
                        hyperlink.to_xml(xml, rel_id)?;
                    } else {
                        // Fallback to placeholder if ID not found (shouldn't happen)
                        let placeholder = format!("{{{{HYPERLINK_{}}}}}", *hyperlink_counter);
                        hyperlink.to_xml(xml, &placeholder)?;
                    }
                    *hyperlink_counter += 1;
                },
                ParagraphElement::InlineImage(image) => {
                    xml.push_str("<w:r>");
                    if let Some(rel_id) = rel_mapper.get_image_id(*image_counter) {
                        image.to_xml(xml, rel_id)?;
                    } else {
                        // Fallback to placeholder if ID not found (shouldn't happen)
                        let placeholder = format!("{{{{IMAGE_{}}}}}", *image_counter);
                        image.to_xml(xml, &placeholder)?;
                    }
                    xml.push_str("</w:r>");
                    *image_counter += 1;
                },
            }
        }

        xml.push_str("</w:p>");
        Ok(())
    }
}

/// Paragraph properties.
#[derive(Debug, Default)]
pub(crate) struct ParagraphProperties {
    pub(crate) alignment: Option<ParagraphAlignment>,
    pub(crate) numbering: Option<NumberingProperties>,
    pub(crate) space_before: Option<u32>,
    pub(crate) space_after: Option<u32>,
    pub(crate) line_spacing: Option<LineSpacing>,
    pub(crate) indent_left: Option<u32>,
    pub(crate) indent_right: Option<u32>,
    pub(crate) indent_first_line: Option<i32>,
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
