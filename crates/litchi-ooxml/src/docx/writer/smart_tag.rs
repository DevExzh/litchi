//! Mutable Word smart tags.

use super::bookmark::MutableBookmark;
use super::field::MutableField;
use super::hyperlink::MutableHyperlink;
use super::image::MutableInlineImage;
use super::paragraph::ParagraphElement;
use super::relmap::RelationshipMapper;
use super::revision::RevisionTextMode;
use super::run::MutableRun;
use crate::error::{OoxmlError, Result};
use litchi_core::xml::escape_xml;
use std::fmt::Write as _;

/// A custom property on a mutable smart tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MutableSmartTagAttribute {
    uri: Option<String>,
    name: String,
    value: String,
}

impl MutableSmartTagAttribute {
    /// Set the optional property namespace URI.
    pub fn set_uri(&mut self, uri: impl Into<String>) -> &mut Self {
        self.uri = Some(uri.into());
        self
    }

    /// Return the optional property namespace URI.
    pub fn uri(&self) -> Option<&str> {
        self.uri.as_deref()
    }

    /// Return the property name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return the property value.
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// A mutable run-level Word smart tag.
#[derive(Debug)]
pub struct MutableSmartTag {
    uri: Option<String>,
    element: String,
    attributes: Vec<MutableSmartTagAttribute>,
    pub(crate) elements: Vec<ParagraphElement>,
}

impl MutableSmartTag {
    /// Create a smart tag with its required element name.
    pub fn new(element: impl Into<String>) -> Self {
        Self {
            uri: None,
            element: element.into(),
            attributes: Vec::new(),
            elements: Vec::new(),
        }
    }

    /// Set the optional smart-tag vocabulary namespace URI.
    pub fn set_uri(&mut self, uri: impl Into<String>) -> &mut Self {
        self.uri = Some(uri.into());
        self
    }

    /// Add a custom smart-tag property.
    pub fn add_attribute(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> &mut MutableSmartTagAttribute {
        self.attributes.push(MutableSmartTagAttribute {
            uri: None,
            name: name.into(),
            value: value.into(),
        });
        self.attributes
            .last_mut()
            .expect("attribute was just added")
    }

    /// Add an empty run to the tagged content.
    pub fn add_run(&mut self) -> &mut MutableRun {
        self.elements.push(ParagraphElement::Run(MutableRun::new()));
        match self.elements.last_mut() {
            Some(ParagraphElement::Run(run)) => run,
            _ => unreachable!(),
        }
    }

    /// Add a text run to the tagged content.
    pub fn add_run_with_text(&mut self, text: &str) -> &mut MutableRun {
        let run = self.add_run();
        run.set_text(text);
        run
    }

    /// Add an external hyperlink to the tagged content.
    pub fn add_hyperlink(&mut self, url: &str, text: &str) -> &mut MutableHyperlink {
        self.elements
            .push(ParagraphElement::Hyperlink(MutableHyperlink::new(
                url.to_owned(),
                text.to_owned(),
            )));
        match self.elements.last_mut() {
            Some(ParagraphElement::Hyperlink(hyperlink)) => hyperlink,
            _ => unreachable!(),
        }
    }

    /// Add an inline image from bytes to the tagged content.
    pub fn add_picture_from_bytes(
        &mut self,
        data: Vec<u8>,
        width_emu: Option<i64>,
        height_emu: Option<i64>,
    ) -> Result<&mut MutableInlineImage> {
        let image = MutableInlineImage::from_bytes(data, width_emu, height_emu)?;
        self.elements.push(ParagraphElement::InlineImage(image));
        match self.elements.last_mut() {
            Some(ParagraphElement::InlineImage(image)) => Ok(image),
            _ => unreachable!(),
        }
    }

    /// Add a nested smart tag.
    pub fn add_smart_tag(&mut self, element: impl Into<String>) -> &mut MutableSmartTag {
        self.elements
            .push(ParagraphElement::SmartTag(Self::new(element)));
        match self.elements.last_mut() {
            Some(ParagraphElement::SmartTag(tag)) => tag,
            _ => unreachable!(),
        }
    }

    /// Add a bookmark start marker to the tagged content.
    pub fn add_bookmark_start(&mut self, id: u32, name: &str) {
        self.elements
            .push(ParagraphElement::BookmarkStart(MutableBookmark::new(
                id,
                name.to_owned(),
            )));
    }

    /// Add a bookmark end marker to the tagged content.
    pub fn add_bookmark_end(&mut self, id: u32) {
        self.elements.push(ParagraphElement::BookmarkEnd(id));
    }

    /// Add a field to the tagged content.
    pub fn add_field(&mut self, field: MutableField) {
        self.elements.push(ParagraphElement::Field(field));
    }

    /// Return the optional vocabulary namespace URI.
    pub fn uri(&self) -> Option<&str> {
        self.uri.as_deref()
    }

    /// Return the required smart-tag element name.
    pub fn element(&self) -> &str {
        &self.element
    }

    /// Return the smart-tag properties.
    pub fn attributes(&self) -> &[MutableSmartTagAttribute] {
        &self.attributes
    }

    pub(crate) fn write_placeholder_mode(
        &self,
        xml: &mut String,
        hyperlink_index: &mut usize,
        image_index: &mut usize,
        mode: RevisionTextMode,
    ) -> Result<()> {
        self.write_start(xml)?;
        for element in &self.elements {
            element.write_placeholder_mode(xml, hyperlink_index, image_index, mode)?;
        }
        xml.push_str("</w:smartTag>");
        Ok(())
    }

    pub(crate) fn write_with_rels_mode(
        &self,
        xml: &mut String,
        rel_mapper: &RelationshipMapper,
        hyperlink_index: &mut usize,
        image_index: &mut usize,
        mode: RevisionTextMode,
    ) -> Result<()> {
        self.write_start(xml)?;
        for element in &self.elements {
            element.write_with_rels_mode(xml, rel_mapper, hyperlink_index, image_index, mode)?;
        }
        xml.push_str("</w:smartTag>");
        Ok(())
    }

    fn write_start(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:smartTag");
        if let Some(uri) = &self.uri {
            write!(xml, " w:uri=\"{}\"", escape_xml(uri))
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        write!(xml, " w:element=\"{}\">", escape_xml(&self.element))
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;

        if !self.attributes.is_empty() {
            xml.push_str("<w:smartTagPr>");
            for attribute in &self.attributes {
                xml.push_str("<w:attr");
                if let Some(uri) = &attribute.uri {
                    write!(xml, " w:uri=\"{}\"", escape_xml(uri))
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                }
                write!(
                    xml,
                    " w:name=\"{}\" w:val=\"{}\"/>",
                    escape_xml(&attribute.name),
                    escape_xml(&attribute.value)
                )
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            }
            xml.push_str("</w:smartTagPr>");
        }
        Ok(())
    }

    pub(crate) fn collect_hyperlink_urls(&self, urls: &mut Vec<String>) {
        for element in &self.elements {
            element.collect_hyperlink_urls(urls);
        }
    }

    pub(crate) fn collect_images<'a>(&'a self, images: &mut Vec<(&'a [u8], super::ImageFormat)>) {
        for element in &self.elements {
            element.collect_images(images);
        }
    }

    pub(crate) fn append_run_text(&self, text: &mut String) {
        for element in &self.elements {
            element.append_run_text(text);
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::docx::Paragraph;
    use crate::docx::writer::MutableDocument;
    use crate::docx::writer::paragraph::MutableParagraph;

    #[test]
    fn generated_nested_smart_tags_round_trip_through_reader() {
        let mut paragraph = MutableParagraph::new();
        let tag = paragraph.add_smart_tag("person");
        tag.set_uri("urn:contacts");
        tag.add_attribute("kind", "friend & peer")
            .set_uri("urn:meta");
        tag.add_run_with_text("A & ");
        tag.add_smart_tag("givenName").add_run_with_text("Bob");

        let mut xml = String::new();
        paragraph.to_xml(&mut xml).unwrap();
        let paragraph = Paragraph::new(xml.into_bytes());
        let tags = paragraph.smart_tags().unwrap();

        assert_eq!(tags.len(), 2);
        assert_eq!(tags[0].element, "person");
        assert_eq!(tags[0].attributes[0].value, "friend & peer");
        assert_eq!(tags[0].text().unwrap(), "A & Bob");
        assert_eq!(tags[1].element, "givenName");
    }

    #[test]
    fn nested_smart_tag_hyperlinks_are_collected_for_relationships() {
        let mut document = MutableDocument::new();
        document
            .add_paragraph()
            .add_smart_tag("link")
            .add_hyperlink("https://example.test", "example");

        assert_eq!(
            document.collect_hyperlink_urls(),
            vec!["https://example.test"]
        );
    }
}
