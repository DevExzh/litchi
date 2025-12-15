//! Presentation sections support for PPTX.
//!
//! Sections are used to organize slides into logical groups within a presentation.
//! This module provides both reading and writing support for sections.

use crate::common::xml::escape_xml;
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fmt::Write as FmtWrite;

/// A section in a presentation.
///
/// Sections are logical groupings of slides that help organize large presentations.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Section {
    /// Section name
    pub name: String,
    /// Section ID (GUID format)
    pub id: String,
    /// Slide IDs in this section
    pub slide_ids: Vec<u32>,
}

impl Section {
    /// Create a new section.
    ///
    /// # Arguments
    /// * `name` - Display name for the section
    /// * `id` - Unique ID (typically GUID format like `{12345678-1234-1234-1234-123456789012}`)
    pub fn new(name: impl Into<String>, id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            id: id.into(),
            slide_ids: Vec::new(),
        }
    }

    /// Add a slide ID to this section.
    pub fn add_slide(&mut self, slide_id: u32) {
        self.slide_ids.push(slide_id);
    }

    /// Create a section with slide IDs.
    pub fn with_slides(mut self, slide_ids: impl IntoIterator<Item = u32>) -> Self {
        self.slide_ids.extend(slide_ids);
        self
    }

    /// Generate XML for this section.
    pub fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(256);

        write!(
            xml,
            r#"<p14:section name="{}" id="{}">"#,
            escape_xml(&self.name),
            escape_xml(&self.id)
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;

        // Section slide ID list
        xml.push_str("<p14:sldIdLst>");
        for slide_id in &self.slide_ids {
            write!(xml, r#"<p14:sldId id="{}"/>"#, slide_id)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }
        xml.push_str("</p14:sldIdLst>");

        xml.push_str("</p14:section>");

        Ok(xml)
    }
}

/// A collection of sections in a presentation.
#[derive(Debug, Clone, Default)]
pub struct SectionList {
    /// Sections in the presentation
    sections: Vec<Section>,
}

impl SectionList {
    /// Create a new empty section list.
    pub fn new() -> Self {
        Self {
            sections: Vec::new(),
        }
    }

    /// Add a section to the list.
    pub fn add_section(&mut self, section: Section) {
        self.sections.push(section);
    }

    /// Get all sections.
    pub fn sections(&self) -> &[Section] {
        &self.sections
    }

    /// Get the number of sections.
    pub fn len(&self) -> usize {
        self.sections.len()
    }

    /// Check if the section list is empty.
    pub fn is_empty(&self) -> bool {
        self.sections.is_empty()
    }

    /// Parse sections from presentation XML.
    ///
    /// Looks for the `<p14:sectionLst>` element in the presentation XML.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut sections = Vec::new();
        let mut current_section: Option<Section> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let tag_name = e.local_name();

                    match tag_name.as_ref() {
                        b"section" => {
                            let mut name = String::new();
                            let mut id = String::new();

                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"name" => {
                                        name = std::str::from_utf8(&attr.value)
                                            .map(|s| s.to_string())
                                            .unwrap_or_default();
                                    },
                                    b"id" => {
                                        id = std::str::from_utf8(&attr.value)
                                            .map(|s| s.to_string())
                                            .unwrap_or_default();
                                    },
                                    _ => {},
                                }
                            }

                            current_section = Some(Section::new(name, id));
                        },
                        b"sldId" if current_section.is_some() => {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"id"
                                    && let Some(ref mut section) = current_section
                                    && let Ok(id_str) = std::str::from_utf8(&attr.value)
                                    && let Ok(id) = id_str.parse::<u32>()
                                {
                                    section.slide_ids.push(id);
                                }
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"section"
                        && let Some(section) = current_section.take()
                    {
                        sections.push(section);
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(SectionList { sections })
    }

    /// Generate XML for the section list.
    ///
    /// This generates the `<p14:sectionLst>` element that goes inside the presentation.xml extLst.
    pub fn to_xml(&self) -> Result<String> {
        if self.sections.is_empty() {
            return Ok(String::new());
        }

        let mut xml = String::with_capacity(1024);

        // Section list extension
        xml.push_str(r#"<p:extLst>"#);
        xml.push_str(r#"<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">"#);
        xml.push_str(r#"<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">"#);

        for section in &self.sections {
            xml.push_str(&section.to_xml()?);
        }

        xml.push_str("</p14:sectionLst>");
        xml.push_str("</p:ext>");
        xml.push_str("</p:extLst>");

        Ok(xml)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_section_creation() {
        let section = Section::new("Introduction", "{12345678-1234-1234-1234-123456789012}")
            .with_slides(vec![256, 257, 258]);

        assert_eq!(section.name, "Introduction");
        assert_eq!(section.slide_ids.len(), 3);
    }

    #[test]
    fn test_section_xml() {
        let section = Section::new("Test Section", "{ABCD1234-5678-90AB-CDEF-123456789ABC}")
            .with_slides(vec![256, 257]);

        let xml = section.to_xml().unwrap();
        assert!(xml.contains("Test Section"));
        assert!(xml.contains("ABCD1234"));
        assert!(xml.contains("256"));
        assert!(xml.contains("257"));
    }

    #[test]
    fn test_section_list_xml() {
        let mut sections = SectionList::new();
        sections.add_section(
            Section::new("Part 1", "{11111111-1111-1111-1111-111111111111}").with_slides(vec![256]),
        );
        sections.add_section(
            Section::new("Part 2", "{22222222-2222-2222-2222-222222222222}")
                .with_slides(vec![257, 258]),
        );

        let xml = sections.to_xml().unwrap();
        assert!(xml.contains("<p14:sectionLst"));
        assert!(xml.contains("Part 1"));
        assert!(xml.contains("Part 2"));
    }
}
