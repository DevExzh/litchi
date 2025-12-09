//! Custom slide show support for PowerPoint presentations.
//!
//! Custom slide shows allow defining named subsets of slides that can be
//! presented independently of the full presentation.

use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

/// A custom slide show definition.
#[derive(Debug, Clone)]
pub struct CustomShow {
    /// Unique ID for the custom show
    pub id: u32,
    /// Display name of the custom show
    pub name: String,
    /// List of slide IDs included in the show (in presentation order)
    pub slide_ids: Vec<u32>,
}

impl CustomShow {
    /// Create a new custom show.
    pub fn new(id: u32, name: impl Into<String>) -> Self {
        Self {
            id,
            name: name.into(),
            slide_ids: Vec::new(),
        }
    }

    /// Add a slide to the custom show.
    pub fn add_slide(&mut self, slide_id: u32) {
        self.slide_ids.push(slide_id);
    }

    /// Add multiple slides to the custom show.
    pub fn add_slides(&mut self, slide_ids: impl IntoIterator<Item = u32>) {
        self.slide_ids.extend(slide_ids);
    }

    /// Set slides with builder pattern.
    pub fn with_slides(mut self, slide_ids: Vec<u32>) -> Self {
        self.slide_ids = slide_ids;
        self
    }

    /// Get the number of slides in the custom show.
    pub fn slide_count(&self) -> usize {
        self.slide_ids.len()
    }
}

/// Collection of custom slide shows for a presentation.
#[derive(Debug, Clone, Default)]
pub struct CustomShowList {
    /// List of custom shows
    pub shows: Vec<CustomShow>,
    /// Next available ID for new shows
    next_id: u32,
}

impl CustomShowList {
    /// Create a new empty custom show list.
    pub fn new() -> Self {
        Self {
            shows: Vec::new(),
            next_id: 0,
        }
    }

    /// Add a custom show to the list.
    pub fn add(&mut self, show: CustomShow) {
        if show.id >= self.next_id {
            self.next_id = show.id + 1;
        }
        self.shows.push(show);
    }

    /// Create and add a new custom show.
    pub fn create(&mut self, name: impl Into<String>, slide_ids: Vec<u32>) -> &CustomShow {
        let show = CustomShow::new(self.next_id, name).with_slides(slide_ids);
        self.next_id += 1;
        self.shows.push(show);
        self.shows.last().unwrap()
    }

    /// Get a custom show by name.
    pub fn get_by_name(&self, name: &str) -> Option<&CustomShow> {
        self.shows.iter().find(|s| s.name == name)
    }

    /// Get a custom show by ID.
    pub fn get_by_id(&self, id: u32) -> Option<&CustomShow> {
        self.shows.iter().find(|s| s.id == id)
    }

    /// Remove a custom show by name.
    pub fn remove_by_name(&mut self, name: &str) -> Option<CustomShow> {
        if let Some(pos) = self.shows.iter().position(|s| s.name == name) {
            Some(self.shows.remove(pos))
        } else {
            None
        }
    }

    /// Get the number of custom shows.
    pub fn len(&self) -> usize {
        self.shows.len()
    }

    /// Check if the list is empty.
    pub fn is_empty(&self) -> bool {
        self.shows.is_empty()
    }

    /// Parse custom shows from presentation XML.
    pub fn parse_xml(xml: &str) -> Result<Self> {
        let mut list = Self::new();
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut current_show: Option<CustomShow> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"custShow" {
                        let mut name = String::new();
                        let mut id = 0u32;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"name" => {
                                    name =
                                        std::str::from_utf8(&attr.value).unwrap_or("").to_string();
                                },
                                b"id" => {
                                    id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok())
                                        .unwrap_or(0);
                                },
                                _ => {},
                            }
                        }
                        current_show = Some(CustomShow::new(id, name));
                    }
                },
                Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"sld"
                        && let Some(ref mut show) = current_show
                    {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r:id" || attr.key.as_ref() == b"id" {
                                // Extract slide relationship ID or actual ID
                                if let Ok(id_str) = std::str::from_utf8(&attr.value) {
                                    // Try to parse as number, or extract from rId format
                                    if let Ok(id) = id_str.trim_start_matches("rId").parse::<u32>()
                                    {
                                        show.add_slide(id);
                                    }
                                }
                            }
                        }
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"custShow"
                        && let Some(show) = current_show.take()
                    {
                        list.add(show);
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(list)
    }

    /// Generate custom shows XML.
    ///
    /// Note: This generates XML with slide IDs as relationship IDs, which is incorrect.
    /// Use `to_xml_with_rel_map` instead when you have the slide ID to relationship ID mapping.
    pub fn to_xml(&self) -> String {
        // Return empty if no custom shows - this avoids corruption
        // Full support requires relationship ID mapping
        if self.is_empty() {
            return String::new();
        }

        // Without a proper slide ID to relationship ID mapping, we cannot
        // generate valid custom show XML. Return empty to avoid corruption.
        String::new()
    }

    /// Generate custom shows XML with proper relationship ID mapping.
    ///
    /// # Arguments
    /// * `slide_id_to_rel_id` - Mapping from slide ID (e.g., 256) to relationship ID (e.g., "rId6")
    pub fn to_xml_with_rel_map(
        &self,
        slide_id_to_rel_id: &std::collections::HashMap<u32, String>,
    ) -> String {
        if self.is_empty() {
            return String::new();
        }

        let mut xml = String::with_capacity(1024);

        xml.push_str("<p:custShowLst>");

        for show in &self.shows {
            xml.push_str(&format!(
                r#"<p:custShow name="{}" id="{}">"#,
                escape_xml(&show.name),
                show.id
            ));
            xml.push_str("<p:sldLst>");
            for slide_id in &show.slide_ids {
                // Look up the relationship ID for this slide ID
                if let Some(rel_id) = slide_id_to_rel_id.get(slide_id) {
                    xml.push_str(&format!(r#"<p:sld r:id="{}"/>"#, rel_id));
                }
            }
            xml.push_str("</p:sldLst>");
            xml.push_str("</p:custShow>");
        }

        xml.push_str("</p:custShowLst>");

        xml
    }
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_custom_show_creation() {
        let show = CustomShow::new(0, "Executive Summary").with_slides(vec![256, 257, 262]);

        assert_eq!(show.name, "Executive Summary");
        assert_eq!(show.slide_count(), 3);
    }

    #[test]
    fn test_custom_show_list() {
        let mut list = CustomShowList::new();
        list.create("Short Version", vec![256, 262]);
        list.create("Full Presentation", vec![256, 257, 258, 259, 260, 261, 262]);

        assert_eq!(list.len(), 2);
        assert!(list.get_by_name("Short Version").is_some());
    }

    #[test]
    fn test_custom_shows_xml() {
        let mut list = CustomShowList::new();
        list.create("Demo", vec![1, 2, 3]);

        let xml = list.to_xml();
        assert!(xml.contains("Demo"));
        assert!(xml.contains("custShow"));
    }
}
