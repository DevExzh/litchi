//! Custom slide show support for PowerPoint presentations.
//!
//! Custom slide shows allow defining named subsets of slides that can be
//! presented independently of the full presentation.
use super::model::*;
use crate::presentation_properties::metadata::escape_xml;
use crate::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;

impl List {
    /// Parse custom shows from presentation XML.
    pub fn parse_xml(xml: &str) -> Result<Self> {
        let mut list = Self::new();
        let xml = litchi_ooxml_common::mce::process_str(xml)?;
        let mut reader = Reader::from_str(xml.as_ref());
        reader.config_mut().trim_text(true);

        let mut current_show: Option<Show> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"custShow" => {
                    let mut name = String::new();
                    let mut id = 0u32;
                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                name = std::str::from_utf8(&attr.value).unwrap_or("").to_string();
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
                    current_show = Some(Show::new(id, name));
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
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(list)
    }

    /// Generate custom shows XML.
    ///
    /// This is a convenience wrapper that emits a best-effort mapping by treating
    /// numeric slide IDs as their corresponding `rId` values (e.g. slide 5 → `rId5`).
    /// For full fidelity, prefer [`Self::to_xml_with_rel_map`] and pass the workbook's
    /// actual slide relationship map.
    pub fn to_xml(&self) -> String {
        if self.is_empty() {
            return String::new();
        }

        let mut fallback_map = HashMap::new();
        for show in &self.shows {
            for slide_id in &show.slide_ids {
                fallback_map
                    .entry(*slide_id)
                    .or_insert_with(|| format!("rId{slide_id}"));
            }
        }

        self.to_xml_with_rel_map(&fallback_map)
    }

    /// Generate custom shows XML with proper relationship ID mapping.
    ///
    /// # Arguments
    /// * `slide_id_to_rel_id` - Mapping from slide ID (e.g., 256) to relationship ID (e.g., "rId6")
    pub fn to_xml_with_rel_map(&self, slide_id_to_rel_id: &HashMap<u32, String>) -> String {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_custom_show_creation() {
        let show = Show::new(0, "Executive Summary").with_slides(vec![256, 257, 262]);

        assert_eq!(show.name, "Executive Summary");
        assert_eq!(show.slide_count(), 3);
    }

    #[test]
    fn test_custom_show_list() {
        let mut list = List::new();
        list.create("Short Version", vec![256, 262]);
        list.create("Full Presentation", vec![256, 257, 258, 259, 260, 261, 262]);

        assert_eq!(list.len(), 2);
        assert!(list.get_by_name("Short Version").is_some());
    }

    #[test]
    fn test_custom_shows_xml() {
        let mut list = List::new();
        list.create("Demo", vec![1, 2, 3]);

        let xml = list.to_xml();
        assert!(xml.contains("Demo"));
        assert!(xml.contains("custShow"));
    }
}
