//! Custom slide show support for `PowerPoint` presentations.
//!
//! Custom slide shows allow defining named subsets of slides that can be
//! presented independently of the full presentation.
use super::super::{BoundedString, Limits, check_string, process_str};
use super::model::{List, Show};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;

impl List {
    /// Parse custom shows from presentation XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_xml(xml: &str, limits: Limits) -> Result<Self> {
        if xml.len() > limits.max_input_bytes() {
            return Err(Error::Limit {
                resource: "custom-show XML input",
                limit: limits.max_input_bytes(),
            });
        }
        let mut list = Self::new();
        let xml = process_str(xml, limits)?;
        if xml.len() > limits.max_output_bytes() {
            return Err(Error::Limit {
                resource: "processed custom-show XML",
                limit: limits.max_output_bytes(),
            });
        }
        let mut reader = Reader::from_str(xml.as_ref());
        reader.config_mut().trim_text(true);

        let mut current_show: Option<Show> = None;
        let mut nodes = 0usize;
        let mut depth = 0usize;
        let mut shows = 0usize;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"custShow" => {
                    check_attributes(&e, reader.decoder(), limits)?;
                    nodes = nodes.checked_add(1).ok_or(Error::Limit {
                        resource: "custom-show XML nodes",
                        limit: limits.max_nodes(),
                    })?;
                    depth = depth.checked_add(1).ok_or(Error::Limit {
                        resource: "custom-show XML depth",
                        limit: limits.max_depth(),
                    })?;
                    if nodes > limits.max_nodes() || depth > limits.max_depth() {
                        return Err(Error::Limit {
                            resource: "custom-show XML nodes or depth",
                            limit: limits.max_nodes(),
                        });
                    }
                    shows = shows.checked_add(1).ok_or(Error::Limit {
                        resource: "custom-show entries",
                        limit: limits.max_custom_shows(),
                    })?;
                    if shows > limits.max_custom_shows() {
                        return Err(Error::Limit {
                            resource: "custom-show entries",
                            limit: limits.max_custom_shows(),
                        });
                    }
                    if current_show.is_some() {
                        return Err(Error::Invalid(
                            "nested custom-show elements are not supported".into(),
                        ));
                    }
                    let mut name = String::new();
                    let mut id = 0u32;
                    for attr in first_wins(&e).flatten() {
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
                Ok(Event::Start(e)) => {
                    nodes = nodes.checked_add(1).ok_or(Error::Limit {
                        resource: "custom-show XML nodes",
                        limit: limits.max_nodes(),
                    })?;
                    depth = depth.checked_add(1).ok_or(Error::Limit {
                        resource: "custom-show XML depth",
                        limit: limits.max_depth(),
                    })?;
                    if nodes > limits.max_nodes() || depth > limits.max_depth() {
                        return Err(Error::Limit {
                            resource: "custom-show XML nodes or depth",
                            limit: limits.max_nodes(),
                        });
                    }
                    check_attributes(&e, reader.decoder(), limits)?;
                },
                Ok(Event::Empty(e)) => {
                    check_attributes(&e, reader.decoder(), limits)?;
                    nodes = nodes.checked_add(1).ok_or(Error::Limit {
                        resource: "custom-show XML nodes",
                        limit: limits.max_nodes(),
                    })?;
                    if nodes > limits.max_nodes() {
                        return Err(Error::Limit {
                            resource: "custom-show XML nodes",
                            limit: limits.max_nodes(),
                        });
                    }
                    if e.local_name().as_ref() == b"sld"
                        && let Some(ref mut show) = current_show
                    {
                        for attr in first_wins(&e).flatten() {
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
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid("custom-show XML has an unmatched end tag".into())
                    })?;
                    if e.local_name().as_ref() == b"custShow"
                        && let Some(show) = current_show.take()
                    {
                        list.add(show);
                    } else if e.local_name().as_ref() == b"custShow" {
                        return Err(Error::Invalid(
                            "custom-show XML closes an unopened show".into(),
                        ));
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        if depth != 0 || current_show.is_some() {
            return Err(Error::Invalid(
                "custom-show XML has an unterminated element".into(),
            ));
        }
        validate(&list, limits)?;
        Ok(list)
    }

    /// Generate custom shows XML.
    ///
    /// This is a convenience wrapper that emits a best-effort mapping by treating
    /// numeric slide IDs as their corresponding `rId` values (e.g. slide 5 → `rId5`).
    /// For full fidelity, prefer [`Self::to_xml_with_rel_map`] and pass the presentation's
    /// actual slide relationship map.
    #[must_use]
    pub fn to_xml(&self, limits: Limits) -> Result<String> {
        validate(self, limits)?;
        if self.is_empty() {
            return Ok(String::new());
        }

        let mut fallback_map = HashMap::new();
        for show in &self.shows {
            for slide_id in &show.slide_ids {
                let relationship_bytes =
                    3usize
                        .checked_add(decimal_len(*slide_id))
                        .ok_or(Error::CapacityOverflow {
                            resource: "custom-show relationship ID",
                        })?;
                if relationship_bytes > limits.max_string_bytes() {
                    return Err(Error::Limit {
                        resource: "custom-show relationship ID",
                        limit: limits.max_string_bytes(),
                    });
                }
                fallback_map
                    .entry(*slide_id)
                    .or_insert_with(|| format!("rId{slide_id}"));
            }
        }

        self.to_xml_with_rel_map(&fallback_map, limits)
    }

    /// Generate custom shows XML with proper relationship ID mapping.
    ///
    /// # Arguments
    /// * `slide_id_to_rel_id` - Mapping from slide ID (e.g., 256) to relationship ID (e.g., "rId6")
    #[must_use]
    pub fn to_xml_with_rel_map(
        &self,
        slide_id_to_rel_id: &HashMap<u32, String>,
        limits: Limits,
    ) -> Result<String> {
        validate(self, limits)?;
        if self.is_empty() {
            return Ok(String::new());
        }

        let mut xml = BoundedString::new(limits, "serialized custom-show XML");

        xml.push_str("<p:custShowLst>")?;

        for show in &self.shows {
            xml.push_str(r#"<p:custShow name="#)?;
            xml.push_escaped_xml(&show.name)?;
            xml.write_fmt(format_args!(r#"" id="{}">"#, show.id))?;
            xml.push_str("<p:sldLst>")?;
            for slide_id in &show.slide_ids {
                // Look up the relationship ID for this slide ID
                if let Some(rel_id) = slide_id_to_rel_id.get(slide_id) {
                    check_string(rel_id, limits, "custom-show relationship ID")?;
                    xml.push_str(r#"<p:sld r:id="#)?;
                    xml.push_escaped_xml(rel_id)?;
                    xml.push_str(r#""/>"#)?;
                }
            }
            xml.push_str("</p:sldLst>")?;
            xml.push_str("</p:custShow>")?;
        }

        xml.push_str("</p:custShowLst>")?;
        Ok(xml.finish())
    }
}

fn validate(value: &List, limits: Limits) -> Result<()> {
    if value.shows.len() > limits.max_custom_shows() {
        return Err(Error::Limit {
            resource: "custom-show entries",
            limit: limits.max_custom_shows(),
        });
    }
    for show in &value.shows {
        if show.name.len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                resource: "custom-show name",
                limit: limits.max_string_bytes(),
            });
        }
        if show.slide_ids.len() > limits.max_slide_references() {
            return Err(Error::Limit {
                resource: "custom-show slide references",
                limit: limits.max_slide_references(),
            });
        }
    }
    Ok(())
}

fn decimal_len(mut value: u32) -> usize {
    let mut length = 1;
    while value >= 10 {
        value /= 10;
        length += 1;
    }
    length
}

fn check_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count.checked_add(1).ok_or(Error::Limit {
            resource: "custom-show XML attributes",
            limit: limits.max_attributes(),
        })?;
        if count > limits.max_attributes() {
            return Err(Error::Limit {
                resource: "custom-show XML attributes",
                limit: limits.max_attributes(),
            });
        }
        if attribute.key.as_ref().len() > limits.max_string_bytes()
            || attribute.value.as_ref().len() > limits.max_string_bytes()
        {
            return Err(Error::Limit {
                resource: "custom-show XML attribute",
                limit: limits.max_string_bytes(),
            });
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                resource: "custom-show XML attribute",
                limit: limits.max_string_bytes(),
            });
        }
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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

        let xml = list.to_xml(Limits::DEFAULT).unwrap();
        assert!(xml.contains("Demo"));
        assert!(xml.contains("custShow"));
    }

    /// `19_999` distinct filler attributes, then `name="first"`, then `20_000`
    /// repeats of `name="repeat"`: a checked attribute iterator that skips its
    /// duplicate errors scans every earlier name for each repeat.
    fn crowded(name: &str, first: &str, repeat: &str) -> String {
        const DISTINCT: usize = 20_000;
        let mut attributes = String::new();
        for index in 0..DISTINCT - 1 {
            attributes.push_str(&format!(" n{index}=\"{index}\""));
        }
        attributes.push_str(&format!(" {name}=\"{first}\""));
        let repeated = format!(" {name}=\"{repeat}\"");
        for _ in 0..DISTINCT {
            attributes.push_str(&repeated);
        }
        attributes
    }

    #[test]
    fn parsed_custom_show_attributes_keep_their_first_occurrence() {
        let single = List::parse_xml(
            r#"<p:custShowLst><p:custShow name="Recap" id="8"><p:sldLst><p:sld r:id="rId2"/></p:sldLst></p:custShow></p:custShowLst>"#,
            Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(
            single.shows,
            vec![Show::new(8, "Recap").with_slides(vec![2])]
        );

        // `id` and `r:id` are distinct names, so a slide reference may carry
        // both; a repeated name is read once, at its first occurrence.
        let repeated = List::parse_xml(
            r#"<p:custShowLst><p:custShow name="First" id="7" name="Second" id="9"><p:sldLst><p:sld r:id="rId3" r:id="rId4"/><p:sld id="5" r:id="rId6" id="10"/></p:sldLst></p:custShow></p:custShowLst>"#,
            Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(
            repeated.shows,
            vec![Show::new(7, "First").with_slides(vec![3, 5, 6])]
        );
    }

    #[test]
    fn crowded_custom_show_attributes_are_read_first_wins() {
        let xml = format!(
            r#"<p:custShowLst><p:custShow{} id="9"><p:sldLst><p:sld{}/></p:sldLst></p:custShow></p:custShowLst>"#,
            crowded("name", "First", "Later"),
            crowded("r:id", "rId3", "rId4"),
        );
        let list = List::parse_xml(&xml, Limits::DEFAULT).unwrap();
        assert_eq!(list.shows, vec![Show::new(9, "First").with_slides(vec![3])]);
    }
}
