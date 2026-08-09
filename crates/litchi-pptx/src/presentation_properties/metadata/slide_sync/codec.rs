//! Bounded XML codec for Slide Synchronization Data parts.

use litchi_ooxml_common::mce::process_ooxml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use super::model::{DateTime, Properties};
use crate::presentation_properties::metadata::escape_xml;
use crate::presentation_properties::metadata::is_presentationml_name;
use crate::{Error, Result};

const MAX_PART_XML_BYTES: usize = 1024 * 1024;
const MAX_XML_NODES: usize = 4096;
const MAX_XML_DEPTH: usize = 32;
const MAX_SERVER_SLIDE_ID_BYTES: usize = 4096;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(what: &str) -> Error {
    Error::Invalid(format!("exceeded maximum {what}"))
}

impl Properties {
    /// Parse a complete Slide Synchronization Data part.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_PART_XML_BYTES {
            return Err(limit("slide synchronization XML bytes"));
        }
        let xml = process_ooxml(xml)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut properties = None;
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let is_empty_element = matches!(event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    nodes = increment(nodes, MAX_XML_NODES, "slide synchronization XML nodes")?;
                    depth = increment(depth, MAX_XML_DEPTH, "slide synchronization XML depth")?;
                    let is_root = is_presentationml_name(&namespace, element.name(), b"sldSyncPr");
                    let is_extension_list =
                        is_presentationml_name(&namespace, element.name(), b"extLst");
                    match depth {
                        1 => {
                            if properties.is_some() || closed_root {
                                return Err(invalid(
                                    "slide synchronization part has multiple root elements",
                                ));
                            }
                            if !is_root {
                                return Err(invalid(
                                    "slide synchronization part must have a sldSyncPr root",
                                ));
                            }
                            properties = Some(from_root(&element, decoder)?);
                        },
                        2 if is_extension_list => {
                            let entry = properties.as_mut().ok_or_else(|| {
                                invalid("missing slide synchronization root element")
                            })?;
                            if entry.has_extension_list {
                                return Err(invalid(
                                    "duplicate slide synchronization extension list",
                                ));
                            }
                            entry.has_extension_list = true;
                        },
                        2 => {
                            return Err(invalid(
                                "unexpected child of the slide synchronization root",
                            ));
                        },
                        _ if is_root => {
                            return Err(invalid("nested slide synchronization root element"));
                        },
                        _ => {},
                    }
                    if is_empty_element {
                        depth -= 1;
                        if depth == 0 {
                            closed_root = true;
                        }
                    }
                },
                Event::Text(text) if depth > 0 => {
                    let content = text
                        .xml_content(quick_xml::XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    if depth == 1 && !content.trim().is_empty() {
                        return Err(invalid(
                            "slide synchronization root cannot carry text content",
                        ));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid slide synchronization XML nesting"))?;
                    if depth == 0 {
                        closed_root = true;
                    }
                },
                Event::Eof if depth != 0 => {
                    return Err(invalid("unterminated slide synchronization XML"));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        properties.ok_or_else(|| invalid("slide synchronization part is empty"))
    }

    /// Serialize a complete Slide Synchronization Data part.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        if self.server_slide_id.len() > MAX_SERVER_SLIDE_ID_BYTES {
            return Err(limit("slide synchronization server slide ID bytes"));
        }
        let mut out = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
        out.extend_from_slice(
            br#"<p:sldSyncPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" serverSldId="#,
        );
        out.extend_from_slice(escape_xml(&self.server_slide_id).as_bytes());
        out.extend_from_slice(br#"" serverSldModifiedTime="#);
        out.extend_from_slice(self.server_modified_time.to_lexical().as_bytes());
        out.extend_from_slice(br#"" clientInsertedTime="#);
        out.extend_from_slice(self.client_inserted_time.to_lexical().as_bytes());
        if self.has_extension_list {
            out.extend_from_slice(br#""><p:extLst/></p:sldSyncPr>"#);
        } else {
            out.extend_from_slice(br#""/>"#);
        }
        if out.len() > MAX_PART_XML_BYTES {
            return Err(limit("serialized slide synchronization XML bytes"));
        }
        Self::parse(&out)?;
        Ok(out)
    }
}

fn from_root(element: &BytesStart<'_>, decoder: Decoder) -> Result<Properties> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            continue;
        }
        match attribute.key.local_name().as_ref() {
            b"serverSldId" | b"serverSldModifiedTime" | b"clientInsertedTime" => {},
            _ => {
                return Err(invalid(format!(
                    "unexpected slide synchronization attribute '{}'",
                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                )));
            },
        }
    }
    let server_slide_id = unqualified_attribute_value(element, b"serverSldId", decoder)?
        .ok_or_else(|| invalid("missing slide synchronization serverSldId attribute"))?;
    if server_slide_id.len() > MAX_SERVER_SLIDE_ID_BYTES {
        return Err(limit("slide synchronization server slide ID bytes"));
    }
    let server_modified_time = DateTime::parse(
        &unqualified_attribute_value(element, b"serverSldModifiedTime", decoder)?.ok_or_else(
            || invalid("missing slide synchronization serverSldModifiedTime attribute"),
        )?,
    )?;
    let client_inserted_time = DateTime::parse(
        &unqualified_attribute_value(element, b"clientInsertedTime", decoder)?
            .ok_or_else(|| invalid("missing slide synchronization clientInsertedTime attribute"))?,
    )?;
    Ok(Properties {
        server_slide_id,
        server_modified_time,
        client_inserted_time,
        has_extension_list: false,
    })
}

fn increment(value: usize, max: usize, label: &str) -> Result<usize> {
    let next = value.checked_add(1).ok_or_else(|| limit(label))?;
    if next > max {
        return Err(limit(label));
    }
    Ok(next)
}
