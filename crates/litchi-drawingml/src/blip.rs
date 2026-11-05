//! `DrawingML` image-reference primitives.

use std::{fmt, fmt::Write as _};

use quick_xml::{Reader, events::BytesStart, events::Event};

use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

/// Caller-selected bounds for the arbitrary-XML `blip` scanner.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct Limits {
    /// Maximum XML bytes inspected by [`find_first_embed`].
    pub max_input_bytes: usize,
    /// Maximum relationship identifier retained by the scanner.
    pub max_relationship_id_bytes: usize,
}

impl Limits {
    /// Historical scanner policy.
    pub const DEFAULT: Self = Self {
        max_input_bytes: 64 * 1024 * 1024,
        max_relationship_id_bytes: 256,
    };

    /// Construct a scanner policy.
    #[must_use]
    pub const fn new(max_input_bytes: usize, max_relationship_id_bytes: usize) -> Self {
        Self {
            max_input_bytes,
            max_relationship_id_bytes,
        }
    }

    fn validate(self) -> Result<Self> {
        if self.max_input_bytes == 0 {
            return Err(Error::invalid_limit(
                crate::LimitKind::InputBytes,
                self.max_input_bytes,
                Self::DEFAULT.max_input_bytes,
            ));
        }
        if self.max_relationship_id_bytes == 0 || self.max_relationship_id_bytes > 256 {
            return Err(Error::invalid_limit(
                crate::LimitKind::AttributeValueBytes,
                self.max_relationship_id_bytes,
                256,
            ));
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Write an embedded-image reference, escaping the relationship ID.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write_embed(xml: &mut String, relationship_id: &str, include_xmlns_r: bool) -> fmt::Result {
    let relationship_id = quick_xml::escape::escape(relationship_id);
    if include_xmlns_r {
        write!(
            xml,
            r#"<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{relationship_id}"/>"#,
        )
    } else {
        write!(xml, r#"<a:blip r:embed="{relationship_id}"/>"#)
    }
}

/// Write an `rIdN` embedded-image reference without allocating the ID.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write_embed_id(
    xml: &mut String,
    relationship_number: u32,
    include_xmlns_r: bool,
) -> fmt::Result {
    if include_xmlns_r {
        write!(
            xml,
            r#"<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId{relationship_number}"/>"#,
        )
    } else {
        write!(xml, r#"<a:blip r:embed="rId{relationship_number}"/>"#)
    }
}

/// Read an embedded relationship ID from a `blip` start element.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read_embed(element: &BytesStart<'_>) -> Result<Option<String>> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"embed" {
            continue;
        }

        let relationship_id =
            std::str::from_utf8(&attribute.value).map_err(|error| Error::Xml(error.to_string()))?;
        return Ok(Some(relationship_id.to_owned()));
    }
    Ok(None)
}

/// Find the first embedded-image relationship ID in `DrawingML` XML.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn find_first_embed(xml: &[u8], limits: Limits) -> Result<Option<String>> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_input_bytes {
        return Err(Error::limit(
            crate::LimitKind::InputBytes,
            xml.len(),
            limits.max_input_bytes,
        ));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                if element.local_name().as_ref() == b"blip" {
                    for attribute in element.checked_attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        if attribute.key.local_name().as_ref() == b"embed"
                            && attribute.value.len() > limits.max_relationship_id_bytes
                        {
                            return Err(Error::limit(
                                crate::LimitKind::AttributeValueBytes,
                                attribute.value.len(),
                                limits.max_relationship_id_bytes,
                            ));
                        }
                    }
                    if let Some(relationship_id) = read_embed(&element)? {
                        return Ok(Some(relationship_id));
                    }
                }
            },
            Ok(Event::Eof) => return Ok(None),
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_escaped_ids_and_reads_normal_ids() {
        let mut xml = String::new();
        write_embed(&mut xml, "rId<&\"", false).expect("String formatting is infallible");
        assert_eq!(xml, r#"<a:blip r:embed="rId&lt;&amp;&quot;"/>"#);

        assert_eq!(
            find_first_embed(br#"<a:blip r:embed="rId7"/>"#, Limits::DEFAULT).expect("valid blip"),
            Some("rId7".to_owned())
        );
    }

    #[test]
    fn malformed_attributes_are_not_silently_discarded() {
        assert!(find_first_embed(br#"<a:blip r:embed="unterminated/>"#, Limits::DEFAULT).is_err());
    }
}
