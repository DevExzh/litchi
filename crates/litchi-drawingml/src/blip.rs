//! DrawingML image-reference primitives.

use std::{fmt, fmt::Write as _};

use quick_xml::{Reader, events::BytesStart, events::Event};

use crate::{Error, Result};

/// Write an embedded-image reference, escaping the relationship ID.
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
pub fn read_embed(element: &BytesStart<'_>) -> Result<Option<String>> {
    for attribute in element.attributes() {
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

/// Find the first embedded-image relationship ID in DrawingML XML.
pub fn find_first_embed(xml: &[u8]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                if element.local_name().as_ref() == b"blip"
                    && let Some(relationship_id) = read_embed(&element)?
                {
                    return Ok(Some(relationship_id));
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
            find_first_embed(br#"<a:blip r:embed="rId7"/>"#).expect("valid blip"),
            Some("rId7".to_owned())
        );
    }

    #[test]
    fn malformed_attributes_are_not_silently_discarded() {
        assert!(find_first_embed(br#"<a:blip r:embed="unterminated/>"#).is_err());
    }
}
