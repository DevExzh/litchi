use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use std::fmt;
use std::fmt::Write as _;

pub fn write_a_blip_embed(xml: &mut String, rid: &str, include_xmlns_r: bool) -> fmt::Result {
    if include_xmlns_r {
        write!(
            xml,
            r#"<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"{}\"/>"#,
            rid
        )
    } else {
        write!(xml, r#"<a:blip r:embed=\"{}\"/>"#, rid)
    }
}

pub fn write_a_blip_embed_rid_num(
    xml: &mut String,
    rid_num: u32,
    include_xmlns_r: bool,
) -> fmt::Result {
    if include_xmlns_r {
        write!(
            xml,
            r#"<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId{}\"/>"#,
            rid_num
        )
    } else {
        write!(xml, r#"<a:blip r:embed=\"rId{}\"/>"#, rid_num)
    }
}

pub fn read_blip_embed_attr(e: &BytesStart<'_>) -> Result<Option<String>> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() != b"embed" {
            continue;
        }

        let rid = std::str::from_utf8(&attr.value).map_err(|e| OoxmlError::Xml(e.to_string()))?;
        return Ok(Some(rid.to_string()));
    }

    Ok(None)
}

pub fn find_first_blip_embed(xml_bytes: &[u8]) -> Result<Option<String>> {
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() != b"blip" {
                    continue;
                }

                if let Some(rid) = read_blip_embed_attr(&e)? {
                    return Ok(Some(rid));
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(e.to_string())),
            _ => {},
        }
    }

    Ok(None)
}
