//! Strict non-executing chart styles-part validation.

use litchi_core::{Error, Result};
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

pub(crate) fn validate(xml: &str, limits: crate::Limits) -> Result<()> {
    if xml.len() > limits.max_content_bytes() {
        return Err(invalid(
            "ODC styles exceed the caller-selected content limit",
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODC styles XML: {error}")))?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth, limits.max_depth())?;
                if depth == 1 {
                    if root_seen
                        || !resolved(&namespace, OFFICE)
                        || element.local_name().as_ref() != b"document-styles"
                    {
                        return Err(invalid(
                            "ODC styles require one office:document-styles root",
                        ));
                    }
                    root_seen = true;
                }
                if resolved(&namespace, OFFICE) && element.local_name().as_ref() == b"scripts" {
                    return Err(Error::Unsupported(
                        "ODC styles refuse executable office:scripts content".into(),
                    ));
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if root_seen
                        || !resolved(&namespace, OFFICE)
                        || element.local_name().as_ref() != b"document-styles"
                    {
                        return Err(invalid(
                            "ODC styles require one office:document-styles root",
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                }
                if resolved(&namespace, OFFICE) && element.local_name().as_ref() == b"scripts" {
                    return Err(Error::Unsupported(
                        "ODC styles refuse executable office:scripts content".into(),
                    ));
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("ODC styles XML depth underflow"));
                }
                if depth == 1 {
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODC styles")),
            Event::Eof => break,
            Event::Text(text) => {
                if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("ODC styles contain text outside the root"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(invalid("ODC styles contain data outside the root"));
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 || !root_seen || !root_closed {
        return Err(invalid("ODC styles structure is incomplete"));
    }
    Ok(())
}

fn checked_depth(depth: usize, limit: usize) -> Result<usize> {
    let next = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODC styles depth overflow"))?;
    if next > limit {
        return Err(invalid("ODC styles exceed the caller-selected depth limit"));
    }
    Ok(next)
}

fn resolved(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
