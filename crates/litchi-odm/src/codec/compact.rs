//! Lossless compaction of producer XML formatting whitespace.

use litchi_core::{Error, Result};
use quick_xml::{Reader, Writer, events::Event};

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;

pub(crate) fn compact_source_xml(source: &str) -> Result<String> {
    if source.len() > MAX_XML_BYTES {
        return Err(invalid(
            "ODM XML source exceeds the 64 MiB compaction limit",
        ));
    }
    let mut reader = Reader::from_reader(source.as_bytes());
    reader.config_mut().check_end_names = true;
    let mut writer = Writer::new(Vec::new());
    let mut preserve = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid ODM XML during compaction: {error}")))?;
        match event {
            Event::Start(element) => {
                if preserve.len() >= MAX_DEPTH {
                    return Err(invalid("ODM XML compaction depth exceeds the limit"));
                }
                preserve
                    .try_reserve(1)
                    .map_err(|allocation_error| Error::Allocation {
                        resource: "ODM XML compaction stack",
                        source: allocation_error,
                    })?;
                let inherited = preserve.last().copied().unwrap_or(false);
                preserve.push(xml_space(&element).unwrap_or(inherited));
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(|error| invalid(format!("ODM XML compaction failed: {error}")))?;
            },
            Event::End(element) => {
                preserve
                    .pop()
                    .ok_or_else(|| invalid("ODM XML compaction depth underflow"))?;
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(|error| invalid(format!("ODM XML compaction failed: {error}")))?;
            },
            Event::Text(text)
                if !preserve.last().copied().unwrap_or(false)
                    && text.as_ref().iter().all(u8::is_ascii_whitespace)
                    && text
                        .as_ref()
                        .iter()
                        .any(|byte| matches!(byte, b'\n' | b'\r' | b'\t')) => {},
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM XML")),
            Event::GeneralRef(reference)
                if !matches!(
                    reference.as_ref(),
                    b"amp" | b"apos" | b"gt" | b"lt" | b"quot"
                ) =>
            {
                return Err(invalid("named entities are not allowed in ODM XML"));
            },
            Event::Eof => break,
            remaining @ (Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_)) => writer
                .write_event(remaining.into_owned())
                .map_err(|error| invalid(format!("ODM XML compaction failed: {error}")))?,
        }
    }
    if !preserve.is_empty() {
        return Err(invalid("ODM XML compaction nesting is incomplete"));
    }
    let bytes = writer.into_inner();
    let compact = String::from_utf8(bytes)
        .map_err(|error| invalid(format!("ODM compact XML is not UTF-8: {error}")))?;
    litchi_odf_common::compact_xml::validate(compact.as_bytes()).map_err(Error::from)?;
    Ok(compact)
}

fn xml_space(element: &quick_xml::events::BytesStart<'_>) -> Option<bool> {
    element.attributes().flatten().find_map(|attribute| {
        (attribute.key.as_ref() == b"xml:space").then(|| attribute.value.as_ref() == b"preserve")
    })
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
