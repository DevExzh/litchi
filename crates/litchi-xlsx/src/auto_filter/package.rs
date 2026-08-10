//! Worksheet-context extraction for `SpreadsheetML` auto-filters.

use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::Writer;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::model::{CORE, Definition, MAX_FRAGMENT_BYTES, STRICT};
use crate::error::{Error, Result, invalid};

pub fn parse_auto_filter(xml: &[u8]) -> Result<Option<Definition>> {
    let processed = process_ooxml(xml)?;
    let Some(fragment) = capture(processed.as_ref())? else {
        return Ok(None);
    };
    super::codec::parse_auto_filter_fragment(&fragment).map(Some)
}

fn capture(xml: &[u8]) -> Result<Option<Vec<u8>>> {
    let mut reader = NsReader::from_reader(xml);
    let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    let mut result = None;
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if let Some((depth, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            match event {
                Event::Start(_) => *depth += 1,
                Event::End(_) => *depth -= 1,
                Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_)
                | Event::Eof => {},
            }
            if *depth == 0 {
                let (_, writer) = capture.take().unwrap_or_else(|| {
                    crate::error::panic_missing_invariant(
                        "required value was checked before extraction",
                    )
                });
                let value = writer.into_inner();
                if value.len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("autoFilter is too large"));
                }
                if result.replace(value).is_some() {
                    return Err(invalid("duplicate worksheet autoFilter"));
                }
            }
            continue;
        }
        match event {
            Event::Start(e)
                if spreadsheet(&namespace) && e.local_name().as_ref() == b"autoFilter" =>
            {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Start(e)).map_err(xml_error)?;
                capture = Some((1, writer));
            },
            Event::Empty(e)
                if spreadsheet(&namespace) && e.local_name().as_ref() == b"autoFilter" =>
            {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Empty(e)).map_err(xml_error)?;
                if result.replace(writer.into_inner()).is_some() {
                    return Err(invalid("duplicate worksheet autoFilter"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if capture.is_some() {
        return Err(invalid("unterminated autoFilter"));
    }
    Ok(result)
}

fn spreadsheet(ns: &ResolveResult<'_>) -> bool {
    matches!(ns, ResolveResult::Bound(v) if v.as_ref() == CORE || v.as_ref() == STRICT)
}

fn xml_error(e: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(e.to_string()))
}
