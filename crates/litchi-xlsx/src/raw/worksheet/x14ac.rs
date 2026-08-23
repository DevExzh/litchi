//! Narrow capture of supported x14ac attributes before MCE preprocessing.
//!
//! Treating the complete extension namespace as understood would make
//! `MustUnderstand` unsound. This scanner instead records only direct
//! `dyDescent` attributes whose core parent structure is already modeled.

use std::collections::BTreeMap;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::ROWS;
use quick_xml::Writer;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::{super::namespace::is_spreadsheetml_name, parse_one_based_row};
use crate::error::{Result, allocation, invalid};
use crate::layout::Descent;

pub(crate) const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const MAX_XML_DEPTH: usize = 256;
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MARKER: &[u8] = b"litchi_x14ac_dyDescent";

#[derive(Debug, Default)]
pub(crate) struct Values {
    pub(crate) defaults: Option<Descent>,
    pub(crate) rows: BTreeMap<u32, Descent>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Worksheet,
    SheetData,
    Row,
    Other,
}

pub(crate) fn capture(content: &[u8]) -> Result<Values> {
    if has_markup_compatibility(content)
        && has_descent_attribute(content)
        && has_alternate_content(content)
    {
        return capture_active(content, true);
    }
    capture_inner(content, true, false)
}

pub(crate) fn capture_defaults(content: &[u8]) -> Result<Option<Descent>> {
    if has_markup_compatibility(content)
        && has_descent_attribute(content)
        && has_alternate_content(content)
    {
        return Ok(capture_active(content, false)?.defaults);
    }
    Ok(capture_inner(content, false, false)?.defaults)
}

fn capture_active(content: &[u8], capture_rows: bool) -> Result<Values> {
    let rewritten = rewrite_descent_attributes(content)?;
    let processed = litchi_ooxml_common::mce::process_ooxml(&rewritten)?;
    capture_inner(processed.as_ref(), capture_rows, true)
}

fn capture_inner(content: &[u8], capture_rows: bool, marker_only: bool) -> Result<Values> {
    let mut reader = NsReader::from_reader(content);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::<Context>::new();
    let mut values = Values::default();
    let mut previous_row = 0u32;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid worksheet extension XML: {error}")))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let decoder = reader.decoder();
        let resolver = reader.resolver();
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "worksheet extension XML exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
                let context = start(
                    stack.last().copied(),
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    &mut previous_row,
                    &mut values,
                    capture_rows,
                    marker_only,
                )?;
                stack.push(context);
            },
            Event::Empty(element) => {
                start(
                    stack.last().copied(),
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    &mut previous_row,
                    &mut values,
                    capture_rows,
                    marker_only,
                )?;
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| invalid("worksheet extension XML has an unexpected end"))?;
            },
            Event::Eof if !stack.is_empty() => {
                return Err(invalid(
                    "worksheet extension XML has an unterminated element",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(values)
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments correspond directly to the x14ac worksheet attributes"
)]
fn start(
    parent: Option<Context>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    previous_row: &mut u32,
    values: &mut Values,
    capture_rows: bool,
    marker_only: bool,
) -> Result<Context> {
    if parent.is_none() && is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
        return Ok(Context::Worksheet);
    }
    if parent == Some(Context::Worksheet)
        && is_spreadsheetml_name(namespace, element.name(), b"sheetFormatPr")
    {
        if let Some(value) = descent(element, decoder, resolver, marker_only)?
            && values.defaults.replace(value).is_some()
        {
            return Err(invalid("duplicate worksheet default dyDescent"));
        }
        return Ok(Context::Other);
    }
    if parent == Some(Context::Worksheet)
        && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
    {
        return Ok(Context::SheetData);
    }
    if parent == Some(Context::SheetData)
        && is_spreadsheetml_name(namespace, element.name(), b"row")
    {
        if !capture_rows {
            return Ok(Context::Row);
        }
        let number = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(value) => parse_one_based_row(&value)?,
            None => previous_row
                .checked_add(1)
                .filter(|value| *value <= ROWS)
                .ok_or_else(|| invalid("inferred extension row exceeds the spreadsheet grid"))?,
        };
        *previous_row = number;
        if let Some(value) = descent(element, decoder, resolver, marker_only)?
            && values.rows.insert(number, value).is_some()
        {
            return Err(invalid(format!(
                "duplicate worksheet row {number} dyDescent"
            )));
        }
        return Ok(Context::Row);
    }
    Ok(Context::Other)
}

fn has_markup_compatibility(content: &[u8]) -> bool {
    content
        .windows(MCE_NAMESPACE.len())
        .any(|window| window == MCE_NAMESPACE)
}

fn has_descent_attribute(content: &[u8]) -> bool {
    may_contain_descent(content)
}

pub(super) fn may_contain_descent(content: &[u8]) -> bool {
    memchr::memmem::find(content, b"dyDescent").is_some()
}

fn has_alternate_content(content: &[u8]) -> bool {
    content
        .windows(b"AlternateContent".len())
        .any(|window| window == b"AlternateContent")
}

fn rewrite_descent_attributes(content: &[u8]) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(content);
    reader.config_mut().check_end_names = true;
    let mut output = Vec::new();
    output
        .try_reserve_exact(content.len())
        .map_err(|source| allocation("worksheet extension rewrite", source))?;
    let mut writer = Writer::new(output);

    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid worksheet extension XML: {error}")))?;
        let resolver = reader.resolver();
        match event {
            Event::Start(element) => writer
                .write_event(Event::Start(rewrite_element(&element, &resolver)?))
                .map_err(|error| invalid(format!("could not rewrite worksheet XML: {error}")))?,
            Event::Empty(element) => writer
                .write_event(Event::Empty(rewrite_element(&element, &resolver)?))
                .map_err(|error| invalid(format!("could not rewrite worksheet XML: {error}")))?,
            Event::Eof => break,
            other @ (Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)) => writer
                .write_event(other)
                .map_err(|error| invalid(format!("could not rewrite worksheet XML: {error}")))?,
        }
    }
    Ok(writer.into_inner())
}

fn rewrite_element(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<BytesStart<'static>> {
    let element_name = element.name();
    let name = std::str::from_utf8(element_name.as_ref())
        .map_err(|error| invalid(format!("worksheet element name is not UTF-8: {error}")))?;
    let mut rewritten = BytesStart::new(name.to_owned());
    let mut replaced = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if attribute.key.as_ref() == MARKER {
            return Err(invalid("worksheet XML uses a reserved internal marker"));
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() == b"dyDescent"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == NAMESPACE)
        {
            if replaced {
                return Err(invalid("duplicate x14ac:dyDescent attribute"));
            }
            rewritten.push_attribute((MARKER, attribute.value.as_ref()));
            replaced = true;
        } else {
            rewritten.push_attribute((attribute.key.as_ref(), attribute.value.as_ref()));
        }
    }
    Ok(rewritten)
}

pub(crate) fn descent(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    marker_only: bool,
) -> Result<Option<Descent>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let is_target = if marker_only {
            local.as_ref() == MARKER && matches!(namespace, ResolveResult::Unbound)
        } else {
            local.as_ref() == b"dyDescent"
                && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == NAMESPACE)
        };
        if !is_target {
            continue;
        }
        let lexical = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        let parsed = lexical
            .parse::<f64>()
            .map_err(|_source| invalid(format!("invalid x14ac:dyDescent value '{lexical}'")))?;
        let parsed = Descent::new(parsed)?;
        if result.replace(parsed).is_some() {
            return Err(invalid("duplicate x14ac:dyDescent attribute"));
        }
    }
    Ok(result)
}

pub(crate) fn attribute_name(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
) -> Result<Option<Box<str>>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() != b"dyDescent"
            || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == NAMESPACE)
        {
            continue;
        }
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("x14ac attribute name is not UTF-8: {error}")))?
            .to_owned()
            .into_boxed_str();
        if result.replace(name).is_some() {
            return Err(invalid("duplicate x14ac:dyDescent attribute"));
        }
    }
    Ok(result)
}
