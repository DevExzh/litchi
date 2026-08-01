//! Narrow capture of supported x14ac attributes before MCE preprocessing.
//!
//! Treating the complete extension namespace as understood would make
//! `MustUnderstand` unsound. This scanner instead records only direct
//! `dyDescent` attributes whose core parent structure is already modeled.

use std::collections::BTreeMap;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::ROWS;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::{super::namespace::is_spreadsheetml_name, parse_one_based_row};
use crate::error::{Result, invalid};
use crate::layout::Descent;

pub(crate) const NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const MAX_XML_DEPTH: usize = 256;

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
    let mut reader = NsReader::from_reader(content);
    let mut stack = Vec::<Context>::new();
    let mut values = Values::default();
    let mut previous_row = 0u32;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid worksheet extension XML: {error}")))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
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
                )?;
            },
            Event::End(_) => {
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(values)
}

#[allow(clippy::too_many_arguments)]
fn start(
    parent: Option<Context>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    previous_row: &mut u32,
    values: &mut Values,
) -> Result<Context> {
    if parent.is_none() && is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
        return Ok(Context::Worksheet);
    }
    if parent == Some(Context::Worksheet)
        && is_spreadsheetml_name(namespace, element.name(), b"sheetFormatPr")
    {
        if let Some(value) = descent(element, decoder, resolver)?
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
        let number = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(value) => parse_one_based_row(&value)?,
            None => previous_row
                .checked_add(1)
                .filter(|value| *value <= ROWS)
                .ok_or_else(|| invalid("inferred extension row exceeds the spreadsheet grid"))?,
        };
        *previous_row = number;
        if let Some(value) = descent(element, decoder, resolver)?
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

pub(crate) fn descent(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<Descent>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() != b"dyDescent"
            || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == NAMESPACE)
        {
            continue;
        }
        let lexical = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        let parsed = lexical
            .parse::<f64>()
            .map_err(|_| invalid(format!("invalid x14ac:dyDescent value '{lexical}'")))?;
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
