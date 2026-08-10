//! Compact replacement of validated empty direct worksheet properties.

use std::ops::Range;

use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, allocation, invalid};

const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(crate) fn replace_direct_empty(
    xml: &[u8],
    local: &str,
    attributes: Option<&str>,
    successors: &[&[u8]],
    context: &'static str,
) -> Result<Vec<u8>> {
    let layout = scan(xml, local.as_bytes(), successors, context)?;
    if layout.alternate_content {
        return Err(invalid(format!(
            "{context} projected through markup compatibility cannot be edited"
        )));
    }
    if layout.span.is_none() && attributes.is_none() {
        return Ok(xml.to_vec());
    }
    let replacement = attributes.map_or_else(Vec::new, |attributes| {
        let prefix = layout
            .root_name
            .split_once(':')
            .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
        format!("<{prefix}{local}{attributes}/>").into_bytes()
    });
    let span = layout.span.unwrap_or(layout.insertion..layout.insertion);
    let capacity = xml
        .len()
        .checked_sub(span.len())
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| invalid(format!("{context} replacement size overflow")))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation(context, source))?;
    output.extend_from_slice(&xml[..span.start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&xml[span.end..]);
    crate::raw::compact::changed(&output, context)
}

struct Layout {
    root_name: String,
    span: Option<Range<usize>>,
    insertion: usize,
    alternate_content: bool,
}

fn scan(xml: &[u8], selected: &[u8], successors: &[&[u8]], context: &str) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_name = None;
    let mut open = None;
    let mut span = None;
    let mut insertion = None;
    let mut root_close = None;
    let mut alternate_content = false;
    loop {
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if mce(&namespace) && element.local_name().as_ref() == b"AlternateContent" {
                    alternate_content = true;
                }
                if depth == 0 {
                    if !sml(&namespace) || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid(format!("{context} requires one worksheet root")));
                    }
                    let qualified_name = element.name();
                    root_name = Some(
                        std::str::from_utf8(qualified_name.as_ref())
                            .map_err(|error| {
                                invalid(format!("worksheet name is not UTF-8: {error}"))
                            })?
                            .to_owned(),
                    );
                } else if depth == 1 && sml(&namespace) {
                    if element.local_name().as_ref() == selected {
                        if span.is_some() || open.is_some() {
                            return Err(invalid(format!("duplicate worksheet {context}")));
                        }
                        open = Some(start);
                    } else if successors.contains(&element.local_name().as_ref())
                        && insertion.is_none()
                    {
                        insertion = Some(start);
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid(format!("{context} XML nesting overflow")))?;
            },
            Event::Empty(element) => {
                if mce(&namespace) && element.local_name().as_ref() == b"AlternateContent" {
                    alternate_content = true;
                }
                if depth == 1 && sml(&namespace) {
                    if element.local_name().as_ref() == selected {
                        if span.is_some() || open.is_some() {
                            return Err(invalid(format!("duplicate worksheet {context}")));
                        }
                        span = Some(start..end);
                    } else if successors.contains(&element.local_name().as_ref())
                        && insertion.is_none()
                    {
                        insertion = Some(start);
                    }
                }
            },
            Event::End(element) => {
                if depth == 2 && sml(&namespace) && element.local_name().as_ref() == selected {
                    let selected_start = open
                        .take()
                        .ok_or_else(|| invalid(format!("{context} close has no start")))?;
                    span = Some(selected_start..end);
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid(format!("unexpected {context} XML end element")))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(format!(
                    "{context} rejects DTD and processing instructions"
                )));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || open.is_some() {
        return Err(invalid(format!("incomplete worksheet {context} XML")));
    }
    Ok(Layout {
        root_name: root_name.ok_or_else(|| invalid("worksheet XML has no root"))?,
        span,
        insertion: insertion
            .or(root_close)
            .ok_or_else(|| invalid(format!("worksheet has no {context} insertion point")))?,
        alternate_content,
    })
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("worksheet XML position does not fit usize"))
}

fn sml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == SML || *value == STRICT_SML)
}

fn mce(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!("worksheet property XML error: {error}"))
}
