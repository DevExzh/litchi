//! Minimal workbook `calcPr` invalidation after semantic cell edits.

use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use crate::error::{Result, allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;

#[derive(Debug)]
struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug)]
struct Tag {
    name: Box<str>,
    attributes: Box<[Attribute]>,
}

#[derive(Debug)]
struct Existing {
    start: usize,
    end: usize,
    tag: Tag,
}

#[derive(Debug)]
struct Pending {
    start: usize,
    tag: Tag,
}

/// Force consumers to recalculate formulas while retaining the workbook's
/// chosen automatic/manual calculation mode.
pub(crate) fn invalidate(content: &[u8]) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(content);
    let mut depth = 0usize;
    let mut root_name = None::<Box<str>>;
    let mut root_close = None;
    let mut insertion = None;
    let mut existing = None;
    let mut pending = None::<Pending>;

    loop {
        let event_start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if pending.is_some() && depth >= 2 {
                    return Err(invalid("workbook calcPr must not contain child elements"));
                }
                if depth == 0 {
                    if root_name.is_some()
                        || !is_spreadsheetml_name(&namespace, element.name(), b"workbook")
                    {
                        return Err(invalid("calculation invalidation requires a workbook root"));
                    }
                    root_name = Some(element_name(&element)?);
                } else if depth == 1 && is_spreadsheetml_name(&namespace, element.name(), b"calcPr")
                {
                    if existing.is_some() || pending.is_some() {
                        return Err(invalid("workbook has duplicate calcPr elements"));
                    }
                    pending = Some(Pending {
                        start: event_start,
                        tag: tag(&element, decoder)?,
                    });
                } else if depth == 1
                    && insertion.is_none()
                    && after_calc_properties(element.name().local_name().as_ref())
                    && is_spreadsheetml_name(
                        &namespace,
                        element.name(),
                        element.name().local_name().as_ref(),
                    )
                {
                    insertion = Some(event_start);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("workbook XML depth overflow"))?;
            },
            Event::Empty(element) => {
                if pending.is_some() && depth >= 2 {
                    return Err(invalid("workbook calcPr must not contain child elements"));
                }
                if depth == 0 {
                    return Err(invalid("workbook root cannot be empty"));
                }
                if depth == 1 && is_spreadsheetml_name(&namespace, element.name(), b"calcPr") {
                    if existing.is_some() || pending.is_some() {
                        return Err(invalid("workbook has duplicate calcPr elements"));
                    }
                    existing = Some(Existing {
                        start: event_start,
                        end: event_end,
                        tag: tag(&element, decoder)?,
                    });
                } else if depth == 1
                    && insertion.is_none()
                    && after_calc_properties(element.name().local_name().as_ref())
                    && is_spreadsheetml_name(
                        &namespace,
                        element.name(),
                        element.name().local_name().as_ref(),
                    )
                {
                    insertion = Some(event_start);
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("workbook has an unmatched closing element"));
                }
                if depth == 2 && pending.is_some() {
                    let value = pending
                        .take()
                        .ok_or_else(|| invalid("workbook calcPr state was lost"))?;
                    existing = Some(Existing {
                        start: value.start,
                        end: event_end,
                        tag: value.tag,
                    });
                } else if depth == 1 {
                    root_close = Some(event_start);
                }
                depth -= 1;
            },
            Event::Text(text) if pending.is_some() && depth >= 2 => {
                if !text
                    .decode()
                    .map_err(|error| invalid(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid("workbook calcPr must not contain text"));
                }
            },
            Event::Comment(_) | Event::CData(_) | Event::PI(_) if pending.is_some() => {
                return Err(invalid("workbook calcPr contains unsupported markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || root_name.is_none() || root_close.is_none() || pending.is_some() {
        return Err(invalid("workbook XML is incomplete"));
    }

    let mut replacement = Vec::new();
    let generated;
    let tag = if let Some(existing) = existing.as_ref() {
        &existing.tag
    } else {
        let name = sibling_name(
            root_name
                .as_deref()
                .ok_or_else(|| invalid("workbook root name was lost"))?,
            "calcPr",
        );
        generated = Tag {
            name: name.into_boxed_str(),
            attributes: Box::new([]),
        };
        &generated
    };
    write_calc_properties(&mut replacement, tag);

    let (start, end) = existing.map_or_else(
        || {
            let at = insertion.or(root_close).unwrap_or(content.len());
            (at, at)
        },
        |existing| (existing.start, existing.end),
    );
    let mut output = Vec::new();
    output
        .try_reserve(content.len().saturating_add(replacement.len()))
        .map_err(|source| allocation("workbook edit output", source))?;
    output.extend_from_slice(&content[..start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&content[end..]);
    Ok(output)
}

fn write_calc_properties(output: &mut Vec<u8>, tag: &Tag) {
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        if matches!(
            attribute.name.as_ref(),
            "calcId" | "fullCalcOnLoad" | "forceFullCalc" | "calcCompleted" | "calcOnSave"
        ) {
            continue;
        }
        output.extend_from_slice(b" ");
        output.extend_from_slice(attribute.name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(&attribute.value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    output.extend_from_slice(
        b" calcId=\"0\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\" calcCompleted=\"0\" calcOnSave=\"1\"/>",
    );
}

fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = element_name(element)?;
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("workbook attribute name is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        attributes.push(Attribute {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        });
    }
    Ok(Tag {
        name,
        attributes: attributes.into_boxed_slice(),
    })
}

fn element_name(element: &BytesStart<'_>) -> Result<Box<str>> {
    std::str::from_utf8(element.name().as_ref())
        .map(str::to_owned)
        .map(String::into_boxed_str)
        .map_err(|error| invalid(format!("workbook element name is not UTF-8: {error}")))
}

fn sibling_name(name: &str, local: &str) -> String {
    name.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

fn after_calc_properties(local: &[u8]) -> bool {
    matches!(
        local,
        b"oleSize"
            | b"customWorkbookViews"
            | b"pivotCaches"
            | b"smartTagPr"
            | b"smartTagTypes"
            | b"webPublishing"
            | b"fileRecoveryPr"
            | b"webPublishObjects"
            | b"extLst"
    )
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("workbook XML position does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn inserts_or_updates_calc_properties_without_changing_other_markup() {
        let source = format!(
            r#"<x:workbook xmlns:x="{S}" xmlns:z="urn:future"><x:sheets/><x:pivotCaches z:keep="yes"/><x:extLst><z:data/></x:extLst></x:workbook>"#
        );
        let updated = invalidate(source.as_bytes()).expect("invalidate");
        let updated = std::str::from_utf8(&updated).expect("UTF-8");
        let calc = updated.find("<x:calcPr").expect("calcPr");
        let pivot = updated.find("<x:pivotCaches").expect("pivot caches");
        assert!(calc < pivot);
        assert!(updated.contains("calcId=\"0\""));
        assert!(updated.contains("<x:extLst><z:data/></x:extLst>"));

        let source = format!(
            r#"<workbook xmlns="{S}"><sheets/><calcPr calcMode="manual" calcId="42" z:future="kept" xmlns:z="urn:future"/></workbook>"#
        );
        let updated = invalidate(source.as_bytes()).expect("update");
        let updated = std::str::from_utf8(&updated).expect("UTF-8");
        assert!(updated.contains("calcMode=\"manual\""));
        assert!(updated.contains("z:future=\"kept\""));
        assert_eq!(updated.matches("calcId=").count(), 1);
    }
}
