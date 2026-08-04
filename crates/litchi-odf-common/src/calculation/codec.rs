//! Bounded XML codec for ODF calculation settings.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::num::NonZeroUsize;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

use super::model::{Iteration, IterationStatus, NullDate, Settings};

pub fn write(out: &mut String, settings: Option<&Settings>) -> Result<()> {
    let Some(settings) = settings else {
        return Ok(());
    };
    settings.validate()?;
    out.push_str("<table:calculation-settings");
    write_optional_bool(out, "table:case-sensitive", settings.case_sensitive);
    write_optional_bool(out, "table:precision-as-shown", settings.precision_as_shown);
    write_optional_bool(
        out,
        "table:search-criteria-must-apply-to-whole-cell",
        settings.search_criteria_must_apply_to_whole_cell,
    );
    write_optional_bool(
        out,
        "table:automatic-find-labels",
        settings.automatic_find_labels,
    );
    write_optional_bool(
        out,
        "table:use-regular-expressions",
        settings.use_regular_expressions,
    );
    write_optional_bool(out, "table:use-wildcards", settings.use_wildcards);
    if let Some(year) = settings.null_year {
        write_attribute(out, "table:null-year", &year.get().to_string());
    }
    if settings.null_date.is_none() && settings.iteration.is_none() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(null_date) = &settings.null_date {
        out.push_str("<table:null-date");
        if null_date.value_type_date {
            write_attribute(out, "table:value-type", "date");
        }
        if let Some(value) = &null_date.date_value {
            write_attribute(out, "table:date-value", value);
        }
        out.push_str("/>");
    }
    if let Some(iteration) = &settings.iteration {
        out.push_str("<table:iteration");
        if let Some(status) = iteration.status {
            write_attribute(out, "table:status", status.as_str());
        }
        if let Some(steps) = iteration.steps {
            write_attribute(out, "table:steps", &steps.get().to_string());
        }
        if let Some(value) = &iteration.maximum_difference {
            write_attribute(out, "table:maximum-difference", value);
        }
        out.push_str("/>");
    }
    out.push_str("</table:calculation-settings>");
    Ok(())
}

pub fn parse(xml: &str) -> Result<Option<Settings>> {
    if xml.len() > super::MAX_XML_BYTES {
        return Err(Error::InvalidFormat(
            "calculation settings XML exceeds the size limit".to_string(),
        ));
    }
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut current = None;
    let mut result = None;
    let mut child_open = false;
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut document_body_depth = None;
    loop {
        events = events.saturating_add(1);
        if events > super::MAX_EVENTS {
            return Err(Error::InvalidFormat(
                "calculation settings XML exceeds the event limit".to_string(),
            ));
        }
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_table = is_namespace(&namespace);
        let is_start = matches!(&event, Event::Start(_));
        let is_end = matches!(&event, Event::End(_));
        if let Event::Start(element) = &event
            && is_namespace_uri(&namespace, OFFICE_NAMESPACE)
            && matches!(
                element.local_name().as_ref(),
                b"chart" | b"drawing" | b"presentation" | b"spreadsheet" | b"text"
            )
        {
            document_body_depth = Some(depth);
        }
        let is_document_body_child = document_body_depth.is_some_and(|value| depth == value + 1);
        match event {
            Event::Start(element)
                if is_table && element.local_name().as_ref() == b"calculation-settings" =>
            {
                if !is_document_body_child {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings must be a direct office document-body child"
                            .to_string(),
                    ));
                }
                if current.is_some() || result.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate table:calculation-settings".to_string(),
                    ));
                }
                current = Some(parse_settings_attributes(
                    reader.resolver(),
                    reader.decoder(),
                    &element,
                )?);
            },
            Event::Empty(element)
                if is_table && element.local_name().as_ref() == b"calculation-settings" =>
            {
                if !is_document_body_child {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings must be a direct office document-body child"
                            .to_string(),
                    ));
                }
                if current.is_some() || result.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate table:calculation-settings".to_string(),
                    ));
                }
                result = Some(parse_settings_attributes(
                    reader.resolver(),
                    reader.decoder(),
                    &element,
                )?);
            },
            Event::Start(element) if current.is_some() => {
                if child_open {
                    return Err(Error::InvalidFormat(
                        "calculation setting children must be empty".to_string(),
                    ));
                }
                parse_settings_child(
                    current.as_mut().expect("settings were checked"),
                    reader.resolver(),
                    reader.decoder(),
                    is_table,
                    &element,
                )?;
                child_open = true;
            },
            Event::Empty(element) if current.is_some() => {
                if child_open {
                    return Err(Error::InvalidFormat(
                        "calculation setting children must be empty".to_string(),
                    ));
                }
                parse_settings_child(
                    current.as_mut().expect("settings were checked"),
                    reader.resolver(),
                    reader.decoder(),
                    is_table,
                    &element,
                )?;
            },
            Event::End(element) if current.is_some() => {
                if child_open {
                    child_open = false;
                } else if is_table && element.local_name().as_ref() == b"calculation-settings" {
                    result = current.take();
                }
            },
            Event::Text(text) if current.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid calculation settings text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings cannot contain text".to_string(),
                    ));
                }
            },
            Event::CData(text) if current.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid calculation settings CDATA: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings cannot contain CDATA".to_string(),
                    ));
                }
            },
            Event::GeneralRef(_) if current.is_some() => {
                return Err(Error::InvalidFormat(
                    "table:calculation-settings cannot contain entity references".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        if is_start {
            if depth >= super::MAX_DEPTH {
                return Err(Error::InvalidFormat(
                    "calculation settings XML exceeds the nesting limit".to_string(),
                ));
            }
            depth = depth.saturating_add(1);
        } else if is_end {
            depth = depth.saturating_sub(1);
        }
        buf.clear();
    }
    if current.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated table:calculation-settings".to_string(),
        ));
    }
    if let Some(settings) = &result {
        settings.validate()?;
    }
    Ok(result)
}

fn parse_settings_attributes(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
) -> Result<Settings> {
    Ok(Settings {
        case_sensitive: optional_bool(resolver, decoder, element, b"case-sensitive")?,
        precision_as_shown: optional_bool(resolver, decoder, element, b"precision-as-shown")?,
        search_criteria_must_apply_to_whole_cell: optional_bool(
            resolver,
            decoder,
            element,
            b"search-criteria-must-apply-to-whole-cell",
        )?,
        automatic_find_labels: optional_bool(resolver, decoder, element, b"automatic-find-labels")?,
        use_regular_expressions: optional_bool(
            resolver,
            decoder,
            element,
            b"use-regular-expressions",
        )?,
        use_wildcards: optional_bool(resolver, decoder, element, b"use-wildcards")?,
        null_year: optional_positive(resolver, decoder, element, b"null-year")?,
        null_date: None,
        iteration: None,
    })
}

fn parse_settings_child(
    settings: &mut Settings,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    is_table: bool,
    element: &BytesStart<'_>,
) -> Result<()> {
    if !is_table {
        return Err(Error::InvalidFormat(
            "unsupported calculation settings child".to_string(),
        ));
    }
    match element.local_name().as_ref() {
        b"null-date" => {
            if settings.null_date.is_some() || settings.iteration.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate or out-of-order table:null-date".to_string(),
                ));
            }
            let value_type = optional_attribute(resolver, decoder, element, b"value-type")?;
            if value_type.as_deref().is_some_and(|value| value != "date") {
                return Err(Error::InvalidFormat(
                    "table:null-date value type must be 'date'".to_string(),
                ));
            }
            settings.null_date = Some(NullDate {
                value_type_date: value_type.is_some(),
                date_value: optional_attribute(resolver, decoder, element, b"date-value")?,
            });
        },
        b"iteration" => {
            if settings.iteration.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate table:iteration".to_string(),
                ));
            }
            settings.iteration = Some(Iteration {
                status: optional_attribute(resolver, decoder, element, b"status")?
                    .map(|value| IterationStatus::parse(&value))
                    .transpose()?,
                steps: optional_positive(resolver, decoder, element, b"steps")?,
                maximum_difference: optional_attribute(
                    resolver,
                    decoder,
                    element,
                    b"maximum-difference",
                )?,
            });
        },
        _ => {
            return Err(Error::InvalidFormat(
                "unsupported calculation settings child".to_string(),
            ));
        },
    }
    Ok(())
}

fn optional_bool(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<bool>> {
    optional_attribute(resolver, decoder, element, name)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid Boolean value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_positive(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<NonZeroUsize>> {
    optional_attribute(resolver, decoder, element, name)?
        .map(|value| {
            value
                .parse::<usize>()
                .ok()
                .and_then(NonZeroUsize::new)
                .ok_or_else(|| Error::InvalidFormat(format!("invalid positive integer '{value}'")))
        })
        .transpose()
}

fn optional_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        if attribute.value.len() > super::MAX_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "calculation settings attribute exceeds the size limit".to_string(),
            ));
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace) && local.as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")));
        }
    }
    Ok(None)
}

fn is_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace_uri(namespace, TABLE_NAMESPACE)
}

fn is_namespace_uri(namespace: &ResolveResult<'_>, uri: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == uri)
}

fn write_optional_bool(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        write_attribute(out, name, if value { "true" } else { "false" });
    }
}

fn write_attribute(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    out.push_str(&escape_xml(value));
    out.push('"');
}
