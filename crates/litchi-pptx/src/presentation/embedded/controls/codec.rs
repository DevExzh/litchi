use super::model::Persistence;
use super::{ACTIVEX_NAMESPACE, MAX_CONTROLS, MAX_SLIDE_XML_BYTES};
use crate::Result;
use crate::presentation::embedded::{
    MAX_XML_ATTRIBUTES, MAX_XML_DEPTH, bounded, increment_nodes, invalid, is_presentationml_name,
    limit, relationship_value, validate_root,
};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::attributes::count_up_to;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

#[derive(Default)]
pub(crate) struct Parsed {
    pub(crate) shape_id: Option<String>,
    pub(crate) name: Option<String>,
    pub(crate) show_as_icon: Option<bool>,
    pub(crate) image_width: Option<u32>,
    pub(crate) image_height: Option<u32>,
    pub(crate) relationship_id: Option<String>,
}

pub(crate) fn scan(xml_bytes: &[u8], count: &mut usize) -> Result<Vec<Parsed>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml_bytes.len(),
            MAX_SLIDE_XML_BYTES,
        ));
    }
    let mce = Limits {
        max_input_bytes: MAX_SLIDE_XML_BYTES,
        max_output_bytes: MAX_SLIDE_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    };
    let xml = process_markup_compatibility(xml_bytes, &Capabilities::ooxml_baseline(), &mce)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut values = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut c_sld_depth = None;
    let mut controls_depth = None;
    let mut open_control_depth = None;
    let mut saw_container = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if depth == 1 {
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                } else if depth == 2 && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    c_sld_depth = Some(depth);
                } else if c_sld_depth == Some(depth - 1)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    if saw_container {
                        return Err(invalid("slide contains multiple control containers"));
                    }
                    saw_container = true;
                    controls_depth = Some(depth);
                } else if controls_depth == Some(depth - 1)
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    add_control(count)?;
                    values.push(parse_control(&element, decoder, &resolver)?);
                    open_control_depth = Some(depth);
                }
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        child_depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if child_depth == 1 {
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                    root_closed = true;
                } else if c_sld_depth == Some(child_depth - 1)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    if saw_container {
                        return Err(invalid("slide contains multiple control containers"));
                    }
                    saw_container = true;
                } else if controls_depth == Some(child_depth - 1)
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    add_control(count)?;
                    values.push(parse_control(&element, decoder, &resolver)?);
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid control XML nesting"));
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("control XML must close with p:sld"));
                    }
                    root_closed = true;
                }
                if open_control_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    open_control_depth = None;
                }
                if controls_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    controls_depth = None;
                }
                if c_sld_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    c_sld_depth = None;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "control XML rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 || open_control_depth.is_some() {
                    return Err(invalid("unterminated PresentationML control slide"));
                }
                return Ok(values);
            },
            _ => {},
        }
    }
}

fn add_control(count: &mut usize) -> Result<()> {
    *count = count.checked_add(1).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::Controls,
            usize::MAX,
            MAX_CONTROLS,
        )
    })?;
    if *count > MAX_CONTROLS {
        return Err(limit(
            crate::error::CodecLimitKind::Controls,
            *count,
            MAX_CONTROLS,
        ));
    }
    Ok(())
}

fn parse_control(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Parsed> {
    if count_up_to(element, MAX_XML_ATTRIBUTES) > MAX_XML_ATTRIBUTES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlAttributes,
            count_up_to(element, MAX_XML_ATTRIBUTES),
            MAX_XML_ATTRIBUTES,
        ));
    }
    let optional = |name: &[u8], _label: &'static str| -> Result<Option<String>> {
        let value = unqualified_attribute_value(element, name, decoder)?;
        if let Some(value) = &value {
            bounded(value, crate::error::CodecLimitKind::XmlAttributeBytes)?;
        }
        Ok(value)
    };
    let show_as_icon = optional(b"showAsIcon", "control show-as-icon")?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(invalid("invalid control show-as-icon flag")),
        })
        .transpose()?;
    let number = |name: &[u8], label: &'static str| -> Result<Option<u32>> {
        optional(name, label)?
            .map(|value| {
                value
                    .parse()
                    .map_err(|_err| invalid(format!("invalid {label}")))
            })
            .transpose()
    };
    Ok(Parsed {
        shape_id: optional(b"spid", "control shape ID")?,
        name: optional(b"name", "control name")?,
        show_as_icon,
        image_width: number(b"imgW", "control image width")?,
        image_height: number(b"imgH", "control image height")?,
        relationship_id: relationship_value(element, b"id", decoder, resolver)?
            .filter(|value| !value.is_empty()),
    })
}

pub(crate) fn parse_descriptor(
    xml: &[u8],
) -> Result<(String, Option<String>, Persistence, Option<String>)> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml.len(),
            MAX_SLIDE_XML_BYTES,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut root = false;
    let mut result = None;
    loop {
        let decoder = reader.decoder();
        let (_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                increment_nodes(&mut depth)?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if !root {
                    root = true;
                    if element.name().local_name().as_ref() != b"ocx" {
                        return Err(invalid("control descriptor must have an ax:ocx root"));
                    }
                    let class_id = attribute(
                        &element,
                        b"classid",
                        decoder,
                        reader.resolver(),
                        ACTIVEX_NAMESPACE,
                    )?
                    .ok_or_else(|| invalid("control descriptor is missing ax:classid"))?;
                    let license = attribute(
                        &element,
                        b"license",
                        decoder,
                        reader.resolver(),
                        ACTIVEX_NAMESPACE,
                    )?;
                    let persistence = match attribute(
                        &element,
                        b"persistence",
                        decoder,
                        reader.resolver(),
                        ACTIVEX_NAMESPACE,
                    )?
                    .as_deref()
                    {
                        Some("persistPropertyBag") => Persistence::PropertyBag,
                        Some("persistStream") => Persistence::Stream,
                        Some("persistStreamInit") => Persistence::StreamInit,
                        Some("persistStorage") => Persistence::Storage,
                        Some(_) => Persistence::Unknown,
                        None => Persistence::Unknown,
                    };
                    let relationship_id =
                        relationship_value(&element, b"id", decoder, reader.resolver())?;
                    result = Some((class_id, license, persistence, relationship_id));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                increment_nodes(&mut depth)?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if !root {
                    if element.name().local_name().as_ref() != b"ocx" {
                        return Err(invalid("control descriptor must have an ax:ocx root"));
                    }
                    let class_id = attribute(
                        &element,
                        b"classid",
                        decoder,
                        reader.resolver(),
                        ACTIVEX_NAMESPACE,
                    )?
                    .ok_or_else(|| invalid("control descriptor is missing ax:classid"))?;
                    let license = attribute(
                        &element,
                        b"license",
                        decoder,
                        reader.resolver(),
                        ACTIVEX_NAMESPACE,
                    )?;
                    let persistence = match attribute(
                        &element,
                        b"persistence",
                        decoder,
                        reader.resolver(),
                        ACTIVEX_NAMESPACE,
                    )?
                    .as_deref()
                    {
                        Some("persistPropertyBag") => Persistence::PropertyBag,
                        Some("persistStream") => Persistence::Stream,
                        Some("persistStreamInit") => Persistence::StreamInit,
                        Some("persistStorage") => Persistence::Storage,
                        Some(_) | None => Persistence::Unknown,
                    };
                    let relationship_id =
                        relationship_value(&element, b"id", decoder, reader.resolver())?;
                    result = Some((class_id, license, persistence, relationship_id));
                }
                break;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("invalid control descriptor nesting"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "control descriptor rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    result.ok_or_else(|| invalid("control descriptor is empty"))
}

fn attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    expected_namespace: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| crate::Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == *expected_namespace)
        {
            continue;
        }
        if value.is_some() {
            return Err(invalid(format!(
                "duplicate ActiveX attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(|error| crate::Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn slide(control_attributes: &str) -> Vec<u8> {
        format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:controls><p:control{control_attributes}/></p:controls></p:cSld></p:sld>"#
        )
        .into_bytes()
    }

    /// `count` distinct attributes.
    fn distinct(count: usize) -> String {
        (0..count)
            .map(|index| format!(" x{index}=\"{index}\""))
            .collect()
    }

    fn is_attribute_limit(result: &Result<Vec<Parsed>>) -> bool {
        matches!(
            result,
            Err(crate::Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
                observed,
                maximum: MAX_XML_ATTRIBUTES,
            }) if *observed == MAX_XML_ATTRIBUTES + 1
        )
    }

    #[test]
    fn control_attribute_limit_counts_every_attribute() {
        let at_limit = format!(" name=\"CheckBox\"{}", distinct(MAX_XML_ATTRIBUTES - 1));
        let parsed = scan(&slide(&at_limit), &mut 0).unwrap();
        assert_eq!(parsed[0].name.as_deref(), Some("CheckBox"));

        let over_limit = format!(" name=\"CheckBox\"{}", distinct(MAX_XML_ATTRIBUTES));
        assert!(is_attribute_limit(&scan(&slide(&over_limit), &mut 0)));
    }

    #[test]
    fn crowded_and_duplicate_control_attributes_are_refused_as_before() {
        // 20,000 distinct names followed by 20,000 repeats of the last one.
        let mut crowded = distinct(20_000);
        for _ in 0..20_000 {
            crowded.push_str(" x19999=\"repeat\"");
        }
        assert!(is_attribute_limit(&scan(&slide(&crowded), &mut 0)));

        // Under the limit, a duplicate is refused where the attributes are read.
        assert!(matches!(
            scan(&slide(r#" name="A" name="B""#), &mut 0),
            Err(crate::Error::Decode(
                litchi_ooxml_common::XmlError::Malformed(_)
            ))
        ));
    }
}

#[cfg(test)]
mod envelope_tests {
    //! What the control scan accepts before, around and after `p:sld`, and the exact error reported otherwise. Each expected
    //! line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&[u8]) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(xml.as_bytes())))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The control scan of one slide.
    #[test]
    fn control_scan_envelope() {
        let root = Root {
            open: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            body: r"<p:cSld><p:spTree/></p:cSld>",
            close: r"</p:sld>",
            empty: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"unterminated PresentationML control slide\")",
                "whitespace only: Invalid(\"unterminated PresentationML control slide\")",
                "declaration only: Invalid(\"unterminated PresentationML control slide\")",
                "empty-element root: Ok(0)",
                "start and end root: Ok(0)",
                "root then whitespace: Ok(0)",
                "root then comment: Ok(0)",
                "root then PI: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "root then text: Ok(0)",
                "root then second root: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "root then foreign start: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "root then foreign empty: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "declaration after root: Ok(0)",
                "two declarations: Ok(0)",
                "comment before declaration: Ok(0)",
                "text before root: Ok(0)",
                "unterminated root: Invalid(\"unterminated PresentationML control slide\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "top-level CDATA: Ok(0)",
                "declaration then root: Ok(0)",
                "whitespace before declaration: Ok(0)",
                "PI before declaration: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "comment then root: Ok(0)",
                "PI then root: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "root then text then second root: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "empty-element root then root: Invalid(\"embedded-resource XML must have one PresentationML sld root\")",
                "declaration after empty-element root: Ok(0)",
            ],
            |xml| outcome(super::scan(xml, &mut 0), |values| values.len().to_string()),
        );
    }
}
