//! Typed run-property decoding.

use crate::UnderlineStyle;
use crate::color::Theme;
use crate::error::{Error, Result};
use litchi_core::VerticalPosition;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{XmlVersion, encoding::Decoder};

use super::super::model::{RunProperties, RunUnderline, RunUnderlineColor};
use super::xml::is_fragment_word_name;

pub(super) fn update_run_properties(
    props: &mut RunProperties,
    element: &BytesStart<'_>,
) -> Result<()> {
    let property = element.local_name();
    if !matches!(
        property.as_ref(),
        b"b" | b"i" | b"strike" | b"u" | b"vertAlign"
    ) {
        return Ok(());
    }

    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() == b"val" {
            value = Some(attribute.value);
            break;
        }
    }

    match property.as_ref() {
        b"b" => props.bold = Some(value.as_deref().is_none_or(is_on)),
        b"i" => props.italic = Some(value.as_deref().is_none_or(is_on)),
        b"strike" => props.strikethrough = Some(value.as_deref().is_none_or(is_on)),
        b"u" => {
            props.underline = Some(match value.as_deref() {
                None => UnderlineStyle::Single,
                Some(value) => UnderlineStyle::from_xml(
                    std::str::from_utf8(value)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?,
                )
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "invalid Word underline style '{}'",
                        String::from_utf8_lossy(value)
                    ))
                })?,
            });
        },
        b"vertAlign" => {
            props.vertical_position = match value.as_deref() {
                Some(b"superscript") => Some(VerticalPosition::Superscript),
                Some(b"subscript") => Some(VerticalPosition::Subscript),
                _ => None,
            };
        },
        _ => {},
    }
    Ok(())
}

pub(super) fn parse_run_underline(xml_bytes: &[u8]) -> Result<Option<RunUnderline>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut properties_depth = None;
    let mut saw_root = false;
    let mut saw_properties = false;
    let mut underline = None;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        if fragment_prefix.is_none()
            && depth == 0
            && let Event::Start(element) | Event::Empty(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word run XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                    &fragment_prefix,
                );
                if depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != b"r" {
                        return Err(Error::InvalidFormat(
                            "Word underline XML has an invalid run root".into(),
                        ));
                    }
                    saw_root = true;
                } else if depth == 2 && is_word && element.local_name().as_ref() == b"rPr" {
                    if saw_properties {
                        return Err(Error::InvalidFormat(
                            "duplicate Word run property container".into(),
                        ));
                    }
                    saw_properties = true;
                    properties_depth = Some(depth);
                } else if depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == b"u"
                {
                    set_run_underline(
                        &mut underline,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                    )?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word run XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                    &fragment_prefix,
                );
                if child_depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != b"r" {
                        return Err(Error::InvalidFormat(
                            "Word underline XML has an invalid run root".into(),
                        ));
                    }
                    saw_root = true;
                } else if child_depth == 2 && is_word && element.local_name().as_ref() == b"rPr" {
                    if saw_properties {
                        return Err(Error::InvalidFormat(
                            "duplicate Word run property container".into(),
                        ));
                    }
                    saw_properties = true;
                } else if child_depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == b"u"
                {
                    set_run_underline(
                        &mut underline,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                    )?;
                }
            },
            Event::End(_) => {
                if properties_depth == Some(depth) {
                    properties_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word run XML nesting".into()))?;
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat("unterminated Word run XML".into()));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !saw_root {
        return Err(Error::InvalidFormat(
            "Word underline XML has no run root".into(),
        ));
    }
    Ok(underline)
}

fn set_run_underline(
    slot: &mut Option<RunUnderline>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<()> {
    if slot.is_some() {
        return Err(Error::InvalidFormat(
            "duplicate Word underline property".into(),
        ));
    }
    let style = run_underline_attribute(element, b"val", decoder, resolver, fragment_prefix)?
        .map(|value| {
            UnderlineStyle::from_xml(&value).ok_or_else(|| {
                Error::InvalidFormat(format!("invalid Word underline style '{value}'"))
            })
        })
        .transpose()?
        .unwrap_or(UnderlineStyle::Single);
    let color = run_underline_attribute(element, b"color", decoder, resolver, fragment_prefix)?
        .map(|value| parse_run_underline_color(&value))
        .transpose()?;
    let theme_color =
        run_underline_attribute(element, b"themeColor", decoder, resolver, fragment_prefix)?
            .map(|value| {
                Theme::parse(&value).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid Word underline theme color '{value}'"))
                })
            })
            .transpose()?;
    let theme_tint =
        run_underline_attribute(element, b"themeTint", decoder, resolver, fragment_prefix)?
            .map(|value| parse_run_underline_hex_byte(&value, "theme tint"))
            .transpose()?;
    let theme_shade =
        run_underline_attribute(element, b"themeShade", decoder, resolver, fragment_prefix)?
            .map(|value| parse_run_underline_hex_byte(&value, "theme shade"))
            .transpose()?;

    *slot = Some(RunUnderline {
        style,
        color,
        theme_color,
        theme_tint,
        theme_shade,
    });
    Ok(())
}

fn run_underline_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_fragment_word_name(&namespace, attribute.key, name, fragment_prefix) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate Word underline attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn parse_run_underline_color(value: &str) -> Result<RunUnderlineColor> {
    if value == "auto" {
        return Ok(RunUnderlineColor::Auto);
    }
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::InvalidFormat(format!(
            "invalid Word underline color '{value}'"
        )));
    }
    let mut rgb = [0u8; 3];
    for (index, component) in rgb.iter_mut().enumerate() {
        *component = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16)
            .map_err(|_| Error::InvalidFormat(format!("invalid Word underline color '{value}'")))?;
    }
    Ok(RunUnderlineColor::Rgb(rgb))
}

fn parse_run_underline_hex_byte(value: &str, description: &str) -> Result<u8> {
    if value.len() != 2 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::InvalidFormat(format!(
            "invalid Word underline {description} '{value}'"
        )));
    }
    u8::from_str_radix(value, 16).map_err(|_| {
        Error::InvalidFormat(format!("invalid Word underline {description} '{value}'"))
    })
}

#[inline]
fn is_on(value: &[u8]) -> bool {
    matches!(value, b"true" | b"1" | b"on")
}
