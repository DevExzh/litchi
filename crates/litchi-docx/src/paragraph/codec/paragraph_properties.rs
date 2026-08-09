#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Typed paragraph-property decoding.

use crate::error::{Error, Result};
use crate::namespace::{direct_word_property_value, normalize_xml_integer};
use quick_xml::encoding::Decoder;
use quick_xml::events::Event;
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::model::{LineSpacingRule, Paragraph, ParagraphSpacing};
use super::xml::{
    element_prefix, is_fragment_word_name, paragraph_attribute, same_word_prefix,
    same_word_prefix_end, set_paragraph_property, word_attribute_value,
};

impl Paragraph {
    /// Return direct paragraph numbering properties, including `numId=0` cancellation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn numbering(&self) -> Result<Option<crate::numbering::Paragraph>> {
        Ok(self.list_properties()?.0)
    }

    /// Return the paragraph style identifier from `<w:pPr>`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn style_id(&self) -> Result<Option<String>> {
        Ok(self.list_properties()?.1)
    }

    fn list_properties(&self) -> Result<(Option<crate::numbering::Paragraph>, Option<String>)> {
        let mut reader = NsReader::from_reader(self.xml_bytes());
        let mut depth = 0usize;
        let mut word_prefix: Option<Vec<u8>> = None;
        let mut ppr_depth = None;
        let mut numpr_depth = None;
        let mut saw_numpr = false;
        let mut num_id = None;
        let mut level = None;
        let mut style_id = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                    })?;
                    if depth == 1 && element.local_name().as_ref() == b"p" {
                        word_prefix = Some(element_prefix(&element));
                    }
                    if !same_word_prefix(&element, word_prefix.as_deref()) {
                        continue;
                    }
                    match element.local_name().as_ref() {
                        b"pPr" if depth == 2 => ppr_depth = Some(depth),
                        b"numPr" if ppr_depth.is_some_and(|value| depth == value + 1) => {
                            if saw_numpr {
                                return Err(Error::InvalidFormat(
                                    "paragraph has duplicate numPr".to_owned(),
                                ));
                            }
                            saw_numpr = true;
                            numpr_depth = Some(depth);
                        },
                        b"pStyle" if ppr_depth.is_some_and(|value| depth == value + 1) => {
                            set_paragraph_property(
                                &mut style_id,
                                paragraph_attribute(&element, b"val", decoder)?,
                                "pStyle",
                            )?;
                        },
                        b"numId" if numpr_depth.is_some_and(|value| depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw.parse::<u32>().map_err(|_source_error| {
                                Error::InvalidFormat(format!("invalid paragraph numId '{raw}'"))
                            })?;
                            set_paragraph_property(&mut num_id, parsed, "numId")?;
                        },
                        b"ilvl" if numpr_depth.is_some_and(|value| depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw
                                .parse::<u8>()
                                .ok()
                                .filter(|value| *value <= 8)
                                .ok_or_else(|| {
                                    Error::InvalidFormat(format!("invalid paragraph ilvl '{raw}'"))
                                })?;
                            set_paragraph_property(&mut level, parsed, "ilvl")?;
                        },
                        _ => {},
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                    })?;
                    if !same_word_prefix(&element, word_prefix.as_deref()) {
                        continue;
                    }
                    match element.local_name().as_ref() {
                        b"pStyle" if ppr_depth.is_some_and(|value| child_depth == value + 1) => {
                            set_paragraph_property(
                                &mut style_id,
                                paragraph_attribute(&element, b"val", decoder)?,
                                "pStyle",
                            )?;
                        },
                        b"numPr" if ppr_depth.is_some_and(|value| child_depth == value + 1) => {
                            return Err(Error::InvalidFormat(
                                "paragraph numPr is missing numId".to_owned(),
                            ));
                        },
                        b"numId" if numpr_depth.is_some_and(|value| child_depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw.parse::<u32>().map_err(|_source_error| {
                                Error::InvalidFormat(format!("invalid paragraph numId '{raw}'"))
                            })?;
                            set_paragraph_property(&mut num_id, parsed, "numId")?;
                        },
                        b"ilvl" if numpr_depth.is_some_and(|value| child_depth == value + 1) => {
                            let raw = paragraph_attribute(&element, b"val", decoder)?;
                            let parsed = raw
                                .parse::<u8>()
                                .ok()
                                .filter(|value| *value <= 8)
                                .ok_or_else(|| {
                                    Error::InvalidFormat(format!("invalid paragraph ilvl '{raw}'"))
                                })?;
                            set_paragraph_property(&mut level, parsed, "ilvl")?;
                        },
                        _ => {},
                    }
                },
                Event::End(element) => {
                    if same_word_prefix_end(&element, word_prefix.as_deref()) {
                        match element.local_name().as_ref() {
                            b"numPr" if numpr_depth == Some(depth) => numpr_depth = None,
                            b"pPr" if ppr_depth == Some(depth) => ppr_depth = None,
                            _ => {},
                        }
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid paragraph XML nesting".to_owned())
                    })?;
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
        let numbering = if saw_numpr {
            Some(crate::numbering::Paragraph {
                num_id: num_id.ok_or_else(|| {
                    Error::InvalidFormat("paragraph numPr is missing numId".to_owned())
                })?,
                level: level.unwrap_or(0),
            })
        } else {
            None
        };
        Ok((numbering, style_id))
    }
}

impl Paragraph {
    /// Return the HTML division ID referenced by this paragraph, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn division_id(&self) -> Result<Option<String>> {
        direct_word_property_value(self.xml_bytes(), b"p", b"pPr", b"divId")?
            .map(|value| normalize_xml_integer(value, "Word paragraph division ID"))
            .transpose()
    }

    /// Return the direct spacing properties of this paragraph.
    ///
    /// The returned values are validated against the `WordprocessingML` types
    /// used by `CT_Spacing`: before/after are non-negative twips, line is a
    /// signed value, and lineRule is one of the schema-defined tokens.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn spacing(&self) -> Result<Option<ParagraphSpacing>> {
        parse_spacing(self.xml_bytes())
    }
}

fn parse_spacing(xml_bytes: &[u8]) -> Result<Option<ParagraphSpacing>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut ppr_depth = None;
    let mut saw_ppr = false;
    let mut spacing = None;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                })?;

                if depth == 1 {
                    if !matches!(namespace, ResolveResult::Bound(_)) {
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                    }
                    if saw_root
                        || element.local_name().as_ref() != b"p"
                        || !is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"p",
                            &fragment_prefix,
                        )
                    {
                        return Err(Error::InvalidFormat(
                            "paragraph spacing XML has an invalid root".to_owned(),
                        ));
                    }
                    saw_root = true;
                }

                let is_word = is_fragment_word_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                    &fragment_prefix,
                );
                if depth == 2 && is_word && element.local_name().as_ref() == b"pPr" {
                    if saw_ppr {
                        return Err(Error::InvalidFormat(
                            "paragraph has duplicate pPr".to_owned(),
                        ));
                    }
                    saw_ppr = true;
                    ppr_depth = Some(depth);
                } else if depth == 3
                    && ppr_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == b"spacing"
                {
                    if spacing.is_some() {
                        return Err(Error::InvalidFormat(
                            "paragraph has duplicate spacing".to_owned(),
                        ));
                    }
                    spacing = Some(parse_spacing_element(
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                    )?);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph XML nesting is too deep".to_owned())
                })?;

                if child_depth == 1 {
                    if !matches!(namespace, ResolveResult::Bound(_)) {
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                    }
                    if saw_root
                        || element.local_name().as_ref() != b"p"
                        || !is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"p",
                            &fragment_prefix,
                        )
                    {
                        return Err(Error::InvalidFormat(
                            "paragraph spacing XML has an invalid root".to_owned(),
                        ));
                    }
                    saw_root = true;
                } else {
                    let is_word = is_fragment_word_name(
                        &namespace,
                        element.name(),
                        element.local_name().as_ref(),
                        &fragment_prefix,
                    );
                    if child_depth == 2 && is_word && element.local_name().as_ref() == b"pPr" {
                        if saw_ppr {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate pPr".to_owned(),
                            ));
                        }
                        saw_ppr = true;
                    } else if child_depth == 3
                        && ppr_depth == Some(2)
                        && is_word
                        && element.local_name().as_ref() == b"spacing"
                    {
                        if spacing.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate spacing".to_owned(),
                            ));
                        }
                        spacing = Some(parse_spacing_element(
                            &element,
                            decoder,
                            &resolver,
                            &fragment_prefix,
                        )?);
                    }
                }
            },
            Event::End(_) => {
                if ppr_depth == Some(depth) {
                    ppr_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid paragraph XML nesting".to_owned())
                })?;
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat(
                    "unterminated paragraph spacing XML".to_owned(),
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

    if !saw_root {
        return Err(Error::InvalidFormat(
            "paragraph spacing XML has no root".to_owned(),
        ));
    }
    Ok(spacing)
}

fn parse_spacing_element(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<ParagraphSpacing> {
    let before = parse_u64(
        word_attribute_value(element, b"before", decoder, resolver, fragment_prefix)?,
        "paragraph spacing before",
    )?;
    let before_lines = parse_i32(
        word_attribute_value(element, b"beforeLines", decoder, resolver, fragment_prefix)?,
        "paragraph spacing beforeLines",
    )?;
    let before_auto_spacing = parse_on_off(
        word_attribute_value(
            element,
            b"beforeAutospacing",
            decoder,
            resolver,
            fragment_prefix,
        )?,
        "paragraph spacing beforeAutospacing",
    )?;
    let after = parse_u64(
        word_attribute_value(element, b"after", decoder, resolver, fragment_prefix)?,
        "paragraph spacing after",
    )?;
    let after_lines = parse_i32(
        word_attribute_value(element, b"afterLines", decoder, resolver, fragment_prefix)?,
        "paragraph spacing afterLines",
    )?;
    let after_auto_spacing = parse_on_off(
        word_attribute_value(
            element,
            b"afterAutospacing",
            decoder,
            resolver,
            fragment_prefix,
        )?,
        "paragraph spacing afterAutospacing",
    )?;
    let line = parse_i32(
        word_attribute_value(element, b"line", decoder, resolver, fragment_prefix)?,
        "paragraph spacing line",
    )?;
    let line_rule = word_attribute_value(element, b"lineRule", decoder, resolver, fragment_prefix)?
        .map(|value| {
            LineSpacingRule::from_xml(&value).ok_or_else(|| {
                Error::InvalidFormat(format!("invalid paragraph spacing lineRule '{value}'"))
            })
        })
        .transpose()?;

    Ok(ParagraphSpacing {
        before,
        before_lines,
        before_auto_spacing,
        after,
        after_lines,
        after_auto_spacing,
        line,
        line_rule,
    })
}

fn parse_u64(value: Option<String>, name: &str) -> Result<Option<u64>> {
    value
        .map(|value| {
            value.parse::<u64>().map_err(|_source_error| {
                Error::InvalidFormat(format!("invalid {name} value '{value}'"))
            })
        })
        .transpose()
}

fn parse_i32(value: Option<String>, name: &str) -> Result<Option<i32>> {
    value
        .map(|value| {
            value.parse::<i32>().map_err(|_source_error| {
                Error::InvalidFormat(format!("invalid {name} value '{value}'"))
            })
        })
        .transpose()
}

fn parse_on_off(value: Option<String>, name: &str) -> Result<Option<bool>> {
    value
        .map(|value| match value.as_str() {
            "true" | "1" | "on" => Ok(true),
            "false" | "0" | "off" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid {name} value '{value}'"
            ))),
        })
        .transpose()
}
