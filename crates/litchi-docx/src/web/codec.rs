//! Bounded WordprocessingML web-settings XML codec.

use super::model::{
    Border, Borders, Child, Color, Conformance, Div, Frame, Frameset, Id, Key, Layout, Screen,
    Scrollbar, Settings, SplitBar, Twips,
};
use super::{
    CONTENT_TYPE, MAX_FRAMESET_NESTING, MAX_TEXT_BYTES, MAX_XML_BYTES, MAX_XML_EVENTS, ParseBudget,
    invalid, is_wordprocessing_namespace, reserve_one, word_attribute_value,
};
use crate::color::Theme;
use crate::{Error, Result};
use litchi_opc::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::escape::escape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as _;

impl Settings {
    fn encode(&self, conformance: Conformance) -> Result<Vec<u8>> {
        if conformance == Conformance::Strict && self.rely_on_vml.is_some() {
            return Err(invalid("relyOnVML is not valid in Strict web settings"));
        }
        let capacity = validate_value(self)?;
        let mut xml = String::new();
        xml.try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "web-settings XML",
                source,
            })?;
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str("<w:webSettings xmlns:w=\"");
        xml.push_str(conformance.wordprocessingml());
        xml.push_str("\" xmlns:r=\"");
        xml.push_str(conformance.relationships());
        xml.push_str("\">");

        if let Some(frameset) = &self.frameset {
            write_frameset(&mut xml, frameset, 1)?;
        }
        if let Some(divs) = &self.divs {
            xml.push_str("<w:divs>");
            for div in divs {
                write_html_div(&mut xml, div, 1)?;
            }
            xml.push_str("</w:divs>");
        }
        if let Some(value) = &self.encoding {
            write_value_element(&mut xml, "encoding", value)?;
        }
        write_optional_on_off(&mut xml, "optimizeForBrowser", self.optimize_for_browser)?;
        if conformance == Conformance::Transitional {
            write_optional_on_off(&mut xml, "relyOnVML", self.rely_on_vml)?;
        }
        write_optional_on_off(&mut xml, "allowPNG", self.allow_png)?;
        write_optional_on_off(&mut xml, "doNotRelyOnCSS", self.do_not_rely_on_css)?;
        write_optional_on_off(
            &mut xml,
            "doNotSaveAsSingleFile",
            self.do_not_save_as_single_file,
        )?;
        write_optional_on_off(
            &mut xml,
            "doNotOrganizeInFolder",
            self.do_not_organize_in_folder,
        )?;
        write_optional_on_off(
            &mut xml,
            "doNotUseLongFileNames",
            self.do_not_use_long_file_names,
        )?;
        if let Some(value) = self.pixels_per_inch {
            write!(xml, "<w:pixelsPerInch w:val=\"{value}\"/>")?;
        }
        if let Some(value) = self.target_screen_size {
            write_value_element(&mut xml, "targetScreenSz", value.as_str())?;
        }
        write_optional_on_off(&mut xml, "saveSmartTagsAsXml", self.save_smart_tags_as_xml)?;

        xml.push_str("</w:webSettings>");
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        Ok(xml.into_bytes())
    }

    fn read_part(part: &dyn Part) -> Result<(Self, Conformance)> {
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        if part.blob().len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        let xml = process_web_xml(part.blob())?;
        let (settings, conformance) = Self::parse_xml(xml.as_ref())?;
        validate_frame_relationships(part, &settings, conformance)?;
        Ok((settings, conformance))
    }

    pub(super) fn parse_xml(xml: &[u8]) -> Result<(Self, Conformance)> {
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        let mut reader = NsReader::from_reader(xml);
        let mut settings = Self::default();
        let mut depth = 0usize;
        let mut saw_root = false;
        let mut conformance = None;
        let mut last_child_rank = None;
        let mut budget = ParseBudget::default();

        loop {
            budget.event()?;
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
                        Error::Invalid("Word web-settings XML nesting is too deep".into())
                    })?;
                    if depth > MAX_FRAMESET_NESTING {
                        return Err(invalid(
                            "Word web-settings XML nesting exceeds the safety limit",
                        ));
                    }
                    if depth == 1 {
                        conformance = Some(validate_root(&namespace, &element, saw_root)?);
                        saw_root = true;
                    } else if depth == 2 && saw_root && is_wordprocessing_namespace(&namespace) {
                        let profile = conformance.ok_or_else(|| {
                            invalid("web-settings root conformance was not resolved")
                        })?;
                        validate_web_child(&namespace, &element, profile, &mut last_child_rank)?;
                        if element.local_name().as_ref() == b"frameset" {
                            let frameset = parse_frameset(&mut reader, 1, &mut budget)?;
                            set_once(&mut settings.frameset, frameset, "frameset")?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                Error::Invalid("invalid Word web-settings XML nesting".into())
                            })?;
                        } else if element.local_name().as_ref() == b"divs" {
                            let divs = parse_div_container(&mut reader, b"divs", 1, &mut budget)?;
                            set_once(&mut settings.divs, divs, "divs")?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                Error::Invalid("invalid Word web-settings XML nesting".into())
                            })?;
                        } else if is_scalar_setting(element.local_name().as_ref()) {
                            parse_setting(&element, decoder, &resolver, &mut settings)?;
                            finish_leaf(
                                &mut reader,
                                element.local_name().as_ref(),
                                "web setting",
                                &mut budget,
                            )?;
                            depth = depth.checked_sub(1).ok_or_else(|| {
                                Error::Invalid("invalid Word web-settings XML nesting".into())
                            })?;
                        }
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("Word web-settings XML nesting is too deep".into())
                    })?;
                    if child_depth > MAX_FRAMESET_NESTING {
                        return Err(invalid(
                            "Word web-settings XML nesting exceeds the safety limit",
                        ));
                    }
                    if child_depth == 1 {
                        conformance = Some(validate_root(&namespace, &element, saw_root)?);
                        saw_root = true;
                    } else if child_depth == 2
                        && saw_root
                        && is_wordprocessing_namespace(&namespace)
                    {
                        let profile = conformance.ok_or_else(|| {
                            invalid("web-settings root conformance was not resolved")
                        })?;
                        validate_web_child(&namespace, &element, profile, &mut last_child_rank)?;
                        if element.local_name().as_ref() == b"frameset" {
                            set_once(&mut settings.frameset, Frameset::default(), "frameset")?;
                        } else if element.local_name().as_ref() == b"divs" {
                            return Err(invalid("Word HTML division container must not be empty"));
                        } else if is_scalar_setting(element.local_name().as_ref()) {
                            parse_setting(&element, decoder, &resolver, &mut settings)?;
                        }
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid("invalid Word web-settings XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(Error::Invalid("unterminated Word web-settings XML".into()));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(invalid("web-settings part has no webSettings root"));
        }
        let conformance =
            conformance.ok_or_else(|| invalid("web-settings root conformance was not resolved"))?;
        validate_value(&settings)?;
        Ok((settings, conformance))
    }
}

/// Parse bounded web-settings XML without resolving frame relationships.
pub fn parse(xml: &[u8]) -> Result<(Settings, Conformance)> {
    let processed = process_web_xml(xml)?;
    Settings::parse_xml(processed.as_ref())
}

fn process_web_xml(xml: &[u8]) -> Result<std::borrow::Cow<'_, [u8]>> {
    let limits = litchi_ooxml_common::MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        ..litchi_ooxml_common::MceLimits::default()
    };
    litchi_ooxml_common::process_markup_compatibility(
        xml,
        &litchi_ooxml_common::MceCapabilities::default(),
        &limits,
    )
    .map(|output| output.xml)
    .map_err(Error::from)
}

/// Serialize a checked web-settings model.
pub fn write(value: &Settings, conformance: Conformance) -> Result<Vec<u8>> {
    value.encode(conformance)
}

/// Read one bounded web-settings part and validate its frame relationships.
pub fn read(part: &dyn Part) -> Result<(Settings, Conformance)> {
    Settings::read_part(part)
}

#[derive(Default)]
struct Budget {
    nodes: usize,
    encoded_bytes: usize,
}

impl Budget {
    fn node(&mut self, description: &str) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("web-settings node count overflow"))?;
        if self.nodes > MAX_XML_EVENTS / 2 {
            return Err(invalid(format!(
                "web-settings {description} exceeds the node limit"
            )));
        }
        self.bytes(256, description)
    }

    fn text(&mut self, value: &str, description: &str) -> Result<()> {
        validate_text(value, description, true)?;
        let escaped = value
            .len()
            .checked_mul(6)
            .ok_or_else(|| invalid("web-settings escaped text size overflow"))?;
        self.bytes(escaped, description)
    }

    fn bytes(&mut self, bytes: usize, description: &str) -> Result<()> {
        self.encoded_bytes = self
            .encoded_bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid("web-settings output size overflow"))?;
        if self.encoded_bytes > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings {description} exceeds {MAX_XML_BYTES} output bytes"
            )));
        }
        Ok(())
    }
}

fn validate_value(value: &Settings) -> Result<usize> {
    let mut budget = Budget {
        nodes: 1,
        encoded_bytes: 512,
    };
    if let Some(encoding) = &value.encoding {
        validate_encoding(encoding)?;
        budget.node("encoding")?;
        budget.text(encoding, "encoding")?;
    }
    if let Some(pixels) = value.pixels_per_inch {
        validate_pixels_per_inch(pixels)?;
        budget.node("pixels-per-inch")?;
    }
    if let Some(frameset) = &value.frameset {
        validate_frameset(frameset, 1, &mut budget)?;
    }
    if let Some(divs) = &value.divs {
        if divs.is_empty() {
            return Err(invalid("Word HTML division container must not be empty"));
        }
        budget.node("division container")?;
        validate_div_slice(divs, 1, &mut budget)?;
    }
    for present in [
        value.optimize_for_browser,
        value.rely_on_vml,
        value.allow_png,
        value.do_not_rely_on_css,
        value.do_not_save_as_single_file,
        value.do_not_organize_in_folder,
        value.do_not_use_long_file_names,
        value.target_screen_size.map(|_| true),
        value.save_smart_tags_as_xml,
    ] {
        if present.is_some() {
            budget.node("scalar setting")?;
        }
    }
    Ok(budget.encoded_bytes)
}

fn validate_frameset(value: &Frameset, depth: usize, budget: &mut Budget) -> Result<()> {
    if depth > MAX_FRAMESET_NESTING {
        return Err(invalid("web frameset nesting exceeds the safety limit"));
    }
    budget.node("frameset")?;
    if let Some(size) = &value.size {
        budget.node("frameset size")?;
        budget.text(size, "frameset size")?;
    }
    if let Some(split) = &value.split_bar {
        budget.node("frameset split bar")?;
        if let Some(color) = &split.color {
            validate_color(&color.value, "frameset splitter color")?;
            budget.node("frameset splitter color")?;
            budget.text(&color.value, "frameset splitter color")?;
        }
    }
    if value.layout.is_some() {
        budget.node("frameset layout")?;
    }
    for child in &value.children {
        match child {
            Child::Frameset(nested) => validate_frameset(nested, depth + 1, budget)?,
            Child::Frame(frame) => {
                budget.node("frame")?;
                for (text, description) in [
                    (frame.size.as_deref(), "frame size"),
                    (frame.name.as_deref(), "frame name"),
                    (
                        frame.source_file_relationship_id.as_deref(),
                        "frame relationship ID",
                    ),
                ] {
                    if let Some(text) = text {
                        if description == "frame relationship ID" {
                            validate_relationship_id(text)?;
                        }
                        budget.node(description)?;
                        budget.text(text, description)?;
                    }
                }
            },
        }
    }
    Ok(())
}

pub(super) fn validate_divs(divs: &[Div], depth: usize) -> Result<()> {
    if divs.is_empty() {
        return Err(invalid("Word HTML division container must not be empty"));
    }
    let mut budget = Budget::default();
    validate_div_slice(divs, depth, &mut budget)
}

fn validate_div_slice(divs: &[Div], depth: usize, budget: &mut Budget) -> Result<()> {
    if depth > MAX_FRAMESET_NESTING {
        return Err(invalid("HTML division nesting exceeds the safety limit"));
    }
    let mut ids = std::collections::HashSet::new();
    ids.try_reserve(divs.len())
        .map_err(|source| Error::Allocation {
            resource: "HTML division identifier index",
            source,
        })?;
    for div in divs {
        if !ids.insert(div.id) {
            return Err(invalid(format!("HTML division '{}' is ambiguous", div.id)));
        }
        budget.node("HTML division")?;
        budget.bytes(decimal_len(div.id.get()), "HTML division ID")?;
        for margin in [div.left, div.right, div.top, div.bottom] {
            budget.node("HTML division margin")?;
            budget.bytes(decimal_len(margin.get()), "HTML division margin")?;
        }
        if let Some(borders) = &div.borders {
            budget.node("HTML division borders")?;
            for border in [&borders.top, &borders.left, &borders.bottom, &borders.right]
                .into_iter()
                .flatten()
            {
                validate_border_style(&border.style)?;
                budget.node("HTML division border")?;
                budget.text(&border.style, "HTML division border style")?;
                if let Some(color) = &border.color {
                    validate_color(color, "HTML division border color")?;
                    budget.text(color, "HTML division border color")?;
                }
            }
        }
        if !div.children.is_empty() {
            validate_div_slice(&div.children, depth + 1, budget)?;
        }
    }
    Ok(())
}

pub(super) fn div_position(divs: &[Div], key: Key) -> Result<Option<usize>> {
    match key {
        Key::Index(index) => {
            if index >= divs.len() {
                Err(invalid(format!(
                    "HTML division position {index} is outside 0..{}",
                    divs.len()
                )))
            } else {
                Ok(Some(index))
            }
        },
        Key::Id(id) => {
            let mut matches = divs
                .iter()
                .enumerate()
                .filter(|(_, div)| div.id == id)
                .map(|(index, _)| index);
            let first = matches.next();
            if first.is_some() && matches.next().is_some() {
                Err(invalid(format!("HTML division ID '{id}' is ambiguous")))
            } else {
                Ok(first)
            }
        },
    }
}

pub(super) fn validate_text(value: &str, description: &str, allow_empty: bool) -> Result<()> {
    if (!allow_empty && value.is_empty()) || value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("invalid {description} length")));
    }
    if value
        .chars()
        .any(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    {
        return Err(invalid(format!(
            "{description} contains a control character"
        )));
    }
    Ok(())
}

pub(super) fn validate_encoding(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 255
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.'))
    {
        return Err(invalid("web encoding is not a bounded character-set name"));
    }
    Ok(())
}

pub(super) fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 255
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.' | b':'))
    {
        return Err(invalid(
            "frame relationship ID is not a safe XML identifier",
        ));
    }
    Ok(())
}

pub(super) fn validate_pixels_per_inch(value: u16) -> Result<()> {
    if value <= 1023 {
        Ok(())
    } else {
        Err(invalid("pixels-per-inch must be in the range 0..=1023"))
    }
}

pub(super) fn parse_i64(value: &str, description: &str) -> Result<i64> {
    let value = value.trim();
    if value.is_empty() || value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!("invalid {description} length")));
    }
    value
        .parse::<i64>()
        .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
}

fn decimal_len(value: i64) -> usize {
    let magnitude = value.unsigned_abs();
    let digits = if magnitude == 0 {
        1
    } else {
        magnitude.ilog10() as usize + 1
    };
    digits + usize::from(value.is_negative())
}

pub(super) fn validate_border_style(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 64
        || !value.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(invalid("HTML division border style is not a schema token"));
    }
    Ok(())
}

fn validate_color(value: &str, description: &str) -> Result<()> {
    if value.eq_ignore_ascii_case("auto")
        || (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "invalid {description} '{value}'; expected auto or six hexadecimal digits"
        )))
    }
}

fn write_value_element(xml: &mut String, name: &str, value: &str) -> Result<()> {
    write!(xml, "<w:{name} w:val=\"{}\"/>", escape(value))
        .map_err(|error| Error::Xml(error.to_string()))
}

fn write_optional_on_off(xml: &mut String, name: &str, value: Option<bool>) -> Result<()> {
    match value {
        Some(true) => {
            write!(xml, "<w:{name}/>")?;
        },
        Some(false) => {
            write!(xml, "<w:{name} w:val=\"false\"/>")?;
        },
        None => {},
    }
    Ok(())
}

/// Write the explicit numeric form required by desktop Word for `CT_Div`
/// role markers. Word rejects the otherwise schema-valid empty true form for
/// both `blockQuote` and `bodyDiv`.
fn write_explicit_on_off(xml: &mut String, name: &str, value: Option<bool>) -> Result<()> {
    if let Some(value) = value {
        write!(xml, "<w:{name} w:val=\"{}\"/>", u8::from(value))?;
    }
    Ok(())
}

fn write_frameset(xml: &mut String, frameset: &Frameset, nesting: usize) -> Result<()> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word web frameset nesting exceeds the supported safety limit".into(),
        ));
    }
    xml.push_str("<w:frameset>");
    if let Some(value) = &frameset.size {
        write_value_element(xml, "sz", value)?;
    }
    if let Some(split_bar) = &frameset.split_bar {
        write_frameset_split_bar(xml, split_bar)?;
    }
    if let Some(layout) = frameset.layout {
        write_value_element(xml, "frameLayout", layout.as_str())?;
    }
    for child in &frameset.children {
        match child {
            Child::Frameset(nested) => write_frameset(xml, nested, nesting + 1)?,
            Child::Frame(frame) => write_frame(xml, frame)?,
        }
    }
    xml.push_str("</w:frameset>");
    Ok(())
}

fn write_frame(xml: &mut String, frame: &Frame) -> Result<()> {
    xml.push_str("<w:frame>");
    if let Some(value) = &frame.size {
        write_value_element(xml, "sz", value)?;
    }
    if let Some(value) = &frame.name {
        write_value_element(xml, "name", value)?;
    }
    if let Some(value) = &frame.source_file_relationship_id {
        write!(xml, "<w:sourceFileName r:id=\"{}\"/>", escape(value))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = frame.margin_width {
        write_value_element(xml, "marW", &value.to_string())?;
    }
    if let Some(value) = frame.margin_height {
        write_value_element(xml, "marH", &value.to_string())?;
    }
    if let Some(value) = frame.scrollbar {
        write_value_element(xml, "scrollbar", value.as_str())?;
    }
    write_optional_on_off(xml, "noResizeAllowed", frame.no_resize_allowed)?;
    write_optional_on_off(xml, "linkedToFile", frame.linked_to_file)?;
    xml.push_str("</w:frame>");
    Ok(())
}

fn write_frameset_split_bar(xml: &mut String, split_bar: &SplitBar) -> Result<()> {
    xml.push_str("<w:framesetSplitbar>");
    if let Some(value) = split_bar.width_twips {
        write_value_element(xml, "w", &value.to_string())?;
    }
    if let Some(color) = &split_bar.color {
        xml.push_str("<w:color");
        write_color_attributes(
            xml,
            &color.value,
            color.theme_color,
            color.theme_tint,
            color.theme_shade,
        )?;
        xml.push_str("/>");
    }
    write_optional_on_off(xml, "noBorder", split_bar.no_border)?;
    write_optional_on_off(xml, "flatBorders", split_bar.flat_borders)?;
    xml.push_str("</w:framesetSplitbar>");
    Ok(())
}

fn write_html_div(xml: &mut String, div: &Div, nesting: usize) -> Result<()> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word HTML division nesting exceeds the supported safety limit".into(),
        ));
    }
    write!(xml, "<w:div w:id=\"{}\">", div.id)?;
    write_explicit_on_off(xml, "blockQuote", div.block_quote)?;
    write_explicit_on_off(xml, "bodyDiv", div.body_div)?;
    for (name, value) in [
        ("marLeft", div.left),
        ("marRight", div.right),
        ("marTop", div.top),
        ("marBottom", div.bottom),
    ] {
        write!(xml, "<w:{name} w:val=\"{value}\"/>")?;
    }
    if let Some(borders) = &div.borders {
        write_html_div_borders(xml, borders)?;
    }
    if !div.children.is_empty() {
        xml.push_str("<w:divsChild>");
        for child in &div.children {
            write_html_div(xml, child, nesting + 1)?;
        }
        xml.push_str("</w:divsChild>");
    }
    xml.push_str("</w:div>");
    Ok(())
}

fn write_html_div_borders(xml: &mut String, borders: &Borders) -> Result<()> {
    xml.push_str("<w:divBdr>");
    for (name, border) in [
        ("top", &borders.top),
        ("left", &borders.left),
        ("bottom", &borders.bottom),
        ("right", &borders.right),
    ] {
        let Some(border) = border else {
            continue;
        };
        write!(xml, "<w:{name} w:val=\"{}\"", escape(&border.style))
            .map_err(|error| Error::Xml(error.to_string()))?;
        if let Some(color) = &border.color {
            write!(xml, " w:color=\"{}\"", escape(color))
                .map_err(|error| Error::Xml(error.to_string()))?;
        }
        write_theme_attributes(
            xml,
            border.theme_color,
            border.theme_tint,
            border.theme_shade,
        )?;
        if let Some(value) = border.size_eighth_points {
            write!(xml, " w:sz=\"{value}\"").map_err(|error| Error::Xml(error.to_string()))?;
        }
        if let Some(value) = border.space_points {
            write!(xml, " w:space=\"{value}\"").map_err(|error| Error::Xml(error.to_string()))?;
        }
        write_optional_on_off_attribute(xml, "shadow", border.shadow)?;
        write_optional_on_off_attribute(xml, "frame", border.frame)?;
        xml.push_str("/>");
    }
    xml.push_str("</w:divBdr>");
    Ok(())
}

fn write_color_attributes(
    xml: &mut String,
    value: &str,
    theme_color: Option<Theme>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
) -> Result<()> {
    write!(xml, " w:val=\"{}\"", escape(value)).map_err(|error| Error::Xml(error.to_string()))?;
    write_theme_attributes(xml, theme_color, theme_tint, theme_shade)
}

fn write_theme_attributes(
    xml: &mut String,
    theme_color: Option<Theme>,
    theme_tint: Option<u8>,
    theme_shade: Option<u8>,
) -> Result<()> {
    if let Some(value) = theme_color {
        write!(xml, " w:themeColor=\"{}\"", value.as_str())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = theme_tint {
        write!(xml, " w:themeTint=\"{value:02X}\"")
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = theme_shade {
        write!(xml, " w:themeShade=\"{value:02X}\"")
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}

fn write_optional_on_off_attribute(
    xml: &mut String,
    name: &str,
    value: Option<bool>,
) -> Result<()> {
    if let Some(value) = value {
        write!(
            xml,
            " w:{name}=\"{}\"",
            if value { "true" } else { "false" }
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}

fn validate_frame_relationships(
    part: &dyn Part,
    settings: &Settings,
    conformance: Conformance,
) -> Result<()> {
    const FRAME_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame";
    const STRICT_FRAME_RELATIONSHIP: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/frame";

    fn validate(part: &dyn Part, frameset: &Frameset, expected_type: &str) -> Result<()> {
        for child in &frameset.children {
            match child {
                Child::Frameset(nested) => {
                    validate(part, nested, expected_type)?;
                },
                Child::Frame(frame) => {
                    let Some(id) = &frame.source_file_relationship_id else {
                        continue;
                    };
                    let relationship = part.rels().get(id).ok_or_else(|| {
                        Error::Invalid(format!("frame source relationship '{id}' does not exist"))
                    })?;
                    if relationship.reltype() != expected_type {
                        return Err(Error::Invalid(format!(
                            "frame source relationship '{id}' has an invalid type"
                        )));
                    }
                },
            }
        }
        Ok(())
    }

    if let Some(frameset) = &settings.frameset {
        let expected = match conformance {
            Conformance::Transitional => FRAME_RELATIONSHIP,
            Conformance::Strict => STRICT_FRAME_RELATIONSHIP,
        };
        validate(part, frameset, expected)?;
    }
    Ok(())
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<Conformance> {
    if saw_root || element.local_name().as_ref() != b"webSettings" {
        return Err(invalid(
            "web-settings part has an invalid or trailing root element",
        ));
    }
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return Err(invalid(
            "web-settings root has no WordprocessingML namespace",
        ));
    };
    Conformance::from_word_namespace(namespace)
        .ok_or_else(|| invalid("web-settings root uses an unsupported namespace"))
}

fn validate_web_child(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    conformance: Conformance,
    last_rank: &mut Option<u8>,
) -> Result<()> {
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return Err(invalid(
            "web-settings child has no WordprocessingML namespace",
        ));
    };
    if *namespace != conformance.wordprocessingml().as_bytes() {
        return Err(invalid(
            "web-settings child namespace does not match the root conformance",
        ));
    }

    let name = element.local_name();
    let name = name.as_ref();
    if conformance == Conformance::Strict && name == b"relyOnVML" {
        return Err(invalid("relyOnVML is not valid in Strict web settings"));
    }
    let rank = match name {
        b"frameset" => Some(0),
        b"divs" => Some(1),
        b"encoding" => Some(2),
        b"optimizeForBrowser" => Some(3),
        b"relyOnVML" => Some(4),
        b"allowPNG" => Some(5),
        b"doNotRelyOnCSS" => Some(6),
        b"doNotSaveAsSingleFile" => Some(7),
        b"doNotOrganizeInFolder" => Some(8),
        b"doNotUseLongFileNames" => Some(9),
        b"pixelsPerInch" => Some(10),
        b"targetScreenSz" => Some(11),
        b"saveSmartTagsAsXml" => Some(12),
        _ => None,
    };
    if let Some(rank) = rank {
        if last_rank.is_some_and(|last| rank < last) {
            return Err(invalid("web-settings children are out of schema order"));
        }
        *last_rank = Some(rank);
    }
    Ok(())
}

fn is_scalar_setting(name: &[u8]) -> bool {
    matches!(
        name,
        b"encoding"
            | b"optimizeForBrowser"
            | b"relyOnVML"
            | b"allowPNG"
            | b"doNotRelyOnCSS"
            | b"doNotSaveAsSingleFile"
            | b"doNotOrganizeInFolder"
            | b"doNotUseLongFileNames"
            | b"pixelsPerInch"
            | b"targetScreenSz"
            | b"saveSmartTagsAsXml"
    )
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut Settings,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"encoding" => set_once(
            &mut settings.encoding,
            required_value(element, decoder, resolver, "web encoding")?,
            "encoding",
        ),
        b"optimizeForBrowser" => set_on_off(
            &mut settings.optimize_for_browser,
            element,
            decoder,
            resolver,
            "optimizeForBrowser",
        ),
        b"relyOnVML" => set_on_off(
            &mut settings.rely_on_vml,
            element,
            decoder,
            resolver,
            "relyOnVML",
        ),
        b"allowPNG" => set_on_off(
            &mut settings.allow_png,
            element,
            decoder,
            resolver,
            "allowPNG",
        ),
        b"doNotRelyOnCSS" => set_on_off(
            &mut settings.do_not_rely_on_css,
            element,
            decoder,
            resolver,
            "doNotRelyOnCSS",
        ),
        b"doNotSaveAsSingleFile" => set_on_off(
            &mut settings.do_not_save_as_single_file,
            element,
            decoder,
            resolver,
            "doNotSaveAsSingleFile",
        ),
        b"doNotOrganizeInFolder" => set_on_off(
            &mut settings.do_not_organize_in_folder,
            element,
            decoder,
            resolver,
            "doNotOrganizeInFolder",
        ),
        b"doNotUseLongFileNames" => set_on_off(
            &mut settings.do_not_use_long_file_names,
            element,
            decoder,
            resolver,
            "doNotUseLongFileNames",
        ),
        b"pixelsPerInch" => {
            let value = required_value(element, decoder, resolver, "pixels per inch")?;
            let value = value
                .trim()
                .parse::<u16>()
                .map_err(|_| invalid(format!("invalid pixels-per-inch value '{value}'")))?;
            validate_pixels_per_inch(value)?;
            set_once(&mut settings.pixels_per_inch, value, "pixelsPerInch")
        },
        b"targetScreenSz" => {
            let value = required_value(element, decoder, resolver, "target screen size")?;
            let value = Screen::from_xml(&value).ok_or_else(|| {
                Error::Invalid(format!("invalid target-screen-size value '{value}'"))
            })?;
            set_once(&mut settings.target_screen_size, value, "targetScreenSz")
        },
        b"saveSmartTagsAsXml" => set_on_off(
            &mut settings.save_smart_tags_as_xml,
            element,
            decoder,
            resolver,
            "saveSmartTagsAsXml",
        ),
        _ => Ok(()),
    }
}

fn parse_frameset(
    reader: &mut NsReader<&[u8]>,
    nesting: usize,
    budget: &mut ParseBudget,
) -> Result<Frameset> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word web frameset nesting exceeds the supported safety limit".into(),
        ));
    }
    let mut frameset = Frameset::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    b"sz" => {
                        let value = required_value(&element, decoder, &resolver, "frame size")?;
                        set_once(&mut frameset.size, value, "frameset size")?;
                        finish_leaf(
                            reader,
                            element.local_name().as_ref(),
                            "frameset size",
                            budget,
                        )?;
                    },
                    b"framesetSplitbar" => {
                        let split_bar = parse_frameset_split_bar(reader, budget)?;
                        set_once(&mut frameset.split_bar, split_bar, "frameset split bar")?;
                    },
                    b"frameLayout" => {
                        let layout = parse_frame_layout(&element, decoder, &resolver)?;
                        set_once(&mut frameset.layout, layout, "frame layout")?;
                        finish_leaf(
                            reader,
                            element.local_name().as_ref(),
                            "frame layout",
                            budget,
                        )?;
                    },
                    b"frameset" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        let child = parse_frameset(reader, nesting + 1, budget)?;
                        frameset.children.push(Child::Frameset(child));
                    },
                    b"frame" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        let child = parse_frame(reader, budget)?;
                        frameset.children.push(Child::Frame(child));
                    },
                    _ => skip_element(reader, budget)?,
                }
            },
            Event::Empty(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    b"sz" => set_once(
                        &mut frameset.size,
                        required_value(&element, decoder, &resolver, "frame size")?,
                        "frameset size",
                    )?,
                    b"framesetSplitbar" => set_once(
                        &mut frameset.split_bar,
                        SplitBar::default(),
                        "frameset split bar",
                    )?,
                    b"frameLayout" => set_once(
                        &mut frameset.layout,
                        parse_frame_layout(&element, decoder, &resolver)?,
                        "frame layout",
                    )?,
                    b"frameset" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        frameset.children.push(Child::Frameset(Frameset::default()));
                    },
                    b"frame" => {
                        reserve_one(&mut frameset.children, "parsed frameset child")?;
                        frameset.children.push(Child::Frame(Frame::default()));
                    },
                    _ => {},
                }
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"frameset" =>
            {
                return Ok(frameset);
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word web frameset".into()));
            },
            _ => {},
        }
    }
}

fn parse_div_container(
    reader: &mut NsReader<&[u8]>,
    end_name: &[u8],
    nesting: usize,
    budget: &mut ParseBudget,
) -> Result<Vec<Div>> {
    if nesting > MAX_FRAMESET_NESTING {
        return Err(Error::Invalid(
            "Word HTML division nesting exceeds the supported safety limit".into(),
        ));
    }
    let mut divs = Vec::new();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                reserve_one(&mut divs, "parsed HTML division")?;
                let div = parse_html_div(reader, &element, decoder, &resolver, nesting, budget)?;
                divs.push(div);
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                let _ = parse_div_id(&element, decoder, &resolver)?;
                return Err(invalid(
                    "Word HTML division is missing its four required margins",
                ));
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == end_name =>
            {
                if divs.is_empty() {
                    return Err(invalid("Word HTML division container must not be empty"));
                }
                return Ok(divs);
            },
            Event::Eof => {
                return Err(Error::Invalid(
                    "unterminated Word HTML division container".into(),
                ));
            },
            _ => {},
        }
    }
}

fn parse_div_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Id> {
    let id = word_attribute_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| Error::Invalid("Word HTML division ID is required".into()))?;
    Id::parse(&id)
}

struct DivParse {
    value: Div,
    margins: u8,
    last_rank: Option<u8>,
}

impl DivParse {
    const LEFT: u8 = 1;
    const RIGHT: u8 = 2;
    const TOP: u8 = 4;
    const BOTTOM: u8 = 8;
    const ALL_MARGINS: u8 = Self::LEFT | Self::RIGHT | Self::TOP | Self::BOTTOM;

    fn new(id: Id) -> Self {
        Self {
            value: Div::new(id),
            margins: 0,
            last_rank: None,
        }
    }

    fn advance(&mut self, rank: u8) -> Result<()> {
        if self.last_rank.is_some_and(|last| rank < last) {
            return Err(invalid("HTML division children are out of schema order"));
        }
        self.last_rank = Some(rank);
        Ok(())
    }

    fn set_margin(&mut self, bit: u8, value: Twips, description: &'static str) -> Result<()> {
        if self.margins & bit != 0 {
            return Err(invalid(format!("duplicate Word {description}")));
        }
        self.margins |= bit;
        match bit {
            Self::LEFT => self.value.left = value,
            Self::RIGHT => self.value.right = value,
            Self::TOP => self.value.top = value,
            Self::BOTTOM => self.value.bottom = value,
            _ => return Err(invalid("invalid HTML division margin selector")),
        }
        Ok(())
    }

    fn append_children(&mut self, mut children: Vec<Div>) -> Result<()> {
        self.value
            .children
            .try_reserve(children.len())
            .map_err(|source| Error::Allocation {
                resource: "parsed HTML child divisions",
                source,
            })?;
        self.value.children.append(&mut children);
        Ok(())
    }

    fn finish(self) -> Result<Div> {
        if self.margins != Self::ALL_MARGINS {
            return Err(invalid(
                "Word HTML division is missing one or more required margins",
            ));
        }
        Ok(self.value)
    }
}

fn parse_html_div(
    reader: &mut NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    nesting: usize,
    budget: &mut ParseBudget,
) -> Result<Div> {
    let mut div = DivParse::new(parse_div_id(element, decoder, resolver)?);
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    name if is_html_div_leaf(name) => {
                        parse_html_div_leaf(&element, decoder, &resolver, &mut div)?;
                        finish_leaf(
                            reader,
                            element.local_name().as_ref(),
                            "HTML division property",
                            budget,
                        )?;
                    },
                    b"divBdr" => {
                        div.advance(6)?;
                        let borders = parse_html_div_borders(reader, budget)?;
                        set_once(&mut div.value.borders, borders, "HTML division borders")?;
                    },
                    b"divsChild" => {
                        div.advance(7)?;
                        let children =
                            parse_div_container(reader, b"divsChild", nesting + 1, budget)?;
                        div.append_children(children)?;
                    },
                    _ => skip_element(reader, budget)?,
                }
            },
            Event::Empty(element) if is_wordprocessing_namespace(&namespace) => {
                match element.local_name().as_ref() {
                    name if is_html_div_leaf(name) => {
                        parse_html_div_leaf(&element, decoder, &resolver, &mut div)?;
                    },
                    b"divBdr" => {
                        div.advance(6)?;
                        set_once(
                            &mut div.value.borders,
                            Borders::default(),
                            "HTML division borders",
                        )?;
                    },
                    b"divsChild" => {
                        return Err(invalid(
                            "Word HTML child division container must not be empty",
                        ));
                    },
                    _ => {},
                }
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"div" =>
            {
                return div.finish();
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word HTML division".into()));
            },
            _ => {},
        }
    }
}

fn is_html_div_leaf(name: &[u8]) -> bool {
    matches!(
        name,
        b"blockQuote" | b"bodyDiv" | b"marLeft" | b"marRight" | b"marTop" | b"marBottom"
    )
}

fn parse_html_div_leaf(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    div: &mut DivParse,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"blockQuote" => {
            div.advance(0)?;
            set_on_off(
                &mut div.value.block_quote,
                element,
                decoder,
                resolver,
                "HTML blockQuote",
            )
        },
        b"bodyDiv" => {
            div.advance(1)?;
            set_on_off(
                &mut div.value.body_div,
                element,
                decoder,
                resolver,
                "HTML bodyDiv",
            )
        },
        b"marLeft" => set_signed_twips(
            div,
            DivParse::LEFT,
            2,
            element,
            decoder,
            resolver,
            "HTML division left margin",
        ),
        b"marRight" => set_signed_twips(
            div,
            DivParse::RIGHT,
            3,
            element,
            decoder,
            resolver,
            "HTML division right margin",
        ),
        b"marTop" => set_signed_twips(
            div,
            DivParse::TOP,
            4,
            element,
            decoder,
            resolver,
            "HTML division top margin",
        ),
        b"marBottom" => set_signed_twips(
            div,
            DivParse::BOTTOM,
            5,
            element,
            decoder,
            resolver,
            "HTML division bottom margin",
        ),
        _ => Ok(()),
    }
}

fn set_signed_twips(
    div: &mut DivParse,
    bit: u8,
    rank: u8,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &'static str,
) -> Result<()> {
    div.advance(rank)?;
    let value = required_value(element, decoder, resolver, description)?;
    let value = Twips::parse(&value)?;
    div.set_margin(bit, value, description)
}

fn parse_html_div_borders(
    reader: &mut NsReader<&[u8]>,
    budget: &mut ParseBudget,
) -> Result<Borders> {
    let mut borders = Borders::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_html_div_border_side(element.local_name().as_ref()) =>
            {
                set_html_div_border(&mut borders, &element, decoder, &resolver)?;
                finish_leaf(
                    reader,
                    element.local_name().as_ref(),
                    "HTML division border",
                    budget,
                )?;
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_html_div_border_side(element.local_name().as_ref()) =>
            {
                set_html_div_border(&mut borders, &element, decoder, &resolver)?;
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"divBdr" =>
            {
                return Ok(borders);
            },
            Event::Eof => {
                return Err(Error::Invalid(
                    "unterminated Word HTML division borders".into(),
                ));
            },
            _ => {},
        }
    }
}

fn is_html_div_border_side(name: &[u8]) -> bool {
    matches!(name, b"top" | b"left" | b"bottom" | b"right")
}

fn set_html_div_border(
    borders: &mut Borders,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let border = parse_html_div_border(element, decoder, resolver)?;
    let (slot, description) = match element.local_name().as_ref() {
        b"top" => (&mut borders.top, "top HTML division border"),
        b"left" => (&mut borders.left, "left HTML division border"),
        b"bottom" => (&mut borders.bottom, "bottom HTML division border"),
        b"right" => (&mut borders.right, "right HTML division border"),
        _ => return Ok(()),
    };
    set_once(slot, border, description)
}

fn parse_html_div_border(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Border> {
    let style = required_value(element, decoder, resolver, "HTML division border style")?;
    let color = word_attribute_value(element, b"color", decoder, resolver)?
        .map(|value| validate_word_color(value, "HTML division border color"))
        .transpose()?;
    let theme_color = word_attribute_value(element, b"themeColor", decoder, resolver)?
        .map(|value| {
            Theme::parse(&value)
                .ok_or_else(|| Error::Invalid(format!("invalid theme color '{value}'")))
        })
        .transpose()?;
    Ok(Border {
        style,
        color,
        theme_color,
        theme_tint: optional_hex_byte(element, b"themeTint", decoder, resolver)?,
        theme_shade: optional_hex_byte(element, b"themeShade", decoder, resolver)?,
        size_eighth_points: optional_unsigned_long_attribute(element, b"sz", decoder, resolver)?,
        space_points: optional_unsigned_long_attribute(element, b"space", decoder, resolver)?,
        shadow: optional_on_off_attribute(element, b"shadow", decoder, resolver)?,
        frame: optional_on_off_attribute(element, b"frame", decoder, resolver)?,
    })
}

fn parse_frame(reader: &mut NsReader<&[u8]>, budget: &mut ParseBudget) -> Result<Frame> {
    let mut frame = Frame::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                if is_frame_property(element.local_name().as_ref()) {
                    parse_frame_property(&element, decoder, &resolver, &mut frame)?;
                    finish_leaf(
                        reader,
                        element.local_name().as_ref(),
                        "frame property",
                        budget,
                    )?;
                } else {
                    skip_element(reader, budget)?;
                }
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_frame_property(element.local_name().as_ref()) =>
            {
                parse_frame_property(&element, decoder, &resolver, &mut frame)?;
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"frame" =>
            {
                return Ok(frame);
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word web frame".into()));
            },
            _ => {},
        }
    }
}

fn is_frame_property(name: &[u8]) -> bool {
    matches!(
        name,
        b"sz"
            | b"name"
            | b"sourceFileName"
            | b"marW"
            | b"marH"
            | b"scrollbar"
            | b"noResizeAllowed"
            | b"linkedToFile"
    )
}

fn parse_frame_property(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    frame: &mut Frame,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"sz" => set_once(
            &mut frame.size,
            required_value(element, decoder, resolver, "frame size")?,
            "frame size",
        ),
        b"name" => set_once(
            &mut frame.name,
            required_value(element, decoder, resolver, "frame name")?,
            "frame name",
        ),
        b"sourceFileName" => set_once(
            &mut frame.source_file_relationship_id,
            required_relationship_id(element, decoder, resolver)?,
            "frame source file",
        ),
        b"marW" => set_once(
            &mut frame.margin_width,
            required_unsigned_long(element, decoder, resolver, "frame margin width")?,
            "frame margin width",
        ),
        b"marH" => set_once(
            &mut frame.margin_height,
            required_unsigned_long(element, decoder, resolver, "frame margin height")?,
            "frame margin height",
        ),
        b"scrollbar" => {
            let value = required_value(element, decoder, resolver, "frame scrollbar")?;
            let value = Scrollbar::from_xml(&value).ok_or_else(|| {
                Error::Invalid(format!("invalid frame scrollbar value '{value}'"))
            })?;
            set_once(&mut frame.scrollbar, value, "frame scrollbar")
        },
        b"noResizeAllowed" => set_on_off(
            &mut frame.no_resize_allowed,
            element,
            decoder,
            resolver,
            "frame noResizeAllowed",
        ),
        b"linkedToFile" => set_on_off(
            &mut frame.linked_to_file,
            element,
            decoder,
            resolver,
            "frame linkedToFile",
        ),
        _ => Ok(()),
    }
}

fn parse_frameset_split_bar(
    reader: &mut NsReader<&[u8]>,
    budget: &mut ParseBudget,
) -> Result<SplitBar> {
    let mut split_bar = SplitBar::default();
    loop {
        budget.event()?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if is_wordprocessing_namespace(&namespace) => {
                if is_split_bar_property(element.local_name().as_ref()) {
                    parse_split_bar_property(&element, decoder, &resolver, &mut split_bar)?;
                    finish_leaf(
                        reader,
                        element.local_name().as_ref(),
                        "frameset split-bar property",
                        budget,
                    )?;
                } else {
                    skip_element(reader, budget)?;
                }
            },
            Event::Empty(element)
                if is_wordprocessing_namespace(&namespace)
                    && is_split_bar_property(element.local_name().as_ref()) =>
            {
                parse_split_bar_property(&element, decoder, &resolver, &mut split_bar)?;
            },
            Event::Start(_) => skip_element(reader, budget)?,
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"framesetSplitbar" =>
            {
                return Ok(split_bar);
            },
            Event::Eof => {
                return Err(Error::Invalid(
                    "unterminated Word frameset split bar".into(),
                ));
            },
            _ => {},
        }
    }
}

fn is_split_bar_property(name: &[u8]) -> bool {
    matches!(name, b"w" | b"color" | b"noBorder" | b"flatBorders")
}

fn parse_split_bar_property(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    split_bar: &mut SplitBar,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"w" => set_once(
            &mut split_bar.width_twips,
            required_unsigned_long(element, decoder, resolver, "split-bar width")?,
            "split-bar width",
        ),
        b"color" => set_once(
            &mut split_bar.color,
            parse_frameset_color(element, decoder, resolver)?,
            "split-bar color",
        ),
        b"noBorder" => set_on_off(
            &mut split_bar.no_border,
            element,
            decoder,
            resolver,
            "split-bar noBorder",
        ),
        b"flatBorders" => set_on_off(
            &mut split_bar.flat_borders,
            element,
            decoder,
            resolver,
            "split-bar flatBorders",
        ),
        _ => Ok(()),
    }
}

fn parse_frame_layout(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Layout> {
    let value = required_value(element, decoder, resolver, "frame layout")?;
    Layout::from_xml(&value)
        .ok_or_else(|| Error::Invalid(format!("invalid frame-layout value '{value}'")))
}

fn parse_frameset_color(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Color> {
    let value = validate_word_color(
        required_value(element, decoder, resolver, "frameset splitter color")?,
        "frameset splitter color",
    )?;
    let theme_color = word_attribute_value(element, b"themeColor", decoder, resolver)?
        .map(|value| {
            Theme::parse(&value)
                .ok_or_else(|| Error::Invalid(format!("invalid theme color '{value}'")))
        })
        .transpose()?;
    let theme_tint = optional_hex_byte(element, b"themeTint", decoder, resolver)?;
    let theme_shade = optional_hex_byte(element, b"themeShade", decoder, resolver)?;
    Ok(Color {
        value,
        theme_color,
        theme_tint,
        theme_shade,
    })
}

pub(super) fn validate_word_color(value: String, description: &str) -> Result<String> {
    if value != "auto" && (value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(Error::Invalid(format!("invalid {description} '{value}'")));
    }
    Ok(value)
}

fn required_unsigned_long(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<u64> {
    let value = required_value(element, decoder, resolver, description)?;
    value
        .trim()
        .parse::<u64>()
        .map_err(|_| Error::Invalid(format!("invalid unsigned {description} value '{value}'")))
}

fn optional_unsigned_long_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<u64>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| {
            value.trim().parse::<u64>().map_err(|_| {
                Error::Invalid(format!("invalid unsigned Word attribute value '{value}'"))
            })
        })
        .transpose()
}

fn optional_on_off_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<bool>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| match value.as_str() {
            "true" | "1" | "on" => Ok(true),
            "false" | "0" | "off" => Ok(false),
            _ => Err(Error::Invalid(format!(
                "invalid Word on/off value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_hex_byte(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<u8>> {
    word_attribute_value(element, name, decoder, resolver)?
        .map(|value| {
            if value.len() != 2 {
                return Err(Error::Invalid(format!(
                    "invalid hexadecimal byte '{value}'"
                )));
            }
            u8::from_str_radix(&value, 16)
                .map_err(|_| Error::Invalid(format!("invalid hexadecimal byte '{value}'")))
        })
        .transpose()
}

fn required_relationship_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    const RELATIONSHIPS: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(namespace))
                if namespace == RELATIONSHIPS || namespace == STRICT_RELATIONSHIPS
        );
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(Error::Invalid(
                "duplicate frame source relationship ID".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    value.ok_or_else(|| Error::Invalid("frame source relationship ID is required".into()))
}

fn finish_leaf(
    reader: &mut NsReader<&[u8]>,
    expected_name: &[u8],
    description: &str,
    budget: &mut ParseBudget,
) -> Result<()> {
    loop {
        budget.event()?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::End(element)
                if is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == expected_name =>
            {
                return Ok(());
            },
            Event::End(_) => {
                return Err(invalid(format!(
                    "Word {description} has a mismatched closing element"
                )));
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => {
                return Err(Error::Invalid(format!(
                    "unterminated Word {description} element"
                )));
            },
            _ => {
                return Err(Error::Invalid(format!(
                    "Word {description} must not contain child content"
                )));
            },
        }
    }
}

fn skip_element(reader: &mut NsReader<&[u8]>, budget: &mut ParseBudget) -> Result<()> {
    let mut depth = 1usize;
    loop {
        budget.event()?;
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("Word web XML nesting is too deep".into()))?;
                if depth > MAX_FRAMESET_NESTING {
                    return Err(invalid("Word web XML nesting exceeds the safety limit"));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid Word web XML nesting".into()))?;
                if depth == 0 {
                    return Ok(());
                }
            },
            Event::Eof => {
                return Err(Error::Invalid("unterminated Word web XML element".into()));
            },
            _ => {},
        }
    }
}

fn required_value(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, b"val", decoder, resolver)?
        .ok_or_else(|| Error::Invalid(format!("Word {description} value is required")))
}

fn set_on_off(
    slot: &mut Option<bool>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<()> {
    let value = match word_attribute_value(element, b"val", decoder, resolver)? {
        Some(value) => match value.as_str() {
            "true" | "1" | "on" => true,
            "false" | "0" | "off" => false,
            _ => {
                return Err(Error::Invalid(format!(
                    "invalid Word on/off value '{value}'"
                )));
            },
        },
        None => true,
    };
    set_once(slot, value, description)
}

fn set_once<T>(slot: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::Invalid(format!(
            "duplicate Word web setting '{description}'"
        )));
    }
    Ok(())
}
