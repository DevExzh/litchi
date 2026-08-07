//! Bounded XML codecs and snapshot-preserving edits for outline styles.

use std::collections::HashSet;

use litchi_core::{Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};

use crate::list_label_alignment::{Alignment, FollowedBy, Length};

use super::{
    FO, MAX_DEPTH, MAX_OUTLINE_LEVELS, MAX_STYLES, MAX_TOTAL_ATTRIBUTE_BYTES, MAX_VALUE_BYTES,
    MAX_XML_BYTES, NamespaceKind, OFFICE, STYLE, TEXT, invalid, invalid_error,
    model::{
        Attribute, LevelStyle, ListProperties, NumberFormat, PositionMode, PositiveInteger, Style,
        Styles, TextAlign, TextProperties, validate_text,
    },
    namespace_kind,
};

impl Style {
    /// Serialize this style with complete namespace declarations.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let extension_namespaces = collect_extension_namespaces(self);
        let mut output = String::with_capacity(1_024);
        output.push_str("<text:outline-style xmlns:text=\"");
        output.push_str(std::str::from_utf8(TEXT).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:style=\"");
        output.push_str(std::str::from_utf8(STYLE).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:fo=\"");
        output.push_str(std::str::from_utf8(FO).expect("namespace is UTF-8"));
        output.push_str("\" xmlns:office=\"");
        output.push_str(std::str::from_utf8(OFFICE).expect("namespace is UTF-8"));
        output.push('"');
        for (index, namespace) in extension_namespaces.iter().enumerate() {
            output.push_str(&format!(" xmlns:ext{index}=\"{}\"", escape_xml(namespace)));
        }
        write_attribute(&mut output, "style:name", &self.name);
        write_descriptors(&mut output, &self.extensions, &extension_namespaces)?;
        output.push('>');
        for level in &self.levels {
            write_level(&mut output, level, &extension_namespaces)?;
        }
        output.push_str("</text:outline-style>");
        if output.len() > MAX_XML_BYTES {
            return invalid("serialized outline style exceeds the resource limit");
        }
        Ok(output)
    }
}

fn write_level(output: &mut String, level: &LevelStyle, namespaces: &[String]) -> Result<()> {
    output.push_str("<text:outline-level-style");
    write_attribute(output, "text:level", &level.level.to_string());
    write_optional_attribute(output, "text:style-name", level.text_style_name.as_deref());
    write_optional_attribute(
        output,
        "style:num-format",
        level.number_format.as_ref().map(NumberFormat::as_str),
    );
    write_optional_attribute(output, "style:num-prefix", level.number_prefix.as_deref());
    write_optional_attribute(output, "style:num-suffix", level.number_suffix.as_deref());
    if let Some(value) = level.letter_sync {
        write_attribute(
            output,
            "style:num-letter-sync",
            if value { "true" } else { "false" },
        );
    }
    write_optional_attribute(
        output,
        "text:display-levels",
        level.display_levels.as_ref().map(PositiveInteger::as_str),
    );
    write_optional_attribute(
        output,
        "text:start-value",
        level.start_value.as_ref().map(PositiveInteger::as_str),
    );
    write_descriptors(output, &level.extensions, namespaces)?;
    if level.list_level_properties.is_none() && level.text_properties.is_none() {
        output.push_str("/>");
        return Ok(());
    }
    output.push('>');
    if let Some(properties) = level.list_level_properties.as_ref() {
        write_list_properties(output, properties, namespaces)?;
    }
    if let Some(properties) = level.text_properties.as_ref() {
        output.push_str("<style:text-properties");
        write_descriptors(output, &properties.attributes, namespaces)?;
        output.push_str("/>");
    }
    output.push_str("</text:outline-level-style>");
    Ok(())
}

fn write_list_properties(
    output: &mut String,
    properties: &ListProperties,
    namespaces: &[String],
) -> Result<()> {
    output.push_str("<style:list-level-properties");
    write_optional_attribute(
        output,
        "fo:text-align",
        properties.text_align.map(TextAlign::as_str),
    );
    for (name, value) in [
        ("text:space-before", properties.space_before.as_ref()),
        (
            "text:min-label-width",
            properties.minimum_label_width.as_ref(),
        ),
        (
            "text:min-label-distance",
            properties.minimum_label_distance.as_ref(),
        ),
        ("fo:width", properties.width.as_ref()),
        ("fo:height", properties.height.as_ref()),
    ] {
        write_optional_attribute(output, name, value.map(Length::as_str));
    }
    write_optional_attribute(output, "style:font-name", properties.font_name.as_deref());
    write_optional_attribute(
        output,
        "style:vertical-rel",
        properties.vertical_relation.as_deref(),
    );
    write_optional_attribute(
        output,
        "style:vertical-pos",
        properties.vertical_position.as_deref(),
    );
    write_optional_attribute(
        output,
        "text:list-level-position-and-space-mode",
        properties.position_mode.map(PositionMode::as_str),
    );
    write_descriptors(output, &properties.extensions, namespaces)?;
    let Some(alignment) = properties.label_alignment.as_ref() else {
        output.push_str("/>");
        return Ok(());
    };
    output.push('>');
    output.push_str("<style:list-level-label-alignment");
    write_attribute(
        output,
        "text:label-followed-by",
        match alignment.label_followed_by {
            FollowedBy::ListTab => "listtab",
            FollowedBy::Space => "space",
            FollowedBy::Nothing => "nothing",
        },
    );
    write_optional_attribute(
        output,
        "text:list-tab-stop-position",
        alignment
            .list_tab_stop_position
            .as_ref()
            .map(Length::as_str),
    );
    write_optional_attribute(
        output,
        "fo:text-indent",
        alignment.text_indent.as_ref().map(Length::as_str),
    );
    write_optional_attribute(
        output,
        "fo:margin-left",
        alignment.margin_left.as_ref().map(Length::as_str),
    );
    output.push_str("/></style:list-level-properties>");
    Ok(())
}

fn collect_extension_namespaces(style: &Style) -> Vec<String> {
    let mut namespaces = Vec::new();
    let mut visit = |attribute: &Attribute| {
        if namespace_kind(attribute.namespace_uri.as_bytes()) == NamespaceKind::Other
            && !namespaces.contains(&attribute.namespace_uri)
        {
            namespaces.push(attribute.namespace_uri.clone());
        }
    };
    for attribute in &style.extensions {
        visit(attribute);
    }
    for level in &style.levels {
        for attribute in &level.extensions {
            visit(attribute);
        }
        if let Some(properties) = level.list_level_properties.as_ref() {
            for attribute in &properties.extensions {
                visit(attribute);
            }
        }
        if let Some(properties) = level.text_properties.as_ref() {
            for attribute in &properties.attributes {
                visit(attribute);
            }
        }
    }
    namespaces
}

fn write_descriptors(
    output: &mut String,
    attributes: &[Attribute],
    namespaces: &[String],
) -> Result<()> {
    let mut seen = HashSet::new();
    for attribute in attributes {
        attribute.validate()?;
        if !seen.insert((&attribute.namespace_uri, &attribute.local_name)) {
            return invalid("duplicate expanded outline formatting attribute");
        }
        let prefix = match namespace_kind(attribute.namespace_uri.as_bytes()) {
            NamespaceKind::Office => "office".to_string(),
            NamespaceKind::Text => "text".to_string(),
            NamespaceKind::Style => "style".to_string(),
            NamespaceKind::Fo => "fo".to_string(),
            NamespaceKind::Other => format!(
                "ext{}",
                namespaces
                    .iter()
                    .position(|namespace| namespace == &attribute.namespace_uri)
                    .ok_or_else(|| invalid_error("missing outline extension namespace"))?
            ),
        };
        write_attribute(
            output,
            &format!("{prefix}:{}", attribute.local_name),
            &attribute.value,
        );
    }
    Ok(())
}

fn write_optional_attribute(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        write_attribute(output, name, value);
    }
}

fn write_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

#[derive(Debug)]
struct ResolvedAttribute {
    namespace: NamespaceKind,
    namespace_uri: String,
    local_name: String,
    value: String,
}

#[derive(Debug)]
struct ActiveOutline {
    depth: usize,
    value: Style,
    levels: HashSet<u16>,
}

#[derive(Debug)]
struct ActiveLevel {
    depth: usize,
    value: LevelStyle,
    child_order: u8,
}

#[derive(Debug)]
struct ActiveProperties {
    depth: usize,
    value: ListProperties,
    alignment_depth: Option<usize>,
}

/// Parse every outline numbering style in `office:styles`.
pub fn parse_outline_styles(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("outline style XML exceeds the resource limit");
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(NamespaceKind, Vec<u8>)> = Vec::new();
    let mut outline: Option<ActiveOutline> = None;
    let mut level: Option<ActiveLevel> = None;
    let mut properties: Option<ActiveProperties> = None;
    let mut text_properties_depth = None;
    let mut styles = Vec::new();
    let mut names = HashSet::new();
    let mut total_attribute_bytes = 0usize;

    loop {
        match reader.read_event() {
            Ok(Event::Decl(declaration)) => {
                version = declaration
                    .xml_version()
                    .map_err(|error| invalid_error(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return invalid("outline style XML nesting is too deep");
                }
                let current = resolve_element(&reader, start.name());
                let parent = stack.last().cloned();
                stack.push(current.clone());
                let depth = stack.len();
                handle_start(
                    &reader,
                    version,
                    &start,
                    current,
                    parent.as_ref(),
                    depth,
                    false,
                    &mut outline,
                    &mut level,
                    &mut properties,
                    &mut text_properties_depth,
                    &mut styles,
                    &mut names,
                    &mut total_attribute_bytes,
                )?;
            },
            Ok(Event::Empty(start)) => {
                let current = resolve_element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                handle_start(
                    &reader,
                    version,
                    &start,
                    current,
                    parent,
                    depth,
                    true,
                    &mut outline,
                    &mut level,
                    &mut properties,
                    &mut text_properties_depth,
                    &mut styles,
                    &mut names,
                    &mut total_attribute_bytes,
                )?;
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if properties
                    .as_ref()
                    .and_then(|active| active.alignment_depth)
                    == Some(depth)
                {
                    properties
                        .as_mut()
                        .expect("properties exist")
                        .alignment_depth = None;
                }
                if properties
                    .as_ref()
                    .is_some_and(|active| active.depth == depth)
                {
                    let completed = properties.take().expect("properties exist").value;
                    level
                        .as_mut()
                        .expect("outline level owns properties")
                        .value
                        .list_level_properties = Some(completed);
                }
                if text_properties_depth == Some(depth) {
                    text_properties_depth = None;
                }
                if level.as_ref().is_some_and(|active| active.depth == depth) {
                    finish_level(&mut outline, level.take().expect("level exists"))?;
                }
                if outline.as_ref().is_some_and(|active| active.depth == depth) {
                    finish_outline(
                        &mut styles,
                        &mut names,
                        outline.take().expect("outline exists"),
                    )?;
                }
                stack.pop();
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if outline.is_some() && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return invalid("outline style elements cannot contain text");
                }
            },
            Ok(Event::CData(data)) => {
                let bytes: &[u8] = data.as_ref();
                if outline.is_some() && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return invalid("outline style elements cannot contain CDATA");
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return invalid("DTD and processing instructions are not allowed");
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return invalid(format!("invalid outline style XML: {error}")),
        }
    }
    if outline.is_some() || level.is_some() || properties.is_some() {
        return invalid("unterminated outline style");
    }
    Ok(Styles { styles })
}

#[derive(Clone)]
struct XmlSpan {
    depth: usize,
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}

/// Insert or replace one outline style while preserving unrelated XML bytes.
pub fn set_outline_style_xml(xml: &str, style: &Style) -> Result<(String, Option<Style>)> {
    style.validate()?;
    let parsed = parse_outline_styles(xml)?;
    let old = parsed.get(&style.name).cloned();
    let (styles, target) = scan_outline_spans(xml, &style.name)?;
    let replacement = style.to_xml()?;
    let updated = if let Some(target) = target {
        replace_span(xml, &target, &replacement)
    } else if styles.empty {
        expand_empty_span(xml, &styles, &replacement)?
    } else {
        let mut updated = xml.to_owned();
        updated.insert_str(styles.end_start, &replacement);
        updated
    };
    parse_outline_styles(&updated)?;
    Ok((updated, old))
}

/// Remove one outline style while preserving unrelated XML bytes.
pub fn remove_outline_style_xml(xml: &str, name: &str) -> Result<(String, Option<Style>)> {
    validate_text(name, "style:name", false)?;
    let parsed = parse_outline_styles(xml)?;
    let Some(old) = parsed.get(name).cloned() else {
        return Ok((xml.to_owned(), None));
    };
    let (_, target) = scan_outline_spans(xml, name)?;
    let target = target.ok_or_else(|| invalid_error("outline style span was not found"))?;
    let updated = replace_span(xml, &target, "");
    parse_outline_styles(&updated)?;
    Ok((updated, Some(old)))
}

fn scan_outline_spans(xml: &str, name: &str) -> Result<(XmlSpan, Option<XmlSpan>)> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(NamespaceKind, Vec<u8>)> = Vec::new();
    let mut styles = None;
    let mut target = None;
    loop {
        match reader.read_event() {
            Ok(Event::Decl(declaration)) => {
                version = declaration
                    .xml_version()
                    .map_err(|error| invalid_error(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                let current = resolve_element(&reader, start.name());
                let parent = stack.last().cloned();
                stack.push(current.clone());
                let depth = stack.len();
                if current.0 == NamespaceKind::Office && current.1 == b"styles" {
                    if styles.is_some() {
                        return invalid("duplicate office:styles container");
                    }
                    styles = Some(XmlSpan {
                        depth,
                        start: begin,
                        end: 0,
                        end_start: 0,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        empty: false,
                    });
                } else if parent.is_some_and(|parent| {
                    parent.0 == NamespaceKind::Office && parent.1 == b"styles"
                }) && current.0 == NamespaceKind::Text
                    && current.1 == b"outline-style"
                {
                    let mut total = 0;
                    let mut attributes = attributes(&reader, version, &start, &mut total)?;
                    if take(&mut attributes, NamespaceKind::Style, "name").as_deref() == Some(name)
                    {
                        if target.is_some() {
                            return invalid("duplicate target outline style");
                        }
                        target = Some(XmlSpan {
                            depth,
                            start: begin,
                            end: 0,
                            end_start: 0,
                            qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                            empty: false,
                        });
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                let current = resolve_element(&reader, start.name());
                if current.0 == NamespaceKind::Office && current.1 == b"styles" {
                    if styles.is_some() {
                        return invalid("duplicate office:styles container");
                    }
                    styles = Some(XmlSpan {
                        depth: stack.len() + 1,
                        start: begin,
                        end,
                        end_start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        empty: true,
                    });
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = event_start(xml, end)?;
                let depth = stack.len();
                if styles.as_ref().is_some_and(|span| span.depth == depth) {
                    let span = styles.as_mut().expect("styles span exists");
                    span.end_start = begin;
                    span.end = end;
                }
                if target.as_ref().is_some_and(|span| span.depth == depth) {
                    let span = target.as_mut().expect("target span exists");
                    span.end_start = begin;
                    span.end = end;
                }
                stack.pop();
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return invalid(format!("invalid outline style XML: {error}")),
        }
    }
    let styles = styles.ok_or_else(|| invalid_error("document has no office:styles container"))?;
    if styles.end == 0 {
        return invalid("unterminated office:styles container");
    }
    if target.as_ref().is_some_and(|span| span.end == 0) {
        return invalid("unterminated outline style");
    }
    Ok((styles, target))
}

fn replace_span(xml: &str, span: &XmlSpan, replacement: &str) -> String {
    let mut output = String::with_capacity(xml.len() - (span.end - span.start) + replacement.len());
    output.push_str(&xml[..span.start]);
    output.push_str(replacement);
    output.push_str(&xml[span.end..]);
    output
}

fn expand_empty_span(xml: &str, span: &XmlSpan, child: &str) -> Result<String> {
    let source = &xml[span.start..span.end];
    let close = source
        .rfind("/>")
        .ok_or_else(|| invalid_error("invalid empty office:styles element"))?;
    let mut replacement = String::with_capacity(source.len() + child.len() + span.qname.len() + 3);
    replacement.push_str(&source[..close]);
    replacement.push('>');
    replacement.push_str(child);
    replacement.push_str("</");
    replacement.push_str(&span.qname);
    replacement.push('>');
    Ok(replace_span(xml, span, &replacement))
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| invalid_error("invalid outline XML event boundary"))
}

#[allow(clippy::too_many_arguments)]
fn handle_start(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    current: (NamespaceKind, Vec<u8>),
    parent: Option<&(NamespaceKind, Vec<u8>)>,
    depth: usize,
    empty: bool,
    outline: &mut Option<ActiveOutline>,
    level: &mut Option<ActiveLevel>,
    properties: &mut Option<ActiveProperties>,
    text_properties_depth: &mut Option<usize>,
    styles: &mut Vec<Style>,
    names: &mut HashSet<String>,
    total_attribute_bytes: &mut usize,
) -> Result<()> {
    let is_outline = current.0 == NamespaceKind::Text && current.1 == b"outline-style";
    if outline.is_none() {
        if is_outline {
            let valid_parent = parent
                .is_some_and(|parent| parent.0 == NamespaceKind::Office && parent.1 == b"styles");
            if !valid_parent || empty {
                return invalid("text:outline-style must be a nonempty child of office:styles");
            }
            let mut attributes = attributes(reader, version, start, total_attribute_bytes)?;
            let name = take(&mut attributes, NamespaceKind::Style, "name")
                .ok_or_else(|| invalid_error("outline style requires style:name"))?;
            validate_text(&name, "style:name", false)?;
            let extensions = finish_attributes(attributes, "text:outline-style")?;
            *outline = Some(ActiveOutline {
                depth,
                value: Style {
                    name,
                    levels: Vec::new(),
                    extensions,
                },
                levels: HashSet::new(),
            });
        }
        return Ok(());
    }

    if current.0 == NamespaceKind::Text && current.1 == b"outline-style" {
        return invalid("nested or misplaced text:outline-style");
    }
    if let Some(active_properties) = properties.as_mut() {
        if depth == active_properties.depth + 1
            && current.0 == NamespaceKind::Style
            && current.1 == b"list-level-label-alignment"
        {
            if active_properties.value.label_alignment.is_some() {
                return invalid("duplicate style:list-level-label-alignment");
            }
            let alignment = parse_alignment(reader, version, start, total_attribute_bytes)?;
            active_properties.value.label_alignment = Some(alignment);
            if !empty {
                active_properties.alignment_depth = Some(depth);
            }
            return Ok(());
        }
        if active_properties.alignment_depth.is_some() || depth > active_properties.depth {
            return invalid("invalid child of style:list-level-properties");
        }
    }
    if text_properties_depth.is_some() {
        return invalid("style:text-properties must be empty");
    }

    let outline_depth = outline.as_ref().expect("outline exists").depth;
    if level.is_none() {
        if depth != outline_depth + 1
            || current.0 != NamespaceKind::Text
            || current.1 != b"outline-level-style"
        {
            return invalid("text:outline-style has an invalid child");
        }
        let value = parse_level(reader, version, start, total_attribute_bytes)?;
        if !outline
            .as_mut()
            .expect("outline exists")
            .levels
            .insert(value.level)
        {
            return invalid("duplicate outline level");
        }
        let active = ActiveLevel {
            depth,
            value,
            child_order: 0,
        };
        if empty {
            finish_level(outline, active)?;
        } else {
            *level = Some(active);
        }
        return Ok(());
    }

    let active_level = level.as_mut().expect("level exists");
    if depth != active_level.depth + 1 {
        return invalid("outline level has invalid nested content");
    }
    if current.0 == NamespaceKind::Style && current.1 == b"list-level-properties" {
        if active_level.child_order != 0 || active_level.value.list_level_properties.is_some() {
            return invalid("duplicate or out-of-order style:list-level-properties");
        }
        active_level.child_order = 1;
        let value = parse_list_properties(reader, version, start, total_attribute_bytes)?;
        if empty {
            active_level.value.list_level_properties = Some(value);
        } else {
            *properties = Some(ActiveProperties {
                depth,
                value,
                alignment_depth: None,
            });
        }
        return Ok(());
    }
    if current.0 == NamespaceKind::Style && current.1 == b"text-properties" {
        if active_level.child_order > 1 || active_level.value.text_properties.is_some() {
            return invalid("duplicate or out-of-order style:text-properties");
        }
        active_level.child_order = 2;
        let attributes = attributes(reader, version, start, total_attribute_bytes)?;
        active_level.value.text_properties = Some(TextProperties {
            attributes: attributes
                .into_iter()
                .map(attribute_descriptor)
                .collect::<Result<Vec<_>>>()?,
        });
        if !empty {
            *text_properties_depth = Some(depth);
        }
        return Ok(());
    }
    let _ = (styles, names);
    invalid("text:outline-level-style has an invalid child")
}

fn parse_level(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    total: &mut usize,
) -> Result<LevelStyle> {
    let mut attributes = attributes(reader, version, start, total)?;
    let level = take(&mut attributes, NamespaceKind::Text, "level")
        .ok_or_else(|| invalid_error("outline level requires text:level"))?
        .parse::<u16>()
        .map_err(|_| invalid_error("text:level is outside the supported range"))?;
    if !(1..=MAX_OUTLINE_LEVELS).contains(&level) {
        return invalid("text:level is outside the supported range");
    }
    let text_style_name = take(&mut attributes, NamespaceKind::Text, "style-name");
    if let Some(value) = text_style_name.as_deref() {
        validate_text(value, "text:style-name", false)?;
    }
    let number_format = take(&mut attributes, NamespaceKind::Style, "num-format")
        .map(NumberFormat::new)
        .transpose()?;
    let number_prefix = take(&mut attributes, NamespaceKind::Style, "num-prefix");
    let number_suffix = take(&mut attributes, NamespaceKind::Style, "num-suffix");
    if let Some(value) = number_prefix.as_deref() {
        validate_text(value, "style:num-prefix", true)?;
    }
    if let Some(value) = number_suffix.as_deref() {
        validate_text(value, "style:num-suffix", true)?;
    }
    let letter_sync = take(&mut attributes, NamespaceKind::Style, "num-letter-sync")
        .map(|value| parse_bool(&value, "style:num-letter-sync"))
        .transpose()?;
    if letter_sync.is_some()
        && !number_format
            .as_ref()
            .is_some_and(|format| matches!(format.as_str(), "a" | "A"))
    {
        return invalid("style:num-letter-sync requires style:num-format 'a' or 'A'");
    }
    let display_levels = take(&mut attributes, NamespaceKind::Text, "display-levels")
        .map(PositiveInteger::new)
        .transpose()?;
    let start_value = take(&mut attributes, NamespaceKind::Text, "start-value")
        .map(PositiveInteger::new)
        .transpose()?;
    Ok(LevelStyle {
        level,
        text_style_name,
        number_format,
        number_prefix,
        number_suffix,
        letter_sync,
        display_levels,
        start_value,
        list_level_properties: None,
        text_properties: None,
        extensions: finish_attributes(attributes, "text:outline-level-style")?,
    })
}

fn parse_list_properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    total: &mut usize,
) -> Result<ListProperties> {
    let mut attributes = attributes(reader, version, start, total)?;
    let text_align = take(&mut attributes, NamespaceKind::Fo, "text-align")
        .map(|value| TextAlign::parse(&value))
        .transpose()?;
    let space_before = length(
        &mut attributes,
        NamespaceKind::Text,
        "space-before",
        false,
        false,
    )?;
    let minimum_label_width = length(
        &mut attributes,
        NamespaceKind::Text,
        "min-label-width",
        true,
        false,
    )?;
    let minimum_label_distance = length(
        &mut attributes,
        NamespaceKind::Text,
        "min-label-distance",
        true,
        false,
    )?;
    let font_name = take(&mut attributes, NamespaceKind::Style, "font-name");
    if let Some(value) = font_name.as_deref() {
        validate_text(value, "style:font-name", true)?;
    }
    let width = length(&mut attributes, NamespaceKind::Fo, "width", true, true)?;
    let height = length(&mut attributes, NamespaceKind::Fo, "height", true, true)?;
    let vertical_relation = take(&mut attributes, NamespaceKind::Style, "vertical-rel");
    let vertical_position = take(&mut attributes, NamespaceKind::Style, "vertical-pos");
    for (value, name) in [
        (vertical_relation.as_deref(), "style:vertical-rel"),
        (vertical_position.as_deref(), "style:vertical-pos"),
    ] {
        if let Some(value) = value {
            validate_text(value, name, false)?;
        }
    }
    let position_mode = take(
        &mut attributes,
        NamespaceKind::Text,
        "list-level-position-and-space-mode",
    )
    .map(|value| PositionMode::parse(&value))
    .transpose()?;
    Ok(ListProperties {
        text_align,
        space_before,
        minimum_label_width,
        minimum_label_distance,
        font_name,
        width,
        height,
        vertical_relation,
        vertical_position,
        position_mode,
        label_alignment: None,
        extensions: finish_attributes(attributes, "style:list-level-properties")?,
    })
}

fn parse_alignment(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    total: &mut usize,
) -> Result<Alignment> {
    let mut attributes = attributes(reader, version, start, total)?;
    let followed_by = take(&mut attributes, NamespaceKind::Text, "label-followed-by")
        .ok_or_else(|| invalid_error("label alignment requires text:label-followed-by"))?;
    let label_followed_by = match followed_by.as_str() {
        "listtab" => FollowedBy::ListTab,
        "space" => FollowedBy::Space,
        "nothing" => FollowedBy::Nothing,
        _ => return invalid("unsupported text:label-followed-by"),
    };
    let alignment = Alignment {
        label_followed_by,
        list_tab_stop_position: take(
            &mut attributes,
            NamespaceKind::Text,
            "list-tab-stop-position",
        )
        .map(Length::new)
        .transpose()?,
        text_indent: take(&mut attributes, NamespaceKind::Fo, "text-indent")
            .map(Length::new)
            .transpose()?,
        margin_left: take(&mut attributes, NamespaceKind::Fo, "margin-left")
            .map(Length::new)
            .transpose()?,
    };
    if !attributes.is_empty() {
        return invalid("unknown style:list-level-label-alignment attribute");
    }
    alignment.validate()?;
    Ok(alignment)
}

fn finish_level(outline: &mut Option<ActiveOutline>, level: ActiveLevel) -> Result<()> {
    if let Some(properties) = level.value.list_level_properties.as_ref()
        && properties.position_mode == Some(PositionMode::LabelAlignment)
        && properties.label_alignment.is_none()
    {
        return invalid("label-alignment mode requires style:list-level-label-alignment");
    }
    outline
        .as_mut()
        .expect("outline owns level")
        .value
        .levels
        .push(level.value);
    Ok(())
}

fn finish_outline(
    styles: &mut Vec<Style>,
    names: &mut HashSet<String>,
    outline: ActiveOutline,
) -> Result<()> {
    if outline.value.levels.is_empty() {
        return invalid("text:outline-style requires at least one level");
    }
    if styles.len() >= MAX_STYLES || !names.insert(outline.value.name.clone()) {
        return invalid("duplicate or excessive outline style");
    }
    styles.push(outline.value);
    Ok(())
}

fn attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    total: &mut usize,
) -> Result<Vec<ResolvedAttribute>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid_error(format!("invalid outline attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_bytes(namespace);
        let local_name = String::from_utf8(local.as_ref().to_vec())
            .map_err(|_| invalid_error("outline attribute name is not UTF-8"))?;
        if !seen.insert((namespace_uri.clone(), local_name.clone())) {
            return invalid("duplicate expanded outline attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| invalid_error(format!("invalid outline attribute value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("outline attribute value exceeds the resource limit");
        }
        *total = total
            .checked_add(namespace_uri.len() + local_name.len() + value.len())
            .ok_or_else(|| invalid_error("outline attribute size overflow"))?;
        if *total > MAX_TOTAL_ATTRIBUTE_BYTES {
            return invalid("outline attributes exceed the aggregate resource limit");
        }
        output.push(ResolvedAttribute {
            namespace: namespace_kind(&namespace_uri),
            namespace_uri: String::from_utf8(namespace_uri)
                .map_err(|_| invalid_error("outline attribute namespace is not UTF-8"))?,
            local_name,
            value,
        });
    }
    Ok(output)
}

fn finish_attributes(attributes: Vec<ResolvedAttribute>, element: &str) -> Result<Vec<Attribute>> {
    attributes
        .into_iter()
        .map(|attribute| {
            if attribute.namespace != NamespaceKind::Other {
                return invalid(format!("unknown {element} attribute"));
            }
            attribute_descriptor(attribute)
        })
        .collect()
}

fn attribute_descriptor(attribute: ResolvedAttribute) -> Result<Attribute> {
    if attribute.namespace_uri.is_empty() {
        return invalid("unqualified outline formatting attributes are not allowed");
    }
    Ok(Attribute {
        namespace_uri: attribute.namespace_uri,
        local_name: attribute.local_name,
        value: attribute.value,
    })
}

fn take(
    attributes: &mut Vec<ResolvedAttribute>,
    namespace: NamespaceKind,
    local_name: &str,
) -> Option<String> {
    attributes
        .iter()
        .position(|attribute| {
            attribute.namespace == namespace && attribute.local_name == local_name
        })
        .map(|index| attributes.remove(index).value)
}

fn length(
    attributes: &mut Vec<ResolvedAttribute>,
    namespace: NamespaceKind,
    local_name: &str,
    nonnegative: bool,
    positive: bool,
) -> Result<Option<Length>> {
    let Some(value) = take(attributes, namespace, local_name) else {
        return Ok(None);
    };
    if nonnegative && value.starts_with('-') {
        return invalid(format!("{local_name} cannot be negative"));
    }
    if positive && length_is_zero(&value) {
        return invalid(format!("{local_name} must be positive"));
    }
    Length::new(value).map(Some)
}

fn length_is_zero(value: &str) -> bool {
    let number = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
        .unwrap_or(value);
    number
        .bytes()
        .all(|byte| byte == b'0' || byte == b'.' || byte == b'+')
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("{name} must be an XML Schema boolean")),
    }
}

fn resolve_element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (NamespaceKind, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (
        namespace_kind(&namespace_bytes(namespace)),
        local.as_ref().to_vec(),
    )
}

fn namespace_bytes(namespace: ResolveResult<'_>) -> Vec<u8> {
    match namespace {
        ResolveResult::Bound(namespace) => namespace.as_ref().to_vec(),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Vec::new(),
    }
}
