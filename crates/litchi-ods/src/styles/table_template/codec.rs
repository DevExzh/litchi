//! XML codec for ODF table-style templates.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{NamespaceResolver, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

use super::semantic::{Axis, Region, Style, Template};
use super::validation::{MAX_AGGREGATE_BYTES, MAX_EXTENSION_DEPTH, MAX_TEMPLATES, MAX_VALUE_BYTES};

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";

impl Template {
    /// Append deterministic ODF XML for this template to an existing buffer.
    pub fn write_xml(&self, output: &mut String) -> Result<()> {
        self.validate()?;
        output.push_str("<table:table-template table:name=\"");
        output.push_str(&escape_xml(&self.name));
        output.push('"');
        write_axis_attribute(
            output,
            "first-row-start-column",
            self.first_row_start_column,
        );
        write_axis_attribute(output, "first-row-end-column", self.first_row_end_column);
        write_axis_attribute(output, "last-row-start-column", self.last_row_start_column);
        write_axis_attribute(output, "last-row-end-column", self.last_row_end_column);
        write_bool_attribute(output, "use-first-row-styles", self.use_first_row_styles);
        write_bool_attribute(output, "use-last-row-styles", self.use_last_row_styles);
        write_bool_attribute(
            output,
            "use-first-column-styles",
            self.use_first_column_styles,
        );
        write_bool_attribute(
            output,
            "use-last-column-styles",
            self.use_last_column_styles,
        );
        write_bool_attribute(
            output,
            "use-banding-rows-styles",
            self.use_banding_rows_styles,
        );
        write_bool_attribute(
            output,
            "use-banding-columns-styles",
            self.use_banding_columns_styles,
        );
        output.push('>');
        for region in Region::ALL {
            if let Some(style) = self.region(region) {
                write_region(output, region.name(), style);
            }
        }
        output.push_str("</table:table-template>");
        Ok(())
    }

    /// Serialize this template as a standalone ODF XML fragment.
    pub fn to_xml(&self) -> Result<String> {
        let mut output = String::new();
        self.write_xml(&mut output)?;
        Ok(output)
    }
}

fn write_axis_attribute(output: &mut String, name: &str, value: Option<Axis>) {
    let Some(value) = value else { return };
    output.push_str(" table:");
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(match value {
        Axis::Row => "row",
        Axis::Column => "column",
    });
    output.push('"');
}

fn write_bool_attribute(output: &mut String, name: &str, value: Option<bool>) {
    let Some(value) = value else { return };
    output.push_str(" table:");
    output.push_str(name);
    output.push_str(if value { "=\"true\"" } else { "=\"false\"" });
}

fn write_region(output: &mut String, name: &str, style: &Style) {
    output.push_str("<table:");
    output.push_str(name);
    output.push_str(" table:style-name=\"");
    output.push_str(&escape_xml(&style.style_name));
    output.push('"');
    if let Some(paragraph) = &style.paragraph_style_name {
        output.push_str(" table:paragraph-style-name=\"");
        output.push_str(&escape_xml(paragraph));
        output.push('"');
    }
    output.push_str("/>");
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Namespace {
    None,
    Office,
    Table,
    Text,
    Other,
}

#[derive(Debug)]
struct Attribute {
    namespace: Namespace,
    local: String,
    value: String,
}

/// Parse table-template elements from one ODF styles XML part.
pub fn parse(xml: &str) -> Result<Vec<Template>> {
    parse_parts(&[xml])
}

/// Parse table-template elements from multiple ODF styles XML parts.
pub fn parse_parts(parts: &[&str]) -> Result<Vec<Template>> {
    let mut templates = Vec::new();
    let mut names = HashSet::new();
    let mut aggregate = 0usize;
    for xml in parts {
        parse_part(xml, &mut templates, &mut names, &mut aggregate)?;
    }
    Ok(templates)
}

fn parse_part(
    xml: &str,
    templates: &mut Vec<Template>,
    names: &mut HashSet<String>,
    aggregate: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut styles_depth = None;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&resolved)?;
        let event = event.into_owned();
        let is_start = matches!(&event, Event::Start(_));
        let is_end = matches!(&event, Event::End(_));
        let mut consumed = false;

        if let Event::Start(element) = &event
            && namespace == Namespace::Office
            && matches!(
                element.local_name().as_ref(),
                b"styles" | b"automatic-styles"
            )
        {
            styles_depth = Some(depth);
        }
        let direct_style_child = styles_depth.is_some_and(|value| depth == value + 1);

        match event {
            Event::Start(element) if element.local_name().as_ref() == b"table-template" => {
                ensure_template_name(namespace, direct_style_child)?;
                if templates.len() >= MAX_TEMPLATES {
                    return Err(Error::InvalidFormat(format!(
                        "document exceeds {MAX_TEMPLATES} table templates"
                    )));
                }
                let template = parse_template(&mut reader, &element, aggregate)?;
                insert_template(templates, names, template)?;
                consumed = true;
            },
            Event::Empty(element) if element.local_name().as_ref() == b"table-template" => {
                ensure_template_name(namespace, direct_style_child)?;
                let template = parse_empty_template(&reader, &element, aggregate)?;
                insert_template(templates, names, template)?;
            },
            Event::End(element)
                if namespace == Namespace::Office
                    && matches!(
                        element.local_name().as_ref(),
                        b"styles" | b"automatic-styles"
                    ) =>
            {
                styles_depth = None;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in table templates"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }

        if is_start && !consumed {
            depth = depth.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("table-template XML depth overflow".to_string())
            })?;
        } else if is_end {
            depth = depth.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat("table-template XML depth underflow".to_string())
            })?;
        }
        buffer.clear();
    }
    Ok(())
}

fn ensure_template_name(namespace: Namespace, direct_style_child: bool) -> Result<()> {
    if namespace != Namespace::Table {
        return Err(Error::InvalidFormat(
            "table-template vocabulary uses the wrong namespace".to_string(),
        ));
    }
    if !direct_style_child {
        return Err(Error::InvalidFormat(
            "table:table-template must be a direct style collection child".to_string(),
        ));
    }
    Ok(())
}

fn insert_template(
    templates: &mut Vec<Template>,
    names: &mut HashSet<String>,
    template: Template,
) -> Result<()> {
    if !names.insert(template.name.clone()) {
        return Err(Error::InvalidFormat(format!(
            "duplicate table template '{}'",
            template.name
        )));
    }
    templates.push(template);
    Ok(())
}

fn parse_empty_template(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Template> {
    let attributes = parse_attributes(reader.resolver(), reader.decoder(), start, aggregate)?;
    build_template(&attributes, TemplateRegions::default())
}

fn parse_template(
    reader: &mut NsReader<&[u8]>,
    start: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Template> {
    let attributes = parse_attributes(reader.resolver(), reader.decoder(), start, aggregate)?;
    let mut regions = TemplateRegions::default();
    let mut buffer = Vec::new();
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&resolved)?;
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                let local = decode_name(element.local_name().as_ref())?;
                if is_region_name(&local) && namespace != Namespace::Table {
                    return Err(Error::InvalidFormat(format!(
                        "table-template region '{local}' uses the wrong namespace"
                    )));
                }
                if namespace == Namespace::Table && is_region_name(&local) {
                    let style = parse_region_attributes(
                        reader.resolver(),
                        reader.decoder(),
                        &element,
                        local == "background",
                        aggregate,
                    )?;
                    consume_empty_region(reader, &local)?;
                    regions.set(&local, style)?;
                } else if namespace == Namespace::Other {
                    skip_extension(reader)?;
                } else {
                    return Err(Error::InvalidFormat(format!(
                        "unexpected table-template child '{local}'"
                    )));
                }
            },
            Event::Empty(element) => {
                let local = decode_name(element.local_name().as_ref())?;
                if is_region_name(&local) && namespace != Namespace::Table {
                    return Err(Error::InvalidFormat(format!(
                        "table-template region '{local}' uses the wrong namespace"
                    )));
                }
                if namespace == Namespace::Table && is_region_name(&local) {
                    let style = parse_region_attributes(
                        reader.resolver(),
                        reader.decoder(),
                        &element,
                        local == "background",
                        aggregate,
                    )?;
                    regions.set(&local, style)?;
                } else if namespace != Namespace::Other {
                    return Err(Error::InvalidFormat(format!(
                        "unexpected table-template child '{local}'"
                    )));
                }
            },
            Event::Text(text) => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid table-template text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:table-template cannot contain character data".to_string(),
                    ));
                }
            },
            Event::End(element) if element.local_name().as_ref() == b"table-template" => break,
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) | Event::CData(_) => {
                return Err(Error::InvalidFormat(
                    "invalid content in table:table-template".to_string(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated table:table-template".to_string(),
                ));
            },
            Event::Comment(_) | Event::Decl(_) | Event::End(_) => {},
        }
        buffer.clear();
    }
    build_template(&attributes, regions)
}

#[derive(Default)]
struct TemplateRegions {
    first_row: Option<Style>,
    last_row: Option<Style>,
    first_column: Option<Style>,
    last_column: Option<Style>,
    body: Option<Style>,
    even_rows: Option<Style>,
    odd_rows: Option<Style>,
    even_columns: Option<Style>,
    odd_columns: Option<Style>,
    background: Option<Style>,
}

impl TemplateRegions {
    fn set(&mut self, local: &str, style: Style) -> Result<()> {
        let slot = match local {
            "first-row" => &mut self.first_row,
            "last-row" => &mut self.last_row,
            "first-column" => &mut self.first_column,
            "last-column" => &mut self.last_column,
            "body" => &mut self.body,
            "even-rows" => &mut self.even_rows,
            "odd-rows" => &mut self.odd_rows,
            "even-columns" => &mut self.even_columns,
            "odd-columns" => &mut self.odd_columns,
            "background" => &mut self.background,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "unknown table-template region '{local}'"
                )));
            },
        };
        if slot.replace(style).is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate table-template region '{local}'"
            )));
        }
        Ok(())
    }
}

fn build_template(attributes: &[Attribute], regions: TemplateRegions) -> Result<Template> {
    reject_template_attributes(attributes)?;
    let name = required_attribute_either(attributes, "name")?.to_string();
    if name.is_empty() {
        return Err(Error::InvalidFormat(
            "table template name must not be empty".to_string(),
        ));
    }
    let template = Template {
        name,
        first_row_start_column: parse_axis_attribute(attributes, "first-row-start-column")?,
        first_row_end_column: parse_axis_attribute(attributes, "first-row-end-column")?,
        last_row_start_column: parse_axis_attribute(attributes, "last-row-start-column")?,
        last_row_end_column: parse_axis_attribute(attributes, "last-row-end-column")?,
        use_first_row_styles: parse_bool_attribute(attributes, "use-first-row-styles")?,
        use_last_row_styles: parse_bool_attribute(attributes, "use-last-row-styles")?,
        use_first_column_styles: parse_bool_attribute(attributes, "use-first-column-styles")?,
        use_last_column_styles: parse_bool_attribute(attributes, "use-last-column-styles")?,
        use_banding_rows_styles: parse_bool_attribute(attributes, "use-banding-rows-styles")?,
        use_banding_columns_styles: parse_bool_attribute(attributes, "use-banding-columns-styles")?,
        first_row: regions.first_row,
        last_row: regions.last_row,
        first_column: regions.first_column,
        last_column: regions.last_column,
        body: regions.body,
        even_rows: regions.even_rows,
        odd_rows: regions.odd_rows,
        even_columns: regions.even_columns,
        odd_columns: regions.odd_columns,
        background: regions.background,
    };
    template.validate()?;
    Ok(template)
}

fn parse_region_attributes(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    start: &BytesStart<'_>,
    background: bool,
    aggregate: &mut usize,
) -> Result<Style> {
    let attributes = parse_attributes(resolver, decoder, start, aggregate)?;
    for attribute in &attributes {
        let allowed = attribute.namespace == Namespace::Table
            && (attribute.local == "style-name"
                || !background && attribute.local == "paragraph-style-name");
        if is_known(attribute.namespace) && !allowed {
            return Err(Error::InvalidFormat(format!(
                "unexpected table-template region attribute '{}'",
                attribute.local
            )));
        }
    }
    let style_name = attribute(&attributes, Namespace::Table, "style-name")
        .ok_or_else(|| Error::InvalidFormat("table-template region requires style-name".into()))?
        .to_string();
    if style_name.is_empty() {
        return Err(Error::InvalidFormat(
            "table-template style name must not be empty".to_string(),
        ));
    }
    Ok(Style {
        style_name,
        paragraph_style_name: attribute(&attributes, Namespace::Table, "paragraph-style-name")
            .map(str::to_string),
    })
}

fn parse_attributes(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    start: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Vec<Attribute>> {
    let mut attributes = Vec::new();
    for source in start.attributes() {
        let source = source.map_err(|error| {
            Error::InvalidFormat(format!("invalid table-template attribute: {error}"))
        })?;
        if source.key.as_ref() == b"xmlns" || source.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = resolver.resolve_attribute(source.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode_name(local.as_ref())?;
        let value = source
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid table-template attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return Err(Error::InvalidFormat(
                "table-template attribute exceeds 64 KiB".to_string(),
            ));
        }
        append_size(aggregate, local.len().saturating_add(value.len()))?;
        if attributes
            .iter()
            .any(|existing: &Attribute| existing.namespace == namespace && existing.local == local)
        {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded table-template attribute '{local}'"
            )));
        }
        attributes.push(Attribute {
            namespace,
            local,
            value,
        });
    }
    Ok(attributes)
}

fn reject_template_attributes(attributes: &[Attribute]) -> Result<()> {
    const ALLOWED: &[&str] = &[
        "name",
        "first-row-start-column",
        "first-row-end-column",
        "last-row-start-column",
        "last-row-end-column",
        "use-first-row-styles",
        "use-last-row-styles",
        "use-first-column-styles",
        "use-last-column-styles",
        "use-banding-rows-styles",
        "use-banding-columns-styles",
    ];
    for attribute in attributes {
        let allowed_namespace = attribute.namespace == Namespace::Table
            || attribute.namespace == Namespace::Text
                && matches!(
                    attribute.local.as_str(),
                    "name"
                        | "first-row-start-column"
                        | "first-row-end-column"
                        | "last-row-start-column"
                        | "last-row-end-column"
                );
        if is_known(attribute.namespace)
            && (!allowed_namespace || !ALLOWED.contains(&attribute.local.as_str()))
        {
            return Err(Error::InvalidFormat(format!(
                "unexpected table-template attribute '{}'",
                attribute.local
            )));
        }
    }
    Ok(())
}

fn required_attribute_either<'a>(attributes: &'a [Attribute], local: &str) -> Result<&'a str> {
    let table = attribute(attributes, Namespace::Table, local);
    let legacy = attribute(attributes, Namespace::Text, local);
    match (table, legacy) {
        (Some(_), Some(_)) => Err(Error::InvalidFormat(format!(
            "duplicate table-template attribute '{local}' across namespaces"
        ))),
        (Some(value), None) | (None, Some(value)) => Ok(value),
        (None, None) => Err(Error::InvalidFormat(format!(
            "table template requires '{local}'"
        ))),
    }
}

fn parse_axis_attribute(attributes: &[Attribute], local: &str) -> Result<Option<Axis>> {
    let table = attribute(attributes, Namespace::Table, local);
    let legacy = attribute(attributes, Namespace::Text, local);
    let value = match (table, legacy) {
        (Some(_), Some(_)) => {
            return Err(Error::InvalidFormat(format!(
                "duplicate table-template attribute '{local}' across namespaces"
            )));
        },
        (Some(value), None) | (None, Some(value)) => Some(value),
        (None, None) => None,
    };
    value
        .map(|value| match value {
            "row" => Ok(Axis::Row),
            "column" => Ok(Axis::Column),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table-template axis '{value}'"
            ))),
        })
        .transpose()
}

fn parse_bool_attribute(attributes: &[Attribute], local: &str) -> Result<Option<bool>> {
    attribute(attributes, Namespace::Table, local)
        .map(|value| match value {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table-template boolean '{value}'"
            ))),
        })
        .transpose()
}

fn attribute<'a>(
    attributes: &'a [Attribute],
    namespace: Namespace,
    local: &str,
) -> Option<&'a str> {
    attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.local == local)
        .map(|attribute| attribute.value.as_str())
}

fn consume_empty_region(reader: &mut NsReader<&[u8]>, local: &str) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::End(element) if element.local_name().as_ref() == local.as_bytes() => {
                return Ok(());
            },
            Event::Text(text) => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid table-template region text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table-template regions must be empty".to_string(),
                    ));
                }
            },
            Event::Comment(_) => {},
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated table-template region".to_string(),
                ));
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "table-template regions must be empty".to_string(),
                ));
            },
        }
        buffer.clear();
    }
}

fn skip_extension(reader: &mut NsReader<&[u8]>) -> Result<()> {
    let mut buffer = Vec::new();
    let mut depth = 1usize;
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("extension element depth overflow".to_string())
                })?;
                if depth > MAX_EXTENSION_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "extension element exceeds depth {MAX_EXTENSION_DEPTH}"
                    )));
                }
            },
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "DTDs and processing instructions are prohibited in extensions".to_string(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "unterminated table-template extension".to_string(),
                ));
            },
            _ => {},
        }
        buffer.clear();
    }
}

fn is_region_name(local: &str) -> bool {
    Region::from_name(local).is_some()
}

fn append_size(aggregate: &mut usize, amount: usize) -> Result<()> {
    *aggregate = aggregate.checked_add(amount).ok_or_else(|| {
        Error::InvalidFormat("table-template aggregate size overflow".to_string())
    })?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return Err(Error::InvalidFormat(
            "table-template metadata exceeds 16 MiB".to_string(),
        ));
    }
    Ok(())
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> Result<Namespace> {
    match namespace {
        ResolveResult::Unbound => Ok(Namespace::None),
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NS => Ok(Namespace::Office),
        ResolveResult::Bound(value) if value.as_ref() == TABLE_NS => Ok(Namespace::Table),
        ResolveResult::Bound(value) if value.as_ref() == TEXT_NS => Ok(Namespace::Text),
        ResolveResult::Bound(_) => Ok(Namespace::Other),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound table-template namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn is_known(namespace: Namespace) -> bool {
    matches!(
        namespace,
        Namespace::Office | Namespace::Table | Namespace::Text
    )
}

fn decode_name(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat("invalid UTF-8 table-template name".to_string()))
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("invalid table-template XML: {error}"))
}
