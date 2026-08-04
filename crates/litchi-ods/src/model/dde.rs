//! Inert Dynamic Data Exchange links and their cached spreadsheet tables.

use super::{
    Cell, Row, Sheet,
    parser::OdsParser,
    protection::{parse_protection, write_sheet_attributes, write_sheet_options},
    scenario::write_sheet_preamble,
    structure::{
        TableStructureAxis, write_columns, write_row_attributes, write_sheet_formatting_attributes,
        write_table_structure,
    },
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::collections::BTreeMap;

const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_DDE_LINKS: usize = 65_536;
const MAX_DDE_SOURCE_VALUE_BYTES: usize = 65_536;
const MAX_DDE_SOURCE_TOTAL_BYTES: usize = 262_144;

/// How an application converts values obtained from a DDE source.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DdeConversionMode {
    /// Apply the default cell style and data style.
    IntoDefaultStyleDataStyle,
    /// Convert numeric text using English number syntax.
    IntoEnglishNumber,
    /// Preserve all incoming values as text.
    KeepText,
}

impl DdeConversionMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "into-default-style-data-style" => Ok(Self::IntoDefaultStyleDataStyle),
            "into-english-number" => Ok(Self::IntoEnglishNumber),
            "keep-text" => Ok(Self::KeepText),
            _ => Err(Error::InvalidFormat(format!(
                "invalid office:conversion-mode '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::IntoDefaultStyleDataStyle => "into-default-style-data-style",
            Self::IntoEnglishNumber => "into-english-number",
            Self::KeepText => "keep-text",
        }
    }
}

/// The non-executing source declaration for an OpenDocument DDE link.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DdeSource {
    /// DDE server/application name.
    pub application: String,
    /// DDE topic, commonly a document or resource name.
    pub topic: String,
    /// DDE item within the topic.
    pub item: String,
    /// Optional declaration name.
    pub name: Option<String>,
    /// Optional conversion policy for cached values.
    pub conversion_mode: Option<DdeConversionMode>,
    /// Whether an application requested automatic refresh.
    ///
    /// Litchi never performs that refresh; the flag is retained as inert metadata.
    pub automatic_update: Option<bool>,
}

impl DdeSource {
    /// Create a DDE source declaration. This never opens or executes the source.
    pub fn new(
        application: impl Into<String>,
        topic: impl Into<String>,
        item: impl Into<String>,
    ) -> Self {
        Self {
            application: application.into(),
            topic: topic.into(),
            item: item.into(),
            name: None,
            conversion_mode: None,
            automatic_update: None,
        }
    }

    /// Validate the source declaration without contacting it.
    pub fn validate(&self) -> Result<()> {
        let values = [
            ("office:dde-application", self.application.as_str(), true),
            ("office:dde-topic", self.topic.as_str(), true),
            ("office:dde-item", self.item.as_str(), true),
        ];
        let mut total = 0usize;
        for (name, value, required) in values {
            if (required && value.is_empty())
                || value.len() > MAX_DDE_SOURCE_VALUE_BYTES
                || value.chars().any(|character| {
                    character == '\0'
                        || (character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
                })
            {
                return Err(Error::InvalidFormat(format!("invalid or oversized {name}")));
            }
            total = total
                .checked_add(value.len())
                .ok_or_else(|| Error::InvalidFormat("DDE source text size overflow".to_string()))?;
        }
        if let Some(name) = &self.name {
            if name.len() > MAX_DDE_SOURCE_VALUE_BYTES
                || name.chars().any(|character| {
                    character == '\0'
                        || (character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
                })
            {
                return Err(Error::InvalidFormat(
                    "invalid or oversized office:name".to_string(),
                ));
            }
            total = total
                .checked_add(name.len())
                .ok_or_else(|| Error::InvalidFormat("DDE source text size overflow".to_string()))?;
        }
        if total > MAX_DDE_SOURCE_TOTAL_BYTES {
            return Err(Error::InvalidFormat(format!(
                "DDE source text exceeds the {MAX_DDE_SOURCE_TOTAL_BYTES} byte safety limit"
            )));
        }
        Ok(())
    }
}

/// A DDE source plus the table of cached values stored in the document.
#[derive(Clone)]
pub struct DdeLink {
    /// Inert source metadata.
    pub source: DdeSource,
    /// Cached table content. Reading this table never contacts the source.
    pub cached_table: Sheet,
}

impl DdeLink {
    /// Create an inert DDE link from source metadata and cached table content.
    pub fn new(source: DdeSource, cached_table: Sheet) -> Result<Self> {
        source.validate()?;
        Ok(Self {
            source,
            cached_table,
        })
    }

    /// Validate this link without contacting or executing its source.
    pub fn validate(&self) -> Result<()> {
        self.source.validate()?;
        self.cached_table
            .rows
            .iter()
            .flat_map(|row| row.cells.iter())
            .try_for_each(Cell::validate_hyperlinks)
    }

    pub(crate) fn has_formulas(&self) -> bool {
        self.cached_table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .any(|cell| cell.formula.is_some())
    }

    pub(crate) fn has_annotations(&self) -> bool {
        self.cached_table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .any(Cell::has_annotation)
    }

    pub(crate) fn has_hyperlinks(&self) -> bool {
        self.cached_table
            .rows
            .iter()
            .flat_map(|row| row.cells.iter())
            .any(Cell::has_hyperlinks)
    }

    pub(crate) fn has_table_sources(&self) -> bool {
        self.cached_table.table_source.is_some()
            || self
                .cached_table
                .rows
                .iter()
                .flat_map(|row| &row.cells)
                .any(|cell| cell.range_source.is_some())
    }
}

#[derive(Default)]
struct DdeLinkBuilder {
    source: Option<DdeSource>,
    cached_table: Option<Sheet>,
}

impl DdeLinkBuilder {
    fn finish(self) -> Result<DdeLink> {
        let source = self.source.ok_or_else(|| {
            Error::InvalidFormat("table:dde-link requires office:dde-source".to_string())
        })?;
        let cached_table = self.cached_table.ok_or_else(|| {
            Error::InvalidFormat("table:dde-link requires a cached table:table".to_string())
        })?;
        DdeLink::new(source, cached_table)
    }
}

/// Parse spreadsheet DDE links as inert metadata and cached table content.
pub(crate) fn parse_dde_links(xml: &str) -> Result<Vec<DdeLink>> {
    let mut reader = Reader::from_str(xml);
    let mut buffer = Vec::new();
    let mut namespaces = BTreeMap::new();
    let mut namespace_scopes = Vec::new();
    let mut element_depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut links_depth = None;
    let mut links_seen = false;
    let mut links_count = 0usize;
    let mut link_depth = None;
    let mut source_depth = None;
    let mut current_link: Option<DdeLinkBuilder> = None;
    let mut captured_table_depth = None;
    let mut captured_table_start = None;
    let mut links = Vec::new();

    loop {
        let event_start = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let event_end = reader.buffer_position() as usize;

        match event {
            Event::Start(ref element) => {
                let scope = push_namespace_scope(element, reader.decoder(), &mut namespaces)?;
                namespace_scopes.push(scope);
                element_depth += 1;

                if captured_table_depth.is_some() {
                    // The normal sheet parser validates the captured table subtree.
                } else if element_name_is(element, &namespaces, OFFICE_NAMESPACE, "spreadsheet") {
                    spreadsheet_depth = Some(element_depth);
                } else if element_name_is(element, &namespaces, TABLE_NAMESPACE, "dde-links")
                    && spreadsheet_depth.is_some_and(|depth| element_depth == depth + 1)
                {
                    if links_seen {
                        return Err(Error::InvalidFormat(
                            "duplicate table:dde-links element".to_string(),
                        ));
                    }
                    links_seen = true;
                    links_depth = Some(element_depth);
                } else if element_name_is(element, &namespaces, TABLE_NAMESPACE, "dde-link")
                    && links_depth.is_some_and(|depth| element_depth == depth + 1)
                {
                    if current_link.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested table:dde-link element".to_string(),
                        ));
                    }
                    current_link = Some(DdeLinkBuilder::default());
                    link_depth = Some(element_depth);
                } else if let Some(builder) = current_link.as_mut()
                    && link_depth.is_some_and(|depth| element_depth == depth + 1)
                    && element_name_is(element, &namespaces, OFFICE_NAMESPACE, "dde-source")
                {
                    if builder.source.is_some() || builder.cached_table.is_some() {
                        return Err(Error::InvalidFormat(
                            "office:dde-source must be the first child of table:dde-link"
                                .to_string(),
                        ));
                    }
                    builder.source = Some(parse_source(element, reader.decoder(), &namespaces)?);
                    source_depth = Some(element_depth);
                } else if let Some(builder) = current_link.as_mut()
                    && link_depth.is_some_and(|depth| element_depth == depth + 1)
                    && qualified_name_is(
                        element.name().as_ref(),
                        &namespaces,
                        TABLE_NAMESPACE,
                        "table",
                    )
                {
                    if builder.source.is_none() || builder.cached_table.is_some() {
                        return Err(Error::InvalidFormat(
                            "cached table:table must follow exactly one office:dde-source"
                                .to_string(),
                        ));
                    }
                    captured_table_depth = Some(element_depth);
                    captured_table_start = Some(event_start);
                } else if source_depth.is_some() || current_link.is_some() {
                    return Err(Error::InvalidFormat(
                        "invalid child element in table:dde-link".to_string(),
                    ));
                } else if links_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "table:dde-links may contain only table:dde-link elements".to_string(),
                    ));
                }
            },
            Event::Empty(ref element) => {
                let scope = push_namespace_scope(element, reader.decoder(), &mut namespaces)?;
                if captured_table_depth.is_some() {
                    // The normal sheet parser validates the captured table subtree.
                } else if let Some(builder) = current_link.as_mut()
                    && link_depth == Some(element_depth)
                    && qualified_name_is(
                        element.name().as_ref(),
                        &namespaces,
                        OFFICE_NAMESPACE,
                        "dde-source",
                    )
                {
                    if builder.source.is_some() || builder.cached_table.is_some() {
                        return Err(Error::InvalidFormat(
                            "office:dde-source must be the first child of table:dde-link"
                                .to_string(),
                        ));
                    }
                    builder.source = Some(parse_source(element, reader.decoder(), &namespaces)?);
                } else if let Some(builder) = current_link.as_mut()
                    && link_depth == Some(element_depth)
                    && element_name_is(element, &namespaces, TABLE_NAMESPACE, "table")
                {
                    if builder.source.is_none() || builder.cached_table.is_some() {
                        return Err(Error::InvalidFormat(
                            "cached table:table must follow exactly one office:dde-source"
                                .to_string(),
                        ));
                    }
                    builder.cached_table = Some(parse_cached_table(
                        &xml[event_start..event_end],
                        &namespaces,
                    )?);
                } else if element_name_is(element, &namespaces, TABLE_NAMESPACE, "dde-links")
                    && spreadsheet_depth == Some(element_depth)
                {
                    return Err(Error::InvalidFormat(
                        "table:dde-links requires at least one link".to_string(),
                    ));
                } else if element_name_is(element, &namespaces, TABLE_NAMESPACE, "dde-link")
                    && links_depth == Some(element_depth)
                {
                    return Err(Error::InvalidFormat(
                        "table:dde-link requires a source and cached table".to_string(),
                    ));
                } else if current_link.is_some() || links_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "invalid empty element in table:dde-links".to_string(),
                    ));
                }
                pop_namespace_scope(&mut namespaces, Some(scope));
            },
            Event::End(ref element) => {
                if captured_table_depth == Some(element_depth)
                    && qualified_name_is(
                        element.name().as_ref(),
                        &namespaces,
                        TABLE_NAMESPACE,
                        "table",
                    )
                {
                    let start = captured_table_start.take().expect("captured table start");
                    let table = parse_cached_table(&xml[start..event_end], &namespaces)?;
                    current_link
                        .as_mut()
                        .expect("captured table belongs to a DDE link")
                        .cached_table = Some(table);
                    captured_table_depth = None;
                } else if source_depth == Some(element_depth)
                    && qualified_name_is(
                        element.name().as_ref(),
                        &namespaces,
                        OFFICE_NAMESPACE,
                        "dde-source",
                    )
                {
                    source_depth = None;
                } else if link_depth == Some(element_depth)
                    && qualified_name_is(
                        element.name().as_ref(),
                        &namespaces,
                        TABLE_NAMESPACE,
                        "dde-link",
                    )
                {
                    links.push(current_link.take().expect("checked link").finish()?);
                    links_count += 1;
                    if links_count > MAX_DDE_LINKS {
                        return Err(Error::InvalidFormat(format!(
                            "DDE link count exceeds the {MAX_DDE_LINKS} link safety limit"
                        )));
                    }
                    link_depth = None;
                } else if links_depth == Some(element_depth)
                    && qualified_name_is(
                        element.name().as_ref(),
                        &namespaces,
                        TABLE_NAMESPACE,
                        "dde-links",
                    )
                {
                    if links_count == 0 {
                        return Err(Error::InvalidFormat(
                            "table:dde-links requires at least one link".to_string(),
                        ));
                    }
                    links_depth = None;
                } else if spreadsheet_depth == Some(element_depth)
                    && qualified_name_is(
                        element.name().as_ref(),
                        &namespaces,
                        OFFICE_NAMESPACE,
                        "spreadsheet",
                    )
                {
                    spreadsheet_depth = None;
                }
                element_depth = element_depth.saturating_sub(1);
                pop_namespace_scope(&mut namespaces, namespace_scopes.pop());
            },
            Event::Text(ref text)
                if captured_table_depth.is_none()
                    && (current_link.is_some() || links_depth.is_some()) =>
            {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid DDE link text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "DDE link containers cannot contain text".to_string(),
                    ));
                }
            },
            Event::CData(ref text)
                if captured_table_depth.is_none()
                    && (current_link.is_some() || links_depth.is_some()) =>
            {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid DDE link CDATA: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "DDE link containers cannot contain CDATA".to_string(),
                    ));
                }
            },
            Event::GeneralRef(_)
                if captured_table_depth.is_none()
                    && (current_link.is_some() || links_depth.is_some()) =>
            {
                return Err(Error::InvalidFormat(
                    "DDE link containers cannot contain entity references".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if current_link.is_some()
        || link_depth.is_some()
        || source_depth.is_some()
        || captured_table_depth.is_some()
        || links_depth.is_some()
    {
        return Err(Error::InvalidFormat(
            "unterminated table:dde-links structure".to_string(),
        ));
    }
    Ok(links)
}

pub(crate) fn parse_source(
    element: &BytesStart<'_>,
    decoder: Decoder,
    namespaces: &BTreeMap<String, String>,
) -> Result<DdeSource> {
    let mut application = None;
    let mut topic = None;
    let mut item = None;
    let mut name = None;
    let mut conversion_mode = None;
    let mut automatic_update = None;

    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid DDE attribute: {error}")))?;
        let attribute_name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|_| Error::InvalidFormat("invalid UTF-8 in DDE attribute name".to_string()))?;
        if attribute_name == "xmlns" || attribute_name.starts_with("xmlns:") {
            continue;
        }
        let Some(local_name) =
            attribute_local_name(attribute.key.as_ref(), namespaces, OFFICE_NAMESPACE)
        else {
            return Err(Error::InvalidFormat(format!(
                "unsupported or spoofed office:dde-source attribute '{attribute_name}'"
            )));
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid DDE attribute: {error}")))?
            .into_owned();
        let slot = match local_name {
            "dde-application" => &mut application,
            "dde-topic" => &mut topic,
            "dde-item" => &mut item,
            "name" => &mut name,
            "conversion-mode" => {
                if conversion_mode
                    .replace(DdeConversionMode::parse(&value)?)
                    .is_some()
                {
                    return Err(Error::InvalidFormat(
                        "duplicate office:conversion-mode attribute".to_string(),
                    ));
                }
                continue;
            },
            "automatic-update" => {
                let parsed = parse_bool(&value)?;
                if automatic_update.replace(parsed).is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate office:automatic-update attribute".to_string(),
                    ));
                }
                continue;
            },
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "unsupported office:dde-source attribute 'office:{local_name}'"
                )));
            },
        };
        if slot.replace(value).is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate office:{local_name} attribute"
            )));
        }
    }

    let source = DdeSource {
        application: application.ok_or_else(|| {
            Error::InvalidFormat("office:dde-source requires office:dde-application".to_string())
        })?,
        topic: topic.ok_or_else(|| {
            Error::InvalidFormat("office:dde-source requires office:dde-topic".to_string())
        })?,
        item: item.ok_or_else(|| {
            Error::InvalidFormat("office:dde-source requires office:dde-item".to_string())
        })?,
        name,
        conversion_mode,
        automatic_update,
    };
    source.validate()?;
    Ok(source)
}

fn parse_cached_table(raw_table: &str, namespaces: &BTreeMap<String, String>) -> Result<Sheet> {
    let office_prefix = unique_prefix("litchi_office", namespaces);
    let table_prefix = unique_prefix("litchi_table", namespaces);
    let text_prefix = unique_prefix("litchi_text", namespaces);
    let mut wrapper = String::with_capacity(raw_table.len() + namespaces.len() * 64 + 256);
    wrapper.push('<');
    wrapper.push_str(&office_prefix);
    wrapper.push_str(":document-content");
    for (prefix, uri) in namespaces {
        if prefix == "xml" || prefix == "xmlns" {
            continue;
        }
        if prefix.is_empty() {
            wrapper.push_str(" xmlns=\"");
        } else {
            wrapper.push_str(" xmlns:");
            wrapper.push_str(prefix);
            wrapper.push_str("=\"");
        }
        wrapper.push_str(&escape_xml(uri));
        wrapper.push('"');
    }
    write_namespace(&mut wrapper, &office_prefix, OFFICE_NAMESPACE);
    write_namespace(&mut wrapper, &table_prefix, TABLE_NAMESPACE);
    write_namespace(&mut wrapper, &text_prefix, TEXT_NAMESPACE);
    wrapper.push('>');
    wrapper.push('<');
    wrapper.push_str(&office_prefix);
    wrapper.push_str(":body><");
    wrapper.push_str(&office_prefix);
    wrapper.push_str(":spreadsheet>");
    wrapper.push_str(raw_table);
    wrapper.push_str("</");
    wrapper.push_str(&office_prefix);
    wrapper.push_str(":spreadsheet></");
    wrapper.push_str(&office_prefix);
    wrapper.push_str(":body></");
    wrapper.push_str(&office_prefix);
    wrapper.push_str(":document-content>");

    let mut sheets = OdsParser::parse_sheets(&wrapper)?;
    if sheets.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "DDE cache must contain exactly one table, found {}",
            sheets.len()
        )));
    }
    let (_, mut protections) = parse_protection(&wrapper)?;
    if protections.len() == 1 {
        sheets[0].protection = protections.remove(0);
    }
    Ok(sheets.remove(0))
}

fn unique_prefix(base: &str, namespaces: &BTreeMap<String, String>) -> String {
    if !namespaces.contains_key(base) {
        return base.to_string();
    }
    (1usize..)
        .map(|index| format!("{base}_{index}"))
        .find(|candidate| !namespaces.contains_key(candidate))
        .expect("an unbounded suffix always yields a free XML prefix")
}

fn write_namespace(output: &mut String, prefix: &str, uri: &str) {
    output.push_str(" xmlns:");
    output.push_str(prefix);
    output.push_str("=\"");
    output.push_str(uri);
    output.push('"');
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid office:automatic-update Boolean '{value}'"
        ))),
    }
}

fn element_name_is(
    element: &BytesStart<'_>,
    namespaces: &BTreeMap<String, String>,
    namespace: &str,
    local_name: &str,
) -> bool {
    qualified_name_is(element.name().as_ref(), namespaces, namespace, local_name)
}

fn qualified_name_is(
    name: &[u8],
    namespaces: &BTreeMap<String, String>,
    namespace: &str,
    local_name: &str,
) -> bool {
    qualified_name_parts(name).is_some_and(|(prefix, local)| {
        local == local_name && namespaces.get(prefix).is_some_and(|uri| uri == namespace)
    })
}

fn attribute_local_name<'a>(
    qualified_name: &'a [u8],
    namespaces: &BTreeMap<String, String>,
    namespace: &str,
) -> Option<&'a str> {
    let (prefix, local) = qualified_name_parts(qualified_name)?;
    namespaces
        .get(prefix)
        .is_some_and(|uri| uri == namespace)
        .then_some(local)
}

fn qualified_name_parts(name: &[u8]) -> Option<(&str, &str)> {
    let name = std::str::from_utf8(name).ok()?;
    Some(name.split_once(':').unwrap_or(("", name)))
}

fn push_namespace_scope(
    element: &BytesStart<'_>,
    decoder: Decoder,
    namespaces: &mut BTreeMap<String, String>,
) -> Result<Vec<(String, Option<String>)>> {
    let mut previous_bindings = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid XML namespace declaration: {error}"))
        })?;
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(|_| {
            Error::InvalidFormat("invalid UTF-8 in XML namespace declaration".to_string())
        })?;
        let prefix = if name == "xmlns" {
            ""
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            prefix
        } else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid XML namespace URI: {error}")))?
            .into_owned();
        let prefix = prefix.to_string();
        let previous = namespaces.insert(prefix.clone(), value);
        previous_bindings.push((prefix, previous));
    }
    Ok(previous_bindings)
}

fn pop_namespace_scope(
    namespaces: &mut BTreeMap<String, String>,
    scope: Option<Vec<(String, Option<String>)>>,
) {
    let Some(scope) = scope else { return };
    for (prefix, previous) in scope.into_iter().rev() {
        if let Some(previous) = previous {
            namespaces.insert(prefix, previous);
        } else {
            namespaces.remove(&prefix);
        }
    }
}

/// Write DDE declarations and cached tables without contacting their sources.
pub(crate) fn write_dde_links(output: &mut String, links: &[DdeLink]) -> Result<()> {
    if links.is_empty() {
        return Ok(());
    }
    if links.len() > MAX_DDE_LINKS {
        return Err(Error::InvalidFormat(format!(
            "DDE link count exceeds the {MAX_DDE_LINKS} link safety limit"
        )));
    }
    output.push_str("<table:dde-links>");
    for link in links {
        link.validate()?;
        output.push_str("<table:dde-link>");
        write_dde_source(output, &link.source)?;
        write_cached_table(output, &link.cached_table)?;
        output.push_str("</table:dde-link>");
    }
    output.push_str("</table:dde-links>");
    Ok(())
}

/// Write one inert DDE source declaration without contacting it.
pub(crate) fn write_dde_source(output: &mut String, source: &DdeSource) -> Result<()> {
    source.validate()?;
    output.push_str("<office:dde-source office:dde-application=\"");
    output.push_str(&escape_xml(&source.application));
    output.push_str("\" office:dde-topic=\"");
    output.push_str(&escape_xml(&source.topic));
    output.push_str("\" office:dde-item=\"");
    output.push_str(&escape_xml(&source.item));
    output.push('"');
    write_optional_attribute(output, "office:name", source.name.as_deref());
    write_optional_attribute(
        output,
        "office:conversion-mode",
        source.conversion_mode.map(DdeConversionMode::as_str),
    );
    if let Some(value) = source.automatic_update {
        write_optional_attribute(
            output,
            "office:automatic-update",
            Some(if value { "true" } else { "false" }),
        );
    }
    output.push_str("/>");
    Ok(())
}

fn write_cached_table(output: &mut String, sheet: &Sheet) -> Result<()> {
    output.push_str("<table:table table:name=\"");
    output.push_str(&escape_xml(&sheet.name));
    output.push('"');
    write_sheet_formatting_attributes(output, &sheet.style, &sheet.print_settings)?;
    write_sheet_attributes(output, &sheet.protection);
    output.push('>');
    write_sheet_preamble(
        output,
        sheet.title.as_deref(),
        sheet.description.as_deref(),
        sheet.table_source.as_ref(),
        sheet.dde_source.as_ref(),
        sheet.scenario.as_ref(),
    )?;
    write_sheet_options(output, &sheet.protection.options);

    let total_columns = sheet_max_cols(sheet).max(sheet.columns.len()).max(1);
    write_table_structure(
        output,
        &sheet.column_structure,
        total_columns,
        TableStructureAxis::Columns,
        |out, range| {
            let explicit_end = range.end.min(sheet.columns.len());
            if range.start < explicit_end {
                write_columns(out, &sheet.columns[range.start..explicit_end]);
            }
            let default_start = range.start.max(sheet.columns.len());
            if default_start < range.end {
                write_default_columns(out, range.end - default_start);
            }
        },
    )?;
    write_table_structure(
        output,
        &sheet.row_structure,
        sheet.rows.len(),
        TableStructureAxis::Rows,
        |out, range| {
            for row in &sheet.rows[range] {
                write_row(out, row);
            }
        },
    )?;
    output.push_str("</table:table>");
    Ok(())
}

fn sheet_max_cols(sheet: &Sheet) -> usize {
    sheet
        .rows
        .iter()
        .map(|row| row.cells.len())
        .max()
        .unwrap_or(0)
}

fn write_default_columns(output: &mut String, count: usize) {
    if count <= 1 {
        output.push_str("<table:table-column/>");
    } else {
        output.push_str("<table:table-column table:number-columns-repeated=\"");
        output.push_str(&count.to_string());
        output.push_str("\"/>");
    }
}

fn write_row(output: &mut String, row: &Row) {
    output.push_str("<table:table-row");
    write_row_attributes(
        output,
        row.style_name.as_deref(),
        row.default_cell_style_name.as_deref(),
        row.visibility,
    );
    output.push('>');
    for cell in &row.cells {
        super::cell::write_cell_xml(output, cell);
    }
    output.push_str("</table:table-row>");
}

fn write_optional_attribute(output: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_namespace_aliases_and_cached_values() {
        let xml = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
          xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
          <o:body><o:spreadsheet><t:dde-links><t:dde-link>
            <o:dde-source o:dde-application="soffice&amp;1" o:dde-topic="topic"
              o:dde-item="item" o:name="Link One" o:conversion-mode="keep-text"
              o:automatic-update="1"/>
            <t:table t:name="Cache"><t:table-row><t:table-cell o:value-type="string"><x:p>cached</x:p></t:table-cell></t:table-row></t:table>
          </t:dde-link></t:dde-links></o:spreadsheet></o:body>
        </o:document-content>"#;

        let links = parse_dde_links(xml).unwrap();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].source.application, "soffice&1");
        assert_eq!(
            links[0].source.conversion_mode,
            Some(DdeConversionMode::KeepText)
        );
        assert_eq!(links[0].source.automatic_update, Some(true));
        assert_eq!(links[0].cached_table.name, "Cache");
        assert_eq!(links[0].cached_table.rows[0].cells[0].text, "cached");
    }

    #[test]
    fn rejects_invalid_shapes_and_values() {
        let wrap = |body: &str| {
            format!(
                r#"<o:spreadsheet xmlns:o="{OFFICE_NAMESPACE}" xmlns:t="{TABLE_NAMESPACE}">{body}</o:spreadsheet>"#
            )
        };
        for body in [
            "<t:dde-links/>",
            "<t:dde-links><t:dde-link/></t:dde-links>",
            "<t:dde-links><t:dde-link><t:table/></t:dde-link></t:dde-links>",
            "<t:dde-links><t:dde-link><o:dde-source o:dde-application=\"a\" o:dde-topic=\"t\"/><t:table/></t:dde-link></t:dde-links>",
            "<t:dde-links><t:dde-link><o:dde-source o:dde-application=\"a\" o:dde-topic=\"t\" o:dde-item=\"i\" o:conversion-mode=\"convert\"/><t:table/></t:dde-link></t:dde-links>",
            "<t:dde-links><t:dde-link><o:dde-source o:dde-application=\"a\" o:dde-topic=\"t\" o:dde-item=\"i\" o:automatic-update=\"yes\"/><t:table/></t:dde-link></t:dde-links>",
            "<t:dde-links>&amp;<t:dde-link><o:dde-source o:dde-application=\"a\" o:dde-topic=\"t\" o:dde-item=\"i\"/><t:table/></t:dde-link></t:dde-links>",
        ] {
            assert!(parse_dde_links(&wrap(body)).is_err(), "{body}");
        }
    }

    #[test]
    fn writer_escapes_source_and_round_trips_cache() {
        let source = DdeSource {
            application: "app&one".to_string(),
            topic: "<topic>".to_string(),
            item: "item\"one".to_string(),
            name: Some("named".to_string()),
            conversion_mode: Some(DdeConversionMode::IntoEnglishNumber),
            automatic_update: Some(false),
        };
        let cached_table = OdsParser::parse_sheets(&format!(
            r#"<o:spreadsheet xmlns:o="{OFFICE_NAMESPACE}" xmlns:t="{TABLE_NAMESPACE}"><t:table t:name="Cache"/></o:spreadsheet>"#
        ))
        .unwrap()
        .remove(0);
        let mut xml = String::new();
        write_dde_links(&mut xml, &[DdeLink::new(source, cached_table).unwrap()]).unwrap();
        assert!(xml.contains("office:dde-application=\"app&amp;one\""));
        assert!(xml.contains("office:dde-topic=\"&lt;topic&gt;\""));

        let wrapped = format!(
            r#"<o:spreadsheet xmlns:o="{OFFICE_NAMESPACE}" xmlns:office="{OFFICE_NAMESPACE}" xmlns:table="{TABLE_NAMESPACE}" xmlns:text="{TEXT_NAMESPACE}">{xml}</o:spreadsheet>"#
        );
        let parsed = parse_dde_links(&wrapped).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].source.application, "app&one");
        assert_eq!(parsed[0].cached_table.name, "Cache");
    }

    #[test]
    fn round_trips_through_builder_and_mutable_packages() {
        let cached_table = OdsParser::parse_sheets(&format!(
            r#"<o:spreadsheet xmlns:o="{OFFICE_NAMESPACE}" xmlns:t="{TABLE_NAMESPACE}" xmlns:x="{TEXT_NAMESPACE}" xmlns:l="http://www.w3.org/1999/xlink"><t:table t:name="Cache"><t:table-row><t:table-cell o:value-type="string"><x:p><x:a l:href="https://example.test/">stored</x:a> cache</x:p></t:table-cell></t:table-row></t:table></o:spreadsheet>"#
        ))
        .unwrap()
        .remove(0);
        let mut source = DdeSource::new("app", "topic", "item");
        source.automatic_update = Some(true);

        let mut builder = crate::SpreadsheetBuilder::new();
        builder.add_sheet("Visible").unwrap();
        builder
            .add_dde_link(DdeLink::new(source, cached_table).unwrap())
            .unwrap();
        let bytes = builder.build().unwrap();

        let mut spreadsheet = crate::Spreadsheet::from_bytes(bytes).unwrap();
        assert!(spreadsheet.content_xml().contains("xmlns:xlink="));
        assert_eq!(spreadsheet.dde_links().len(), 1);
        assert_eq!(
            spreadsheet.dde_links()[0].cached_table.rows[0].cells[0].text,
            "stored cache"
        );
        assert_eq!(
            spreadsheet.dde_links()[0].cached_table.rows[0].cells[0].hyperlinks()[0].range(),
            0.."stored".len()
        );
        assert_eq!(spreadsheet.sheets().unwrap().len(), 1);

        let mut mutable = crate::MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
        mutable.dde_links_mut()[0].source.item = "updated".to_string();
        let reparsed = crate::Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.dde_links()[0].source.item, "updated");
        assert_eq!(
            reparsed.dde_links()[0].cached_table.rows[0].cells[0].text,
            "stored cache"
        );
    }
}
