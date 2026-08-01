//! Typed support for the MS-XLSX Slicers part.

use crate::error::{OoxmlError, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

pub const SLICERS_CONTENT_TYPE: &str = "application/vnd.ms-excel.slicer+xml";
pub const SLICERS_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicer";

const SLICERS_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_SLICERS: usize = 65_536;
const MAX_SLICER_PARTS_PER_WORKSHEET: usize = 1024;
const MAX_DEPTH: usize = 256;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_EXTENSION_BYTES: usize = 1024 * 1024;
const MAX_TOTAL_EXTENSION_BYTES: usize = 8 * 1024 * 1024;
const MAX_EXTENSION_ATTRIBUTES: usize = 128;
const MAX_EXTENSION_ATTRIBUTE_BYTES: usize = 64 * 1024;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;

#[derive(Debug, Clone, PartialEq, Eq)]
struct XmlAttribute {
    qualified_name: String,
    value: String,
}

/// An optional `extLst` subtree retained inertly. Relationship-looking content
/// inside the subtree is never resolved by this module.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlicerExtensionList(Vec<u8>);

impl SlicerExtensionList {
    /// Constructs an extension list from a self-contained `extLst` fragment.
    pub fn new(xml: Vec<u8>) -> Result<Self> {
        validate_extension_fragment(&xml)?;
        Ok(Self(xml))
    }

    pub fn xml(&self) -> &[u8] {
        &self.0
    }
}

/// One slicer view in a worksheet Slicers part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Slicer {
    pub name: String,
    pub cache: String,
    pub caption: Option<String>,
    pub start_item: u32,
    pub column_count: u32,
    pub show_caption: bool,
    pub level: u32,
    pub style: Option<String>,
    pub locked_position: bool,
    pub row_height: u32,
    pub extension_list: Option<SlicerExtensionList>,
    xml_attributes: Vec<XmlAttribute>,
}

impl Slicer {
    pub fn new(name: impl Into<String>, cache: impl Into<String>, row_height: u32) -> Self {
        Self {
            name: name.into(),
            cache: cache.into(),
            caption: None,
            start_item: 0,
            column_count: 1,
            show_caption: true,
            level: 0,
            style: None,
            locked_position: false,
            row_height,
            extension_list: None,
            xml_attributes: Vec::new(),
        }
    }
}

/// The root value of a Slicers part. The schema requires at least one slicer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Slicers {
    pub slicers: Vec<Slicer>,
    xml_attributes: Vec<XmlAttribute>,
}

impl Slicers {
    pub fn new(slicers: Vec<Slicer>) -> Self {
        Self {
            slicers,
            xml_attributes: Vec::new(),
        }
    }
}

/// A Slicers part and its explicit relationship from one worksheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetSlicers {
    pub relationship_id: String,
    pub part_name: String,
    pub slicers: Slicers,
}

/// Parses one MS-XLSX Slicers part.
pub fn parse_slicers(xml: &[u8]) -> Result<Slicers> {
    if xml.len() > MAX_PART_BYTES {
        return Err(limit("part bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let decoder = reader.decoder();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_attributes = Vec::new();
    let mut slicers = Vec::new();
    let mut current: Option<Slicer> = None;
    let mut extension_start: Option<(usize, usize)> = None;
    let mut total_extension_bytes = 0usize;

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("Slicers XML offset overflow"))?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("Slicers XML offset overflow"))?;

        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("content after Slicers root"));
                }
                if extension_start.is_none() {
                    match depth {
                        0 => {
                            if root_seen
                                || !exact(&namespace, SLICERS_NS)
                                || element.local_name().as_ref() != b"slicers"
                            {
                                return Err(invalid("expected x14:slicers root"));
                            }
                            root_seen = true;
                            root_attributes = parse_attributes(&element, &[], decoder)?;
                        },
                        1 => {
                            if !exact(&namespace, SLICERS_NS)
                                || element.local_name().as_ref() != b"slicer"
                            {
                                return Err(invalid("unexpected child of x14:slicers"));
                            }
                            if current.is_some() {
                                return Err(invalid("nested slicer element"));
                            }
                            current = Some(parse_slicer_start(&element, decoder)?);
                        },
                        2 => {
                            if !exact(&namespace, SLICERS_NS)
                                || element.local_name().as_ref() != b"extLst"
                            {
                                return Err(invalid("unexpected child of x14:slicer"));
                            }
                            if current
                                .as_ref()
                                .is_some_and(|value| value.extension_list.is_some())
                            {
                                return Err(invalid("duplicate slicer extLst"));
                            }
                            extension_start = Some((start, depth));
                        },
                        _ => return Err(invalid("unexpected Slicers element depth")),
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| limit("XML depth"))?;
                if depth > MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("content after Slicers root"));
                }
                if extension_start.is_some() {
                    continue;
                }
                match depth {
                    0 => return Err(invalid("Slicers root cannot be empty")),
                    1 => {
                        if !exact(&namespace, SLICERS_NS)
                            || element.local_name().as_ref() != b"slicer"
                        {
                            return Err(invalid("unexpected child of x14:slicers"));
                        }
                        push_slicer(&mut slicers, parse_slicer_start(&element, decoder)?)?;
                    },
                    2 => {
                        if !exact(&namespace, SLICERS_NS)
                            || element.local_name().as_ref() != b"extLst"
                        {
                            return Err(invalid("unexpected child of x14:slicer"));
                        }
                        let current = current
                            .as_mut()
                            .ok_or_else(|| invalid("extLst outside slicer"))?;
                        if current.extension_list.is_some() {
                            return Err(invalid("duplicate slicer extLst"));
                        }
                        retain_extension(xml, start, end, &mut total_extension_bytes, current)?;
                    },
                    _ => return Err(invalid("unexpected Slicers element depth")),
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected Slicers closing element"));
                }
                depth -= 1;
                if let Some((extension_offset, parent_depth)) = extension_start {
                    if depth == parent_depth {
                        let current = current
                            .as_mut()
                            .ok_or_else(|| invalid("extLst outside slicer"))?;
                        retain_extension(
                            xml,
                            extension_offset,
                            end,
                            &mut total_extension_bytes,
                            current,
                        )?;
                        extension_start = None;
                    }
                    continue;
                }
                match depth {
                    0 => {
                        if element.local_name().as_ref() != b"slicers" || !root_seen {
                            return Err(invalid("mismatched Slicers root"));
                        }
                        root_closed = true;
                    },
                    1 => {
                        if element.local_name().as_ref() != b"slicer" {
                            return Err(invalid("mismatched slicer closing element"));
                        }
                        let value = current
                            .take()
                            .ok_or_else(|| invalid("missing slicer start"))?;
                        push_slicer(&mut slicers, value)?;
                    },
                    _ => return Err(invalid("unexpected Slicers closing depth")),
                }
            },
            Event::Text(text) if extension_start.is_none() => {
                let value = text.decode().map_err(xml_error)?;
                if !value.trim().is_empty() {
                    return Err(invalid("unexpected text in Slicers part"));
                }
            },
            Event::CData(_) if extension_start.is_none() => {
                return Err(invalid("unexpected CDATA in Slicers part"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
            _ => {},
        }
    }

    if !root_seen || !root_closed || depth != 0 || extension_start.is_some() || current.is_some() {
        return Err(invalid("incomplete Slicers XML"));
    }
    let value = Slicers {
        slicers,
        xml_attributes: root_attributes,
    };
    validate_slicers(&value)?;
    Ok(value)
}

/// Deterministically serializes one MS-XLSX Slicers part.
pub fn write_slicers(value: &Slicers) -> Result<Vec<u8>> {
    validate_slicers(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    output.extend_from_slice(b"<x14:slicers xmlns:x14=\"");
    escape(&mut output, SLICERS_NS);
    output.push(b'\"');
    write_extension_attributes(&mut output, &value.xml_attributes, true)?;
    output.push(b'>');
    for slicer in &value.slicers {
        output.extend_from_slice(b"<x14:slicer name=\"");
        escape(&mut output, &slicer.name);
        output.extend_from_slice(b"\" cache=\"");
        escape(&mut output, &slicer.cache);
        output.push(b'\"');
        if let Some(caption) = &slicer.caption {
            output.extend_from_slice(b" caption=\"");
            escape(&mut output, caption);
            output.push(b'\"');
        }
        if slicer.start_item != 0 {
            attr_u32(&mut output, "startItem", slicer.start_item);
        }
        if slicer.column_count != 1 {
            attr_u32(&mut output, "columnCount", slicer.column_count);
        }
        if !slicer.show_caption {
            output.extend_from_slice(b" showCaption=\"false\"");
        }
        if slicer.level != 0 {
            attr_u32(&mut output, "level", slicer.level);
        }
        if let Some(style) = &slicer.style {
            output.extend_from_slice(b" style=\"");
            escape(&mut output, style);
            output.push(b'\"');
        }
        if slicer.locked_position {
            output.extend_from_slice(b" lockedPosition=\"true\"");
        }
        attr_u32(&mut output, "rowHeight", slicer.row_height);
        write_extension_attributes(&mut output, &slicer.xml_attributes, false)?;
        if let Some(extension) = &slicer.extension_list {
            output.push(b'>');
            output.extend_from_slice(extension.xml());
            output.extend_from_slice(b"</x14:slicer>");
        } else {
            output.extend_from_slice(b"/>");
        }
    }
    output.extend_from_slice(b"</x14:slicers>");
    if output.len() > MAX_PART_BYTES {
        return Err(limit("serialized part bytes"));
    }
    Ok(output)
}

/// Resolves every Slicers relationship owned by `worksheet_name`, after
/// validating all MS-XLSX Slicers edges in the package.
pub fn load_worksheet_slicers(
    package: &OpcPackage,
    worksheet_name: &PackURI,
) -> Result<Vec<WorksheetSlicers>> {
    validate_package_graph(package)?;
    let worksheet = package.get_part(worksheet_name)?;
    require_worksheet(worksheet)?;
    let relationships: Vec<_> = worksheet
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == SLICERS_RELATIONSHIP_TYPE)
        .collect();
    if relationships.len() > MAX_SLICER_PARTS_PER_WORKSHEET {
        return Err(limit("parts per worksheet"));
    }
    let mut output = Vec::with_capacity(relationships.len());
    for relationship in relationships {
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        output.push(WorksheetSlicers {
            relationship_id: relationship.r_id().to_owned(),
            part_name: target.to_string(),
            slicers: parse_slicers(part.blob())?,
        });
    }
    Ok(output)
}

/// Adds one leaf Slicers part and its explicit worksheet relationship.
pub fn store_worksheet_slicers(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    value: &WorksheetSlicers,
) -> Result<()> {
    validate_package_graph(package)?;
    validate_relationship_id(&value.relationship_id)?;
    let xml = write_slicers(&value.slicers)?;
    let part_name = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    let worksheet = package.get_part(worksheet_name)?;
    require_worksheet(worksheet)?;
    let count = worksheet
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == SLICERS_RELATIONSHIP_TYPE)
        .count();
    if count >= MAX_SLICER_PARTS_PER_WORKSHEET {
        return Err(limit("parts per worksheet"));
    }
    if worksheet.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(format!(
            "worksheet relationship ID '{}' already exists",
            value.relationship_id
        )));
    }
    if package
        .iter_parts()
        .any(|part| part.partname() == &part_name)
    {
        return Err(invalid(format!(
            "Slicers part '{part_name}' already exists"
        )));
    }
    let target = part_name.relative_ref(worksheet_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(
        part_name,
        SLICERS_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(worksheet_name)?
        .rels_mut()
        .add_relationship(
            SLICERS_RELATIONSHIP_TYPE.into(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

pub(crate) fn validate_package_graph(package: &OpcPackage) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == SLICERS_RELATIONSHIP_TYPE)
    {
        return Err(invalid("package root cannot source a Slicers relationship"));
    }
    let mut targets: HashMap<String, String> = HashMap::new();
    let mut relationship_targets = HashSet::new();
    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == SLICERS_RELATIONSHIP_TYPE)
        {
            require_worksheet(source)?;
            if relationship.is_external() {
                return Err(invalid("Slicers relationship must be internal"));
            }
            let target = relationship.target_partname()?;
            let part = package.get_part(&target)?;
            if part.content_type() != SLICERS_CONTENT_TYPE {
                return Err(OoxmlError::InvalidContentType {
                    expected: SLICERS_CONTENT_TYPE.into(),
                    got: part.content_type().into(),
                });
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "Slicers part '{target}' has forbidden outbound relationships"
                )));
            }
            if let Some(previous) =
                targets.insert(target.to_string(), source.partname().to_string())
            {
                return Err(invalid(format!(
                    "Slicers part '{target}' is targeted more than once (first source '{previous}')"
                )));
            }
            relationship_targets.insert(target.to_string());
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == SLICERS_CONTENT_TYPE)
    {
        if !relationship_targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "orphan Slicers part '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn parse_slicer_start(element: &BytesStart<'_>, decoder: Decoder) -> Result<Slicer> {
    let known = [
        "name",
        "cache",
        "caption",
        "startItem",
        "columnCount",
        "showCaption",
        "level",
        "style",
        "lockedPosition",
        "rowHeight",
    ];
    let (values, xml_attributes) = parse_known_attributes(element, &known, decoder)?;
    let required = |name: &str| {
        values
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.clone())
            .ok_or_else(|| invalid(format!("slicer requires '{name}'")))
    };
    let optional = |name: &str| {
        values
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.as_str())
    };
    let name = required("name")?;
    let cache = required("cache")?;
    let row_height = parse_u32(&required("rowHeight")?, "rowHeight")?;
    let caption = optional("caption").map(str::to_owned);
    let start_item = optional("startItem")
        .map(|value| parse_u32(value, "startItem"))
        .transpose()?
        .unwrap_or(0);
    let column_count = optional("columnCount")
        .map(|value| parse_u32(value, "columnCount"))
        .transpose()?
        .unwrap_or(1);
    let show_caption = optional("showCaption")
        .map(|value| parse_bool(value, "showCaption"))
        .transpose()?
        .unwrap_or(true);
    let level = optional("level")
        .map(|value| parse_u32(value, "level"))
        .transpose()?
        .unwrap_or(0);
    let style = optional("style").map(str::to_owned);
    let locked_position = optional("lockedPosition")
        .map(|value| parse_bool(value, "lockedPosition"))
        .transpose()?
        .unwrap_or(false);
    let value = Slicer {
        name,
        cache,
        caption,
        start_item,
        column_count,
        show_caption,
        level,
        style,
        locked_position,
        row_height,
        extension_list: None,
        xml_attributes,
    };
    validate_slicer(&value)?;
    Ok(value)
}

fn parse_attributes(
    element: &BytesStart<'_>,
    known: &[&str],
    decoder: Decoder,
) -> Result<Vec<XmlAttribute>> {
    let (_, attributes) = parse_known_attributes(element, known, decoder)?;
    Ok(attributes)
}

#[allow(clippy::type_complexity)]
fn parse_known_attributes(
    element: &BytesStart<'_>,
    known: &[&str],
    decoder: Decoder,
) -> Result<(Vec<(String, String)>, Vec<XmlAttribute>)> {
    let mut values = Vec::new();
    let mut extensions = Vec::new();
    let mut extension_bytes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if value.len() > MAX_TEXT_BYTES {
            return Err(limit("attribute bytes"));
        }
        if name == "xmlns:x14" && value == SLICERS_NS {
            continue;
        }
        if !name.contains(':') && known.contains(&name.as_str()) {
            values.push((name, value));
        } else if name == "xmlns" || name.starts_with("xmlns:") || name.contains(':') {
            extension_bytes = extension_bytes
                .checked_add(name.len() + value.len())
                .ok_or_else(|| limit("extension attribute bytes"))?;
            if extensions.len() >= MAX_EXTENSION_ATTRIBUTES
                || extension_bytes > MAX_EXTENSION_ATTRIBUTE_BYTES
            {
                return Err(limit("extension attributes"));
            }
            extensions.push(XmlAttribute {
                qualified_name: name,
                value,
            });
        } else {
            return Err(invalid(format!("unexpected Slicers attribute '{name}'")));
        }
    }
    Ok((values, extensions))
}

fn retain_extension(
    xml: &[u8],
    start: usize,
    end: usize,
    total: &mut usize,
    slicer: &mut Slicer,
) -> Result<()> {
    let bytes = xml
        .get(start..end)
        .ok_or_else(|| invalid("invalid extLst XML offsets"))?;
    if bytes.len() > MAX_EXTENSION_BYTES {
        return Err(limit("extension bytes"));
    }
    *total = total
        .checked_add(bytes.len())
        .ok_or_else(|| limit("total extension bytes"))?;
    if *total > MAX_TOTAL_EXTENSION_BYTES {
        return Err(limit("total extension bytes"));
    }
    slicer.extension_list = Some(SlicerExtensionList(bytes.to_vec()));
    Ok(())
}

fn validate_extension_fragment(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_EXTENSION_BYTES {
        return Err(limit("extension bytes"));
    }
    let wrapper = [
        format!("<x14:slicer xmlns:x14=\"{SLICERS_NS}\">").as_bytes(),
        xml,
        b"</x14:slicer>",
    ]
    .concat();
    let parsed = parse_slicers(
        &[
            format!("<x14:slicers xmlns:x14=\"{SLICERS_NS}\">").as_bytes(),
            wrapper.as_slice(),
            b"</x14:slicers>",
        ]
        .concat(),
    )?;
    if parsed.slicers.len() != 1 || parsed.slicers[0].extension_list.is_none() {
        return Err(invalid("expected a single x14:extLst fragment"));
    }
    Ok(())
}

fn validate_slicers(value: &Slicers) -> Result<()> {
    if value.slicers.is_empty() {
        return Err(invalid("Slicers part requires at least one slicer"));
    }
    if value.slicers.len() > MAX_SLICERS {
        return Err(limit("slicer count"));
    }
    validate_extension_attributes(&value.xml_attributes)?;
    let mut names = HashSet::new();
    let mut total = 0usize;
    for slicer in &value.slicers {
        validate_slicer(slicer)?;
        if !names.insert(slicer.name.to_lowercase()) {
            return Err(invalid(format!("duplicate slicer name '{}'", slicer.name)));
        }
        if let Some(extension) = &slicer.extension_list {
            if extension.0.len() > MAX_EXTENSION_BYTES {
                return Err(limit("extension bytes"));
            }
            total = total
                .checked_add(extension.0.len())
                .ok_or_else(|| limit("total extension bytes"))?;
        }
    }
    if total > MAX_TOTAL_EXTENSION_BYTES {
        return Err(limit("total extension bytes"));
    }
    Ok(())
}

fn validate_slicer(value: &Slicer) -> Result<()> {
    validate_required_string(&value.name, "slicer name")?;
    if value.name.chars().count() > 32_767 {
        return Err(invalid("slicer name exceeds 32767 characters"));
    }
    validate_required_string(&value.cache, "slicer cache name")?;
    if let Some(caption) = &value.caption {
        validate_required_string(caption, "slicer caption")?;
    }
    if let Some(style) = &value.style {
        validate_string(style, "slicer style")?;
    }
    if !(1..=20_000).contains(&value.column_count) {
        return Err(invalid("slicer columnCount must be in 1..=20000"));
    }
    validate_extension_attributes(&value.xml_attributes)?;
    Ok(())
}

fn validate_extension_attributes(attributes: &[XmlAttribute]) -> Result<()> {
    if attributes.len() > MAX_EXTENSION_ATTRIBUTES {
        return Err(limit("extension attributes"));
    }
    let mut bytes = 0usize;
    let mut names = HashSet::new();
    for attribute in attributes {
        if !names.insert(attribute.qualified_name.as_str()) {
            return Err(invalid("duplicate retained XML attribute"));
        }
        if attribute.qualified_name.is_empty()
            || (!attribute.qualified_name.contains(':') && attribute.qualified_name != "xmlns")
        {
            return Err(invalid("invalid retained XML attribute name"));
        }
        validate_string(&attribute.value, "retained XML attribute")?;
        bytes = bytes
            .checked_add(attribute.qualified_name.len() + attribute.value.len())
            .ok_or_else(|| limit("extension attribute bytes"))?;
    }
    if bytes > MAX_EXTENSION_ATTRIBUTE_BYTES {
        return Err(limit("extension attribute bytes"));
    }
    Ok(())
}

fn validate_required_string(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{name} cannot be empty")));
    }
    validate_string(value, name)
}

fn validate_string(value: &str, name: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        Err(limit(name))
    } else if value.chars().any(|character| !is_xml_character(character)) {
        Err(invalid(format!("{name} contains an invalid XML character")))
    } else {
        Ok(())
    }
}

fn push_slicer(output: &mut Vec<Slicer>, value: Slicer) -> Result<()> {
    if output.len() >= MAX_SLICERS {
        return Err(limit("slicer count"));
    }
    output.push(value);
    Ok(())
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {name} '{value}'")))
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("invalid {name} '{value}'"))),
    }
}

fn write_extension_attributes(
    output: &mut Vec<u8>,
    attributes: &[XmlAttribute],
    root: bool,
) -> Result<()> {
    validate_extension_attributes(attributes)?;
    for attribute in attributes {
        if attribute.qualified_name == "xmlns:x14" {
            continue;
        }
        if !root && attribute.qualified_name == "xmlns:x14" {
            continue;
        }
        output.push(b' ');
        output.extend_from_slice(attribute.qualified_name.as_bytes());
        output.extend_from_slice(b"=\"");
        escape(output, &attribute.value);
        output.push(b'\"');
    }
    Ok(())
}

fn attr_u32(output: &mut Vec<u8>, name: &str, value: u32) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(value.to_string().as_bytes());
    output.push(b'\"');
}

fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::SML_WORKSHEET {
        Ok(())
    } else {
        Err(invalid(format!(
            "part '{}' is not a worksheet",
            part.partname()
        )))
    }
}

fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid("invalid Slicers relationship ID length"));
    }
    let mut bytes = value.bytes();
    let first = bytes.next().unwrap();
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!(
            "invalid Slicers relationship ID '{value}'"
        )))
    } else {
        Ok(())
    }
}

fn exact(namespace: &ResolveResult<'_>, value: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(namespace)) if {
        let bytes: &[u8] = namespace; bytes == value.as_bytes()
    })
}

fn is_xml_character(character: char) -> bool {
    matches!(character as u32, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("Slicers {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> String {
        format!(
            r#"<x14:slicers xmlns:x14="{SLICERS_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="v" xmlns:v="urn:producer"><x14:slicer name="State" cache="Slicer_State" caption="State &amp; Region" startItem="2" columnCount="3" showCaption="0" level="1" style="SlicerStyleLight1" lockedPosition="1" rowHeight="228600"><x14:extLst><v:ext r:id="rIdNeverFetched" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><v:data href="https://example.invalid/not-opened"/></v:ext></x14:extLst></x14:slicer></x14:slicers>"#
        )
    }

    fn package() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            worksheet.clone(), ct::SML_WORKSHEET.into(),
            b"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>".to_vec(),
        )));
        (package, worksheet)
    }

    #[test]
    fn microsoft_example_and_typed_values_round_trip() {
        let microsoft = format!(
            r#"<slicers xmlns="{SLICERS_NS}"><slicer name="State" cache="Slicer_State" caption="State" rowHeight="228600"/></slicers>"#
        );
        let parsed = parse_slicers(microsoft.as_bytes()).unwrap();
        assert_eq!(parsed.slicers.len(), 1);
        assert_eq!(
            parsed.slicers[0],
            Slicer {
                caption: Some("State".into()),
                ..Slicer::new("State", "Slicer_State", 228600)
            }
        );
        assert_eq!(
            parse_slicers(&write_slicers(&parsed).unwrap()).unwrap(),
            parsed
        );

        let typed = parse_slicers(sample().as_bytes()).unwrap();
        let slicer = &typed.slicers[0];
        assert_eq!(
            (
                slicer.start_item,
                slicer.column_count,
                slicer.show_caption,
                slicer.level,
                slicer.locked_position
            ),
            (2, 3, false, 1, true)
        );
        assert!(
            slicer
                .extension_list
                .as_ref()
                .unwrap()
                .xml()
                .windows(16)
                .any(|window| window == b"rIdNeverFetched\"")
        );
        assert_eq!(
            parse_slicers(&write_slicers(&typed).unwrap()).unwrap(),
            typed
        );
    }

    #[test]
    fn package_store_and_load_round_trip() {
        let (mut package, worksheet) = package();
        let expected = WorksheetSlicers {
            relationship_id: "rIdSlicers".into(),
            part_name: "/xl/slicers/slicer1.xml".into(),
            slicers: parse_slicers(sample().as_bytes()).unwrap(),
        };
        store_worksheet_slicers(&mut package, &worksheet, &expected).unwrap();
        assert_eq!(
            load_worksheet_slicers(&package, &worksheet).unwrap(),
            vec![expected]
        );
    }

    #[test]
    fn rejects_hostile_xml_schema_violations_and_limits() {
        let cases = [
            format!(
                r#"<!DOCTYPE x><slicers xmlns="{SLICERS_NS}"><slicer name="a" cache="b" rowHeight="1"/></slicers>"#
            ),
            r#"<slicers xmlns="urn:wrong"><slicer name="a" cache="b" rowHeight="1"/></slicers>"#
                .to_string(),
            format!(r#"<slicers xmlns="{SLICERS_NS}"/>"#),
            format!(r#"<slicers xmlns="{SLICERS_NS}"><slicer name="a" cache="b"/></slicers>"#),
            format!(
                r#"<slicers xmlns="{SLICERS_NS}"><slicer name="a" name="b" cache="c" rowHeight="1"/></slicers>"#
            ),
            format!(
                r#"<slicers xmlns="{SLICERS_NS}"><slicer name="a" cache="b" rowHeight="1" columnCount="0"/></slicers>"#
            ),
            format!(
                r#"<slicers xmlns="{SLICERS_NS}"><slicer name="a" cache="b" rowHeight="-1"/></slicers>"#
            ),
            format!(
                r#"<slicers xmlns="{SLICERS_NS}"><slicer name="a" cache="b" rowHeight="1" showCaption="yes"/></slicers>"#
            ),
            format!(
                r#"<slicers xmlns="{SLICERS_NS}"><slicer name="a" cache="b" rowHeight="1"><bad/></slicer></slicers>"#
            ),
        ];
        for xml in cases {
            assert!(parse_slicers(xml.as_bytes()).is_err(), "accepted {xml}");
        }
        assert!(parse_slicers(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
        let mut deep = String::new();
        for _ in 0..=MAX_DEPTH {
            deep.push_str("<x14:x>");
        }
        for _ in 0..=MAX_DEPTH {
            deep.push_str("</x14:x>");
        }
        let xml = format!(
            r#"<x14:slicers xmlns:x14="{SLICERS_NS}"><x14:slicer name="a" cache="b" rowHeight="1"><x14:extLst>{deep}</x14:extLst></x14:slicer></x14:slicers>"#
        );
        assert!(parse_slicers(xml.as_bytes()).is_err());
    }

    #[test]
    fn rejects_invalid_package_graphs() {
        let value = WorksheetSlicers {
            relationship_id: "rId1".into(),
            part_name: "/xl/slicers/slicer1.xml".into(),
            slicers: Slicers::new(vec![Slicer::new("a", "cache", 1)]),
        };

        let (mut external, worksheet) = package();
        external
            .get_part_mut(&worksheet)
            .unwrap()
            .rels_mut()
            .add_relationship(
                SLICERS_RELATIONSHIP_TYPE.into(),
                "https://invalid.example/slicer".into(),
                "rId1".into(),
                true,
            );
        assert!(load_worksheet_slicers(&external, &worksheet).is_err());

        let (mut outbound, worksheet) = package();
        store_worksheet_slicers(&mut outbound, &worksheet, &value).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/xl/slicers/slicer1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship("urn:forbidden".into(), "x".into(), "rId9".into(), false);
        assert!(load_worksheet_slicers(&outbound, &worksheet).is_err());

        let (mut orphan, worksheet) = package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/slicers/orphan.xml").unwrap(),
            SLICERS_CONTENT_TYPE.into(),
            write_slicers(&value.slicers).unwrap(),
        )));
        assert!(load_worksheet_slicers(&orphan, &worksheet).is_err());

        let (mut wrong_source, worksheet) = package();
        let other = PackURI::new("/xl/workbook.xml").unwrap();
        let mut workbook = BlobPart::new(
            other.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            Vec::new(),
        );
        workbook.rels_mut().add_relationship(
            SLICERS_RELATIONSHIP_TYPE.into(),
            "slicers/slicer1.xml".into(),
            "rId1".into(),
            false,
        );
        wrong_source.add_part(Box::new(workbook));
        wrong_source.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/slicers/slicer1.xml").unwrap(),
            SLICERS_CONTENT_TYPE.into(),
            write_slicers(&value.slicers).unwrap(),
        )));
        assert!(load_worksheet_slicers(&wrong_source, &worksheet).is_err());
    }
}
